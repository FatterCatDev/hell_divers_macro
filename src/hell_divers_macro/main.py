"""
Macro manager with a Tkinter UI. Assign Helldivers 2 stratagem macros to a numpad-style
grid, customize hotkeys/direction keys, and save/load profiles.
"""

from __future__ import annotations

import json
import threading
import time
from pathlib import Path
from typing import Dict, List, Tuple

import keyboard
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import ImageTk
import sys

if __package__ in (None, ""):
    # Allow running as a script by adding project src to path for absolute imports.
    sys.path.append(str(Path(__file__).resolve().parent.parent))

from hell_divers_macro.config import (
    DEFAULT_DELAY,
    DEFAULT_DURATION,
    DEFAULT_AUTO_PANEL,
    DEFAULT_DIRECTION_KEYS,
    DEFAULT_PANEL_KEY,
    DEFAULT_OVERLAY_LOCK_KEY,
    DEFAULT_OVERLAY_OPACITY,
    DEFAULT_SLOT_HOTKEYS,
    EXIT_HOTKEY,
    NUMPAD_SLOTS,
)
from hell_divers_macro.log_utils import clear_log_callback, log, set_log_callback
from hell_divers_macro.models import Macro, MacroRecord, MacroTemplate
from hell_divers_macro.paths import ensure_saves_dir
from hell_divers_macro.ui.icons import (
    APP_ICON_PATH,
    OVERLAY_ALPHA,
    OVERLAY_ICON_SIZE,
    build_overlay_placeholder,
    load_icon_image,
)
from hell_divers_macro.ui.theme import (
    ACCENT,
    BG,
    BUTTON_ACTIVE,
    BUTTON_BG,
    ENTRY_BG,
    FG,
    MENU_BG,
    apply_dark_theme,
    init_base_theme,
    make_window_clickthrough,
    place_window_near,
    set_dark_titlebar,
)
from hell_divers_macro.stratagems import (
    load_stratagem_templates,
    resolve_template_keys,
    save_stratagem_templates,
)

_macro_lock = threading.Lock()
_auto_panel_state: dict[str, object] = {"enabled": DEFAULT_AUTO_PANEL, "key": DEFAULT_PANEL_KEY}

# --- Layout ------------------------------------------------------------------
SLOT_WIDTH = 160
SLOT_HEIGHT = 150

# --- Macro execution helpers -------------------------------------------------
_macro_progress_callback = None
_slot_hotkey_lookup: dict[str, str] = {}
def _run_macro(macro: Macro, panel_key: str | None = None) -> None:
    with _macro_lock:
        label = macro.name or macro.hotkey
        sequence = macro.keys
        if panel_key:
            sequence = (panel_key, *sequence)
            log(f"{label}: auto panel ON, prepending '{panel_key}'.")
        log(f"{label}: running {len(sequence)} key presses...")
        for key in sequence:
            keyboard.press(key)
            time.sleep(macro.duration)
            keyboard.release(key)
            time.sleep(macro.delay)
        log(f"{label}: done.")


def _notify_macro_progress(event: str, macro: Macro, slot: str | None, total_time: float | None) -> None:
    cb = _macro_progress_callback
    if cb is None:
        return
    try:
        cb(event, macro, slot, total_time)
    except Exception:
        # Never let UI callbacks break macro execution.
        pass


def _launch_macro_from_hotkey(macro: Macro) -> None:
    if _macro_lock.locked():
        log("Another macro is running, ignoring new request.")
        return

    label = macro.name or macro.hotkey
    log(f"Trigger received for hotkey '{macro.hotkey}' ({label}).")
    panel_key_arg = None
    if _auto_panel_state.get("enabled"):
        key = (_auto_panel_state.get("key") or "").strip()
        if key:
            panel_key_arg = str(key)
    slot = _slot_hotkey_lookup.get(macro.hotkey.lower())
    total_len = len(macro.keys) + (1 if panel_key_arg else 0)
    total_time = total_len * (macro.duration + macro.delay)
    _notify_macro_progress("start", macro, slot, total_time)

    def _worker() -> None:
        try:
            _run_macro(macro, panel_key_arg)
        finally:
            _notify_macro_progress("stop", macro, slot, None)

    threading.Thread(target=_worker, daemon=True).start()


# --- Data manager ------------------------------------------------------------
class MacroManager:
    def __init__(self) -> None:
        self.records: List[MacroRecord] = []
        self._held_scancodes: set[int] = set()

    def _register_macro(self, macro: Macro) -> int:
        def on_press(event) -> None:
            if event.event_type != "down":
                return
            # Ignore non-keypad arrows triggering numpad hotkeys.
            if macro.hotkey.startswith("num "):
                is_keypad = getattr(event, "is_keypad", None)
                if is_keypad is False:
                    return
                if is_keypad is None and event.name in ("up", "down", "left", "right"):
                    return
            sc = event.scan_code
            if sc in self._held_scancodes:
                return
            self._held_scancodes.add(sc)
            _launch_macro_from_hotkey(macro)

        def on_release(event) -> None:
            sc = event.scan_code
            self._held_scancodes.discard(sc)

        press_hook = keyboard.on_press_key(macro.hotkey, on_press, suppress=False)
        release_hook = keyboard.on_release_key(macro.hotkey, on_release, suppress=False)
        return (press_hook, release_hook)

    def _hotkey_in_use(self, hotkey: str, ignore_index: int | None = None) -> bool:
        for idx, record in enumerate(self.records):
            if ignore_index is not None and idx == ignore_index:
                continue
            if record.macro.hotkey == hotkey:
                return True
        return False

    def add_macro(self, macro: Macro) -> None:
        hotkey = macro.hotkey.lower()
        if self._hotkey_in_use(hotkey):
            raise ValueError(f"Hotkey '{hotkey}' already in use.")
        normalized = Macro(hotkey, tuple(macro.keys), macro.delay, macro.duration, macro.name)
        handle = self._register_macro(normalized)
        self.records.append(MacroRecord(normalized, handle))

    def update_macro(self, index: int, macro: Macro) -> None:
        if not (0 <= index < len(self.records)):
            raise IndexError("Invalid macro index.")
        hotkey = macro.hotkey.lower()
        if self._hotkey_in_use(hotkey, ignore_index=index):
            raise ValueError(f"Hotkey '{hotkey}' already in use.")
        new_macro = Macro(hotkey, tuple(macro.keys), macro.delay, macro.duration, macro.name)
        press_hook, release_hook = self.records[index].handle
        keyboard.unhook(press_hook)
        keyboard.unhook(release_hook)
        handle = self._register_macro(new_macro)
        self.records[index] = MacroRecord(new_macro, handle)

    def remove_macro(self, index: int) -> None:
        if not (0 <= index < len(self.records)):
            raise IndexError("Invalid macro index.")
        record = self.records.pop(index)
        press_hook, release_hook = record.handle
        keyboard.unhook(press_hook)
        keyboard.unhook(release_hook)

    def clear(self) -> None:
        for record in self.records:
            press_hook, release_hook = record.handle
            keyboard.unhook(press_hook)
        self.records.clear()


# --- UI components -----------------------------------------------------------
class MacroSelectionDialog:
    def __init__(self, parent: tk.Tk, title: str, templates: Tuple[MacroTemplate, ...]) -> None:
        self.parent = parent
        self.templates = templates
        self.result: MacroTemplate | None = None
        self._current_selection: MacroTemplate | None = None

        categories: dict[str, list[MacroTemplate]] = {}
        for tpl in self.templates:
            categories.setdefault(tpl.category, []).append(tpl)
        ordered_categories = list(categories.keys())
        visible: list[MacroTemplate] = []

        self.top = tk.Toplevel(parent, bg=BG)
        self.top.title(title)
        self.top.transient(parent)

        tk.Label(self.top, text="Choose a macro template:").pack(anchor="w", pady=(8, 4), padx=10)

        search_var = tk.StringVar()
        search_frame = tk.Frame(self.top)
        search_frame.pack(fill=tk.X, padx=10, pady=(0, 8))
        tk.Label(search_frame, text="Search").pack(side=tk.LEFT, padx=(0, 6))
        search_entry = tk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tabs_frame = tk.Frame(self.top)
        tabs_frame.pack(fill=tk.X, padx=10, pady=(0, 6))

        list_frame = tk.Frame(self.top)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 6))

        self.listbox = tk.Listbox(list_frame, height=12)
        list_scroll = tk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.config(yscrollcommand=list_scroll.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        current_cat = {"val": ordered_categories[0] if ordered_categories else ""}
        visible = list(self.templates)

        def populate_list(cat: str | None, query: str) -> None:
            self.listbox.delete(0, tk.END)
            q = query.strip().lower()
            nonlocal visible
            if q:
                visible = [tpl for tpl in self.templates if q in tpl.name.lower()]
            elif cat:
                visible = list(categories.get(cat, []))
            else:
                visible = list(self.templates)
            for tpl in visible:
                self.listbox.insert(tk.END, tpl.name)
            self._current_selection = None

        def switch_cat(cat: str) -> None:
            current_cat["val"] = cat
            for btn in tab_buttons.values():
                btn.config(relief=tk.RAISED)
            tab_buttons[cat].config(relief=tk.SUNKEN)
            populate_list(cat, search_var.get())
            layout_tabs()

        tab_buttons_list: list[tuple[str, tk.Button]] = []
        tab_buttons: dict[str, tk.Button] = {}
        for idx, cat in enumerate(ordered_categories):
            btn = tk.Button(tabs_frame, text=cat, command=lambda c=cat: switch_cat(c))
            tab_buttons[cat] = btn
            tab_buttons_list.append((cat, btn))

        def layout_tabs(event=None) -> None:  # noqa: ANN001
            tabs_frame.update_idletasks()
            available = tabs_frame.winfo_width()
            if available <= 1:
                available = self.top.winfo_width() - 20
            x = 0
            row = 0
            col = 0
            for cat, btn in tab_buttons_list:
                w = btn.winfo_reqwidth() + 6
                if col > 0 and x + w > available:
                    row += 1
                    col = 0
                    x = 0
                btn.grid(row=row, column=col, padx=(0, 6), pady=2, sticky="w")
                col += 1
                x += w

        tabs_frame.bind("<Configure>", layout_tabs)

        def handle_select(event=None) -> None:  # noqa: ANN001
            indices = self.listbox.curselection()
            if not indices:
                return
            if not visible or indices[0] >= len(visible):
                return
            tpl = visible[indices[0]]
            self._current_selection = tpl

        self.listbox.bind("<<ListboxSelect>>", handle_select)
        self.listbox.bind("<Double-Button-1>", lambda _: (handle_select(), self.ok()))

        def on_search(*args: str) -> None:  # noqa: ANN001
            populate_list(current_cat["val"], search_var.get())

        search_var.trace_add("write", on_search)

        btn_frame = tk.Frame(self.top)
        btn_frame.pack(fill=tk.X, padx=10, pady=(4, 10))
        tk.Button(btn_frame, text="OK", command=self.ok).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT)

        self.top.protocol("WM_DELETE_WINDOW", self.cancel)
        self.top.grab_set()
        place_window_near(self.top, parent)
        self.top.focus_set()
        apply_dark_theme(self.top)
        set_dark_titlebar(self.top)

        if ordered_categories:
            switch_cat(ordered_categories[0])

        self.top.update_idletasks()
        parent_width = parent.winfo_width()
        if parent_width > 0:
            reqw = self.top.winfo_reqwidth()
            reqh = self.top.winfo_reqheight()
            if reqw > parent_width:
                self.top.geometry(f"{parent_width}x{reqh}")
        parent.wait_window(self.top)

    def ok(self) -> None:
        self.result = self._current_selection
        self.top.destroy()

    def cancel(self) -> None:
        self.result = None
        self.top.destroy()


class TextEntryDialog:
    def __init__(self, parent: tk.Tk, title: str, prompt: str, initial: str = "") -> None:
        self.result: str | None = None
        self.top = tk.Toplevel(parent, bg=BG)
        self.top.title(title)
        self.top.transient(parent)

        tk.Label(self.top, text=prompt).pack(anchor="w", padx=10, pady=(10, 4))
        self.entry_var = tk.StringVar(value=initial)
        entry = tk.Entry(self.top, textvariable=self.entry_var)
        entry.pack(fill=tk.X, padx=10)

        btn_frame = tk.Frame(self.top)
        btn_frame.pack(fill=tk.X, padx=10, pady=(10, 10))
        tk.Button(btn_frame, text="OK", command=self.ok).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT)

        self.top.protocol("WM_DELETE_WINDOW", self.cancel)
        self.top.bind("<Return>", lambda _: self.ok())
        self.top.bind("<Escape>", lambda _: self.cancel())
        self.top.grab_set()
        place_window_near(self.top, parent)
        entry.focus_set()
        apply_dark_theme(self.top)
        set_dark_titlebar(self.top)
        parent.wait_window(self.top)

    def ok(self) -> None:
        self.result = self.entry_var.get()
        self.top.destroy()

    def cancel(self) -> None:
        self.result = None
        self.top.destroy()


# --- App ---------------------------------------------------------------------
def main() -> None:
    manager = MacroManager()
    ensure_saves_dir()
    templates: Tuple[MacroTemplate, ...] = load_stratagem_templates()
    last_profile_marker = ensure_saves_dir() / ".last_profile"

    root = tk.Tk()
    root.title("HELLDIVERS2 Stratagem Macro")
    root.geometry("540x560")
    root.minsize(520, 520)
    try:
        if APP_ICON_PATH.exists():
            root.iconphoto(True, tk.PhotoImage(file=str(APP_ICON_PATH)))
    except Exception:
        pass
    init_base_theme(root)
    apply_dark_theme(root)
    set_dark_titlebar(root)

    assignments: dict[str, MacroTemplate | None] = {slot: None for slot, _ in NUMPAD_SLOTS}
    slot_hotkeys: dict[str, str] = dict(DEFAULT_SLOT_HOTKEYS)
    direction_keys: dict[str, str] = dict(DEFAULT_DIRECTION_KEYS)
    panel_key: str = DEFAULT_PANEL_KEY
    auto_panel_var = tk.BooleanVar(value=DEFAULT_AUTO_PANEL)
    slot_buttons: dict[str, tk.Button] = {}
    slot_icons: dict[str, ImageTk.PhotoImage | None] = {}
    listening = False
    saved_state: dict = {}
    macro_timing = {"delay": DEFAULT_DELAY, "duration": DEFAULT_DURATION}
    panel_key_display = tk.StringVar(value="")
    overlay_lock_key: str = DEFAULT_OVERLAY_LOCK_KEY
    overlay_locked = {"val": False}
    overlay_opacity = tk.DoubleVar(value=OVERLAY_ALPHA)
    overlay_lock_handle: int | None = None
    overlay_drag_state = {"x": 0, "y": 0}
    overlay_dragging = {"val": False}
    overlay_lock_display = tk.StringVar(value="")
    overlay_resize_state = {"x": 0, "y": 0, "w": 0, "h": 0}
    overlay_resize_handle: tk.Canvas | None = None
    overlay_resizing = {"val": False}
    overlay_auto_panel_check: tk.Checkbutton | None = None
    overlay_close_btn: tk.Button | None = None
    overlay_lock_label: tk.Label | None = None

    def sync_auto_panel_state() -> None:
        _auto_panel_state["key"] = panel_key
        _auto_panel_state["enabled"] = bool(auto_panel_var.get())

    def serialize_state() -> dict:
        return {
            "slots": {slot: (tpl.name if tpl else None) for slot, tpl in assignments.items()},
            "hotkeys": dict(slot_hotkeys),
            "direction_keys": dict(direction_keys),
            "timing": {"delay": macro_timing["delay"], "duration": macro_timing["duration"]},
            "panel": {"key": panel_key, "auto": bool(auto_panel_var.get())},
            "overlay": {
                "lock_key": overlay_lock_key,
                "opacity": _clamp_opacity(overlay_opacity.get()),
            },
        }

    def has_unsaved_changes() -> bool:
        return serialize_state() != saved_state

    def _record_last_profile(path: Path) -> None:
        try:
            last_profile_marker.write_text(str(path), encoding="utf-8")
        except OSError:
            pass

    def _display_hotkey_text(raw: str, default: str) -> str:
        if not raw:
            return default
        lower = raw.lower()
        if lower.startswith("num "):
            return lower.split(" ", 1)[1]
        return raw

    def _clamp_opacity(val: float) -> float:
        try:
            return max(0.1, min(1.0, float(val)))
        except Exception:
            return OVERLAY_ALPHA

    def _update_overlay_lock_display() -> None:
        key_text = _display_hotkey_text(overlay_lock_key or "Unset", overlay_lock_key or "Unset")
        state = "Locked" if overlay_locked["val"] else "Unlocked"
        overlay_lock_display.set(f"Overlay: {state} (Key: {key_text})")

    def _hide_widget(widget: tk.Widget | None) -> None:
        if widget is None:
            return
        try:
            if widget.winfo_manager():
                widget.pack_forget()
        except tk.TclError:
            pass

    def _show_widget(widget: tk.Widget | None, **pack_kwargs) -> None:
        if widget is None:
            return
        try:
            if widget.winfo_manager() != "pack":
                widget.pack(**pack_kwargs)
        except tk.TclError:
            pass

    def _update_overlay_header_visibility() -> None:
        for widget in (overlay_auto_panel_check, overlay_close_btn, overlay_lock_label):
            _hide_widget(widget)
        if overlay_locked["val"]:
            _show_widget(overlay_lock_label, side=tk.RIGHT, padx=(0, 6), anchor="e")
        else:
            _show_widget(overlay_close_btn, side=tk.RIGHT, padx=(6, 0))
            _show_widget(overlay_lock_label, side=tk.RIGHT, padx=(0, 6), anchor="e")
            _show_widget(overlay_auto_panel_check, side=tk.LEFT, anchor="w")

    def _overlay_event_target(event) -> tk.Widget | None:  # noqa: ANN001
        """Return widget under cursor for overlay events."""
        if overlay_win is None:
            return None
        try:
            return overlay_win.winfo_containing(event.x_root, event.y_root)
        except Exception:
            return None

    def _is_overlay_interactive_widget(widget: tk.Widget | None) -> bool:
        """Widgets that should receive clicks instead of starting a drag."""
        if widget is None:
            return False
        return isinstance(widget, (tk.Button, tk.Checkbutton, tk.Entry, tk.Scale, tk.Listbox, tk.Text))

    grid_layout = [
        ["7", "8", "9"],
        ["4", "5", "6"],
        ["1", "2", "3"],
    ]

    def _overlay_min_sizes() -> tuple[int, int]:
        return (OVERLAY_ICON_SIZE[0] * 3 + 40, OVERLAY_ICON_SIZE[1] * 3 + 90)

    overlay_win: tk.Toplevel | None = None
    overlay_slot_canvases: dict[str, tk.Canvas] = {}
    overlay_fill_rects: dict[str, int] = {}
    overlay_icons: dict[str, ImageTk.PhotoImage | None] = {}
    overlay_progress: dict[str, dict[str, object]] = {}
    overlay_user_resized = {"val": False}
    overlay_applying_config = {"val": False}

    def _apply_overlay_window_config() -> None:
        if overlay_win is None or not overlay_win.winfo_exists():
            return
        if overlay_applying_config["val"]:
            return
        overlay_applying_config["val"] = True
        try:
            opacity = _clamp_opacity(overlay_opacity.get())
            overlay_opacity.set(opacity)
            try:
                overlay_win.attributes("-topmost", True)
                overlay_win.attributes("-alpha", opacity)
            except tk.TclError:
                pass
            make_window_clickthrough(overlay_win, alpha=opacity, clickthrough=overlay_locked["val"])
            try:
                overlay_win.attributes("-disabled", overlay_locked["val"])
                overlay_win.configure(cursor="" if overlay_locked["val"] else "fleur")
            except tk.TclError:
                pass
            if overlay_resize_handle is not None:
                try:
                    overlay_resize_handle.configure(cursor="size_nw_se" if not overlay_locked["val"] else "")
                except tk.TclError:
                    pass
            _update_overlay_lock_display()
            _update_overlay_header_visibility()
        finally:
            overlay_applying_config["val"] = False

    def _start_overlay_drag(event) -> None:  # noqa: ANN001
        overlay_dragging["val"] = False
        if (
            overlay_locked["val"]
            or overlay_resizing["val"]
            or overlay_win is None
            or not overlay_win.winfo_exists()
        ):
            return
        target = _overlay_event_target(event)
        if target is overlay_resize_handle or _is_overlay_interactive_widget(target):
            return
        overlay_dragging["val"] = True
        overlay_drag_state["x"] = event.x_root
        overlay_drag_state["y"] = event.y_root

    def _drag_overlay(event) -> None:  # noqa: ANN001
        if (
            overlay_locked["val"]
            or overlay_resizing["val"]
            or overlay_win is None
            or not overlay_win.winfo_exists()
            or not overlay_dragging["val"]
        ):
            return
        dx = event.x_root - overlay_drag_state["x"]
        dy = event.y_root - overlay_drag_state["y"]
        overlay_drag_state["x"] = event.x_root
        overlay_drag_state["y"] = event.y_root
        try:
            new_x = overlay_win.winfo_x() + dx
            new_y = overlay_win.winfo_y() + dy
            overlay_win.geometry(f"+{new_x}+{new_y}")
            overlay_user_resized["val"] = True
        except tk.TclError:
            pass

    def _start_overlay_resize(event) -> None:  # noqa: ANN001
        if overlay_locked["val"] or overlay_win is None or not overlay_win.winfo_exists():
            return
        overlay_resizing["val"] = True
        overlay_resize_state["x"] = event.x_root
        overlay_resize_state["y"] = event.y_root
        overlay_resize_state["w"] = overlay_win.winfo_width()
        overlay_resize_state["h"] = overlay_win.winfo_height()

    def _resize_overlay(event) -> None:  # noqa: ANN001
        if overlay_locked["val"] or overlay_win is None or not overlay_win.winfo_exists():
            return
        dx = event.x_root - overlay_resize_state["x"]
        dy = event.y_root - overlay_resize_state["y"]
        new_w = max(_overlay_min_sizes()[0], overlay_resize_state["w"] + dx)
        new_h = max(_overlay_min_sizes()[1], overlay_resize_state["h"] + dy)
        try:
            overlay_win.geometry(f"{new_w}x{new_h}+{overlay_win.winfo_x()}+{overlay_win.winfo_y()}")
            overlay_user_resized["val"] = True
        except tk.TclError:
            pass

    def _stop_overlay_resize(event=None) -> None:  # noqa: ANN001
        overlay_dragging["val"] = False
        overlay_resizing["val"] = False

    def toggle_overlay_lock() -> None:
        overlay_locked["val"] = not overlay_locked["val"]
        state = "locked" if overlay_locked["val"] else "unlocked (drag to move)"
        _apply_overlay_window_config()
        _update_overlay_lock_display()
        _update_overlay_header_visibility()
        status_var.set(f"Overlay {state}.")
        log(f"Overlay {state}.")

    def register_overlay_lock_hotkey() -> None:
        nonlocal overlay_lock_handle
        if overlay_lock_handle is not None:
            try:
                keyboard.remove_hotkey(overlay_lock_handle)
            except Exception:
                pass
            overlay_lock_handle = None
        key = (overlay_lock_key or "").strip()
        if not key:
            return
        try:
            overlay_lock_handle = keyboard.add_hotkey(
                key, lambda: root.after(0, toggle_overlay_lock), suppress=False
            )
        except Exception as exc:  # noqa: BLE001
            log(f"Could not register overlay lock hotkey '{key}': {exc}")

    def _overlay_visible() -> bool:
        return overlay_win is not None and overlay_win.winfo_exists() and overlay_win.state() != "withdrawn"

    def _size_overlay_window(force: bool = False) -> None:
        if overlay_win is None or not overlay_win.winfo_exists():
            return
        if overlay_user_resized["val"] and not force:
            return
        root.update_idletasks()
        overlay_win.update_idletasks()
        min_w, min_h = _overlay_min_sizes()
        width = max(int(root.winfo_width() * 0.55), min_w)
        height = max(int(root.winfo_height() * 0.55), min_h)
        overlay_win.geometry(f"{width}x{height}")
        place_window_near(overlay_win, root)

    def _ensure_overlay_window() -> None:
        nonlocal overlay_win, overlay_resize_handle, overlay_auto_panel_check, overlay_close_btn, overlay_lock_label
        if overlay_win is not None and overlay_win.winfo_exists():
            return
        overlay_slot_canvases.clear()
        overlay_fill_rects.clear()
        overlay_icons.clear()
        overlay_win = tk.Toplevel(root, bg=BG)
        overlay_win.withdraw()
        overlay_win.overrideredirect(True)
        overlay_win.title("Listening Overlay")
        overlay_win.resizable(True, True)
        overlay_win.attributes("-topmost", True)
        _apply_overlay_window_config()
        overlay_win.protocol("WM_DELETE_WINDOW", lambda: handle_overlay_close())
        overlay_win.bind("<ButtonPress-1>", _start_overlay_drag)
        overlay_win.bind("<B1-Motion>", _drag_overlay)
        overlay_win.bind("<ButtonRelease-1>", _stop_overlay_resize)

        container = tk.Frame(overlay_win, bg=BG)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

        auto_frame = tk.Frame(container, bg=BG)
        auto_frame.pack(fill=tk.X, pady=(0, 8))
        overlay_auto_panel_check = tk.Checkbutton(
            auto_frame,
            text="Auto Stratagem Panel",
            variable=auto_panel_var,
            bg=BG,
            fg=FG,
            activebackground=BUTTON_BG,
            activeforeground=FG,
            selectcolor=BG,
            highlightthickness=0,
            anchor="w",
        )
        overlay_close_btn = tk.Button(
            auto_frame,
            text="X",
            command=lambda: root.after(0, stop_listening),
            width=2,
            bg=BUTTON_BG,
            fg=FG,
            activebackground=BUTTON_ACTIVE,
            activeforeground=FG,
            bd=0,
            highlightthickness=0,
        )
        overlay_lock_label = tk.Label(auto_frame, textvariable=overlay_lock_display, anchor="e")

        grid = tk.Frame(container, bg=BG)
        grid.pack(fill=tk.BOTH, expand=True)
        for r, row in enumerate(grid_layout):
            for c, slot in enumerate(row):
                cell = tk.Frame(
                    grid,
                    width=OVERLAY_ICON_SIZE[0] + 12,
                    height=OVERLAY_ICON_SIZE[1] + 12,
                    bg=BG,
                    highlightthickness=0,
                    bd=0,
                )
                cell.grid(row=r, column=c, padx=4, pady=4, sticky="nsew")
                cell.grid_propagate(False)
                canvas = tk.Canvas(
                    cell,
                    width=OVERLAY_ICON_SIZE[0],
                    height=OVERLAY_ICON_SIZE[1],
                    bg=BG,
                    highlightthickness=0,
                    bd=0,
                )
                canvas.pack(expand=True)
                overlay_slot_canvases[slot] = canvas
                rect_id = canvas.create_rectangle(
                    0,
                    0,
                    OVERLAY_ICON_SIZE[0],
                    0,
                    fill="#3a3a3a",
                    outline="",
                    stipple="gray50",
                    tags="fill",
                )
                overlay_fill_rects[slot] = rect_id
                grid.grid_columnconfigure(c, weight=1, uniform="overlay_slots")
            grid.grid_rowconfigure(r, weight=1, uniform="overlay_slots")

        apply_dark_theme(overlay_win)
        set_dark_titlebar(overlay_win)
        overlay_user_resized["val"] = False
        overlay_win.bind(
            "<Configure>",
            lambda event=None: overlay_user_resized.__setitem__("val", True)  # noqa: ANN001
            if overlay_win.state() != "withdrawn"
            else None,
        )
        _size_overlay_window()

        overlay_resize_handle = tk.Canvas(container, width=16, height=16, bg=BG, bd=0, highlightthickness=0)
        overlay_resize_handle.place(relx=1.0, rely=1.0, anchor="se")
        overlay_resize_handle.create_polygon(0, 16, 16, 16, 16, 0, fill=ACCENT, outline="")
        overlay_resize_handle.bind("<ButtonPress-1>", _start_overlay_resize)
        overlay_resize_handle.bind("<B1-Motion>", _resize_overlay)
        overlay_resize_handle.bind("<ButtonRelease-1>", _stop_overlay_resize)
        _update_overlay_header_visibility()

    def _set_overlay_fill(slot: str, progress: float) -> None:
        canvas = overlay_slot_canvases.get(slot)
        rect_id = overlay_fill_rects.get(slot)
        if canvas is None or rect_id is None:
            return
        progress = max(0.0, min(1.0, progress))
        height = int(OVERLAY_ICON_SIZE[1] * progress)
        canvas.coords(rect_id, 0, 0, OVERLAY_ICON_SIZE[0], height)
        canvas.itemconfigure(rect_id, state=tk.NORMAL if height > 0 else tk.HIDDEN)

    def update_overlay_slot(slot: str) -> None:
        if overlay_win is None or not overlay_win.winfo_exists():
            return
        canvas = overlay_slot_canvases.get(slot)
        if canvas is None:
            return
        tpl = assignments.get(slot)
        hotkey = slot_hotkeys.get(slot, slot)
        hotkey_text = _display_hotkey_text(hotkey, slot)
        icon: ImageTk.PhotoImage | None = None
        if tpl:
            icon = load_icon_image(tpl.name, hotkey_text, variant="badge", size=OVERLAY_ICON_SIZE)
        else:
            icon = build_overlay_placeholder(hotkey_text, size=OVERLAY_ICON_SIZE)
        overlay_icons[slot] = icon
        canvas.delete("icon")
        canvas.create_image(
            OVERLAY_ICON_SIZE[0] // 2,
            OVERLAY_ICON_SIZE[1] // 2,
            image=icon,
            tags="icon",
        )
        _set_overlay_fill(slot, 0)

    def refresh_overlay_slots() -> None:
        if overlay_win is None or not overlay_win.winfo_exists():
            return
        for slot in assignments:
            update_overlay_slot(slot)

    def _cancel_overlay_progress(slot: str) -> None:
        data = overlay_progress.pop(slot, None)
        if data is None:
            return
        job = data.get("job")
        if job is not None:
            try:
                root.after_cancel(job)
            except Exception:
                pass

    def _tick_overlay_progress(slot: str) -> None:
        data = overlay_progress.get(slot)
        if data is None:
            return
        duration = float(data["duration"])
        start_ts = float(data["start"])
        elapsed = max(0.0, time.time() - start_ts)
        progress = min(1.0, elapsed / duration) if duration > 0 else 1.0
        _set_overlay_fill(slot, progress)
        if progress >= 1.0:
            _cancel_overlay_progress(slot)
            _set_overlay_fill(slot, 0)
            return
        data["job"] = root.after(50, lambda s=slot: _tick_overlay_progress(s))

    def start_overlay_progress(slot: str, total_time: float | None) -> None:
        if overlay_win is None or not overlay_win.winfo_exists():
            return
        duration = max(total_time or 0.05, 0.05)
        _cancel_overlay_progress(slot)
        overlay_progress[slot] = {"start": time.time(), "duration": duration, "job": None}
        _tick_overlay_progress(slot)

    def stop_overlay_progress(slot: str) -> None:
        _cancel_overlay_progress(slot)
        _set_overlay_fill(slot, 0)

    def handle_macro_progress(event: str, macro: Macro, slot: str | None, total_time: float | None) -> None:
        if not slot:
            return
        if event == "start":
            root.after(0, lambda s=slot, t=total_time: start_overlay_progress(s, t))
        elif event == "stop":
            root.after(0, lambda s=slot: stop_overlay_progress(s))

    global _macro_progress_callback
    _macro_progress_callback = handle_macro_progress

    def show_overlay() -> None:
        _ensure_overlay_window()
        if not overlay_user_resized["val"]:
            _size_overlay_window(force=True)
        _update_overlay_lock_display()
        refresh_overlay_slots()
        if overlay_win is not None:
            _apply_overlay_window_config()
            overlay_win.deiconify()
            overlay_win.lift()
            overlay_win.attributes("-topmost", True)

    def hide_overlay() -> None:
        if overlay_win is None or not overlay_win.winfo_exists():
            return
        overlay_win.withdraw()

    def handle_overlay_close() -> None:
        stop_listening()

    def _on_root_configure(event=None) -> None:  # noqa: ANN001
        if _overlay_visible():
            _size_overlay_window()

    root.bind("<Configure>", _on_root_configure)

    def update_button_label(slot: str) -> None:
        tpl = assignments.get(slot)
        hotkey = slot_hotkeys.get(slot, slot)
        hotkey_text = _display_hotkey_text(hotkey, slot)
        icon: ImageTk.PhotoImage | None = None
        if tpl:
            icon = load_icon_image(tpl.name, hotkey_text)
        slot_icons[slot] = icon
        if icon:
            slot_buttons[slot].config(image=icon, text="", compound=tk.CENTER)
        else:
            name = tpl.name if tpl else "Unassigned"
            slot_buttons[slot].config(text=f"{hotkey_text}\n{name}", image="", compound=tk.NONE)
        update_overlay_slot(slot)

    def update_all_buttons() -> None:
        for slot in assignments:
            update_button_label(slot)

    def refresh_panel_key_display() -> None:
        display_text = _display_hotkey_text(panel_key, panel_key or "Unset")
        panel_key_display.set(display_text)

    refresh_panel_key_display()
    sync_auto_panel_state()
    auto_panel_var.trace_add("write", lambda *_: sync_auto_panel_state())

    def rebuild_listeners() -> None:
        manager.clear()
        _slot_hotkey_lookup.clear()
        if not listening:
            return
        for slot, _ in NUMPAD_SLOTS:
            hotkey = slot_hotkeys.get(slot)
            tpl = assignments.get(slot)
            if tpl is None or not hotkey:
                continue
            _slot_hotkey_lookup[hotkey.lower()] = slot
            try:
                manager.add_macro(
                    Macro(
                        hotkey,
                        resolve_template_keys(tpl, direction_keys),
                        macro_timing["delay"],
                        macro_timing["duration"],
                        name=tpl.name,
                    )
                )
            except ValueError as exc:
                log(f"Cannot register {tpl.name} for {hotkey}: {exc}")
        status_var.set("Listening for numpad keys (7 8 9 / 4 5 6 / 1 2 3).")

    def choose_macro_for_slot(slot: str) -> None:
        dialog = MacroSelectionDialog(root, f"Macro for numpad {slot}", templates)
        tpl = dialog.result
        if tpl is None:
            return
        assignments[slot] = tpl
        update_button_label(slot)
        if listening:
            rebuild_listeners()

    grid_frame = tk.Frame(root)

    auto_panel_frame = tk.Frame(root)
    auto_panel_frame.pack(fill=tk.X, padx=16, pady=(8, 0))
    auto_panel_toggle = tk.Checkbutton(
        auto_panel_frame,
        text="Auto Stratagem Panel",
        variable=auto_panel_var,
        bg=BG,
        fg=FG,
        activebackground=BUTTON_BG,
        activeforeground=FG,
        selectcolor=BG,
        highlightthickness=0,
        anchor="w",
    )
    auto_panel_toggle.pack(side=tk.LEFT)
    tk.Label(auto_panel_frame, text="Panel key:", anchor="w").pack(side=tk.LEFT, padx=(12, 4))
    tk.Label(auto_panel_frame, textvariable=panel_key_display, anchor="w").pack(side=tk.LEFT)

    grid_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)

    for r, row in enumerate(grid_layout):
        for c, slot in enumerate(row):
            cell = tk.Frame(grid_frame, width=SLOT_WIDTH, height=SLOT_HEIGHT, bg=BG, highlightthickness=0, bd=0)
            cell.grid(row=r, column=c, padx=6, pady=6, sticky="nsew")
            cell.grid_propagate(False)
            btn = tk.Button(
                cell,
                text=f"{slot}\nUnassigned",
                command=lambda s=slot: choose_macro_for_slot(s),
            )
            btn.pack(fill=tk.BOTH, expand=True)
            grid_frame.grid_columnconfigure(c, weight=1, uniform="slots")
            grid_frame.grid_rowconfigure(r, weight=1, uniform="slots")
            slot_buttons[slot] = btn

    controls_frame = tk.Frame(root)
    controls_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

    status_var = tk.StringVar(value="Not listening. Assign macros, then start listening.")

    def stop_listening() -> None:
        nonlocal listening
        if not listening:
            hide_overlay()
            return
        listening = False
        manager.clear()
        listen_btn.config(text="Start Listening")
        status_var.set("Not listening. Assign macros, then start listening.")
        log("Listener OFF.")
        hide_overlay()

    def start_listening() -> None:
        nonlocal listening
        if listening:
            return
        listening = True
        rebuild_listeners()
        listen_btn.config(text="Stop Listening")
        status_var.set("Listening for numpad keys (7 8 9 / 4 5 6 / 1 2 3).")
        log("Listener ON: waiting for assigned hotkeys.")
        show_overlay()

    def toggle_listening() -> None:
        if listening:
            stop_listening()
        else:
            start_listening()

    listen_btn = tk.Button(controls_frame, text="Start Listening", command=toggle_listening)
    listen_btn.pack(side=tk.LEFT, padx=(0, 8))

    status_label = tk.Label(controls_frame, textvariable=status_var, anchor="w")
    status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
    _update_overlay_lock_display()
    register_overlay_lock_hotkey()

    log_frame = tk.Frame(root)
    log_frame.pack(fill=tk.BOTH, expand=False, padx=16, pady=(0, 8))
    tk.Label(log_frame, text="Debug Log").pack(anchor="w")
    log_list = tk.Listbox(log_frame, height=6)
    log_scroll = tk.Scrollbar(log_frame, orient="vertical", command=log_list.yview)
    log_list.config(yscrollcommand=log_scroll.set)
    log_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    def ui_log(message: str) -> None:
        print(message)
        log_list.insert(tk.END, message)
        log_list.yview_moveto(1)

    set_log_callback(lambda msg: root.after(0, lambda m=msg: ui_log(m)))

    # --- Settings dialog ---
    def open_settings() -> None:
        settings = tk.Toplevel(root, bg=BG)
        settings.title("Settings - Key Binds")
        settings.resizable(True, True)
        place_window_near(settings, root)
        set_dark_titlebar(settings)

        status_local = tk.StringVar(value="Select a slot and press a key to rebind.")
        change_buttons: dict[str, tk.Button] = {}
        hotkey_labels: dict[str, tk.Label] = {}
        dir_buttons: dict[str, tk.Button] = {}
        dir_labels: dict[str, tk.Label] = {}
        panel_change_btn: tk.Button | None = None
        overlay_lock_btn: tk.Button | None = None
        panel_key_label_var = tk.StringVar(value=_display_hotkey_text(panel_key, panel_key))
        overlay_lock_label_var = tk.StringVar(value=_display_hotkey_text(overlay_lock_key, overlay_lock_key))
        overlay_opacity_label_var = tk.StringVar(value=f"{int(_clamp_opacity(overlay_opacity.get()) * 100)}%")
        capturing = {"active": False}
        pending_hotkeys: dict[str, str] = dict(slot_hotkeys)
        pending_direction_keys: dict[str, str] = dict(direction_keys)
        pending_panel_key = panel_key
        pending_overlay_lock_key = overlay_lock_key
        pending_overlay_opacity = tk.DoubleVar(value=_clamp_opacity(overlay_opacity.get()))

        tabs_frame = tk.Frame(settings)
        tabs_frame.pack(fill=tk.X, padx=10, pady=(10, 0))

        content_frame = tk.Frame(settings)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        slot_section = tk.Frame(content_frame)
        dir_section = tk.Frame(content_frame)
        delay_section = tk.Frame(content_frame)
        panel_section = tk.Frame(content_frame)
        overlay_section = tk.Frame(content_frame)

        active_values = {"delay": macro_timing["delay"], "duration": macro_timing["duration"]}

        current_category = {"val": None}

        def is_dirty() -> bool:
            return (
                pending_hotkeys != slot_hotkeys
                or pending_direction_keys != direction_keys
                or active_values["delay"] != macro_timing["delay"]
                or active_values["duration"] != macro_timing["duration"]
                or pending_panel_key != panel_key
                or pending_overlay_lock_key != overlay_lock_key
                or abs(_clamp_opacity(pending_overlay_opacity.get()) - _clamp_opacity(overlay_opacity.get())) > 1e-4
            )

        def refresh_labels() -> None:
            for slot, label in hotkey_labels.items():
                label.config(text=_display_hotkey_text(pending_hotkeys.get(slot, ""), slot))
            for direction, label in dir_labels.items():
                label.config(text=pending_direction_keys.get(direction, ""))
            panel_key_label_var.set(_display_hotkey_text(pending_panel_key, pending_panel_key))
            overlay_lock_label_var.set(_display_hotkey_text(pending_overlay_lock_key, pending_overlay_lock_key))
            overlay_opacity_label_var.set(f"{int(_clamp_opacity(pending_overlay_opacity.get()) * 100)}%")
            settings.update_idletasks()

        def _all_capture_buttons() -> list[tk.Button]:
            buttons: list[tk.Button] = list(change_buttons.values()) + list(dir_buttons.values())
            if panel_change_btn is not None:
                buttons.append(panel_change_btn)
            if overlay_lock_btn is not None:
                buttons.append(overlay_lock_btn)
            return buttons

        def finish_capture(
            target: str, new_key: str | None, error: str | None, was_listening: bool, kind: str
        ) -> None:
            nonlocal pending_panel_key
            capturing["active"] = False
            for btn in _all_capture_buttons():
                btn.config(state=tk.NORMAL)
            if new_key:
                if kind == "slot":
                    pending_hotkeys[target] = new_key
                    status_local.set(f"Numpad {target} pending bind to '{new_key}'. Click Apply to confirm.")
                elif kind == "direction":
                    pending_direction_keys[target] = new_key
                    status_local.set(f"{target} pending bind to '{new_key}'. Click Apply to confirm.")
                elif kind == "panel":
                    pending_panel_key = new_key
                    status_local.set(f"Stratagem Panel pending bind to '{new_key}'. Click Apply to confirm.")
                else:
                    pending_overlay_lock_key = new_key
                    status_local.set(f"Overlay lock toggle pending bind to '{new_key}'. Click Apply to confirm.")
                refresh_labels()
            elif error:
                status_local.set(error)
            else:
                status_local.set("No key captured.")

            if was_listening:
                rebuild_listeners()

        def start_capture(target: str, kind: str) -> None:
            if capturing["active"]:
                return
            capturing["active"] = True
            was_listening = listening
            if was_listening:
                manager.clear()
            if kind == "slot":
                prompt = f"Press a key to bind to numpad {target}..."
            elif kind == "direction":
                prompt = f"Press a key to bind to {target} direction..."
            elif kind == "panel":
                prompt = "Press a key to open the Stratagem Panel..."
            else:
                prompt = "Press a key to toggle the Overlay Lock..."
            status_local.set(prompt)
            for btn in _all_capture_buttons():
                btn.config(state=tk.DISABLED)

            def worker() -> None:
                try:
                    key = keyboard.read_key(suppress=False)
                    key = key.lower()
                    settings.after(0, lambda: finish_capture(target, key, None, was_listening, kind))
                except Exception as exc:  # noqa: BLE001
                    settings.after(
                        0, lambda: finish_capture(target, None, f"Capture failed: {exc}", was_listening, kind)
                    )

            threading.Thread(target=worker, daemon=True).start()

        def apply_and_stay() -> None:
            nonlocal panel_key, overlay_lock_key
            slot_hotkeys.update(pending_hotkeys)
            direction_keys.update(pending_direction_keys)
            macro_timing["delay"] = active_values["delay"]
            macro_timing["duration"] = active_values["duration"]
            panel_key = pending_panel_key
            overlay_lock_key = pending_overlay_lock_key
            overlay_opacity.set(_clamp_opacity(pending_overlay_opacity.get()))
            refresh_panel_key_display()
            sync_auto_panel_state()
            register_overlay_lock_hotkey()
            _update_overlay_lock_display()
            _apply_overlay_window_config()
            update_all_buttons()
            if listening:
                rebuild_listeners()
            status_local.set("Applied bindings and overlay settings.")

        def reset_pending_from_live() -> None:
            nonlocal pending_panel_key, pending_overlay_lock_key
            pending_hotkeys.clear()
            pending_hotkeys.update(slot_hotkeys)
            pending_direction_keys.clear()
            pending_direction_keys.update(direction_keys)
            pending_panel_key = panel_key
            panel_key_label_var.set(_display_hotkey_text(pending_panel_key, pending_panel_key))
            pending_overlay_lock_key = overlay_lock_key
            overlay_lock_label_var.set(_display_hotkey_text(pending_overlay_lock_key, pending_overlay_lock_key))
            pending_overlay_opacity.set(_clamp_opacity(overlay_opacity.get()))
            opacity_scale.set(pending_overlay_opacity.get())
            active_values["delay"] = macro_timing["delay"]
            active_values["duration"] = macro_timing["duration"]
            refresh_labels()

        def switch_category(target: str) -> None:
            if target == current_category["val"]:
                return
            if is_dirty():
                if messagebox.askyesno("Apply changes?", "Apply changes before switching categories?"):
                    apply_and_stay()
                else:
                    reset_pending_from_live()
            # hide all
            for frame in (slot_section, dir_section, delay_section, panel_section, overlay_section):
                frame.pack_forget()
            for btn in (slot_btn, dir_btn, panel_btn, delay_btn, overlay_btn):
                btn.config(relief=tk.RAISED)
            if target == "slot":
                slot_section.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
                slot_btn.config(relief=tk.SUNKEN)
            elif target == "direction":
                dir_section.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
                dir_btn.config(relief=tk.SUNKEN)
            elif target == "panel":
                panel_section.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
                panel_btn.config(relief=tk.SUNKEN)
            elif target == "overlay":
                overlay_section.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
                overlay_btn.config(relief=tk.SUNKEN)
            else:
                delay_section.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
                delay_btn.config(relief=tk.SUNKEN)
            current_category["val"] = target
            settings.update_idletasks()

        slot_btn = tk.Button(tabs_frame, text="Slot Hotkeys", command=lambda: switch_category("slot"))
        dir_btn = tk.Button(tabs_frame, text="Direction Keys", command=lambda: switch_category("direction"))
        panel_btn = tk.Button(tabs_frame, text="Panel Key", command=lambda: switch_category("panel"))
        delay_btn = tk.Button(tabs_frame, text="Macro Delay", command=lambda: switch_category("delay"))
        slot_btn.pack(side=tk.LEFT, padx=(0, 6))
        dir_btn.pack(side=tk.LEFT)
        panel_btn.pack(side=tk.LEFT, padx=(6, 6))
        delay_btn.pack(side=tk.LEFT)
        overlay_btn = tk.Button(tabs_frame, text="Overlay", command=lambda: switch_category("overlay"))
        overlay_btn.pack(side=tk.LEFT, padx=(6, 0))

        # Build slot section
        for slot, _ in NUMPAD_SLOTS:
            row = tk.Frame(slot_section)
            row.pack(fill=tk.X, pady=2)
            tk.Label(row, text=f"Numpad {slot}", width=12, anchor="w").pack(side=tk.LEFT)
            hotkey_labels[slot] = tk.Label(
                row, text=_display_hotkey_text(pending_hotkeys.get(slot, ""), slot), width=12, anchor="w"
            )
            hotkey_labels[slot].pack(side=tk.LEFT, padx=(0, 6))
            btn = tk.Button(row, text="Change", command=lambda s=slot: start_capture(s, "slot"))
            btn.pack(side=tk.LEFT)
            change_buttons[slot] = btn

        # Build direction section
        for direction in ("Up", "Down", "Left", "Right"):
            row = tk.Frame(dir_section)
            row.pack(fill=tk.X, pady=2)
            tk.Label(row, text=direction, width=12, anchor="w").pack(side=tk.LEFT)
            dir_labels[direction] = tk.Label(row, text=pending_direction_keys.get(direction, ""), width=12, anchor="w")
            dir_labels[direction].pack(side=tk.LEFT, padx=(0, 6))
            btn = tk.Button(row, text="Change", command=lambda d=direction: start_capture(d, "direction"))
            btn.pack(side=tk.LEFT)
            dir_buttons[direction] = btn

        # Build panel section
        row = tk.Frame(panel_section)
        row.pack(fill=tk.X, pady=4)
        tk.Label(row, text="Stratagem Panel", width=16, anchor="w").pack(side=tk.LEFT)
        panel_label = tk.Label(row, textvariable=panel_key_label_var, width=12, anchor="w")
        panel_label.pack(side=tk.LEFT, padx=(0, 6))
        panel_change_btn = tk.Button(row, text="Change", command=lambda: start_capture("panel", "panel"))
        panel_change_btn.pack(side=tk.LEFT)

        # Build overlay section
        row_overlay = tk.Frame(overlay_section)
        row_overlay.pack(fill=tk.X, pady=4)
        tk.Label(row_overlay, text="Overlay Lock Key", width=18, anchor="w").pack(side=tk.LEFT)
        tk.Label(row_overlay, textvariable=overlay_lock_label_var, width=12, anchor="w").pack(side=tk.LEFT, padx=(0, 6))
        overlay_lock_btn = tk.Button(row_overlay, text="Change", command=lambda: start_capture("overlay_lock", "overlay_lock"))
        overlay_lock_btn.pack(side=tk.LEFT)

        tk.Label(overlay_section, text="Overlay Opacity", anchor="w").pack(fill=tk.X, pady=(12, 4))
        op_row = tk.Frame(overlay_section)
        op_row.pack(fill=tk.X, pady=(0, 8))
        opacity_scale = tk.Scale(
            op_row,
            from_=0.3,
            to=1.0,
            resolution=0.01,
            orient=tk.HORIZONTAL,
            showvalue=False,
            length=200,
            bg=BG,
            fg=FG,
            troughcolor=BUTTON_BG,
            highlightthickness=0,
            command=lambda v: None,
        )
        opacity_scale.set(pending_overlay_opacity.get())
        opacity_scale.pack(side=tk.LEFT, padx=(0, 8))
        tk.Label(op_row, textvariable=overlay_opacity_label_var, width=6, anchor="w").pack(side=tk.LEFT)

        def on_overlay_opacity_change(val: str) -> None:
            pending_overlay_opacity.set(_clamp_opacity(float(val)))
            overlay_opacity_label_var.set(f"{int(_clamp_opacity(float(val)) * 100)}%")

        opacity_scale.config(command=on_overlay_opacity_change)

        # Build delay section
        tk.Label(delay_section, text="Milliseconds between key presses:", anchor="w").pack(
            fill=tk.X, pady=(4, 4)
        )
        delay_var = tk.StringVar(value=str(int(macro_timing["delay"] * 1000)))
        delay_entry = tk.Entry(delay_section, textvariable=delay_var, width=10)
        delay_entry.pack(anchor="w", pady=(0, 4))

        tk.Label(delay_section, text="Milliseconds key stays pressed:", anchor="w").pack(
            fill=tk.X, pady=(8, 4)
        )
        duration_var = tk.StringVar(value=str(int(macro_timing["duration"] * 1000)))
        duration_entry = tk.Entry(delay_section, textvariable=duration_var, width=10)
        duration_entry.pack(anchor="w", pady=(0, 4))

        def apply_delay() -> None:
            try:
                ms_delay = float(delay_var.get())
                ms_duration = float(duration_var.get())
                if ms_delay < 0 or ms_duration < 0:
                    raise ValueError
                active_values["delay"] = ms_delay / 1000.0
                active_values["duration"] = ms_duration / 1000.0
                status_local.set(
                    f"Pending delay {ms_delay:.0f} ms, duration {ms_duration:.0f} ms. Click Apply to confirm."
                )
            except ValueError:
                messagebox.showerror("Invalid delay/duration", "Enter non-negative numbers (milliseconds).")

        tk.Button(delay_section, text="Apply Delay", command=apply_delay).pack(anchor="w")

        switch_category("slot")

        status_label_local = tk.Label(settings, textvariable=status_local, anchor="w", wraplength=320)
        status_label_local.pack(fill=tk.X, padx=10, pady=(4, 8))

        btn_frame = tk.Frame(settings)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        tk.Button(btn_frame, text="Apply", command=apply_and_stay).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Close", command=settings.destroy).pack(side=tk.RIGHT)

        apply_dark_theme(settings)
        refresh_labels()

    # --- Persistence ---
    def open_edit_templates() -> None:
        nonlocal templates
        edit = tk.Toplevel(root, bg=BG)
        edit.title("Edit Stratagem Templates")
        edit.resizable(True, True)
        place_window_near(edit, root)
        set_dark_titlebar(edit)

        working: list[MacroTemplate] = list(templates)
        filter_var = tk.StringVar()
        listbox = tk.Listbox(edit, height=16)
        scroll = tk.Scrollbar(edit, orient="vertical", command=listbox.yview)
        listbox.config(yscrollcommand=scroll.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=(0, 10))
        scroll.pack(side=tk.LEFT, fill=tk.Y, pady=(0, 10))

        detail = tk.Frame(edit, bg=BG)
        detail.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        name_var = tk.StringVar()
        category_var = tk.StringVar()
        directions_text = tk.Text(detail, height=4, width=30, bg=ENTRY_BG, fg=FG, insertbackground=FG)
        status_var = tk.StringVar(value="Select a template to edit.")

        # Search bar
        search_frame = tk.Frame(edit, bg=BG)
        search_frame.pack(fill=tk.X, padx=10, pady=(10, 6))
        tk.Label(search_frame, text="Search", anchor="w").pack(side=tk.LEFT, padx=(0, 6))
        search_entry = tk.Entry(search_frame, textvariable=filter_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(detail, text="Name (read-only):", anchor="w").pack(fill=tk.X)
        tk.Label(detail, textvariable=name_var, anchor="w").pack(fill=tk.X, pady=(0, 6))
        tk.Label(detail, text="Category:", anchor="w").pack(fill=tk.X)
        tk.Entry(detail, textvariable=category_var).pack(fill=tk.X, pady=(0, 6))
        tk.Label(detail, text="Directions (comma separated):", anchor="w").pack(fill=tk.X)
        directions_text.pack(fill=tk.BOTH, expand=False, pady=(0, 6))
        tk.Label(detail, textvariable=status_var, anchor="w", wraplength=280).pack(fill=tk.X, pady=(0, 8))

        btns = tk.Frame(detail, bg=BG)
        btns.pack(fill=tk.X)

        current_index = {"val": None}

        def refresh_list(select: int | None = None) -> None:
            listbox.delete(0, tk.END)
            query = filter_var.get().strip().lower()
            filtered: list[MacroTemplate] = []
            for tpl in working:
                if not query or query in tpl.name.lower() or query in tpl.category.lower():
                    filtered.append(tpl)
                    listbox.insert(tk.END, tpl.name)
            listbox._filtered = filtered  # type: ignore[attr-defined]
            if select is not None and listbox.size() and 0 <= select < listbox.size():
                listbox.selection_set(select)
                listbox.see(select)
                load_selection(select)

        def load_selection(idx: int) -> None:
            filtered = getattr(listbox, "_filtered", working)  # type: ignore[attr-defined]
            if not (0 <= idx < len(filtered)):
                return
            tpl = filtered[idx]
            # find real index in working
            try:
                real_idx = working.index(tpl)
            except ValueError:
                real_idx = idx
            current_index["val"] = idx
            name_var.set(tpl.name)
            category_var.set(tpl.category)
            directions_text.delete("1.0", tk.END)
            directions_text.insert(tk.END, ", ".join(tpl.directions))
            status_var.set("Edit category or directions, then Apply.")

        def parse_directions(raw: str) -> list[str]:
            vals = [part.strip() for part in raw.split(",") if part.strip()]
            if not vals:
                raise ValueError("Enter at least one direction (comma separated).")
            return [v.title() for v in vals]

        def apply_current() -> bool:
            idx = current_index["val"]
            filtered = getattr(listbox, "_filtered", working)  # type: ignore[attr-defined]
            if idx is None or not (0 <= idx < len(filtered)):
                status_var.set("Select a template first.")
                return False
            try:
                directions = tuple(parse_directions(directions_text.get("1.0", tk.END)))
            except ValueError as exc:
                messagebox.showerror("Invalid directions", str(exc))
                return False
            tpl = filtered[idx]
            category = category_var.get().strip() or tpl.category
            updated = MacroTemplate(tpl.name, directions, tpl.delay, category=category)
            # replace in working
            for i, item in enumerate(working):
                if item.name == tpl.name:
                    working[i] = updated
                    break
            status_var.set(f"Updated {tpl.name}. Remember to Save Templates.")
            return True

        def save_all_and_refresh() -> None:
            if not apply_current():
                return
            save_stratagem_templates(tuple(working))
            templates = tuple(working)
            name_map = {tpl.name: tpl for tpl in templates}
            for slot, tpl in list(assignments.items()):
                if tpl is None:
                    continue
                assignments[slot] = name_map.get(tpl.name, tpl)
            update_all_buttons()
            if listening:
                rebuild_listeners()
            status_var.set("Templates saved.")
            messagebox.showinfo("Templates saved", "Stratagem templates saved and reloaded.")

        tk.Button(btns, text="Apply Changes", command=apply_current).pack(side=tk.LEFT)
        tk.Button(btns, text="Save Templates", command=save_all_and_refresh).pack(side=tk.RIGHT)

        def on_select(event=None):  # noqa: ANN001
            sel = listbox.curselection()
            if not sel:
                return
            load_selection(sel[0])

        def on_search(*args):  # noqa: ANN001
            refresh_list(0)

        listbox.bind("<<ListboxSelect>>", on_select)
        filter_var.trace_add("write", on_search)
        if working:
            refresh_list(0)
        else:
            status_var.set("No templates to edit.")

        apply_dark_theme(edit)
        edit.grab_set()

    def _save_profile_to_path(path: Path, show_message: bool = True) -> bool:
        data = serialize_state()
        try:
            with path.open("w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
        except OSError as exc:
            messagebox.showerror("Cannot save profile", str(exc))
            return False
        if show_message:
            messagebox.showinfo("Profile saved", f"Saved to {path.name}.")
        _record_last_profile(path)
        return True

    def save_profile_action() -> None:
        nonlocal saved_state
        saves_dir = ensure_saves_dir()
        path_str = filedialog.asksaveasfilename(
            title="Save profile",
            initialdir=saves_dir,
            defaultextension=".json",
            filetypes=[("Profile files", "*.json"), ("All files", "*.*")],
            initialfile="profile.json",
        )
        if not path_str:
            return
        path = Path(path_str)
        if _save_profile_to_path(path):
            saved_state = serialize_state()

    def _load_profile_from_path(path: Path, show_messages: bool = True) -> bool:
        nonlocal saved_state, panel_key, overlay_lock_key
        try:
            with path.open("r", encoding="utf-8") as fh:
                content = json.load(fh)
        except (OSError, json.JSONDecodeError) as exc:
            if show_messages:
                messagebox.showerror("Cannot load profile", f"Failed to read file: {exc}")
            return False

        slots_data = content.get("slots")
        if not isinstance(slots_data, dict):
            if show_messages:
                messagebox.showerror("Cannot load profile", "Invalid profile format.")
            return False

        missing: list[str] = []
        for slot, _ in NUMPAD_SLOTS:
            name = slots_data.get(slot)
            if name is None:
                assignments[slot] = None
                continue
            tpl = next((t for t in templates if t.name == name), None)
            if tpl is None:
                missing.append(name)
                assignments[slot] = None
            else:
                assignments[slot] = tpl

        hotkeys_data = content.get("hotkeys")
        if isinstance(hotkeys_data, dict):
            for slot, _ in NUMPAD_SLOTS:
                hk = hotkeys_data.get(slot)
                if isinstance(hk, str) and hk.strip():
                    slot_hotkeys[slot] = hk.strip()
                else:
                    slot_hotkeys[slot] = DEFAULT_SLOT_HOTKEYS.get(slot, slot_hotkeys[slot])

        direction_data = content.get("direction_keys")
        if isinstance(direction_data, dict):
            for direction in DEFAULT_DIRECTION_KEYS:
                val = direction_data.get(direction)
                if isinstance(val, str) and val.strip():
                    direction_keys[direction] = val.strip()
                else:
                    direction_keys[direction] = DEFAULT_DIRECTION_KEYS[direction]

        timing_data = content.get("timing")
        if isinstance(timing_data, dict):
            delay_val = timing_data.get("delay", macro_timing["delay"])
            duration_val = timing_data.get("duration", macro_timing["duration"])
            try:
                delay_f = float(delay_val)
                duration_f = float(duration_val)
                if delay_f >= 0 and duration_f >= 0:
                    macro_timing["delay"] = delay_f
                    macro_timing["duration"] = duration_f
            except (TypeError, ValueError):
                pass

        panel_data = content.get("panel")
        if isinstance(panel_data, dict):
            key_val = panel_data.get("key")
            if isinstance(key_val, str) and key_val.strip():
                panel_key = key_val.strip().lower()
            else:
                panel_key = DEFAULT_PANEL_KEY
            auto_val = panel_data.get("auto")
            if isinstance(auto_val, bool):
                auto_panel_var.set(auto_val)
            else:
                auto_panel_var.set(DEFAULT_AUTO_PANEL)
        else:
            panel_key = DEFAULT_PANEL_KEY
            auto_panel_var.set(DEFAULT_AUTO_PANEL)
        overlay_data = content.get("overlay")
        if isinstance(overlay_data, dict):
            lock_val = overlay_data.get("lock_key")
            if isinstance(lock_val, str) and lock_val.strip():
                overlay_lock_key = lock_val.strip()
            op_val = overlay_data.get("opacity", OVERLAY_ALPHA)
            try:
                overlay_opacity.set(_clamp_opacity(float(op_val)))
            except (TypeError, ValueError):
                overlay_opacity.set(OVERLAY_ALPHA)
        else:
            overlay_lock_key = DEFAULT_OVERLAY_LOCK_KEY
            overlay_opacity.set(OVERLAY_ALPHA)
        overlay_locked["val"] = False
        register_overlay_lock_hotkey()
        _update_overlay_lock_display()
        _apply_overlay_window_config()
        refresh_panel_key_display()
        sync_auto_panel_state()

        update_all_buttons()
        if listening:
            rebuild_listeners()

        if show_messages:
            if missing:
                messagebox.showwarning(
                    "Profile loaded with missing macros",
                    f"Loaded {path}, but these macros were not found: {', '.join(missing)}.",
                )
            else:
                messagebox.showinfo("Profile loaded", f"Loaded {path}.")
        saved_state = serialize_state()
        _record_last_profile(path)
        return True

    def load_profile_action() -> None:
        saves_dir = ensure_saves_dir()
        file_path = filedialog.askopenfilename(
            title="Load profile",
            initialdir=saves_dir,
            filetypes=[("Profile files", "*.json"), ("All files", "*.*")],
        )
        if not file_path:
            return
        _load_profile_from_path(Path(file_path), show_messages=True)

    def load_blank_profile(show_messages: bool = True) -> None:
        """Reset to defaults and clear last-profile marker."""
        nonlocal saved_state, panel_key, overlay_lock_key
        assignments.clear()
        assignments.update({slot: None for slot, _ in NUMPAD_SLOTS})
        slot_hotkeys.clear()
        slot_hotkeys.update(DEFAULT_SLOT_HOTKEYS)
        direction_keys.clear()
        direction_keys.update(DEFAULT_DIRECTION_KEYS)
        panel_key = DEFAULT_PANEL_KEY
        overlay_lock_key = DEFAULT_OVERLAY_LOCK_KEY
        overlay_locked["val"] = False
        overlay_opacity.set(OVERLAY_ALPHA)
        auto_panel_var.set(DEFAULT_AUTO_PANEL)
        macro_timing["delay"] = DEFAULT_DELAY
        macro_timing["duration"] = DEFAULT_DURATION
        update_all_buttons()
        refresh_panel_key_display()
        sync_auto_panel_state()
        register_overlay_lock_hotkey()
        _update_overlay_lock_display()
        _apply_overlay_window_config()
        if listening:
            rebuild_listeners()
        saved_state = serialize_state()
        try:
            if last_profile_marker.exists():
                last_profile_marker.unlink()
        except OSError:
            pass
        if show_messages:
            messagebox.showinfo("New profile", "Reset to the blank default profile.")

    # --- Menus (custom dark bar) ---
    def _popup_menu(menu: tk.Menu, btn: tk.Button) -> None:
        try:
            x = btn.winfo_rootx()
            y = btn.winfo_rooty() + btn.winfo_height()
            menu.tk_popup(x, y)
        finally:
            try:
                menu.grab_release()
            except tk.TclError:
                pass

    menu_bar = tk.Frame(root, bg=MENU_BG, bd=0, highlightthickness=0)
    # Place the menu bar above the auto-panel row.
    menu_bar.pack(fill=tk.X, side=tk.TOP, before=auto_panel_frame)

    file_menu = tk.Menu(
        root,
        tearoff=False,
        bg=MENU_BG,
        fg=FG,
        activebackground=BUTTON_ACTIVE,
        activeforeground=FG,
        relief=tk.FLAT,
        bd=0,
    )
    file_menu.add_command(label="Save Profile", command=save_profile_action)
    file_menu.add_command(label="Load Profile", command=load_profile_action)
    file_menu.add_command(
        label="New",
        command=lambda: load_blank_profile(show_messages=True),
    )
    file_menu.add_command(label="Settings", command=open_settings)
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=lambda: root.event_generate("<<RequestExit>>"))

    edit_menu = tk.Menu(
        root,
        tearoff=False,
        bg=MENU_BG,
        fg=FG,
        activebackground=BUTTON_ACTIVE,
        activeforeground=FG,
        relief=tk.FLAT,
        bd=0,
    )
    edit_menu.add_command(label="Edit Stratagem Templates", command=open_edit_templates)

    def show_about() -> None:
        about = tk.Toplevel(root, bg=BG)
        about.title("About")
        about.transient(root)
        about.resizable(False, False)
        place_window_near(about, root)

        text = (
            "HELLDIVERS2 Stratagem Macro\nCreated by FatterCatDev\n\n"
            "Assign templates to the numpad grid, then start listening to trigger them with numpad keys.\n"
            "Exit via File > Exit or Ctrl+Shift+Q."
        )

        tk.Label(about, text=text, justify="left", anchor="w", wraplength=320).pack(
            fill=tk.BOTH, expand=True, padx=16, pady=(16, 8)
        )
        tk.Button(about, text="OK", command=about.destroy).pack(pady=(0, 12))

        apply_dark_theme(about)
        set_dark_titlebar(about)
        about.grab_set()
        about.focus_set()

    def maybe_save_before_exit() -> bool:
        nonlocal saved_state
        if not has_unsaved_changes():
            return True
        resp = messagebox.askyesnocancel(
            "Save changes?",
            "Save your profile before exiting?",
            icon=messagebox.QUESTION,
        )
        if resp is None:  # Cancel
            return False
        if resp is False:  # Don't save
            return True
        # Save selected
        saves_dir = ensure_saves_dir()
        path_str = filedialog.asksaveasfilename(
            title="Save profile",
            initialdir=saves_dir,
            defaultextension=".json",
            filetypes=[("Profile files", "*.json"), ("All files", "*.*")],
            initialfile="profile.json",
        )
        if not path_str:
            return False
        path = Path(path_str)
        if _save_profile_to_path(path, show_message=False):
            saved_state = serialize_state()
            messagebox.showinfo("Profile saved", f"Saved to {path.name}.")
            return True
        return False

    help_menu = tk.Menu(
        root,
        tearoff=False,
        bg=MENU_BG,
        fg=FG,
        activebackground=BUTTON_ACTIVE,
        activeforeground=FG,
        relief=tk.FLAT,
        bd=0,
    )
    help_menu.add_command(
        label="About",
        command=show_about,
    )

    def _add_menu_button(label: str, menu: tk.Menu) -> None:
        btn = tk.Button(
            menu_bar,
            text=label,
            bg=MENU_BG,
            fg=FG,
            relief=tk.FLAT,
            bd=0,
            padx=10,
            pady=6,
            highlightthickness=0,
            activebackground=BUTTON_ACTIVE,
            activeforeground=FG,
        )
        btn.config(command=lambda b=btn, m=menu: _popup_menu(m, b))
        btn.pack(side=tk.LEFT)

    _add_menu_button("File", file_menu)
    _add_menu_button("Edit", edit_menu)
    _add_menu_button("Help", help_menu)

    def _refresh_menu_bar_colors() -> None:
        menu_bar.configure(bg=MENU_BG)
        for child in menu_bar.winfo_children():
            try:
                child.configure(
                    bg=MENU_BG,
                    fg=FG,
                    activebackground=BUTTON_ACTIVE,
                    activeforeground=FG,
                    relief=tk.FLAT,
                    bd=0,
                    highlightthickness=0,
                )
            except tk.TclError:
                pass

    _refresh_menu_bar_colors()

    exit_handle: int | None = None

    def close_app() -> None:
        nonlocal exit_handle, overlay_win, overlay_lock_handle
        if exit_handle is not None:
            keyboard.remove_hotkey(exit_handle)
            exit_handle = None
        if overlay_lock_handle is not None:
            try:
                keyboard.remove_hotkey(overlay_lock_handle)
            except Exception:
                pass
            overlay_lock_handle = None
        manager.clear()
        clear_log_callback()
        if overlay_win is not None and overlay_win.winfo_exists():
            overlay_win.destroy()
            overlay_win = None
        global _macro_progress_callback
        _macro_progress_callback = None
        for slot in list(overlay_progress.keys()):
            _cancel_overlay_progress(slot)
        root.destroy()

    def attempt_exit() -> None:
        if maybe_save_before_exit():
            close_app()

    def request_exit() -> None:
        root.after(0, attempt_exit)

    root.protocol("WM_DELETE_WINDOW", attempt_exit)
    root.bind("<<RequestExit>>", lambda _: attempt_exit())
    exit_handle = keyboard.add_hotkey(EXIT_HOTKEY, request_exit, suppress=False)

    saved_state = serialize_state()
    update_all_buttons()
    apply_dark_theme(root)
    _refresh_menu_bar_colors()
    set_dark_titlebar(root)  # Re-apply after widgets are realized.

    # Auto-load last profile if available.
    if last_profile_marker.exists():
        try:
            last_path = Path(last_profile_marker.read_text(encoding="utf-8").strip())
            if last_path.exists():
                if _load_profile_from_path(last_path, show_messages=False):
                    log(f"Loaded last profile: {last_path.name}")
            else:
                log(f"Last profile not found; expected at {last_path}")
        except OSError:
            pass

    root.mainloop()


if __name__ == "__main__":
    main()

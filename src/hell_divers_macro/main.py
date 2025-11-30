"""
Tkinter desktop app for binding HELLDIVERS 2 stratagem macros to a numpad grid.

The UI pieces are split into small helpers:
- MacroManager handles keyboard hooks and running macros.
- OverlayWindow mirrors slot assignments in a floating window.
- Dialog helpers live in ui.dialogs.
"""

from __future__ import annotations

import json
import sys
import threading
from pathlib import Path
from typing import Dict

import keyboard
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import ImageTk

if __package__ in (None, ""):
    # Allow running as a script by adding project src to path for absolute imports.
    sys.path.append(str(Path(__file__).resolve().parent.parent))

from hell_divers_macro.config import (  # noqa: E402
    DEFAULT_AUTO_PANEL,
    DEFAULT_DELAY,
    DEFAULT_DIRECTION_KEYS,
    DEFAULT_DURATION,
    DEFAULT_OVERLAY_LOCK_KEY,
    DEFAULT_OVERLAY_OPACITY,
    DEFAULT_PANEL_KEY,
    DEFAULT_SLOT_HOTKEYS,
    EXIT_HOTKEY,
    NUMPAD_SLOTS,
)
from hell_divers_macro.log_utils import clear_log_callback, log, set_log_callback  # noqa: E402
from hell_divers_macro.macro_manager import MacroManager  # noqa: E402
from hell_divers_macro.models import Macro, MacroTemplate  # noqa: E402
from hell_divers_macro.paths import ensure_saves_dir  # noqa: E402
from hell_divers_macro.state import AppState  # noqa: E402
from hell_divers_macro.stratagems import (  # noqa: E402
    load_stratagem_templates,
    resolve_template_keys,
    save_stratagem_templates,
)
from hell_divers_macro.ui.dialogs import MacroSelectionDialog  # noqa: E402
from hell_divers_macro.ui.icons import (  # noqa: E402
    APP_ICON_PATH,
    ICON_SIZE,
    OVERLAY_ALPHA,
    load_icon_image,
)
from hell_divers_macro.ui.overlay import OverlayWindow  # noqa: E402
from hell_divers_macro.ui.theme import (  # noqa: E402
    BG,
    BUTTON_ACTIVE,
    BUTTON_BG,
    ENTRY_BG,
    FG,
    MENU_BG,
    apply_dark_theme,
    init_base_theme,
    place_window_near,
    set_dark_titlebar,
)


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


class MacroApp:
    def __init__(self) -> None:
        self.state = AppState()
        ensure_saves_dir()
        self.templates: tuple[MacroTemplate, ...] = load_stratagem_templates()
        self.last_profile_marker = ensure_saves_dir() / ".last_profile"
        self.saved_state = self.state.serialize()

        self.root = tk.Tk()
        self.root.title("HELLDIVERS2 Stratagem Macro")
        self.root.geometry("540x560")
        self.root.minsize(520, 520)
        try:
            if APP_ICON_PATH.exists():
                self.root.iconphoto(True, tk.PhotoImage(file=str(APP_ICON_PATH)))
        except Exception:
            pass

        init_base_theme(self.root)

        self.slot_buttons: dict[str, tk.Button] = {}
        self.slot_icons: dict[str, ImageTk.PhotoImage | None] = {}
        self.status_var = tk.StringVar(value="Not listening. Assign macros, then start listening.")
        self.auto_panel_var = tk.BooleanVar(value=self.state.auto_panel)
        self.panel_key_display = tk.StringVar(value="")
        self.overlay_opacity_var = tk.DoubleVar(value=self.state.overlay_opacity)
        self.overlay_lock_handle: int | None = None
        self.exit_handle: int | None = None
        self.listening = False
        self.log_list: tk.Listbox | None = None

        self.manager = MacroManager(
            progress_callback=None,
            auto_panel_key=self.state.panel_key,
            auto_panel_enabled=self.state.auto_panel,
        )
        self.overlay = OverlayWindow(
            self.root,
            auto_panel_var=self.auto_panel_var,
            opacity_var=self.overlay_opacity_var,
            initial_lock_key=self.state.overlay_lock_key,
            status_callback=self.status_var.set,
            hotkey_display=_display_hotkey_text,
            on_close=self.stop_listening,
        )
        self.manager.set_progress_callback(self.overlay.handle_macro_progress)

        self._build_ui()
        self._refresh_menu_bar_colors()
        apply_dark_theme(self.root)
        set_dark_titlebar(self.root)
        self._refresh_panel_display()
        self._sync_auto_panel_state()
        self._register_overlay_lock_hotkey()
        self._register_exit_hotkey()
        self._load_last_profile()
        self.root.protocol("WM_DELETE_WINDOW", self._attempt_exit)
        self.root.bind("<<RequestExit>>", lambda _: self._attempt_exit())

    # -- UI construction -------------------------------------------------- #
    def _build_ui(self) -> None:
        self._build_menu_bar()
        self._build_auto_panel_row()
        self._build_grid()
        self._build_controls()
        self._build_log()
        self.root.bind("<Configure>", lambda event=None: self._on_root_configure())

    def _build_menu_bar(self) -> None:
        self.menu_bar = tk.Frame(self.root, bg=MENU_BG, bd=0, highlightthickness=0)
        self.menu_bar.pack(fill=tk.X, side=tk.TOP)

        file_menu = tk.Menu(
            self.root,
            tearoff=False,
            bg=MENU_BG,
            fg=FG,
            activebackground=BUTTON_ACTIVE,
            activeforeground=FG,
            relief=tk.FLAT,
            bd=0,
        )
        file_menu.add_command(label="Save Profile", command=self._save_profile_action)
        file_menu.add_command(label="Load Profile", command=self._load_profile_action)
        file_menu.add_command(label="New", command=lambda: self._load_blank_profile(show_messages=True))
        file_menu.add_command(label="Settings", command=self._open_settings)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=lambda: self.root.event_generate("<<RequestExit>>"))

        edit_menu = tk.Menu(
            self.root,
            tearoff=False,
            bg=MENU_BG,
            fg=FG,
            activebackground=BUTTON_ACTIVE,
            activeforeground=FG,
            relief=tk.FLAT,
            bd=0,
        )
        edit_menu.add_command(label="Edit Stratagem Templates", command=self._open_edit_templates)

        help_menu = tk.Menu(
            self.root,
            tearoff=False,
            bg=MENU_BG,
            fg=FG,
            activebackground=BUTTON_ACTIVE,
            activeforeground=FG,
            relief=tk.FLAT,
            bd=0,
        )
        help_menu.add_command(label="About", command=self._show_about)

        self._add_menu_button("File", file_menu)
        self._add_menu_button("Edit", edit_menu)
        self._add_menu_button("Help", help_menu)

    def _add_menu_button(self, label: str, menu: tk.Menu) -> None:
        btn = tk.Button(
            self.menu_bar,
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
        btn.config(command=lambda b=btn, m=menu: self._popup_menu(m, b))
        btn.pack(side=tk.LEFT)

    def _build_auto_panel_row(self) -> None:
        self.auto_panel_frame = tk.Frame(self.root, bg=BG)
        self.auto_panel_frame.pack(fill=tk.X, padx=16, pady=(8, 0))
        auto_panel_toggle = tk.Checkbutton(
            self.auto_panel_frame,
            text="Auto Stratagem Panel",
            variable=self.auto_panel_var,
            bg=BG,
            fg=FG,
            activebackground=BUTTON_BG,
            activeforeground=FG,
            selectcolor=BG,
            highlightthickness=0,
            anchor="w",
            command=self._sync_auto_panel_state,
        )
        auto_panel_toggle.pack(side=tk.LEFT)
        tk.Label(self.auto_panel_frame, text="Panel key:", anchor="w").pack(side=tk.LEFT, padx=(12, 4))
        tk.Label(self.auto_panel_frame, textvariable=self.panel_key_display, anchor="w").pack(side=tk.LEFT)

    def _build_grid(self) -> None:
        grid_layout = [["7", "8", "9"], ["4", "5", "6"], ["1", "2", "3"]]
        grid_frame = tk.Frame(self.root, bg=BG)
        grid_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)
        for r, row in enumerate(grid_layout):
            for c, slot in enumerate(row):
                cell = tk.Frame(grid_frame, width=160, height=150, bg=BG, highlightthickness=0, bd=0)
                cell.grid(row=r, column=c, padx=6, pady=6, sticky="nsew")
                cell.grid_propagate(False)
                btn = tk.Button(
                    cell,
                    text=f"{slot}\nUnassigned",
                    command=lambda s=slot: self._choose_macro_for_slot(s),
                )
                btn.pack(fill=tk.BOTH, expand=True)
                grid_frame.grid_columnconfigure(c, weight=1, uniform="slots")
                grid_frame.grid_rowconfigure(r, weight=1, uniform="slots")
                self.slot_buttons[slot] = btn
        self._update_all_buttons()

    def _build_controls(self) -> None:
        controls_frame = tk.Frame(self.root, bg=BG)
        controls_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        self.listen_btn = tk.Button(controls_frame, text="Start Listening", command=self._toggle_listening)
        self.listen_btn.pack(side=tk.LEFT, padx=(0, 8))

        status_label = tk.Label(controls_frame, textvariable=self.status_var, anchor="w")
        status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _build_log(self) -> None:
        log_frame = tk.Frame(self.root, bg=BG)
        log_frame.pack(fill=tk.BOTH, expand=False, padx=16, pady=(0, 8))
        tk.Label(log_frame, text="Debug Log").pack(anchor="w")
        self.log_list = tk.Listbox(log_frame, height=6)
        log_scroll = tk.Scrollbar(log_frame, orient="vertical", command=self.log_list.yview)
        self.log_list.config(yscrollcommand=log_scroll.set)
        self.log_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        def ui_log(message: str) -> None:
            print(message)
            if self.log_list is None:
                return
            self.log_list.insert(tk.END, message)
            self.log_list.yview_moveto(1)

        set_log_callback(lambda msg: self.root.after(0, lambda m=msg: ui_log(m)))

    # -- State helpers ---------------------------------------------------- #
    def _serialize_state(self) -> dict:
        return self.state.serialize()

    def _has_unsaved_changes(self) -> bool:
        return self._serialize_state() != self.saved_state

    def _refresh_panel_display(self) -> None:
        display_text = _display_hotkey_text(self.state.panel_key, self.state.panel_key or "Unset")
        self.panel_key_display.set(display_text)

    def _sync_auto_panel_state(self) -> None:
        self.state.auto_panel = bool(self.auto_panel_var.get())
        self.manager.set_auto_panel(self.state.auto_panel, self.state.panel_key)

    def _update_slot_button(self, slot: str) -> None:
        tpl = self.state.assignments.get(slot)
        hotkey = self.state.slot_hotkeys.get(slot, slot)
        hotkey_text = _display_hotkey_text(hotkey, slot)
        icon: ImageTk.PhotoImage | None = None
        if tpl:
            icon = load_icon_image(tpl.name, hotkey_text)
        self.slot_icons[slot] = icon
        if icon:
            self.slot_buttons[slot].config(image=icon, text="", compound=tk.CENTER)
        else:
            name = tpl.name if tpl else "Unassigned"
            self.slot_buttons[slot].config(text=f"{hotkey_text}\n{name}", image="", compound=tk.NONE)
        self.overlay.update_slot(slot, tpl, hotkey)

    def _update_all_buttons(self) -> None:
        for slot in self.state.assignments:
            self._update_slot_button(slot)

    def _rebuild_listeners(self) -> None:
        self.manager.clear()
        if not self.listening:
            return
        macros: Dict[str, Macro] = {}
        for slot, _ in NUMPAD_SLOTS:
            tpl = self.state.assignments.get(slot)
            hotkey = self.state.slot_hotkeys.get(slot)
            if tpl is None or not hotkey:
                continue
            macros[slot] = Macro(
                hotkey.lower(),
                resolve_template_keys(tpl, self.state.direction_keys),
                self.state.macro_delay,
                self.state.macro_duration,
                name=tpl.name,
            )
        self.manager.set_auto_panel(self.state.auto_panel, self.state.panel_key)
        self.manager.register_macros(macros)
        self.status_var.set("Listening for numpad keys (7 8 9 / 4 5 6 / 1 2 3).")

    # -- Actions ---------------------------------------------------------- #
    def _choose_macro_for_slot(self, slot: str) -> None:
        dialog = MacroSelectionDialog(self.root, f"Macro for numpad {slot}", self.templates)
        tpl = dialog.result
        if tpl is None:
            return
        self.state.assignments[slot] = tpl
        self._update_slot_button(slot)
        if self.listening:
            self._rebuild_listeners()

    def _start_listening(self) -> None:
        if self.listening:
            return
        self.listening = True
        self._rebuild_listeners()
        self.listen_btn.config(text="Stop Listening")
        self.status_var.set("Listening for numpad keys (7 8 9 / 4 5 6 / 1 2 3).")
        log("Listener ON: waiting for assigned hotkeys.")
        self.overlay.show(self.state.assignments, self.state.slot_hotkeys)

    def stop_listening(self) -> None:
        if not self.listening:
            self.overlay.hide()
            return
        self.listening = False
        self.manager.clear()
        self.listen_btn.config(text="Start Listening")
        self.status_var.set("Not listening. Assign macros, then start listening.")
        log("Listener OFF.")
        self.overlay.hide()

    def _toggle_listening(self) -> None:
        if self.listening:
            self.stop_listening()
        else:
            self._start_listening()

    def _save_profile_to_path(self, path: Path, show_message: bool = True) -> bool:
        data = self._serialize_state()
        try:
            with path.open("w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
        except OSError as exc:
            messagebox.showerror("Cannot save profile", str(exc))
            return False
        if show_message:
            messagebox.showinfo("Profile saved", f"Saved to {path.name}.")
        self._record_last_profile(path)
        return True

    def _save_profile_action(self) -> None:
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
        if self._save_profile_to_path(path):
            self.saved_state = self._serialize_state()

    def _load_profile_from_path(self, path: Path, show_messages: bool = True) -> bool:
        try:
            with path.open("r", encoding="utf-8") as fh:
                content = json.load(fh)
        except (OSError, json.JSONDecodeError) as exc:
            if show_messages:
                messagebox.showerror("Cannot load profile", f"Failed to read file: {exc}")
            return False
        missing = self.state.apply_profile(content, self.templates)
        self.auto_panel_var.set(self.state.auto_panel)
        self.overlay_opacity_var.set(_clamp_opacity(self.state.overlay_opacity))
        self.overlay.set_lock_key(self.state.overlay_lock_key)
        self.overlay.set_locked(False)
        self.manager.set_auto_panel(self.state.auto_panel, self.state.panel_key)
        self._sync_auto_panel_state()
        self._register_overlay_lock_hotkey()
        self._refresh_panel_display()
        self._update_all_buttons()
        if self.listening:
            self._rebuild_listeners()
        if show_messages:
            if missing:
                messagebox.showwarning(
                    "Profile loaded with missing macros",
                    f"Loaded {path}, but these macros were not found: {', '.join(missing)}.",
                )
            else:
                messagebox.showinfo("Profile loaded", f"Loaded {path}.")
        self.saved_state = self._serialize_state()
        self._record_last_profile(path)
        return True

    def _load_profile_action(self) -> None:
        saves_dir = ensure_saves_dir()
        file_path = filedialog.askopenfilename(
            title="Load profile",
            initialdir=saves_dir,
            filetypes=[("Profile files", "*.json"), ("All files", "*.*")],
        )
        if not file_path:
            return
        self._load_profile_from_path(Path(file_path), show_messages=True)

    def _load_blank_profile(self, show_messages: bool = True) -> None:
        self.state.reset()
        self.auto_panel_var.set(DEFAULT_AUTO_PANEL)
        self.overlay_opacity_var.set(DEFAULT_OVERLAY_OPACITY)
        self.overlay.set_lock_key(self.state.overlay_lock_key)
        self.overlay.set_locked(False)
        self._register_overlay_lock_hotkey()
        self._refresh_panel_display()
        self._sync_auto_panel_state()
        self._update_all_buttons()
        if self.listening:
            self._rebuild_listeners()
        self.saved_state = self._serialize_state()
        try:
            if self.last_profile_marker.exists():
                self.last_profile_marker.unlink()
        except OSError:
            pass
        if show_messages:
            messagebox.showinfo("New profile", "Reset to the blank default profile.")

    # -- Settings + templates --------------------------------------------- #
    def _open_settings(self) -> None:
        settings = tk.Toplevel(self.root, bg=BG)
        settings.title("Settings - Key Binds")
        settings.resizable(True, True)
        place_window_near(settings, self.root)
        set_dark_titlebar(settings)

        status_local = tk.StringVar(value="Select a slot and press a key to rebind.")
        change_buttons: dict[str, tk.Button] = {}
        hotkey_labels: dict[str, tk.Label] = {}
        dir_buttons: dict[str, tk.Button] = {}
        dir_labels: dict[str, tk.Label] = {}
        panel_change_btn: tk.Button | None = None
        overlay_lock_btn: tk.Button | None = None
        panel_key_label_var = tk.StringVar(value=_display_hotkey_text(self.state.panel_key, self.state.panel_key))
        overlay_lock_label_var = tk.StringVar(
            value=_display_hotkey_text(self.state.overlay_lock_key, self.state.overlay_lock_key)
        )
        overlay_opacity_label_var = tk.StringVar(value=f"{int(_clamp_opacity(self.overlay_opacity_var.get()) * 100)}%")
        capturing = {"active": False}
        pending_hotkeys: dict[str, str] = dict(self.state.slot_hotkeys)
        pending_direction_keys: dict[str, str] = dict(self.state.direction_keys)
        pending_panel_key = self.state.panel_key
        pending_overlay_lock_key = self.state.overlay_lock_key
        pending_overlay_opacity = tk.DoubleVar(value=_clamp_opacity(self.overlay_opacity_var.get()))
        active_values = {"delay": self.state.macro_delay, "duration": self.state.macro_duration}

        tabs_frame = tk.Frame(settings, bg=BG)
        tabs_frame.pack(fill=tk.X, padx=10, pady=(10, 0))

        content_frame = tk.Frame(settings, bg=BG)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        slot_section = tk.Frame(content_frame, bg=BG)
        dir_section = tk.Frame(content_frame, bg=BG)
        delay_section = tk.Frame(content_frame, bg=BG)
        panel_section = tk.Frame(content_frame, bg=BG)
        overlay_section = tk.Frame(content_frame, bg=BG)

        current_category = {"val": None}

        def is_dirty() -> bool:
            return (
                pending_hotkeys != self.state.slot_hotkeys
                or pending_direction_keys != self.state.direction_keys
                or active_values["delay"] != self.state.macro_delay
                or active_values["duration"] != self.state.macro_duration
                or pending_panel_key != self.state.panel_key
                or pending_overlay_lock_key != self.state.overlay_lock_key
                or abs(_clamp_opacity(pending_overlay_opacity.get()) - _clamp_opacity(self.overlay_opacity_var.get()))
                > 1e-4
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
            nonlocal pending_panel_key, pending_overlay_lock_key
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
                self._rebuild_listeners()

        def start_capture(target: str, kind: str) -> None:
            if capturing["active"]:
                return
            capturing["active"] = True
            was_listening = self.listening
            if was_listening:
                self.manager.clear()
            prompt = {
                "slot": f"Press a key to bind to numpad {target}...",
                "direction": f"Press a key to bind to {target} direction...",
                "panel": "Press a key to open the Stratagem Panel...",
                "overlay_lock": "Press a key to toggle the Overlay Lock...",
            }.get(kind, "Press a key...")
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
            nonlocal pending_panel_key, pending_overlay_lock_key
            self.state.slot_hotkeys.update(pending_hotkeys)
            self.state.direction_keys.update(pending_direction_keys)
            self.state.macro_delay = active_values["delay"]
            self.state.macro_duration = active_values["duration"]
            self.state.panel_key = pending_panel_key
            self.state.overlay_lock_key = pending_overlay_lock_key
            self.state.overlay_opacity = _clamp_opacity(pending_overlay_opacity.get())
            self.auto_panel_var.set(self.state.auto_panel)
            self.overlay_opacity_var.set(self.state.overlay_opacity)
            self.overlay.set_lock_key(self.state.overlay_lock_key)
            self.overlay.set_locked(False)
            self._refresh_panel_display()
            self._sync_auto_panel_state()
            self._register_overlay_lock_hotkey()
            self._update_all_buttons()
            if self.listening:
                self._rebuild_listeners()
            status_local.set("Applied bindings and overlay settings.")

        def reset_pending_from_live() -> None:
            nonlocal pending_panel_key, pending_overlay_lock_key
            pending_hotkeys.clear()
            pending_hotkeys.update(self.state.slot_hotkeys)
            pending_direction_keys.clear()
            pending_direction_keys.update(self.state.direction_keys)
            pending_panel_key = self.state.panel_key
            pending_overlay_lock_key = self.state.overlay_lock_key
            pending_overlay_opacity.set(_clamp_opacity(self.overlay_opacity_var.get()))
            active_values["delay"] = self.state.macro_delay
            active_values["duration"] = self.state.macro_duration
            refresh_labels()

        def switch_category(target: str) -> None:
            if target == current_category["val"]:
                return
            if is_dirty():
                if messagebox.askyesno("Apply changes?", "Apply changes before switching categories?"):
                    apply_and_stay()
                else:
                    reset_pending_from_live()
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

        for slot, _ in NUMPAD_SLOTS:
            row = tk.Frame(slot_section, bg=BG)
            row.pack(fill=tk.X, pady=2)
            tk.Label(row, text=f"Numpad {slot}", width=12, anchor="w").pack(side=tk.LEFT)
            hotkey_labels[slot] = tk.Label(
                row, text=_display_hotkey_text(pending_hotkeys.get(slot, ""), slot), width=12, anchor="w"
            )
            hotkey_labels[slot].pack(side=tk.LEFT, padx=(0, 6))
            btn = tk.Button(row, text="Change", command=lambda s=slot: start_capture(s, "slot"))
            btn.pack(side=tk.LEFT)
            change_buttons[slot] = btn

        for direction in ("Up", "Down", "Left", "Right"):
            row = tk.Frame(dir_section, bg=BG)
            row.pack(fill=tk.X, pady=2)
            tk.Label(row, text=direction, width=12, anchor="w").pack(side=tk.LEFT)
            dir_labels[direction] = tk.Label(row, text=pending_direction_keys.get(direction, ""), width=12, anchor="w")
            dir_labels[direction].pack(side=tk.LEFT, padx=(0, 6))
            btn = tk.Button(row, text="Change", command=lambda d=direction: start_capture(d, "direction"))
            btn.pack(side=tk.LEFT)
            dir_buttons[direction] = btn

        row = tk.Frame(panel_section, bg=BG)
        row.pack(fill=tk.X, pady=4)
        tk.Label(row, text="Stratagem Panel", width=16, anchor="w").pack(side=tk.LEFT)
        panel_label = tk.Label(row, textvariable=panel_key_label_var, width=12, anchor="w")
        panel_label.pack(side=tk.LEFT, padx=(0, 6))
        panel_change_btn = tk.Button(row, text="Change", command=lambda: start_capture("panel", "panel"))
        panel_change_btn.pack(side=tk.LEFT)

        row_overlay = tk.Frame(overlay_section, bg=BG)
        row_overlay.pack(fill=tk.X, pady=4)
        tk.Label(row_overlay, text="Overlay Lock Key", width=18, anchor="w").pack(side=tk.LEFT)
        tk.Label(row_overlay, textvariable=overlay_lock_label_var, width=12, anchor="w").pack(
            side=tk.LEFT, padx=(0, 6)
        )
        overlay_lock_btn = tk.Button(
            row_overlay, text="Change", command=lambda: start_capture("overlay_lock", "overlay_lock")
        )
        overlay_lock_btn.pack(side=tk.LEFT)

        tk.Label(overlay_section, text="Overlay Opacity", anchor="w").pack(fill=tk.X, pady=(12, 4))
        op_row = tk.Frame(overlay_section, bg=BG)
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

        tk.Label(delay_section, text="Milliseconds between key presses:", anchor="w").pack(
            fill=tk.X, pady=(4, 4)
        )
        delay_var = tk.StringVar(value=str(int(self.state.macro_delay * 1000)))
        delay_entry = tk.Entry(delay_section, textvariable=delay_var, width=10)
        delay_entry.pack(anchor="w", pady=(0, 4))

        tk.Label(delay_section, text="Milliseconds key stays pressed:", anchor="w").pack(
            fill=tk.X, pady=(8, 4)
        )
        duration_var = tk.StringVar(value=str(int(self.state.macro_duration * 1000)))
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

        btn_frame = tk.Frame(settings, bg=BG)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        tk.Button(btn_frame, text="Apply", command=apply_and_stay).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Close", command=settings.destroy).pack(side=tk.RIGHT)

        apply_dark_theme(settings)
        refresh_labels()

    def _open_edit_templates(self) -> None:
        edit = tk.Toplevel(self.root, bg=BG)
        edit.title("Edit Stratagem Templates")
        edit.resizable(True, True)
        place_window_near(edit, self.root)
        set_dark_titlebar(edit)

        working: list[MacroTemplate] = list(self.templates)
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
            self.templates = tuple(working)
            name_map = {tpl.name: tpl for tpl in self.templates}
            for slot, tpl in list(self.state.assignments.items()):
                if tpl is None:
                    continue
                self.state.assignments[slot] = name_map.get(tpl.name, tpl)
            self._update_all_buttons()
            if self.listening:
                self._rebuild_listeners()
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

    # -- Misc dialogs ----------------------------------------------------- #
    def _show_about(self) -> None:
        about = tk.Toplevel(self.root, bg=BG)
        about.title("About")
        about.transient(self.root)
        about.resizable(False, False)
        place_window_near(about, self.root)

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

    # -- Overlay helpers -------------------------------------------------- #
    def _on_root_configure(self) -> None:
        if self.overlay.is_visible():
            self.overlay.resize_to_parent()

    # -- Menu helpers ----------------------------------------------------- #
    def _popup_menu(self, menu: tk.Menu, btn: tk.Button) -> None:
        try:
            x = btn.winfo_rootx()
            y = btn.winfo_rooty() + btn.winfo_height()
            menu.tk_popup(x, y)
        finally:
            try:
                menu.grab_release()
            except tk.TclError:
                pass

    def _refresh_menu_bar_colors(self) -> None:
        self.menu_bar.configure(bg=MENU_BG)
        for child in self.menu_bar.winfo_children():
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

    # -- Hotkeys & exit --------------------------------------------------- #
    def _register_overlay_lock_hotkey(self) -> None:
        if self.overlay_lock_handle is not None:
            try:
                keyboard.remove_hotkey(self.overlay_lock_handle)
            except Exception:
                pass
            self.overlay_lock_handle = None
        key = (self.state.overlay_lock_key or "").strip()
        if not key:
            return
        try:
            self.overlay_lock_handle = keyboard.add_hotkey(
                key, lambda: self.root.after(0, self.overlay.toggle_lock), suppress=False
            )
        except Exception as exc:  # noqa: BLE001
            log(f"Could not register overlay lock hotkey '{key}': {exc}")

    def _register_exit_hotkey(self) -> None:
        self.exit_handle = keyboard.add_hotkey(EXIT_HOTKEY, lambda: self.root.after(0, self._attempt_exit))

    # -- Persistence helpers --------------------------------------------- #
    def _record_last_profile(self, path: Path) -> None:
        try:
            self.last_profile_marker.write_text(str(path), encoding="utf-8")
        except OSError:
            pass

    def _load_last_profile(self) -> None:
        if self.last_profile_marker.exists():
            try:
                last_path = Path(self.last_profile_marker.read_text(encoding="utf-8").strip())
                if last_path.exists():
                    if self._load_profile_from_path(last_path, show_messages=False):
                        log(f"Loaded last profile: {last_path.name}")
                else:
                    log(f"Last profile not found; expected at {last_path}")
            except OSError:
                pass

    # -- Exit ------------------------------------------------------------- #
    def _maybe_save_before_exit(self) -> bool:
        if not self._has_unsaved_changes():
            return True
        resp = messagebox.askyesnocancel(
            "Save changes?",
            "Save your profile before exiting?",
            icon=messagebox.QUESTION,
        )
        if resp is None:
            return False
        if resp is False:
            return True
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
        if self._save_profile_to_path(path, show_message=False):
            self.saved_state = self._serialize_state()
            messagebox.showinfo("Profile saved", f"Saved to {path.name}.")
            return True
        return False

    def _close_app(self) -> None:
        if self.exit_handle is not None:
            keyboard.remove_hotkey(self.exit_handle)
            self.exit_handle = None
        if self.overlay_lock_handle is not None:
            try:
                keyboard.remove_hotkey(self.overlay_lock_handle)
            except Exception:
                pass
            self.overlay_lock_handle = None
        self.manager.shutdown()
        clear_log_callback()
        self.overlay.hide()
        self.root.destroy()

    def _attempt_exit(self) -> None:
        if self._maybe_save_before_exit():
            self._close_app()

    # -- Run -------------------------------------------------------------- #
    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = MacroApp()
    app.run()


if __name__ == "__main__":
    main()

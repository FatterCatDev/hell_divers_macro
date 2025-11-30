from __future__ import annotations

"""Overlay window that mirrors slot assignments and macro progress."""

import time
import tkinter as tk
from typing import Callable, Dict

from hell_divers_macro.ui.icons import (
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
    FG,
    apply_dark_theme,
    make_window_clickthrough,
    place_window_near,
    set_dark_titlebar,
)


class OverlayWindow:
    """Floating, semi-transparent overlay with slots and progress bars."""

    def __init__(
        self,
        root: tk.Tk,
        *,
        auto_panel_var: tk.BooleanVar,
        opacity_var: tk.DoubleVar,
        initial_lock_key: str,
        status_callback: Callable[[str], None],
        hotkey_display: Callable[[str, str], str],
        on_close: Callable[[], None],
    ) -> None:
        self.root = root
        self.auto_panel_var = auto_panel_var
        self.opacity_var = opacity_var
        self.lock_key = initial_lock_key
        self.status_callback = status_callback
        self.hotkey_display = hotkey_display
        self.on_close = on_close

        self.win: tk.Toplevel | None = None
        self.slot_canvases: dict[str, tk.Canvas] = {}
        self.fill_rects: dict[str, int] = {}
        self.icons: dict[str, tk.PhotoImage | None] = {}
        self.progress: dict[str, dict[str, object]] = {}
        self.user_resized = False
        self.applying_config = False
        self.locked = False
        self.lock_display = tk.StringVar(value="")
        self.resize_handle: tk.Canvas | None = None
        self.auto_panel_check: tk.Checkbutton | None = None
        self.close_btn: tk.Button | None = None
        self.lock_label: tk.Label | None = None
        self.drag_state = {"x": 0, "y": 0}
        self.dragging = False
        self.resize_state = {"x": 0, "y": 0, "w": 0, "h": 0}
        self.resizing = False

    # -- Public API -------------------------------------------------------- #
    def set_lock_key(self, key: str) -> None:
        self.lock_key = key
        self._update_lock_display()

    def set_locked(self, locked: bool) -> None:
        self.locked = locked
        self._apply_window_config()
        self._update_header_visibility()
        self._update_lock_display()

    def toggle_lock(self) -> None:
        self.set_locked(not self.locked)
        state = "locked" if self.locked else "unlocked (drag to move)"
        self.status_callback(f"Overlay {state}.")

    def show(self, assignments: Dict[str, object], slot_hotkeys: Dict[str, str]) -> None:
        self._ensure_window()
        if not self.user_resized:
            self._size_window(force=True)
        self._update_lock_display()
        self.refresh_slots(assignments, slot_hotkeys)
        if self.win is not None:
            self._apply_window_config()
            self.win.deiconify()
            self.win.lift()
            try:
                self.win.attributes("-topmost", True)
            except tk.TclError:
                pass

    def hide(self) -> None:
        if self.win is None or not self.win.winfo_exists():
            return
        self.win.withdraw()

    def is_visible(self) -> bool:
        return self.win is not None and self.win.winfo_exists() and self.win.state() != "withdrawn"

    def refresh_slots(self, assignments: Dict[str, object], slot_hotkeys: Dict[str, str]) -> None:
        if self.win is None or not self.win.winfo_exists():
            return
        for slot in assignments:
            tpl = assignments.get(slot)
            hotkey = slot_hotkeys.get(slot, slot)
            self.update_slot(slot, tpl, hotkey)

    def update_slot(self, slot: str, tpl, hotkey: str) -> None:
        if self.win is None or not self.win.winfo_exists():
            return
        canvas = self.slot_canvases.get(slot)
        if canvas is None:
            return
        hotkey_text = self.hotkey_display(hotkey, slot)
        icon: tk.PhotoImage | None
        if tpl:
            icon = load_icon_image(tpl.name, hotkey_text, variant="badge", size=OVERLAY_ICON_SIZE)
        else:
            icon = build_overlay_placeholder(hotkey_text, size=OVERLAY_ICON_SIZE)
        self.icons[slot] = icon
        canvas.delete("icon")
        canvas.create_image(
            OVERLAY_ICON_SIZE[0] // 2,
            OVERLAY_ICON_SIZE[1] // 2,
            image=icon,
            tags="icon",
        )
        self._set_fill(slot, 0)

    def start_progress(self, slot: str, total_time: float | None) -> None:
        if self.win is None or not self.win.winfo_exists():
            return
        duration = max(total_time or 0.05, 0.05)
        self._cancel_progress(slot)
        self.progress[slot] = {"start": time.time(), "duration": duration, "job": None}
        self._tick_progress(slot)

    def stop_progress(self, slot: str) -> None:
        self._cancel_progress(slot)
        self._set_fill(slot, 0)

    def resize_to_parent(self, force: bool = False) -> None:
        """Match the overlay size to the root window (unless the user resized)."""
        self._size_window(force)

    def handle_macro_progress(self, event: str, macro, slot: str | None, total_time: float | None) -> None:
        if not slot:
            return
        if event == "start":
            self.root.after(0, lambda s=slot, t=total_time: self.start_progress(s, t))
        elif event == "stop":
            self.root.after(0, lambda s=slot: self.stop_progress(s))

    # -- Window construction ---------------------------------------------- #
    def _ensure_window(self) -> None:
        if self.win is not None and self.win.winfo_exists():
            return
        self.slot_canvases.clear()
        self.fill_rects.clear()
        self.icons.clear()
        self.progress.clear()
        self.win = tk.Toplevel(self.root, bg=BG)
        self.win.withdraw()
        self.win.overrideredirect(True)
        self.win.title("Listening Overlay")
        self.win.resizable(True, True)
        try:
            self.win.attributes("-topmost", True)
        except tk.TclError:
            pass
        self._apply_window_config()
        self.win.protocol("WM_DELETE_WINDOW", self.on_close)
        self.win.bind("<ButtonPress-1>", self._start_drag)
        self.win.bind("<B1-Motion>", self._drag)
        self.win.bind("<ButtonRelease-1>", self._stop_resize)

        container = tk.Frame(self.win, bg=BG)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

        auto_frame = tk.Frame(container, bg=BG)
        auto_frame.pack(fill=tk.X, pady=(0, 8))
        self.auto_panel_check = tk.Checkbutton(
            auto_frame,
            text="Auto Stratagem Panel",
            variable=self.auto_panel_var,
            bg=BG,
            fg=FG,
            activebackground=BUTTON_BG,
            activeforeground=FG,
            selectcolor=BG,
            highlightthickness=0,
            anchor="w",
        )
        self.auto_panel_check.pack(side=tk.LEFT)
        self.close_btn = tk.Button(
            auto_frame,
            text="X",
            command=self.on_close,
            width=2,
            bg=BUTTON_BG,
            fg=FG,
            activebackground=BUTTON_ACTIVE,
            activeforeground=FG,
            bd=0,
            highlightthickness=0,
        )
        self.close_btn.pack(side=tk.RIGHT, padx=(6, 0))
        self.lock_label = tk.Label(auto_frame, textvariable=self.lock_display, anchor="e")
        self.lock_label.pack(side=tk.RIGHT, padx=(0, 6), anchor="e")

        grid_layout = [["7", "8", "9"], ["4", "5", "6"], ["1", "2", "3"]]
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
                self.slot_canvases[slot] = canvas
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
                self.fill_rects[slot] = rect_id
                grid.grid_columnconfigure(c, weight=1, uniform="overlay_slots")
            grid.grid_rowconfigure(r, weight=1, uniform="overlay_slots")

        apply_dark_theme(self.win)
        set_dark_titlebar(self.win)
        self.user_resized = False
        self.win.bind(
            "<Configure>",
            lambda event=None: self._mark_user_resized() if self.win and self.win.state() != "withdrawn" else None,
        )
        self._size_window()

        self.resize_handle = tk.Canvas(container, width=16, height=16, bg=BG, bd=0, highlightthickness=0)
        self.resize_handle.place(relx=1.0, rely=1.0, anchor="se")
        self.resize_handle.create_polygon(0, 16, 16, 16, 16, 0, fill=ACCENT, outline="")
        self.resize_handle.bind("<ButtonPress-1>", self._start_resize)
        self.resize_handle.bind("<B1-Motion>", self._resize)
        self.resize_handle.bind("<ButtonRelease-1>", self._stop_resize)
        self._update_header_visibility()

    # -- Layout helpers ---------------------------------------------------- #
    def _overlay_min_sizes(self) -> tuple[int, int]:
        return (OVERLAY_ICON_SIZE[0] * 3 + 40, OVERLAY_ICON_SIZE[1] * 3 + 90)

    def _size_window(self, force: bool = False) -> None:
        if self.win is None or not self.win.winfo_exists():
            return
        if self.user_resized and not force:
            return
        self.root.update_idletasks()
        self.win.update_idletasks()
        min_w, min_h = self._overlay_min_sizes()
        width = max(int(self.root.winfo_width() * 0.55), min_w)
        height = max(int(self.root.winfo_height() * 0.55), min_h)
        self.win.geometry(f"{width}x{height}")
        place_window_near(self.win, self.root)

    def _apply_window_config(self) -> None:
        if self.win is None or not self.win.winfo_exists():
            return
        if self.applying_config:
            return
        self.applying_config = True
        try:
            opacity = self._clamp_opacity(self.opacity_var.get())
            self.opacity_var.set(opacity)
            try:
                self.win.attributes("-topmost", True)
                self.win.attributes("-alpha", opacity)
            except tk.TclError:
                pass
            make_window_clickthrough(self.win, alpha=opacity, clickthrough=self.locked)
            try:
                self.win.attributes("-disabled", self.locked)
                self.win.configure(cursor="" if self.locked else "fleur")
            except tk.TclError:
                pass
            if self.resize_handle is not None:
                try:
                    self.resize_handle.configure(cursor="size_nw_se" if not self.locked else "")
                except tk.TclError:
                    pass
            self._update_lock_display()
            self._update_header_visibility()
        finally:
            self.applying_config = False

    def _update_header_visibility(self) -> None:
        for widget in (self.auto_panel_check, self.close_btn, self.lock_label):
            self._hide_widget(widget)
        if self.locked:
            self._show_widget(self.lock_label, side=tk.RIGHT, padx=(0, 6), anchor="e")
        else:
            self._show_widget(self.close_btn, side=tk.RIGHT, padx=(6, 0))
            self._show_widget(self.lock_label, side=tk.RIGHT, padx=(0, 6), anchor="e")
            self._show_widget(self.auto_panel_check, side=tk.LEFT, anchor="w")

    # -- Progress helpers -------------------------------------------------- #
    def _set_fill(self, slot: str, progress: float) -> None:
        canvas = self.slot_canvases.get(slot)
        rect_id = self.fill_rects.get(slot)
        if canvas is None or rect_id is None:
            return
        progress = max(0.0, min(1.0, progress))
        height = int(OVERLAY_ICON_SIZE[1] * progress)
        canvas.coords(rect_id, 0, 0, OVERLAY_ICON_SIZE[0], height)
        canvas.itemconfigure(rect_id, state=tk.NORMAL if height > 0 else tk.HIDDEN)

    def _cancel_progress(self, slot: str) -> None:
        data = self.progress.pop(slot, None)
        if data is None:
            return
        job = data.get("job")
        if job is not None:
            try:
                self.root.after_cancel(job)
            except Exception:
                pass

    def _tick_progress(self, slot: str) -> None:
        data = self.progress.get(slot)
        if data is None:
            return
        duration = float(data["duration"])
        start_ts = float(data["start"])
        elapsed = max(0.0, time.time() - start_ts)
        progress = min(1.0, elapsed / duration) if duration > 0 else 1.0
        self._set_fill(slot, progress)
        if progress >= 1.0:
            self._cancel_progress(slot)
            self._set_fill(slot, 0)
            return
        data["job"] = self.root.after(50, lambda s=slot: self._tick_progress(s))

    # -- Events ------------------------------------------------------------ #
    def _overlay_event_target(self, event) -> tk.Widget | None:  # noqa: ANN001
        if self.win is None:
            return None
        try:
            return self.win.winfo_containing(event.x_root, event.y_root)
        except Exception:
            return None

    def _start_drag(self, event) -> None:  # noqa: ANN001
        self.dragging = False
        if self.locked or self.resizing or self.win is None or not self.win.winfo_exists():
            return
        target = self._overlay_event_target(event)
        if target is self.resize_handle or self._is_interactive_widget(target):
            return
        self.dragging = True
        self.drag_state["x"] = event.x_root
        self.drag_state["y"] = event.y_root

    def _drag(self, event) -> None:  # noqa: ANN001
        if self.locked or self.resizing or self.win is None or not self.win.winfo_exists() or not self.dragging:
            return
        dx = event.x_root - self.drag_state["x"]
        dy = event.y_root - self.drag_state["y"]
        self.drag_state["x"] = event.x_root
        self.drag_state["y"] = event.y_root
        try:
            new_x = self.win.winfo_x() + dx
            new_y = self.win.winfo_y() + dy
            self.win.geometry(f"+{new_x}+{new_y}")
            self.user_resized = True
        except tk.TclError:
            pass

    def _start_resize(self, event) -> None:  # noqa: ANN001
        if self.locked or self.win is None or not self.win.winfo_exists():
            return
        self.resizing = True
        self.resize_state["x"] = event.x_root
        self.resize_state["y"] = event.y_root
        self.resize_state["w"] = self.win.winfo_width()
        self.resize_state["h"] = self.win.winfo_height()

    def _resize(self, event) -> None:  # noqa: ANN001
        if self.locked or self.win is None or not self.win.winfo_exists():
            return
        dx = event.x_root - self.resize_state["x"]
        dy = event.y_root - self.resize_state["y"]
        new_w = max(self._overlay_min_sizes()[0], self.resize_state["w"] + dx)
        new_h = max(self._overlay_min_sizes()[1], self.resize_state["h"] + dy)
        try:
            self.win.geometry(f"{new_w}x{new_h}+{self.win.winfo_x()}+{self.win.winfo_y()}")
            self.user_resized = True
        except tk.TclError:
            pass

    def _stop_resize(self, event=None) -> None:  # noqa: ANN001
        self.dragging = False
        self.resizing = False

    # -- Utility ----------------------------------------------------------- #
    def _update_lock_display(self) -> None:
        key_text = self.hotkey_display(self.lock_key or "Unset", self.lock_key or "Unset")
        state = "Locked" if self.locked else "Unlocked"
        self.lock_display.set(f"Overlay: {state} (Key: {key_text})")

    def _hide_widget(self, widget: tk.Widget | None) -> None:
        if widget is None:
            return
        try:
            if widget.winfo_manager():
                widget.pack_forget()
        except tk.TclError:
            pass

    def _show_widget(self, widget: tk.Widget | None, **pack_kwargs) -> None:
        if widget is None:
            return
        try:
            if widget.winfo_manager() != "pack":
                widget.pack(**pack_kwargs)
        except tk.TclError:
            pass

    def _is_interactive_widget(self, widget: tk.Widget | None) -> bool:
        if widget is None:
            return False
        return isinstance(widget, (tk.Button, tk.Checkbutton, tk.Entry, tk.Scale, tk.Listbox, tk.Text))

    def _mark_user_resized(self) -> None:
        self.user_resized = True

    def _clamp_opacity(self, val: float) -> float:
        try:
            return max(0.1, min(1.0, float(val)))
        except Exception:
            return OVERLAY_ALPHA

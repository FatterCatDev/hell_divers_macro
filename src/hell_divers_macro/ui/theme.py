from __future__ import annotations

"""Dark theme utilities and window helpers for the Tk UI."""

import platform
import sys
import tkinter as tk
from typing import Iterable

BG = "#121212"
FG = "#e5e5e5"
BUTTON_BG = "#1f1f1f"
BUTTON_ACTIVE = "#2d2d2d"
ENTRY_BG = "#1a1a1a"
ACCENT = "#4e8cff"
MENU_BG = "#161616"
IS_WINDOWS = platform.system() == "Windows"


def apply_dark_theme(widget: tk.Misc) -> None:
    """Recursively apply a dark palette to a widget tree."""
    cls = widget.winfo_class()
    try:
        widget.configure(bg=BG)
    except tk.TclError:
        pass

    try:
        widget.configure(fg=FG)
    except tk.TclError:
        pass

    if cls == "Button":
        try:
            widget.configure(
                bg=BUTTON_BG,
                fg=FG,
                activebackground=BUTTON_ACTIVE,
                activeforeground=FG,
                highlightthickness=0,
                bd=1,
            )
        except tk.TclError:
            pass
    elif cls == "Label":
        try:
            widget.configure(bg=BG, fg=FG)
        except tk.TclError:
            pass
    elif cls in ("Frame", "TFrame"):
        try:
            widget.configure(bg=BG)
        except tk.TclError:
            pass
    elif cls == "Entry":
        try:
            widget.configure(
                bg=ENTRY_BG,
                fg=FG,
                insertbackground=FG,
                disabledforeground="#777777",
            )
        except tk.TclError:
            pass
    elif cls == "Listbox":
        try:
            widget.configure(
                bg=ENTRY_BG,
                fg=FG,
                selectbackground=ACCENT,
                selectforeground=FG,
                highlightthickness=0,
                relief=tk.FLAT,
            )
        except tk.TclError:
            pass
    elif cls == "Scrollbar":
        try:
            widget.configure(bg=BG, troughcolor=BUTTON_BG, activebackground=BUTTON_ACTIVE, highlightthickness=0)
        except tk.TclError:
            pass

    for child in widget.winfo_children():
        apply_dark_theme(child)


def init_base_theme(root: tk.Tk) -> None:
    """Set base palette defaults for new widgets."""
    root.configure(bg=BG)
    root.option_add("*Background", BG)
    root.option_add("*Foreground", FG)
    root.option_add("*Button.Background", BUTTON_BG)
    root.option_add("*Button.Foreground", FG)
    root.option_add("*Entry.Background", ENTRY_BG)
    root.option_add("*Entry.Foreground", FG)
    root.option_add("*Entry.InsertBackground", FG)
    root.option_add("*Listbox.Background", ENTRY_BG)
    root.option_add("*Listbox.Foreground", FG)
    root.option_add("*Menu.Background", MENU_BG)
    root.option_add("*Menu.Foreground", FG)
    root.option_add("*Menu.activeBackground", BUTTON_ACTIVE)
    root.option_add("*Menu.activeForeground", FG)


def place_window_near(child: tk.Toplevel, parent: tk.Tk) -> None:
    """Position child centered over the parent window."""
    child.update_idletasks()
    try:
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        cw = child.winfo_width()
        ch = child.winfo_height()
        x = px + max(0, (pw - cw) // 2)
        y = py + max(0, (ph - ch) // 2)
        child.geometry(f"+{x}+{y}")
        child.lift()
    except tk.TclError:
        # Fall back silently if positioning fails (e.g., parent not mapped yet).
        pass


def _hex_to_colorref(hex_color: str) -> int:
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    return (b << 16) | (g << 8) | r  # COLORREF is 0x00bbggrr


def set_dark_titlebar(win: tk.Tk | tk.Toplevel) -> None:
    """On Windows, request a dark title bar; no-op elsewhere."""
    if not IS_WINDOWS:
        return
    try:
        import ctypes
        from ctypes import wintypes

        win.update_idletasks()
        hwnd = wintypes.HWND(win.winfo_id())
        user32 = ctypes.windll.user32
        GA_ROOT = 2  # GetAncestor flag for root window
        root_hwnd = user32.GetAncestor(hwnd, GA_ROOT)
        if root_hwnd:
            hwnd = wintypes.HWND(root_hwnd)

        # Ask the OS to allow dark chrome even if system setting is light.
        try:
            uxtheme = ctypes.windll.uxtheme

            def _call(func_name: str, *args) -> bool:
                func = getattr(uxtheme, func_name, None)
                if func is None:
                    return False
                func.restype = wintypes.BOOL
                func.argtypes = [type(arg) for arg in args]
                try:
                    func(*args)
                    return True
                except Exception:
                    return False

            # 0=Default, 1=AllowDark, 2=ForceDark (depends on build)
            _call("SetPreferredAppMode", ctypes.c_int(2))
            _call("AllowDarkModeForApp", wintypes.BOOL(True))
            _call("AllowDarkModeForWindow", hwnd, wintypes.BOOL(True))
            _call("RefreshImmersiveColorPolicyState")
            _call("FlushMenuThemes")
        except Exception:
            pass

        def _set_attr(attr: int, val: int) -> None:
            value = wintypes.BOOL(val) if isinstance(val, bool) or val in (0, 1) else ctypes.c_int(val)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, ctypes.c_uint(attr), ctypes.byref(value), ctypes.sizeof(value)
            )

        def _set_color_attr(attr: int, hex_color: str) -> None:
            color = wintypes.DWORD(_hex_to_colorref(hex_color))
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, ctypes.c_uint(attr), ctypes.byref(color), ctypes.sizeof(color)
            )

        # Windows 10/11 dark mode attribute (19 for 1809, 20 for 1903+).
        build = sys.getwindowsversion().build
        attr_order = (20, 19) if build >= 18362 else (19, 20)
        for attr in attr_order:
            _set_attr(attr, 1)

        # Disable bright backdrops (e.g., Mica) and force our colors on Win11+.
        try:
            _set_attr(38, 0)  # DWMWA_SYSTEMBACKDROP_TYPE = None
        except Exception:
            pass

        # Darken the frame/title elements so borders aren't bright on Win11.
        _set_color_attr(34, BUTTON_BG)  # DWMWA_BORDER_COLOR
        _set_color_attr(35, BG)  # DWMWA_CAPTION_COLOR
        _set_color_attr(36, FG)  # DWMWA_TEXT_COLOR
    except Exception:
        pass


def make_window_clickthrough(
    win: tk.Tk | tk.Toplevel, alpha: float | None = None, clickthrough: bool = True
) -> None:
    """On Windows, toggle mouse/focus passthrough for a window."""
    if not IS_WINDOWS:
        return
    try:
        import ctypes
        from ctypes import wintypes

        hwnd = wintypes.HWND(win.winfo_id())
        user32 = ctypes.windll.user32
        GWL_EXSTYLE = -20
        WS_EX_LAYERED = 0x00080000
        WS_EX_TRANSPARENT = 0x00000020
        WS_EX_TOOLWINDOW = 0x00000080
        WS_EX_NOACTIVATE = 0x08000000

        # Ensure we are setting styles on the real root window handle.
        GA_ROOT = 2
        root_hwnd = user32.GetAncestor(hwnd, GA_ROOT)
        if root_hwnd:
            hwnd = wintypes.HWND(root_hwnd)

        current_style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        new_style = current_style | WS_EX_LAYERED | WS_EX_TOOLWINDOW
        if clickthrough:
            new_style |= WS_EX_NOACTIVATE
            new_style |= WS_EX_TRANSPARENT
        else:
            new_style &= ~WS_EX_NOACTIVATE
            new_style &= ~WS_EX_TRANSPARENT
        user32.SetWindowLongW(hwnd, GWL_EXSTYLE, new_style)

        if alpha is not None:
            clamped = max(0.0, min(1.0, alpha))
            user32.SetLayeredWindowAttributes(hwnd, 0, int(clamped * 255), 0x2)

        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_NOZORDER = 0x0004
        SWP_FRAMECHANGED = 0x0020
        SWP_NOACTIVATE = 0x0010
        swp_flags = SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED
        if clickthrough:
            swp_flags |= SWP_NOACTIVATE
        user32.SetWindowPos(
            hwnd,
            None,
            0,
            0,
            0,
            0,
            swp_flags,
        )
    except Exception:
        # If any ctypes call fails, leave the window unchanged.
        pass

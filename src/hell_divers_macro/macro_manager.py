from __future__ import annotations

"""Keyboard macro registration and execution utilities."""

import threading
import time
from typing import Callable, Dict

import keyboard

from hell_divers_macro.config import DEFAULT_AUTO_PANEL, DEFAULT_PANEL_KEY
from hell_divers_macro.log_utils import log
from hell_divers_macro.models import Macro, MacroRecord

MacroProgressCallback = Callable[[str, Macro, str | None, float | None], None]


class MacroManager:
    """Manage keyboard listeners for macro hotkeys."""

    def __init__(
        self,
        progress_callback: MacroProgressCallback | None = None,
        *,
        auto_panel_key: str = DEFAULT_PANEL_KEY,
        auto_panel_enabled: bool = DEFAULT_AUTO_PANEL,
    ) -> None:
        self.records: list[MacroRecord] = []
        self._held_scancodes: set[int] = set()
        self._progress_callback = progress_callback
        self._slot_hotkey_lookup: dict[str, str] = {}
        self.auto_panel_key = auto_panel_key
        self.auto_panel_enabled = auto_panel_enabled
        self._lock = threading.Lock()

    # -- Public config ----------------------------------------------------- #
    def set_auto_panel(self, enabled: bool, key: str) -> None:
        self.auto_panel_enabled = enabled
        self.auto_panel_key = key

    def set_progress_callback(self, callback: MacroProgressCallback | None) -> None:
        self._progress_callback = callback

    # -- Macro lifecycle --------------------------------------------------- #
    def register_macros(self, macros_by_slot: Dict[str, Macro]) -> None:
        """Clear existing listeners and register a fresh mapping."""
        self.clear()
        self._slot_hotkey_lookup.clear()
        for slot, macro in macros_by_slot.items():
            hotkey = macro.hotkey.lower()
            if hotkey in self._slot_hotkey_lookup:
                log(f"Hotkey '{hotkey}' already in use; skipping {macro.name or slot}.")
                continue
            self._slot_hotkey_lookup[hotkey] = slot
            self._add_macro(macro)

    def clear(self) -> None:
        """Remove all listeners."""
        for record in self.records:
            press_hook, release_hook = record.handle
            keyboard.unhook(press_hook)
            keyboard.unhook(release_hook)
        self.records.clear()
        self._slot_hotkey_lookup.clear()

    def shutdown(self) -> None:
        """Alias for clear for symmetry with startup."""
        self.clear()

    # -- Internals --------------------------------------------------------- #
    def _add_macro(self, macro: Macro) -> None:
        handle = self._register_macro(macro)
        self.records.append(MacroRecord(macro, handle))

    def _register_macro(self, macro: Macro) -> int:
        def on_press(event) -> None:
            if event.event_type != "down":
                return
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
            self._launch_macro(macro)

        def on_release(event) -> None:
            self._held_scancodes.discard(event.scan_code)

        press_hook = keyboard.on_press_key(macro.hotkey, on_press, suppress=False)
        release_hook = keyboard.on_release_key(macro.hotkey, on_release, suppress=False)
        return (press_hook, release_hook)

    def _launch_macro(self, macro: Macro) -> None:
        label = macro.name or macro.hotkey
        log(f"Trigger received for hotkey '{macro.hotkey}' ({label}).")
        slot = self._slot_hotkey_lookup.get(macro.hotkey.lower())
        panel_key_arg = None
        if self.auto_panel_enabled:
            key = (self.auto_panel_key or "").strip()
            if key:
                panel_key_arg = str(key)

        total_len = len(macro.keys) + (1 if panel_key_arg else 0)
        total_time = total_len * (macro.duration + macro.delay)
        self._notify_progress("start", macro, slot, total_time)

        def _worker() -> None:
            try:
                self._run_macro(macro, panel_key_arg)
            finally:
                self._notify_progress("stop", macro, slot, None)

        threading.Thread(target=_worker, daemon=True).start()

    def _run_macro(self, macro: Macro, panel_key: str | None) -> None:
        with self._lock:
            label = macro.name or macro.hotkey
            sequence: Iterable[str] = (panel_key, *macro.keys) if panel_key else tuple(macro.keys)
            if panel_key:
                log(f"{label}: auto panel ON, prepending '{panel_key}'.")
            seq_tuple = tuple(sequence)
            log(f"{label}: running {len(seq_tuple)} key presses...")
            for key in seq_tuple:
                keyboard.press(key)
                time.sleep(macro.duration)
                keyboard.release(key)
                time.sleep(macro.delay)
            log(f"{label}: done.")

    def _notify_progress(self, event: str, macro: Macro, slot: str | None, total_time: float | None) -> None:
        cb = self._progress_callback
        if cb is None:
            return
        try:
            cb(event, macro, slot, total_time)
        except Exception:
            pass

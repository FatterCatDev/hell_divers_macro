from __future__ import annotations

"""Centralized application state and profile serialization."""

from dataclasses import dataclass, field
from typing import Dict, Tuple

from hell_divers_macro.config import (
    DEFAULT_AUTO_PANEL,
    DEFAULT_DELAY,
    DEFAULT_DIRECTION_KEYS,
    DEFAULT_DURATION,
    DEFAULT_OVERLAY_LOCK_KEY,
    DEFAULT_OVERLAY_OPACITY,
    DEFAULT_PANEL_KEY,
    DEFAULT_SLOT_HOTKEYS,
    NUMPAD_SLOTS,
)
from hell_divers_macro.models import MacroTemplate


@dataclass
class AppState:
    assignments: Dict[str, MacroTemplate | None] = field(
        default_factory=lambda: {slot: None for slot, _ in NUMPAD_SLOTS}
    )
    slot_hotkeys: Dict[str, str] = field(default_factory=lambda: dict(DEFAULT_SLOT_HOTKEYS))
    direction_keys: Dict[str, str] = field(default_factory=lambda: dict(DEFAULT_DIRECTION_KEYS))
    panel_key: str = DEFAULT_PANEL_KEY
    auto_panel: bool = DEFAULT_AUTO_PANEL
    overlay_lock_key: str = DEFAULT_OVERLAY_LOCK_KEY
    overlay_opacity: float = DEFAULT_OVERLAY_OPACITY
    macro_delay: float = DEFAULT_DELAY
    macro_duration: float = DEFAULT_DURATION

    def serialize(self) -> dict:
        return {
            "slots": {slot: (tpl.name if tpl else None) for slot, tpl in self.assignments.items()},
            "hotkeys": dict(self.slot_hotkeys),
            "direction_keys": dict(self.direction_keys),
            "timing": {"delay": self.macro_delay, "duration": self.macro_duration},
            "panel": {"key": self.panel_key, "auto": bool(self.auto_panel)},
            "overlay": {
                "lock_key": self.overlay_lock_key,
                "opacity": self.overlay_opacity,
            },
        }

    def reset(self) -> None:
        self.assignments = {slot: None for slot, _ in NUMPAD_SLOTS}
        self.slot_hotkeys = dict(DEFAULT_SLOT_HOTKEYS)
        self.direction_keys = dict(DEFAULT_DIRECTION_KEYS)
        self.panel_key = DEFAULT_PANEL_KEY
        self.auto_panel = DEFAULT_AUTO_PANEL
        self.overlay_lock_key = DEFAULT_OVERLAY_LOCK_KEY
        self.overlay_opacity = DEFAULT_OVERLAY_OPACITY
        self.macro_delay = DEFAULT_DELAY
        self.macro_duration = DEFAULT_DURATION

    def apply_profile(self, data: dict, templates: Tuple[MacroTemplate, ...]) -> list[str]:
        """Load state from a profile dict; returns missing macro names."""
        slots_data = data.get("slots", {})
        missing: list[str] = []
        for slot, _ in NUMPAD_SLOTS:
            name = slots_data.get(slot)
            if name is None:
                self.assignments[slot] = None
                continue
            tpl = next((t for t in templates if t.name == name), None)
            if tpl is None:
                missing.append(name)
                self.assignments[slot] = None
            else:
                self.assignments[slot] = tpl

        hotkeys_data = data.get("hotkeys", {})
        if isinstance(hotkeys_data, dict):
            for slot, _ in NUMPAD_SLOTS:
                hk = hotkeys_data.get(slot)
                if isinstance(hk, str) and hk.strip():
                    self.slot_hotkeys[slot] = hk.strip()
                else:
                    self.slot_hotkeys[slot] = DEFAULT_SLOT_HOTKEYS.get(slot, self.slot_hotkeys[slot])

        direction_data = data.get("direction_keys", {})
        if isinstance(direction_data, dict):
            for direction in DEFAULT_DIRECTION_KEYS:
                val = direction_data.get(direction)
                if isinstance(val, str) and val.strip():
                    self.direction_keys[direction] = val.strip()
                else:
                    self.direction_keys[direction] = DEFAULT_DIRECTION_KEYS[direction]

        timing_data = data.get("timing", {})
        if isinstance(timing_data, dict):
            delay_val = timing_data.get("delay", self.macro_delay)
            duration_val = timing_data.get("duration", self.macro_duration)
            try:
                delay_f = float(delay_val)
                duration_f = float(duration_val)
                if delay_f >= 0 and duration_f >= 0:
                    self.macro_delay = delay_f
                    self.macro_duration = duration_f
            except (TypeError, ValueError):
                pass

        panel_data = data.get("panel", {})
        if isinstance(panel_data, dict):
            key_val = panel_data.get("key")
            if isinstance(key_val, str) and key_val.strip():
                self.panel_key = key_val.strip().lower()
            else:
                self.panel_key = DEFAULT_PANEL_KEY
            auto_val = panel_data.get("auto")
            if isinstance(auto_val, bool):
                self.auto_panel = auto_val
            else:
                self.auto_panel = DEFAULT_AUTO_PANEL
        else:
            self.panel_key = DEFAULT_PANEL_KEY
            self.auto_panel = DEFAULT_AUTO_PANEL

        overlay_data = data.get("overlay", {})
        if isinstance(overlay_data, dict):
            lock_val = overlay_data.get("lock_key")
            if isinstance(lock_val, str) and lock_val.strip():
                self.overlay_lock_key = lock_val.strip()
            op_val = overlay_data.get("opacity", self.overlay_opacity)
            try:
                self.overlay_opacity = max(0.1, min(1.0, float(op_val)))
            except (TypeError, ValueError):
                self.overlay_opacity = DEFAULT_OVERLAY_OPACITY
        else:
            self.overlay_lock_key = DEFAULT_OVERLAY_LOCK_KEY
            self.overlay_opacity = DEFAULT_OVERLAY_OPACITY
        return missing

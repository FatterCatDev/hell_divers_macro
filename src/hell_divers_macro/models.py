from dataclasses import dataclass
from typing import Callable, Tuple

from .config import DEFAULT_DELAY, DEFAULT_DURATION

HotkeyHandle = Callable[[], None]

@dataclass(frozen=True)
class Macro:
    hotkey: str
    keys: Tuple[str, ...]
    delay: float = DEFAULT_DELAY
    duration: float = DEFAULT_DURATION
    name: str | None = None


@dataclass
class MacroRecord:
    macro: Macro
    handle: tuple[HotkeyHandle, HotkeyHandle]


@dataclass(frozen=True)
class MacroTemplate:
    name: str
    directions: Tuple[str, ...]
    delay: float = DEFAULT_DELAY
    category: str = "Misc"

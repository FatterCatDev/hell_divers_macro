from typing import Callable

_log_callback: Callable[[str], None] | None = None


def set_log_callback(callback: Callable[[str], None] | None) -> None:
    global _log_callback
    _log_callback = callback


def clear_log_callback() -> None:
    set_log_callback(None)


def log(message: str) -> None:
    if _log_callback:
        try:
            _log_callback(message)
            return
        except Exception:
            # Fall back to stdout if the UI logger fails.
            pass
    print(message)

"""Application-wide constants and default mappings."""

DEFAULT_DELAY = 0.05  # seconds between key presses
DEFAULT_DURATION = 0.05  # seconds a key stays pressed
EXIT_HOTKEY = "ctrl+shift+q"
SAVES_DIR_NAME = "saves"

# Describes the slot label and the keyboard hotkey used by the listener.
NUMPAD_SLOTS = (
    ("7", "num 7"),
    ("8", "num 8"),
    ("9", "num 9"),
    ("4", "num 4"),
    ("5", "num 5"),
    ("6", "num 6"),
    ("1", "num 1"),
    ("2", "num 2"),
    ("3", "num 3"),
)

DEFAULT_SLOT_HOTKEYS = {slot: key for slot, key in NUMPAD_SLOTS}
# Default arrow-key mapping used by stratagem direction sequences.
DEFAULT_DIRECTION_KEYS = {
    "Up": "up",
    "Down": "down",
    "Left": "left",
    "Right": "right",
}

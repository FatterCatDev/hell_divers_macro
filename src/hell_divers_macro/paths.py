import sys
from pathlib import Path

from .config import SAVES_DIR_NAME


def get_base_dir() -> Path:
    """Return the project root for locating saves and data files."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent.parent


def ensure_saves_dir() -> Path:
    path = get_base_dir() / SAVES_DIR_NAME
    path.mkdir(parents=True, exist_ok=True)
    return path


def stratagem_md_path() -> Path:
    return get_base_dir() / "data" / "helldivers2_stratagem_codes.md"

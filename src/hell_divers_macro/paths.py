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
    # When frozen, packaged data lives under the PyInstaller extraction dir (_MEIPASS).
    resource_base = getattr(sys, "_MEIPASS", None)
    if resource_base:
        candidate = Path(resource_base) / "data" / "helldivers2_stratagem_codes.md"
        if candidate.exists():
            return candidate
    return get_base_dir() / "data" / "helldivers2_stratagem_codes.md"

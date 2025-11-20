from typing import Dict, Tuple

from .config import DEFAULT_DELAY, DEFAULT_DIRECTION_KEYS, DEFAULT_DURATION
from .models import MacroTemplate
from .paths import stratagem_md_path


def load_stratagem_templates() -> Tuple[MacroTemplate, ...]:
    """Load stratagems from helldivers2_stratagem_codes.md or fall back to placeholders."""
    path = stratagem_md_path()
    templates: list[MacroTemplate] = []
    current_category = "Stratagems"
    if path.exists():
        lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
        for line in lines:
            line = line.strip()
            if line.startswith("##"):
                current_category = line.lstrip("#").strip()
                continue
            if not line.startswith("- **"):
                continue
            try:
                name_part, seq_part = line.split("**:", 1)
                name = name_part.replace("- **", "").strip()
                seq_text = seq_part.strip()
                if not name or not seq_text:
                    continue
                directions = [
                    part.strip().title()
                    for part in seq_text.split(",")
                    if part.strip()
                ]
                if directions:
                    templates.append(
                        MacroTemplate(name, tuple(directions), DEFAULT_DELAY, category=current_category)
                    )
            except ValueError:
                continue
    if not templates:
        templates = [
            MacroTemplate(f"macro_Place_Holder_{i}", (), DEFAULT_DELAY, category="Misc")
            for i in range(1, 13)
        ]
    return tuple(templates)


def resolve_template_keys(
    template: MacroTemplate, direction_keys: Dict[str, str]
) -> Tuple[str, ...]:
    keys: list[str] = []
    for direction in template.directions:
        mapped = direction_keys.get(direction) or DEFAULT_DIRECTION_KEYS.get(direction)
        if mapped:
            keys.append(mapped)
        else:
            keys.append(direction.lower())
    return tuple(keys)

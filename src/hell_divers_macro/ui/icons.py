from __future__ import annotations

"""Icon loading and placeholder helpers."""

import io
import re
import sys
from pathlib import Path
from typing import Dict, Tuple

from PIL import Image, ImageDraw, ImageFont, ImageTk

from hell_divers_macro.config import DEFAULT_OVERLAY_OPACITY

ICON_SIZE: Tuple[int, int] = (120, 110)
OVERLAY_ICON_SIZE: Tuple[int, int] = (96, 88)
OVERLAY_ALPHA = DEFAULT_OVERLAY_OPACITY


def _resolve_assets_dir() -> Path:
    """Locate assets, preferring packaged paths when frozen."""
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path.cwd()))
        for candidate in (base / "hell_divers_macro" / "assets", base / "assets"):
            if candidate.exists():
                return candidate
    return Path(__file__).resolve().parent.parent / "assets"


ASSETS_DIR = _resolve_assets_dir()
APP_ICON_PATH = ASSETS_DIR / "helldivers_2_macro_icon.png"

_cairosvg_mod = None
_cairosvg_error = None
_HAS_CAIRO = False

_svglib_mod = None
_svglib_error = None
_HAS_SVGLIB = False

_icon_cache: dict[tuple[str, str, str, tuple[int, int]], ImageTk.PhotoImage] = {}
_overlay_placeholder_cache: dict[tuple[str, tuple[int, int]], ImageTk.PhotoImage] = {}


def _normalize_name(name: str) -> str:
    """Normalize names for matching against asset filenames."""
    return re.sub(r"[^a-z0-9]", "", name.lower())


def _build_asset_map() -> dict[str, Path]:
    mapping: dict[str, Path] = {}
    if ASSETS_DIR.exists():
        for pattern in ("*.png", "*.svg"):
            for path in ASSETS_DIR.rglob(pattern):
                key = _normalize_name(path.stem)
                mapping.setdefault(key, path)
    return mapping


ASSET_MAP = _build_asset_map()


def _get_cairosvg():
    global _cairosvg_mod, _HAS_CAIRO, _cairosvg_error
    if _cairosvg_mod is not None:
        return _cairosvg_mod
    try:
        import cairosvg as _cs  # type: ignore[import-not-found]

        _cairosvg_mod = _cs
        _HAS_CAIRO = True
        _cairosvg_error = None
    except Exception as exc:
        _cairosvg_error = str(exc)
        _HAS_CAIRO = False
        print(f"cairosvg import failed: {exc}")
    return _cairosvg_mod


def _get_svglib():
    global _svglib_mod, _HAS_SVGLIB, _svglib_error
    if _svglib_mod is not None:
        return _svglib_mod
    try:
        from svglib.svglib import svg2rlg  # type: ignore[import-not-found]
        from reportlab.graphics import renderPM  # type: ignore[import-not-found]

        _svglib_mod = (svg2rlg, renderPM)
        _HAS_SVGLIB = True
        _svglib_error = None
    except Exception as exc:
        _svglib_error = str(exc)
        _HAS_SVGLIB = False
        print(f"svglib import failed: {exc}")
    return _svglib_mod


def _svg_to_png_bytes(svg_bytes: bytes, target_size: tuple[int, int]) -> bytes | None:
    cs = _get_cairosvg()
    if cs is not None:
        return cs.svg2png(bytestring=svg_bytes, output_width=target_size[0], output_height=target_size[1])
    svglib_mod = _get_svglib()
    if svglib_mod is not None:
        svg2rlg, renderPM = svglib_mod
        try:
            drawing = svg2rlg(io.BytesIO(svg_bytes))
            png_bytes = renderPM.drawToString(drawing, fmt="PNG")
            return png_bytes
        except Exception as exc:
            print(f"svglib render failed: {exc}")
    return None


def _draw_key_badge(draw: ImageDraw.ImageDraw, font: ImageFont.ImageFont, key_text: str) -> None:
    """Draw a small dark circle badge with the hotkey text."""
    if not key_text:
        return
    bbox = draw.textbbox((0, 0), key_text, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    radius = int(max(text_w, text_h) / 2 + 6)
    cx = 10 + radius
    cy = 10 + radius
    draw.ellipse(
        (cx - radius, cy - radius, cx + radius, cy + radius),
        fill=(18, 18, 18, 210),
    )
    draw.text(
        (cx - text_w / 2, cy - text_h / 2),
        key_text,
        font=font,
        fill=(229, 229, 229, 255),
    )


def load_icon_image(
    name: str, hotkey_text: str, *, variant: str = "full", size: tuple[int, int] | None = None
) -> ImageTk.PhotoImage | None:
    """Return a PhotoImage with icon overlays, or None if unavailable.

    variant: "full" keeps the name ribbon; "badge" keeps only the key badge.
    """
    target_size = size or ICON_SIZE
    cache_key = (name, hotkey_text, variant, target_size)
    if cache_key in _icon_cache:
        return _icon_cache[cache_key]

    key = _normalize_name(name)
    asset_path = ASSET_MAP.get(key)
    if not asset_path or not asset_path.exists():
        return None

    try:
        if asset_path.suffix.lower() == ".png":
            image = Image.open(asset_path).convert("RGBA")
            if image.size != target_size:
                image = image.resize(target_size, Image.LANCZOS)
        else:
            svg_bytes = asset_path.read_bytes()
            png_bytes = _svg_to_png_bytes(svg_bytes, target_size)
            if png_bytes is None:
                if not getattr(load_icon_image, "_warned", False):
                    msg = "SVG rasterization unavailable; icons will not display."
                    details = _cairosvg_error or _svglib_error
                    if details:
                        msg += f" ({details})"
                    print(msg)
                    load_icon_image._warned = True  # type: ignore[attr-defined]
                return None
            image = Image.open(io.BytesIO(png_bytes)).convert("RGBA")
            if image.size != target_size:
                image = image.resize(target_size, Image.LANCZOS)
        draw = ImageDraw.Draw(image)
        font = ImageFont.load_default()

        # Top-left hotkey label.
        key_text = hotkey_text.strip()
        if key_text:
            if variant == "badge":
                _draw_key_badge(draw, font, key_text)
            else:
                draw.text((6, 4), key_text, font=font, fill=(229, 229, 229, 255))

        # Bottom name overlay.
        if variant == "full":
            name_w, name_h = draw.textbbox((0, 0), name, font=font)[2:]
            overlay_height = name_h + 8
            y0 = image.height - overlay_height
            draw.rectangle([0, y0, image.width, image.height], fill=(18, 18, 18, 180))
            draw.text(
                ((image.width - name_w) / 2, y0 + 4),
                name,
                font=font,
                fill=(229, 229, 229, 255),
            )

        photo = ImageTk.PhotoImage(image)
        _icon_cache[cache_key] = photo
        return photo
    except Exception:
        return None


def build_overlay_placeholder(
    hotkey_text: str, size: tuple[int, int] = OVERLAY_ICON_SIZE
) -> ImageTk.PhotoImage:
    """Create a dimmed placeholder tile with just the hotkey badge."""
    label = hotkey_text.strip() or "?"
    cache_key = (label, size)
    if cache_key in _overlay_placeholder_cache:
        return _overlay_placeholder_cache[cache_key]
    image = Image.new("RGBA", size, (26, 26, 26, 180))
    draw = ImageDraw.Draw(image)
    draw.rectangle((1, 1, size[0] - 2, size[1] - 2), outline=(70, 70, 70, 210), width=2)
    font = ImageFont.load_default()
    _draw_key_badge(draw, font, label)
    photo = ImageTk.PhotoImage(image)
    _overlay_placeholder_cache[cache_key] = photo
    return photo

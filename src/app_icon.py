from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image

from src.app_paths import data_dir, logo_path


def icon_path() -> Path:
    return data_dir() / "logo_cuesta.ico"


def ensure_app_icon() -> Path | None:
    """Gera logo_cuesta.ico a partir do PNG (vários tamanhos para Windows)."""
    png = logo_path()
    ico = icon_path()
    if ico.is_file():
        if not png.is_file() or ico.stat().st_mtime >= png.stat().st_mtime:
            return ico
    if getattr(sys, "frozen", False):
        return ico if ico.is_file() else None
    if not png.is_file():
        return None
    img = Image.open(png)
    if img.mode != "RGBA":
        img = img.convert("RGBA")
    base = img.copy()
    if base.width < 256 or base.height < 256:
        base = base.resize((256, 256), Image.Resampling.LANCZOS)
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    ico.parent.mkdir(parents=True, exist_ok=True)
    base.save(ico, format="ICO", sizes=sizes)
    return ico

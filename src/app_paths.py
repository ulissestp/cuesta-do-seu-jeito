from __future__ import annotations

import sys
from pathlib import Path


def _frozen() -> bool:
    return getattr(sys, "frozen", False)


def project_root() -> Path:
    if _frozen():
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent.parent


def writable_root() -> Path:
    """Pasta gravável (planilhas e PDFs de saída)."""
    if _frozen():
        return Path(sys.executable).parent
    return project_root()


def data_dir() -> Path:
    return project_root() / "data"


def logo_path() -> Path:
    return data_dir() / "logo_cuesta.png"


def output_dir() -> Path:
    path = writable_root() / "output"
    path.mkdir(parents=True, exist_ok=True)
    return path

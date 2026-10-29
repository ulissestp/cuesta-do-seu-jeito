"""Persistência da planilha de trabalho da última sessão."""

from __future__ import annotations

import json
from pathlib import Path

from src.app_paths import output_dir

_SESSION_FILE = "ultima_sessao.json"


def _session_file() -> Path:
    return output_dir() / _SESSION_FILE


def load_last_workbook() -> Path | None:
    path = _session_file()
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        workbook = Path(data.get("workbook", ""))
    except (json.JSONDecodeError, OSError, TypeError):
        return None
    if workbook.is_file():
        return workbook.resolve()
    return None


def remember_workbook(workbook_path: Path) -> None:
    out = output_dir()
    out.mkdir(parents=True, exist_ok=True)
    payload = {"workbook": str(workbook_path.resolve())}
    _session_file().write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

"""Caminhos e leitura da documentação de ajuda (dev e executável empacotado)."""

from __future__ import annotations

from pathlib import Path

from src.app_paths import docs_dir


def help_markdown_path() -> Path:
    return docs_dir() / "AJUDA.md"


def user_manual_md_path() -> Path:
    return docs_dir() / "MANUAL_DO_USUARIO.md"


def user_manual_pdf_path() -> Path:
    return docs_dir() / "Manual_CUESTA_RT_v1.1.pdf"


def load_help_markdown() -> str:
    path = help_markdown_path()
    if path.is_file():
        return path.read_text(encoding="utf-8")
    return (
        "Arquivo de ajuda não encontrado.\n\n"
        "Esperado em: docs/AJUDA.md (desenvolvimento) ou pasta docs no pacote do programa."
    )

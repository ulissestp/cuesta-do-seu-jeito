"""Constantes compartilhadas do formulário de perfil."""

from src.excel_io import CIDADES_ROTEIRO, PERFIL_OPCOES, RITMO_OPCOES

FORM_BLANK = "—"

CIDADE_FORM_LABELS = [f"{idx} - {city}" for idx, city in enumerate(CIDADES_ROTEIRO, start=1)]

SCORE_OPCOES = ["1", "2", "3", "4", "5"]
SCORE_FORM_OPCOES = [FORM_BLANK, *SCORE_OPCOES]

PERFIL_FORM_OPCOES = [FORM_BLANK, *PERFIL_OPCOES]
RITMO_FORM_OPCOES = [FORM_BLANK, *RITMO_OPCOES]


def is_form_blank(value: str | None) -> bool:
    if value is None:
        return True
    text = str(value).strip()
    return not text or text == FORM_BLANK


def normalize_form_choice(value: str) -> str:
    return "" if is_form_blank(value) else value.strip()


def parse_score_choice(value: str) -> int:
    if is_form_blank(value):
        return 0
    return int(value.strip())

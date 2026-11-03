from __future__ import annotations

from datetime import date, time

from src.brazil_holidays import is_brazilian_holiday
from src.excel_io import Attraction

WEEKDAY_KEYS = ("segunda", "terca", "quarta", "quinta", "sexta")
SCHEDULE_KEYS = WEEKDAY_KEYS + ("feriado",)
DEFAULT_OPEN = time(8, 0)
DEFAULT_CLOSE = time(18, 0)


def schedule_key_for(day: date, horarios: dict | None = None) -> str | None:
    """Coluna de horário da data. Sem sábado/domingo na planilha, reaproveita sexta e segunda."""
    columns = horarios or {}
    if is_brazilian_holiday(day) and (not columns or "feriado" in columns):
        return "feriado"
    weekday = day.weekday()
    if weekday <= 4:
        return WEEKDAY_KEYS[weekday]
    if weekday == 5:
        if "sabado" in columns:
            return "sabado"
        if "sexta" in columns or not columns:
            return "sexta"
    if weekday == 6:
        if "domingo" in columns:
            return "domingo"
        if "segunda" in columns or not columns:
            return "segunda"
    return None


def resolved_hours(attraction: Attraction, day: date) -> tuple[time, time] | None:
    """Horário do atrativo na data. None = fechado nesse dia.

    Inventário antigo, sem grade semanal, mantém o horário único em qualquer dia.
    Sem colunas de sábado e domingo, o fim de semana reaproveita sexta e segunda.
    """
    horarios = attraction.horarios
    if horarios:
        key = schedule_key_for(day, horarios)
        if not key:
            return None
        slot = horarios.get(key)
        if not slot:
            return None
        abertura, fechamento = slot
        if abertura is None or fechamento is None or fechamento <= abertura:
            return None
        return abertura, fechamento
    if attraction.horario_abertura or attraction.horario_fechamento:
        return (
            attraction.horario_abertura or DEFAULT_OPEN,
            attraction.horario_fechamento or DEFAULT_CLOSE,
        )
    return DEFAULT_OPEN, DEFAULT_CLOSE

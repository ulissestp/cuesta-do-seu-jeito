from __future__ import annotations

from datetime import date, timedelta
from functools import lru_cache


def easter_sunday(year: int) -> date:
    """Domingo de Páscoa (algoritmo gregoriano)."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    ell = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * ell) // 451
    month = (h + ell - 7 * m + 114) // 31
    day = ((h + ell - 7 * m + 114) % 31) + 1
    return date(year, month, day)


@lru_cache(maxsize=16)
def brazilian_holidays(year: int) -> frozenset[date]:
    """Feriados nacionais, Carnaval, Corpus Christi e 9 de julho (São Paulo)."""
    easter = easter_sunday(year)
    days = {
        date(year, 1, 1),
        easter - timedelta(days=48),  # segunda de Carnaval
        easter - timedelta(days=47),  # terça de Carnaval
        easter - timedelta(days=2),  # Sexta-feira Santa
        date(year, 4, 21),
        date(year, 5, 1),
        easter + timedelta(days=60),  # Corpus Christi
        date(year, 7, 9),
        date(year, 9, 7),
        date(year, 10, 12),
        date(year, 11, 2),
        date(year, 11, 15),
        date(year, 11, 20),
        date(year, 12, 25),
    }
    return frozenset(days)


def is_brazilian_holiday(day: date) -> bool:
    return day in brazilian_holidays(day.year)

from __future__ import annotations

from pathlib import Path

from src.excel_io import (
    read_inventory,
    read_profile,
    format_profile_validation_errors,
    validate_profile_for_roteiro,
    filter_inventory_by_cities,
)
from src.group_scores import GroupScores
from src.itinerary_engine import ItineraryResult, build_itinerary, itinerary_to_markdown


def gerar_roteiro_da_planilha(workbook_path: Path) -> ItineraryResult:
    profile = read_profile(workbook_path)
    missing = validate_profile_for_roteiro(profile)
    if missing:
        raise ValueError(format_profile_validation_errors(missing))
    inventory = read_inventory(workbook_path)
    if not inventory:
        raise ValueError("Inventário turístico vazio. Importe a planilha na etapa 2.")
    inventory = filter_inventory_by_cities(inventory, profile.cidades)
    if not inventory:
        cities = ", ".join(profile.cidades)
        raise ValueError(
            f"Nenhuma atração encontrada para as cidades selecionadas ({cities}). "
            "Verifique a coluna Cidade (K) do inventário."
        )
    scores = GroupScores.from_profile(profile)
    return build_itinerary(profile, inventory, scores)


def gerar_roteiro_markdown(workbook_path: Path) -> str:
    return itinerary_to_markdown(gerar_roteiro_da_planilha(workbook_path))

from __future__ import annotations

from pathlib import Path

from src.excel_io import (
    read_inventory,
    read_profile,
    duplicate_attraction_ids,
    duplicate_ids_message,
    format_profile_validation_errors,
    validate_profile_for_roteiro,
    filter_inventory_by_categories,
    filter_inventory_by_cities,
)
from src.group_scores import GroupScores
from src.itinerary_engine import (
    ItineraryResult,
    ProgressCallback,
    _report,
    build_itinerary,
    itinerary_opening_markdown,
    itinerary_to_markdown,
)


def gerar_roteiro_da_planilha(
    workbook_path: Path,
    pinned_ids: list[str] | None = None,
    on_progress: ProgressCallback | None = None,
) -> ItineraryResult:
    _report(on_progress, "Lendo o perfil e o inventário…", 0.08)
    profile = read_profile(workbook_path)
    missing = validate_profile_for_roteiro(profile)
    if missing:
        raise ValueError(format_profile_validation_errors(missing))
    _report(on_progress, "Conferindo o inventário…", 0.14)
    inventory = read_inventory(workbook_path)
    if not inventory:
        raise ValueError("Inventário turístico vazio. Importe a planilha na etapa 2.")
    duplicated = duplicate_attraction_ids(inventory)
    if duplicated:
        raise ValueError(duplicate_ids_message(duplicated))
    _report(on_progress, "Filtrando cidades e categorias…", 0.2)
    inventory = filter_inventory_by_cities(inventory, profile.cidades)
    inventory = filter_inventory_by_categories(inventory, profile.categorias)
    if not inventory:
        cities = ", ".join(profile.cidades)
        categories = ", ".join(profile.categorias)
        raise ValueError(
            "Nenhuma atração encontrada para as cidades e categorias selecionadas "
            f"({cities}; {categories})."
        )
    if pinned_ids:
        found = {item.id for item in inventory}
        missing = [item_id for item_id in pinned_ids if item_id not in found]
        if missing:
            raise ValueError(
                "ID não encontrado nas cidades e categorias do perfil: " + ", ".join(missing)
            )
    scores = GroupScores.from_profile(profile)
    cities = ", ".join(profile.cidades) if profile.cidades else "—"
    _report(
        on_progress,
        (
            f"Perfil do grupo calculado. {len(inventory)} atrativos em {cities}. "
            "Montando a programação…"
        ),
        0.26,
        itinerary_opening_markdown(profile, scores),
    )
    return build_itinerary(
        profile,
        inventory,
        scores,
        pinned_ids=pinned_ids,
        on_progress=on_progress,
    )


def gerar_roteiro_markdown(workbook_path: Path) -> str:
    return itinerary_to_markdown(gerar_roteiro_da_planilha(workbook_path))

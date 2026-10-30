from __future__ import annotations

import random
from dataclasses import dataclass, field
from datetime import datetime, time, timedelta

from src.excel_io import Attraction, TouristProfile, canonical_city_name
from src.geo_distance import GeoContext, LatLng, haversine_km
from src.group_scores import GroupScores, build_technical_analysis
from src.maps_links import maps_markdown_line, try_shorten_maps_url

DAY_END_HOUR = 18
DAY_LATEST_FIRST_START = time(9, 30)
DEFAULT_OPEN = time(8, 0)
DEFAULT_CLOSE = time(18, 0)
MIN_VISIT_MINUTES = 45
LODGING_TIMELINE_TITLE = "Hospedagem"


@dataclass
class ScheduleBlock:
    start: datetime
    end: datetime
    title: str
    detail: str
    kind: str  # visit | transit | lunch | end | lodging
    maps_url: str = ""


@dataclass
class DayPlan:
    day_number: int
    theme: str
    blocks: list[ScheduleBlock] = field(default_factory=list)


@dataclass
class ItineraryResult:
    profile: TouristProfile
    scores: GroupScores
    selected_attractions: list[Attraction]
    days: list[DayPlan]
    justification: str
    media_plog: float
    media_cohen: float
    lodging_suggestions: list[Attraction] = field(default_factory=list)


def _passes_fator_restritivo(attraction: Attraction, profile: TouristProfile, fator: str) -> bool:
    if "Alto" in fator:
        if attraction.peso_cohen >= 4.5:
            return False
        if "extrema" in attraction.categoria.lower():
            return False
        if attraction.capacidade < profile.numero_pessoas + 5:
            return False
    if fator == "Médio":
        if attraction.capacidade < profile.numero_pessoas + 2:
            return False
    return True


def _is_gastronomia(attraction: Attraction) -> bool:
    return "gastronomia" in attraction.categoria.lower()


def _is_hospedagem(attraction: Attraction) -> bool:
    return "hospedagem" in attraction.categoria.lower()


def filter_attractions(
    profile: TouristProfile,
    inventory: list[Attraction],
    scores: GroupScores,
) -> list[Attraction]:
    plog = scores.media_plog
    cohen = scores.media_cohen
    filtered: list[Attraction] = []
    for a in inventory:
        if a.capacidade < profile.numero_pessoas:
            continue
        if _is_hospedagem(a):
            continue
        if not _passes_fator_restritivo(a, profile, scores.fator_restritivo):
            continue
        if abs(a.peso_plog - plog) > 1.0 or abs(a.peso_cohen - cohen) > 1.0:
            continue
        filtered.append(a)

    if not filtered:
        for a in inventory:
            if a.capacidade >= profile.numero_pessoas and _passes_fator_restritivo(
                a, profile, scores.fator_restritivo
            ) and not _is_hospedagem(a):
                filtered.append(a)

    filtered.sort(
        key=lambda x: (abs(x.peso_plog - plog) + abs(x.peso_cohen - cohen), -x.tempo_horas)
    )
    return filtered


def _max_visits_for_ritmo(ritmo: str) -> int:
    """Limite de visitas turísticas por dia (almoço não conta)."""
    if "Relaxado" in ritmo:
        return 2
    if "Intenso" in ritmo:
        return 6
    return 4  # Moderado (Equilibrado) — várias atividades


def _city_label(attraction: Attraction) -> str:
    raw = (attraction.cidade or "").strip()
    if raw:
        return raw
    canon = canonical_city_name(raw)
    return canon or "Cidade não informada"


def _visit_title(attraction: Attraction) -> str:
    return f"{attraction.nome} — {_city_label(attraction)}"


def _schedule_visits_phase(
    cursor: datetime,
    *,
    pool: list[Attraction],
    day_anchor: datetime,
    used_ids: set[str],
    blocks: list[ScheduleBlock],
    scheduled: list[Attraction],
    max_visits: int,
    visits_so_far: int,
    stop_before: datetime | None,
    geo: GeoContext,
    scores: GroupScores,
    day_start: LatLng | None,
    start_from: Attraction | None = None,
) -> tuple[datetime, int, Attraction | None]:
    """Agenda visitas até o limite do ritmo ou horário, otimizando ordem por distância no dia."""
    day_limit = _combine(day_anchor, time(DAY_END_HOUR, 0))
    cursor_local = cursor
    count = visits_so_far
    last_attr: Attraction | None = start_from
    while count < max_visits:
        limit = stop_before or day_limit
        if cursor_local >= min(limit, day_limit):
            break
        prev = cursor_local
        cursor_local, last_attr = _schedule_from_pool(
            cursor_local,
            pool,
            day_anchor,
            used_ids,
            blocks,
            scheduled,
            geo=geo,
            scores=scores,
            last_attr=last_attr,
            day_start=day_start,
            time_limit=min(limit, day_limit),
        )
        if cursor_local == prev:
            break
        count += 1
    return cursor_local, count, last_attr


def _open_for_lunch(attraction: Attraction, day_anchor: datetime) -> bool:
    open_dt = _combine(day_anchor, attraction.horario_abertura or DEFAULT_OPEN)
    close_dt = _combine(day_anchor, attraction.horario_fechamento or DEFAULT_CLOSE)
    lunch_start = _combine(day_anchor, time(12, 0))
    lunch_end = _combine(day_anchor, time(13, 30))
    return open_dt <= lunch_start and close_dt >= lunch_end


def _score_distance(a: Attraction, plog: float, cohen: float) -> float:
    return abs(a.peso_plog - plog) + abs(a.peso_cohen - cohen)


def _geo_sort_key(
    attraction: Attraction,
    reference: LatLng | None,
    geo: GeoContext,
    plog: float,
    cohen: float,
) -> tuple[float, float]:
    """Menor = melhor: distância geográfica primeiro, depois score Plog/Cohen."""
    point = geo.coords(attraction)
    if reference and point:
        geo_km = haversine_km(reference, point)
    elif point and not reference:
        geo_km = 0.0
    else:
        geo_km = 9999.0
    return (geo_km, _score_distance(attraction, plog, cohen))


def _pick_lunch_gastronomy(
    profile: TouristProfile,
    inventory: list[Attraction],
    scores: GroupScores,
    used_ids: set[str],
    day_anchor: datetime,
    *,
    geo: GeoContext | None = None,
    near: LatLng | None = None,
) -> Attraction | None:
    plog, cohen = scores.media_plog, scores.media_cohen

    strict: list[Attraction] = []
    loose: list[Attraction] = []
    for a in inventory:
        if a.id in used_ids or not _is_gastronomia(a):
            continue
        if a.capacidade < profile.numero_pessoas:
            continue
        if not _passes_fator_restritivo(a, profile, scores.fator_restritivo):
            continue
        if not _open_for_lunch(a, day_anchor):
            continue
        loose.append(a)
        if abs(a.peso_plog - plog) <= 1.0 and abs(a.peso_cohen - cohen) <= 1.0:
            strict.append(a)

    pool = strict or loose
    if not pool:
        return None

    def sort_key(a: Attraction) -> tuple[float, float]:
        if geo and near:
            return _geo_sort_key(a, near, geo, plog, cohen)
        return (_score_distance(a, plog, cohen), 0.0)

    pool.sort(key=sort_key)
    best = sort_key(pool[0])
    tier = [a for a in pool if sort_key(a) <= (best[0] + 0.51, best[1] + 0.51)]
    if len(tier) == 1 and len(pool) > 1:
        tier = pool[: min(5, len(pool))]
    return tier[0]


def _pick_lodging_suggestions(
    profile: TouristProfile,
    inventory: list[Attraction],
    scores: GroupScores,
    *,
    count: int = 2,
) -> list[Attraction]:
    plog, cohen = scores.media_plog, scores.media_cohen

    def distance(a: Attraction) -> float:
        return abs(a.peso_plog - plog) + abs(a.peso_cohen - cohen)

    pool = [
        a
        for a in inventory
        if _is_hospedagem(a)
        and a.capacidade >= profile.numero_pessoas
        and _passes_fator_restritivo(a, profile, scores.fator_restritivo)
    ]
    if not pool:
        return []

    picks: list[Attraction] = []
    remaining = list(pool)
    for _ in range(min(count, len(remaining))):
        remaining.sort(key=distance)
        best = distance(remaining[0])
        tier = [a for a in remaining if distance(a) <= best + 0.51]
        if len(tier) == 1 and len(remaining) > 1:
            tier = remaining[: min(5, len(remaining))]
        choice = random.choice(tier)
        picks.append(choice)
        remaining = [a for a in remaining if a.id != choice.id]
    return picks


def _visit_pool(selected: list[Attraction]) -> list[Attraction]:
    non_gastro = [
        a for a in selected if not _is_gastronomia(a) and not _is_hospedagem(a)
    ]
    return non_gastro or list(selected)


def _theme_for_day(attractions: list[Attraction]) -> str:
    if not attractions:
        return "Exploração local"
    cats = {a.categoria.split("/")[0].strip() for a in attractions}
    if len(cats) == 1:
        return list(cats)[0]
    return " e ".join(sorted(cats)[:2])


def _day_anchor(day_number: int) -> datetime:
    return datetime(2026, 1, day_number, DAY_LATEST_FIRST_START.hour, DAY_LATEST_FIRST_START.minute)


def _combine(day_anchor: datetime, hour_minute: time) -> datetime:
    return datetime.combine(day_anchor.date(), hour_minute)


def _maps_plain_suffix(maps_url: str) -> str:
    url = try_shorten_maps_url(maps_url)
    return f" Localização: {url}" if url else ""


def _visit_detail(attraction: Attraction, start: datetime, end: datetime) -> str:
    elapsed_h = (end - start).total_seconds() / 3600
    city = _city_label(attraction)
    return (
        f"Cidade: {city}. {attraction.descricao} "
        f"Tempo médio de visitação: {attraction.tempo_horas:g} h "
        f"(neste roteiro: {elapsed_h:.1f} h)."
    )


def _lunch_detail(attraction: Attraction, profile: TouristProfile) -> str:
    return (
        f"Pausa para almoço (1h30min, entre 12h00 e 13h30). "
        f"Cidade: {_city_label(attraction)}. {attraction.descricao}"
    )


def _lodging_options_summary(attractions: list[Attraction]) -> str:
    if not attractions:
        return "—"
    lines = ["Sugestões para pernoite — não entram na programação de visitas do dia."]
    for i, attraction in enumerate(attractions, 1):
        maps_part = _maps_plain_suffix(attraction.maps_url)
        lines.append(f"Opção {i}: {_visit_title(attraction)}. {attraction.descricao}{maps_part}")
    return "\n".join(lines)


def _fmt_time(dt: datetime) -> str:
    return dt.strftime("%H:%M")


def _schedule_visit(
    cursor: datetime,
    attraction: Attraction,
    day_anchor: datetime,
) -> tuple[datetime, datetime] | None:
    day_limit = _combine(day_anchor, time(DAY_END_HOUR, 0))
    open_dt = _combine(day_anchor, attraction.horario_abertura or DEFAULT_OPEN)
    close_dt = _combine(day_anchor, attraction.horario_fechamento or DEFAULT_CLOSE)

    start = max(cursor, open_dt)
    if start >= min(close_dt, day_limit):
        return None

    planned_end = start + timedelta(hours=attraction.tempo_horas)
    end = min(planned_end, close_dt, day_limit)

    if (end - start).total_seconds() < MIN_VISIT_MINUTES * 60:
        return None
    return start, end


def _transit_detail_km(km: float | None, minutes: int, label: str) -> str:
    if km is not None:
        return f"Deslocamento {label} (~{km:.1f} km, {minutes} minutos)."
    return f"Deslocamento {label} ({minutes} minutos)."


def _transit_detail(
    origin: Attraction,
    destination: Attraction,
    minutes: int,
    geo: GeoContext,
) -> str:
    km = geo.distance_km(origin, destination)
    if km is not None:
        return (
            f"Deslocamento de {_visit_title(origin)} até {_visit_title(destination)} "
            f"(~{km:.1f} km, {minutes} minutos)."
        )
    return f"Trânsito entre as localidades ({minutes} minutos)."


def _append_transit_from_reference(
    blocks: list[ScheduleBlock],
    visit_start: datetime,
    minutes: int,
    destination: Attraction,
    geo: GeoContext,
    origin: LatLng | None,
) -> None:
    if minutes <= 0:
        return
    start = visit_start - timedelta(minutes=minutes)
    km = None
    dest = geo.coords_resolved(destination)
    if origin and dest:
        km = haversine_km(origin, dest)
    label = f"até {_visit_title(destination)}"
    blocks.append(
        ScheduleBlock(
            start=start,
            end=visit_start,
            title="Deslocamento",
            detail=_transit_detail_km(km, minutes, label),
            kind="transit",
        )
    )


def _append_transit(
    blocks: list[ScheduleBlock],
    cursor: datetime,
    minutes: int,
    origin: Attraction,
    destination: Attraction,
    geo: GeoContext,
) -> datetime:
    transit_end = cursor + timedelta(minutes=minutes)
    blocks.append(
        ScheduleBlock(
            start=cursor,
            end=transit_end,
            title="Deslocamento",
            detail=_transit_detail(origin, destination, minutes, geo),
            kind="transit",
        )
    )
    return transit_end


def _schedule_from_pool(
    cursor: datetime,
    pool: list[Attraction],
    day_anchor: datetime,
    used_ids: set[str],
    blocks: list[ScheduleBlock],
    scheduled: list[Attraction],
    *,
    geo: GeoContext,
    scores: GroupScores,
    last_attr: Attraction | None,
    day_start: LatLng | None,
    time_limit: datetime,
) -> tuple[datetime, Attraction | None]:
    plog, cohen = scores.media_plog, scores.media_cohen
    reference = geo.coords(last_attr) if last_attr else day_start
    candidates: list[tuple[Attraction, datetime, datetime, int]] = []

    for attraction in pool:
        if attraction.id in used_ids:
            continue
        if last_attr is None:
            open_t = attraction.horario_abertura or DEFAULT_OPEN
            if open_t > DAY_LATEST_FIRST_START:
                continue
            transit_min = geo.travel_minutes_from_reference(day_start, attraction)
            cursor_after_transit = _combine(day_anchor, DAY_LATEST_FIRST_START)
        else:
            transit_min = geo.travel_minutes(last_attr, attraction)
            cursor_after_transit = cursor + timedelta(minutes=transit_min)
        if cursor_after_transit >= time_limit:
            continue
        slot = _schedule_visit(cursor_after_transit, attraction, day_anchor)
        if not slot:
            continue
        start, end = slot
        if last_attr is None and start > _combine(day_anchor, DAY_LATEST_FIRST_START):
            continue
        if end > time_limit:
            continue
        candidates.append((attraction, start, end, transit_min))

    if not candidates:
        return cursor, None

    candidates.sort(key=lambda item: _geo_sort_key(item[0], reference, geo, plog, cohen))
    attraction, start, end, transit_min = candidates[0]

    if last_attr is None and transit_min > 0:
        _append_transit_from_reference(blocks, start, transit_min, attraction, geo, day_start)
    elif last_attr and transit_min > 0:
        cursor = _append_transit(blocks, cursor, transit_min, last_attr, attraction, geo)

    blocks.append(
        ScheduleBlock(
            start=start,
            end=end,
            title=_visit_title(attraction),
            detail=_visit_detail(attraction, start, end),
            kind="visit",
            maps_url=attraction.maps_url,
        )
    )
    used_ids.add(attraction.id)
    scheduled.append(attraction)
    return end, attraction


def build_itinerary(
    profile: TouristProfile,
    inventory: list[Attraction],
    scores: GroupScores | None = None,
) -> ItineraryResult:
    scores = scores or GroupScores.from_profile(profile)
    selected = filter_attractions(profile, inventory, scores)
    visit_candidates = _visit_pool(selected)
    days_count = max(1, profile.duracao_dias)
    max_visits = _max_visits_for_ritmo(profile.ritmo)
    geo = GeoContext(inventory)

    days: list[DayPlan] = []
    used_global: set[str] = set()
    pool = list(visit_candidates)
    trip_lodging = _pick_lodging_suggestions(profile, inventory, scores)

    for day in range(1, days_count + 1):
        day_anchor = _day_anchor(day)
        day_start = geo.day_start_reference(
            [a for a in pool if a.id not in used_global],
            profile.cidades,
        )
        blocks: list[ScheduleBlock] = []
        scheduled_today: list[Attraction] = []
        cursor = day_anchor
        lunch_cutoff = _combine(day_anchor, time(12, 0))

        cursor, visit_count, morning_last = _schedule_visits_phase(
            cursor,
            pool=pool,
            day_anchor=day_anchor,
            used_ids=used_global,
            blocks=blocks,
            scheduled=scheduled_today,
            max_visits=max_visits,
            visits_so_far=0,
            stop_before=lunch_cutoff,
            geo=geo,
            scores=scores,
            day_start=day_start,
        )

        lunch_start = max(cursor, lunch_cutoff)
        lunch_end = lunch_start + timedelta(hours=1, minutes=30)
        lunch_near = geo.coords(morning_last) if morning_last else day_start
        lunch_attr = _pick_lunch_gastronomy(
            profile,
            inventory,
            scores,
            used_global,
            day_anchor,
            geo=geo,
            near=lunch_near,
        )
        if lunch_attr:
            used_global.add(lunch_attr.id)
            scheduled_today.append(lunch_attr)
            lunch_title = f"Almoço — {_visit_title(lunch_attr)}"
            lunch_detail = _lunch_detail(lunch_attr, profile)
        else:
            lunch_title = "Pausa para Almoço"
            lunch_detail = (
                "Pausa de 1h30min entre 12h00 e 13h30. "
                "Nenhuma atração de gastronomia compatível com o score no inventário."
            )
        blocks.append(
            ScheduleBlock(
                start=lunch_start,
                end=lunch_end,
                title=lunch_title,
                detail=lunch_detail,
                kind="lunch",
                maps_url=lunch_attr.maps_url if lunch_attr else "",
            )
        )
        cursor = lunch_end
        afternoon_start = geo.coords(lunch_attr) if lunch_attr else lunch_near

        cursor, visit_count, _ = _schedule_visits_phase(
            cursor,
            pool=pool,
            day_anchor=day_anchor,
            used_ids=used_global,
            blocks=blocks,
            scheduled=scheduled_today,
            max_visits=max_visits,
            visits_so_far=visit_count,
            stop_before=None,
            geo=geo,
            scores=scores,
            day_start=afternoon_start or day_start,
            start_from=lunch_attr,
        )

        theme = _theme_for_day(scheduled_today)
        if cursor.hour < DAY_END_HOUR:
            blocks.append(
                ScheduleBlock(
                    start=cursor,
                    end=cursor,
                    title="Encerramento do Dia",
                    detail="Fim das atividades oficiais do roteiro, respeitando a regra de conclusão antes das 18h00.",
                    kind="end",
                )
            )
        elif scheduled_today:
            blocks.append(
                ScheduleBlock(
                    start=_combine(day_anchor, time(17, 0)),
                    end=_combine(day_anchor, time(17, 0)),
                    title="Tarde Livre",
                    detail="Descanso ou exploração autônoma após as atividades programadas.",
                    kind="end",
                )
            )

        if trip_lodging:
            night_marker = _combine(day_anchor, time(18, 0))
            blocks.append(
                ScheduleBlock(
                    start=night_marker,
                    end=night_marker,
                    title=LODGING_TIMELINE_TITLE,
                    detail="",
                    kind="lodging",
                )
            )

        days.append(DayPlan(day_number=day, theme=theme, blocks=blocks))

    used_attractions = [
        a for a in selected if a.id in used_global and not _is_hospedagem(a)
    ]
    weights = [(a.peso_plog, a.peso_cohen) for a in used_attractions]
    justification = build_technical_analysis(scores, weights)

    return ItineraryResult(
        profile=profile,
        scores=scores,
        selected_attractions=used_attractions or selected,
        days=days,
        justification=justification,
        media_plog=scores.media_plog,
        media_cohen=scores.media_cohen,
        lodging_suggestions=trip_lodging,
    )


def _block_detail_markdown(block: ScheduleBlock) -> str:
    detail = block.detail
    if block.maps_url:
        line = maps_markdown_line(block.maps_url)
        if line:
            detail = f"{detail} {line}"
    return detail


def itinerary_to_markdown(result: ItineraryResult) -> str:
    p = result.profile
    s = result.scores
    lines = [
        f"# Roteiro de Viagem — {p.nome_grupo}",
        f"Duração: {p.duracao_dias} dia(s)",
        f"Média Plog: {s.media_plog:.2f} | Média Cohen: {s.media_cohen:.2f} | Fator: {s.fator_restritivo}",
        "",
        "### Quadro-resumo — Sugestões de hospedagem (pernoite)",
        _lodging_options_summary(result.lodging_suggestions),
        "",
    ]
    for day in result.days:
        lines.append(f"## DIA {day.day_number}: {day.theme}")
        for block in day.blocks:
            if block.kind == "lodging" and block.start == block.end:
                lines.append(f"**Noite — {block.title}**")
            elif block.kind == "end" and block.start == block.end:
                lines.append(f"**{_fmt_time(block.start)} — {block.title}** — {block.detail}")
            else:
                lines.append(
                    f"**{_fmt_time(block.start)} – {_fmt_time(block.end)}** — "
                    f"{block.title}: {_block_detail_markdown(block)}"
                )
        lines.append("")
    lines.append("## Análise e validação técnica")
    lines.append(result.justification)
    return "\n".join(lines)

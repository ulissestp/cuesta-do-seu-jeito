from __future__ import annotations

import random
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta

from src.brazil_holidays import is_brazilian_holiday
from src.excel_io import Attraction, TouristProfile, canonical_city_name, trip_dates
from src.opening_hours import resolved_hours
from src.geo_distance import GeoContext, LatLng, haversine_km, travel_minutes_between
from src.group_scores import GroupScores, build_technical_analysis
from src.maps_links import maps_markdown_line, try_shorten_maps_url

DAY_END_HOUR = 18
DAY_LATEST_FIRST_START = time(9, 30)
DEFAULT_OPEN = time(8, 0)
DEFAULT_CLOSE = time(18, 0)
MIN_VISIT_MINUTES = 45
MAX_TRANSIT_MINUTES = 40
LODGING_TIMELINE_TITLE = "Hospedagem"
EVENING_DINING_START = time(19, 0)
EVENING_DINING_TIMELINE_TITLE = "Sugestões para o jantar (após 19h)"
EVENING_DINING_SUGGESTION_COUNT = 2


@dataclass
class ScheduleBlock:
    start: datetime
    end: datetime
    title: str
    detail: str
    kind: str  # visit | transit | lunch | end | lodging | evening_dining
    maps_url: str = ""
    image_url: str = ""


@dataclass
class DayPlan:
    day_number: int
    theme: str
    blocks: list[ScheduleBlock] = field(default_factory=list)
    service_date: datetime | None = None


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
    fora_do_perfil: bool = False
    atrativos_fora_do_perfil: list[str] = field(default_factory=list)


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


def _fora_clause(attraction: Attraction, scores: GroupScores) -> str:
    if _within_profile(attraction, scores):
        return ""
    return " Fora do perfil, incluído por estar entre os mais próximos da média do grupo."


def _within_profile(attraction: Attraction, scores: GroupScores) -> bool:
    return (
        abs(attraction.peso_plog - scores.media_plog) <= 1.0
        and abs(attraction.peso_cohen - scores.media_cohen) <= 1.0
    )


def _eligible_attractions(
    profile: TouristProfile,
    inventory: list[Attraction],
    scores: GroupScores,
) -> list[Attraction]:
    return [
        a
        for a in inventory
        if a.capacidade >= profile.numero_pessoas
        and not _is_hospedagem(a)
        and _passes_fator_restritivo(a, profile, scores.fator_restritivo)
    ]


def filter_attractions(
    profile: TouristProfile,
    inventory: list[Attraction],
    scores: GroupScores,
) -> list[Attraction]:
    """Atrativos dentro da margem ±1 da média Plog e Cohen."""
    plog = scores.media_plog
    cohen = scores.media_cohen
    filtered = [a for a in _eligible_attractions(profile, inventory, scores) if _within_profile(a, scores)]
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


def _morning_visit_goal(max_visits: int) -> int:
    """Metade das visitas do ritmo fica de manhã; o restante, à tarde."""
    return max(1, (max_visits + 1) // 2)


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
    backup: list[Attraction] | None = None,
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
            reserve_after=max(0, max_visits - count - 1),
            backup=backup,
        )
        if cursor_local == prev:
            break
        count += 1
    return cursor_local, count, last_attr


def _cluster_anchor(attractions: list[Attraction], geo: GeoContext) -> LatLng | None:
    """Ponto central do dia: atrativo com menor soma de distâncias aos demais."""
    points = [(geo.coords(attraction), attraction) for attraction in attractions]
    located = [(point, attraction) for point, attraction in points if point]
    if not located:
        return None
    if len(located) == 1:
        return located[0][0]

    def spread(origin: LatLng) -> float:
        return sum(haversine_km(origin, point) for point, _ in located)

    return min(located, key=lambda item: spread(item[0]))[0]


def _km_between(origin: Attraction, destination: Attraction, geo: GeoContext) -> float:
    start = geo.coords(origin)
    end = geo.coords(destination)
    if not start or not end:
        return 9999.0
    return haversine_km(start, end)


def _tight_cluster(attractions: list[Attraction], size: int, geo: GeoContext) -> list[Attraction]:
    """Grupo do dia: os atrativos abertos que ficam mais perto uns dos outros."""
    if len(attractions) <= size:
        return list(attractions)

    def cost(seed: Attraction) -> float:
        nearest = sorted(
            _km_between(seed, other, geo) for other in attractions if other.id != seed.id
        )
        return sum(nearest[: size - 1])

    seed = min(attractions, key=cost)
    rest = sorted(
        (attraction for attraction in attractions if attraction.id != seed.id),
        key=lambda attraction: _km_between(seed, attraction, geo),
    )
    return [seed, *rest[: size - 1]]


def _day_visit_pool(
    service_date: date,
    in_profile: list[Attraction],
    outside: list[Attraction],
    used_ids: set[str],
    max_visits: int,
    geo: GeoContext,
    scores: GroupScores,
) -> tuple[list[Attraction], LatLng | None]:
    """Atrativos abertos no dia, completados pelos mais próximos do grupo daquele dia."""

    def is_open(attraction: Attraction) -> bool:
        return attraction.id not in used_ids and bool(resolved_hours(attraction, service_date))

    # Dois atrativos a mais que a meta: se um não couber no horário, o dia ainda completa o ritmo.
    target = max_visits + 2
    primary = [attraction for attraction in in_profile if is_open(attraction)]
    pool = _tight_cluster(primary, target, geo)
    anchor = _cluster_anchor(pool, geo)
    shortage = max(0, target - len(pool))
    if shortage:
        candidates = [attraction for attraction in outside if is_open(attraction)]

        def extra_key(attraction: Attraction) -> tuple[float, float]:
            point = geo.coords(attraction)
            km = haversine_km(anchor, point) if anchor and point else 9999.0
            return (km, _score_distance(attraction, scores.media_plog, scores.media_cohen))

        candidates.sort(key=extra_key)
        pool = pool + candidates[:shortage]
        if anchor is None:
            anchor = _cluster_anchor(pool, geo)
    return pool, anchor


def _open_for_lunch(attraction: Attraction, day_anchor: datetime) -> bool:
    hours = resolved_hours(attraction, day_anchor.date())
    if not hours:
        return False
    open_dt = _combine(day_anchor, hours[0])
    close_dt = _combine(day_anchor, hours[1])
    lunch_start = _combine(day_anchor, time(12, 0))
    lunch_end = _combine(day_anchor, time(13, 30))
    return open_dt <= lunch_start and close_dt >= lunch_end


def _open_for_evening_dining(attraction: Attraction, day_anchor: datetime) -> bool:
    """Gastronomia aberta às 19h (ainda em funcionamento após 19:00)."""
    hours = resolved_hours(attraction, day_anchor.date())
    if not hours:
        return False
    open_dt = _combine(day_anchor, hours[0])
    close_dt = _combine(day_anchor, hours[1])
    evening = _combine(day_anchor, EVENING_DINING_START)
    return open_dt <= evening and close_dt >= evening


def _open_on_trip_evening(attraction: Attraction, dates: list[date]) -> bool:
    return any(
        _open_for_evening_dining(attraction, datetime.combine(day, time(12, 0)))
        for day in dates
    )


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
    only_ids: set[str] | None = None,
) -> Attraction | None:
    plog, cohen = scores.media_plog, scores.media_cohen

    strict: list[Attraction] = []
    loose: list[Attraction] = []
    for a in inventory:
        if a.id in used_ids or not _is_gastronomia(a):
            continue
        if only_ids is not None and a.id not in only_ids:
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

    if not loose:
        return None

    def sort_key(a: Attraction) -> tuple[float, float]:
        if geo and near:
            return _geo_sort_key(a, near, geo, plog, cohen)
        return (_score_distance(a, plog, cohen), 0.0)

    if geo and near:
        reachable: list[Attraction] = []
        for attraction in loose:
            point = geo.coords_resolved(attraction)
            minutes = travel_minutes_between(near, point) if point else 999
            if minutes <= MAX_TRANSIT_MINUTES:
                reachable.append(attraction)
        loose = reachable
        if not loose:
            return None

    loose.sort(key=sort_key)
    nearest_km = sort_key(loose[0])[0]
    nearby = [a for a in loose if sort_key(a)[0] <= nearest_km + 3.0]
    in_profile = [a for a in nearby if a in strict or _within_profile(a, scores)]
    return (in_profile or nearby)[0]


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


def _pick_evening_gastronomy_suggestions(
    profile: TouristProfile,
    inventory: list[Attraction],
    scores: GroupScores,
    used_ids: set[str],
    dates: list[date],
    *,
    count: int = EVENING_DINING_SUGGESTION_COUNT,
) -> list[Attraction]:
    plog, cohen = scores.media_plog, scores.media_cohen

    def distance(a: Attraction) -> float:
        return abs(a.peso_plog - plog) + abs(a.peso_cohen - cohen)

    pool = [
        a
        for a in inventory
        if _is_gastronomia(a)
        and a.id not in used_ids
        and a.capacidade >= profile.numero_pessoas
        and _passes_fator_restritivo(a, profile, scores.fator_restritivo)
        and _open_on_trip_evening(a, dates)
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


def _valor_clause(attraction: Attraction) -> str:
    valor = (attraction.valor or "").strip()
    if not valor:
        return ""
    if valor.lower().startswith("valor"):
        return f" {valor}"
    return f" Valor: {valor}."


def _visit_detail(attraction: Attraction, start: datetime, end: datetime) -> str:
    elapsed_h = (end - start).total_seconds() / 3600
    city = _city_label(attraction)
    return (
        f"Cidade: {city}. {attraction.descricao} "
        f"Tempo médio de visitação: {attraction.tempo_horas:g} h "
        f"(neste roteiro: {elapsed_h:.1f} h)."
        f"{_valor_clause(attraction)}"
    )


def _lunch_detail(attraction: Attraction, profile: TouristProfile) -> str:
    return (
        f"Pausa para almoço (1h30min, entre 12h00 e 13h30). "
        f"Cidade: {_city_label(attraction)}. {attraction.descricao}"
        f"{_valor_clause(attraction)}"
    )


def _lodging_options_summary(attractions: list[Attraction]) -> str:
    if not attractions:
        return "—"
    lines = ["Sugestões para pernoite — não entram na programação de visitas do dia."]
    for i, attraction in enumerate(attractions, 1):
        maps_part = _maps_plain_suffix(attraction.maps_url)
        lines.append(
            f"Opção {i}: {_visit_title(attraction)}. {attraction.descricao}{_valor_clause(attraction)}{maps_part}"
        )
    return "\n".join(lines)


def _evening_dining_detail(attraction: Attraction) -> str:
    return (
        f"Sugestão para o jantar. Cidade: {_city_label(attraction)}. "
        f"{attraction.descricao}{_valor_clause(attraction)}"
    )


def _insert_evening_dining_timeline_blocks(
    blocks: list[ScheduleBlock],
    insert_at: int,
    day_anchor: datetime,
    suggestions: list[Attraction],
) -> None:
    if not suggestions:
        return
    marker = _combine(day_anchor, EVENING_DINING_START)
    segment: list[ScheduleBlock] = [
        ScheduleBlock(
            start=marker,
            end=marker,
            title=EVENING_DINING_TIMELINE_TITLE,
            detail="Opções abaixo — sem deslocamento entre as sugestões.",
            kind="evening_dining",
        )
    ]
    for index, attraction in enumerate(suggestions, start=1):
        segment.append(
            ScheduleBlock(
                start=marker,
                end=marker,
                title=f"Opção {index} — {_visit_title(attraction)}",
                detail=_evening_dining_detail(attraction),
                kind="evening_dining",
                maps_url=attraction.maps_url,
                image_url=attraction.image_url,
            )
        )
    for offset, block in enumerate(segment):
        blocks.insert(insert_at + offset, block)


def _fmt_time(dt: datetime) -> str:
    return dt.strftime("%H:%M")


def _schedule_visit(
    cursor: datetime,
    attraction: Attraction,
    day_anchor: datetime,
    *,
    phase_limit: datetime | None = None,
    reserve_after: int = 0,
) -> tuple[datetime, datetime] | None:
    day_limit = _combine(day_anchor, time(DAY_END_HOUR, 0))
    limit = phase_limit or day_limit
    hours = resolved_hours(attraction, day_anchor.date())
    if not hours:
        return None
    open_dt = _combine(day_anchor, hours[0])
    close_dt = _combine(day_anchor, hours[1])

    start = max(cursor, open_dt)
    if start >= min(close_dt, day_limit, limit):
        return None

    natural_end = min(
        start + timedelta(hours=attraction.tempo_horas),
        close_dt,
        day_limit,
        limit,
    )
    end = natural_end
    own_minutes = max(20, attraction.tempo_horas * 60)
    minimum_minutes = min(MIN_VISIT_MINUTES, own_minutes)
    for slots in range(reserve_after, 0, -1):
        latest = limit - timedelta(minutes=slots * (MIN_VISIT_MINUTES + 15))
        if latest <= start:
            continue
        shortened = min(natural_end, latest)
        if (shortened - start).total_seconds() >= minimum_minutes * 60:
            end = shortened
            break
    if (end - start).total_seconds() < minimum_minutes * 60:
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
    reserve_after: int = 0,
    backup: list[Attraction] | None = None,
) -> tuple[datetime, Attraction | None]:
    plog, cohen = scores.media_plog, scores.media_cohen
    reference = geo.coords(last_attr) if last_attr else day_start

    def _collect(source: list[Attraction]) -> list[tuple[Attraction, datetime, datetime, int]]:
        found: list[tuple[Attraction, datetime, datetime, int]] = []
        for attraction in source:
            if attraction.id in used_ids:
                continue
            if last_attr is None:
                hours = resolved_hours(attraction, day_anchor.date())
                open_t = hours[0] if hours else time(23, 59)
                # A primeira visita do dia começa de manhã, mesmo que abra depois das 9h30.
                if not hours or open_t >= time(12, 0):
                    continue
                transit_min = geo.travel_minutes_from_reference(day_start, attraction)
                cursor_after_transit = _combine(
                    day_anchor,
                    max(DAY_LATEST_FIRST_START, open_t),
                )
            else:
                transit_min = geo.travel_minutes(last_attr, attraction)
                cursor_after_transit = cursor + timedelta(minutes=transit_min)
            if transit_min > MAX_TRANSIT_MINUTES:
                continue
            if cursor_after_transit >= time_limit:
                continue
            slot = _schedule_visit(
                cursor_after_transit,
                attraction,
                day_anchor,
                phase_limit=time_limit,
                reserve_after=reserve_after,
            )
            if not slot:
                continue
            start, end = slot
            if end > time_limit:
                continue
            found.append((attraction, start, end, transit_min))
        return found

    candidates = _collect(pool)
    if not candidates and backup:
        candidates = _collect(backup)

    if not candidates:
        return cursor, None

    def _candidate_order(item: tuple[Attraction, datetime, datetime, int]) -> tuple[int, time, float]:
        attraction = item[0]
        km = _geo_sort_key(attraction, reference, geo, plog, cohen)[0]
        hours = resolved_hours(attraction, day_anchor.date())
        close = hours[1] if hours else time(23, 59)
        # Mesma faixa de 3 km: visita primeiro quem fecha mais cedo.
        # A primeira do dia prefere quem já está aberto às 9h30.
        starts_after_morning = 0
        if last_attr is None and item[1].time() > DAY_LATEST_FIRST_START:
            starts_after_morning = 1
        return (starts_after_morning, int(km // 3), item[1].time(), close, km)

    candidates.sort(key=_candidate_order)
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
            detail=_visit_detail(attraction, start, end) + _fora_clause(attraction, scores),
            kind="visit",
            maps_url=attraction.maps_url,
            image_url=attraction.image_url,
        )
    )
    used_ids.add(attraction.id)
    scheduled.append(attraction)
    return end, attraction


def build_itinerary(
    profile: TouristProfile,
    inventory: list[Attraction],
    scores: GroupScores | None = None,
    pinned_ids: list[str] | None = None,
) -> ItineraryResult:
    scores = scores or GroupScores.from_profile(profile)
    selected = filter_attractions(profile, inventory, scores)
    in_profile_visits = _visit_pool(selected)
    eligible_visits = _visit_pool(_eligible_attractions(profile, inventory, scores))
    pinned_lunch_ids: set[str] | None = None
    if pinned_ids is not None:
        by_pinned = {item.id: item for item in inventory}
        pinned = [by_pinned[item_id] for item_id in pinned_ids if item_id in by_pinned]
        in_profile_visits = [
            item for item in pinned if not _is_gastronomia(item) and not _is_hospedagem(item)
        ]
        eligible_visits = list(in_profile_visits)
        pinned_lunch_ids = {item.id for item in pinned if _is_gastronomia(item)}
    dates = trip_dates(profile)
    if not dates:
        dates = [_day_anchor(day).date() for day in range(1, max(1, profile.duracao_dias) + 1)]
    max_visits = _max_visits_for_ritmo(profile.ritmo)
    in_profile_ids = {attraction.id for attraction in in_profile_visits}
    outside_visits = [attraction for attraction in eligible_visits if attraction.id not in in_profile_ids]
    geo = GeoContext(inventory)

    days: list[DayPlan] = []
    used_global: set[str] = set()
    used_evening: set[str] = set()
    trip_lodging = _pick_lodging_suggestions(profile, inventory, scores)

    for day, service_date in enumerate(dates, start=1):
        day_anchor = datetime.combine(
            service_date,
            time(DAY_LATEST_FIRST_START.hour, DAY_LATEST_FIRST_START.minute),
        )
        day_pool, cluster = _day_visit_pool(
            service_date,
            in_profile_visits,
            outside_visits,
            used_global,
            max_visits,
            geo,
            scores,
        )
        day_start = cluster or geo.day_start_reference(day_pool, profile.cidades)
        pool_ids = {attraction.id for attraction in day_pool}
        backup = [
            attraction
            for attraction in [*in_profile_visits, *outside_visits]
            if attraction.id not in pool_ids
            and attraction.id not in used_global
            and resolved_hours(attraction, service_date)
        ]
        blocks: list[ScheduleBlock] = []
        scheduled_today: list[Attraction] = []
        cursor = day_anchor
        lunch_cutoff = _combine(day_anchor, time(12, 0))

        cursor, visit_count, morning_last = _schedule_visits_phase(
            cursor,
            pool=day_pool,
            day_anchor=day_anchor,
            used_ids=used_global,
            blocks=blocks,
            scheduled=scheduled_today,
            max_visits=_morning_visit_goal(max_visits),
            visits_so_far=0,
            stop_before=lunch_cutoff,
            geo=geo,
            scores=scores,
            day_start=day_start,
            backup=backup,
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
            only_ids=pinned_lunch_ids,
        )
        if lunch_attr:
            used_global.add(lunch_attr.id)
            scheduled_today.append(lunch_attr)
            lunch_title = f"Almoço — {_visit_title(lunch_attr)}"
            lunch_detail = _lunch_detail(lunch_attr, profile) + _fora_clause(lunch_attr, scores)
        else:
            lunch_title = "Pausa para Almoço"
            lunch_detail = (
                "Pausa de 1h30min entre 12h00 e 13h30. "
                "Nenhuma gastronomia em funcionamento a até 40 minutos de deslocamento neste dia."
            )
        blocks.append(
            ScheduleBlock(
                start=lunch_start,
                end=lunch_end,
                title=lunch_title,
                detail=lunch_detail,
                kind="lunch",
                maps_url=lunch_attr.maps_url if lunch_attr else "",
                image_url=lunch_attr.image_url if lunch_attr else "",
            )
        )
        cursor = lunch_end
        afternoon_start = geo.coords(lunch_attr) if lunch_attr else lunch_near

        cursor, visit_count, _ = _schedule_visits_phase(
            cursor,
            pool=day_pool,
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
            start_from=lunch_attr or morning_last,
            backup=backup,
        )

        theme = _theme_for_day(scheduled_today)
        day_end = _combine(day_anchor, time(DAY_END_HOUR, 0))
        close_at = min(cursor, day_end)
        blocks.append(
            ScheduleBlock(
                start=close_at,
                end=close_at,
                title="Encerramento do Dia",
                detail="Fim das atividades oficiais do roteiro, com conclusão até as 18h00.",
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

        evening_dining = _pick_evening_gastronomy_suggestions(
            profile,
            inventory,
            scores,
            used_global | used_evening,
            [service_date],
        )
        if evening_dining:
            insert_at = len(blocks)
            for index, block in enumerate(blocks):
                if block.kind == "lodging":
                    insert_at = index + 1
                    break
                if block.kind == "end":
                    insert_at = index + 1
            _insert_evening_dining_timeline_blocks(
                blocks, insert_at, day_anchor, evening_dining
            )
            used_evening.update(a.id for a in evening_dining)

        days.append(DayPlan(day_number=day, theme=theme, blocks=blocks, service_date=day_anchor))

    by_id = {a.id: a for a in inventory}
    used_attractions = [
        by_id[item_id]
        for item_id in used_global
        if item_id in by_id and not _is_hospedagem(by_id[item_id])
    ]
    outside = [a for a in used_attractions if not _within_profile(a, scores)]
    outside_names = [_visit_title(a) for a in outside]
    weights = [
        (a.peso_plog, a.peso_cohen) for a in used_attractions if _within_profile(a, scores)
    ]
    justification = build_technical_analysis(
        scores,
        weights,
        fora_do_perfil=bool(outside),
        atrativos_fora=outside_names,
    )

    return ItineraryResult(
        profile=profile,
        scores=scores,
        selected_attractions=used_attractions or selected,
        days=days,
        justification=justification,
        media_plog=scores.media_plog,
        media_cohen=scores.media_cohen,
        lodging_suggestions=trip_lodging,
        fora_do_perfil=bool(outside),
        atrativos_fora_do_perfil=outside_names,
    )


def _block_detail_markdown(block: ScheduleBlock) -> str:
    detail = block.detail
    if block.maps_url:
        line = maps_markdown_line(block.maps_url)
        if line:
            detail = f"{detail} {line}"
    return detail


_WEEKDAY_PT = (
    "segunda-feira",
    "terça-feira",
    "quarta-feira",
    "quinta-feira",
    "sexta-feira",
    "sábado",
    "domingo",
)


def day_heading(day: DayPlan) -> str:
    when = ""
    if day.service_date:
        holiday = " (feriado)" if is_brazilian_holiday(day.service_date.date()) else ""
        when = (
            f" — {_WEEKDAY_PT[day.service_date.weekday()]}, "
            f"{day.service_date.strftime('%d/%m/%Y')}{holiday}"
        )
    return f"DIA {day.day_number}{when} — {day.theme}"


def itinerary_to_markdown(result: ItineraryResult) -> str:
    p = result.profile
    s = result.scores
    lines = [
        f"# Roteiro de Viagem — {p.nome_grupo}",
        f"Duração: {p.duracao_dias} dia(s)"
        + (
            f" ({p.check_in.strftime('%d/%m/%Y')} a {p.check_out.strftime('%d/%m/%Y')})"
            if p.check_in and p.check_out
            else ""
        ),
        f"Média Plog: {s.media_plog:.2f} | Média Cohen: {s.media_cohen:.2f} | Fator: {s.fator_restritivo}",
        "",
    ]
    lines.extend(
        [
            "### Quadro-resumo — Sugestões de hospedagem (pernoite)",
            _lodging_options_summary(result.lodging_suggestions),
            "",
        ]
    )
    for day in result.days:
        lines.append(f"## {day_heading(day)}")
        for block in day.blocks:
            if block.kind == "lodging" and block.start == block.end:
                lines.append(f"**Noite — {block.title}**")
            elif block.kind == "evening_dining" and block.start == block.end:
                detail = _block_detail_markdown(block) if block.detail else ""
                suffix = f" — {detail}" if detail else ""
                lines.append(f"**{_fmt_time(block.start)} — {block.title}**{suffix}")
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

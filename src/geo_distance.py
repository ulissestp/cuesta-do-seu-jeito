from __future__ import annotations

import math
import re
from dataclasses import dataclass

from src.excel_io import Attraction, canonical_city_name
from src.maps_links import clean_maps_url, expand_maps_url, maps_url_has_embedded_coords

# Centróides aproximados (lat, lng) para fallback quando o link não traz coordenadas.
CITY_CENTROIDS: dict[str, tuple[float, float]] = {
    "Botucatu": (-22.8858, -48.4449),
    "Pardinho": (-23.0811, -48.3733),
    "Bofete": (-23.2836, -48.6847),
}

_COORD_AT_PATTERN = re.compile(r"@(-?\d+\.\d+),(-?\d+\.\d+)")
_COORD_BANG_PATTERN = re.compile(r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)")
_COORD_Q_PATTERN = re.compile(r"[?&]q=(-?\d+\.\d+),(-?\d+\.\d+)")


@dataclass(frozen=True)
class LatLng:
    lat: float
    lng: float


def _parse_coords_from_url(url: str) -> LatLng | None:
    text = url or ""
    for pattern in (_COORD_AT_PATTERN, _COORD_BANG_PATTERN, _COORD_Q_PATTERN):
        match = pattern.search(text)
        if match:
            return LatLng(float(match.group(1)), float(match.group(2)))
    return None


def coords_for_maps_url(maps_url: str) -> LatLng | None:
    """Coordenadas na URL ou após expandir link curto do Google Maps (com cache)."""
    url = clean_maps_url(maps_url)
    if not url:
        return None
    point = _parse_coords_from_url(url)
    if point:
        return point
    if maps_url_has_embedded_coords(url):
        return None
    expanded = expand_maps_url(url)
    if expanded and expanded != url:
        point = _parse_coords_from_url(expanded)
        if point:
            return point
    return _parse_coords_from_url(expanded)


def coords_for_city_label(cidade: str) -> LatLng | None:
    canon = canonical_city_name(cidade)
    if not canon:
        return None
    lat, lng = CITY_CENTROIDS.get(canon, (None, None))
    if lat is None:
        return None
    return LatLng(lat, lng)


def coords_for_attraction(attraction: Attraction) -> LatLng | None:
    from_maps = coords_for_maps_url(attraction.maps_url)
    if from_maps:
        return from_maps
    return coords_for_city_label(attraction.cidade)


def haversine_km(a: LatLng, b: LatLng) -> float:
    r = 6371.0
    lat1, lon1 = math.radians(a.lat), math.radians(a.lng)
    lat2, lon2 = math.radians(b.lat), math.radians(b.lng)
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    h = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return 2 * r * math.asin(min(1.0, math.sqrt(h)))


def travel_minutes_between(
    origin: LatLng | None,
    destination: LatLng | None,
    *,
    same_city_hint: bool | None = None,
) -> int:
    """Tempo de deslocamento a partir da distância (km) entre os pontos."""
    if origin and destination:
        km = haversine_km(origin, destination)
        minutes = int(km / 45.0 * 60.0 + 10)
        if km < 0.02:
            return 15
        if km < 0.3:
            return max(10, minutes)
        return max(15, min(120, minutes))
    if same_city_hint is True:
        return 15
    if same_city_hint is False:
        return 35
    return 25


def same_city(a: Attraction, b: Attraction) -> bool | None:
    ca = canonical_city_name(a.cidade)
    cb = canonical_city_name(b.cidade)
    if ca and cb:
        return ca == cb
    return None


class GeoContext:
    """Coordenadas resolvidas por atração (cache por id, sob demanda)."""

    def __init__(self, attractions: list[Attraction]) -> None:
        self._by_id: dict[str, LatLng | None] = {}

    def coords(self, attraction: Attraction) -> LatLng | None:
        if attraction.id in self._by_id:
            return self._by_id[attraction.id]
        point = coords_for_attraction(attraction)
        self._by_id[attraction.id] = point
        return point

    def coords_resolved(self, attraction: Attraction) -> LatLng | None:
        point = self.coords(attraction)
        if point:
            return point
        return coords_for_city_label(attraction.cidade)

    def distance_km(self, a: Attraction, b: Attraction) -> float | None:
        pa, pb = self.coords_resolved(a), self.coords_resolved(b)
        if pa and pb:
            return haversine_km(pa, pb)
        return None

    def travel_minutes_from_reference(self, origin: LatLng | None, to_attr: Attraction) -> int:
        pb = self.coords_resolved(to_attr)
        if origin and pb:
            return travel_minutes_between(origin, pb)
        return 25

    def travel_minutes(self, from_attr: Attraction | None, to_attr: Attraction) -> int:
        if from_attr is None:
            return self.travel_minutes_from_reference(None, to_attr)
        pa, pb = self.coords_resolved(from_attr), self.coords_resolved(to_attr)
        if pa and pb:
            return travel_minutes_between(pa, pb)
        hint = same_city(from_attr, to_attr) if from_attr else None
        return travel_minutes_between(pa, pb, same_city_hint=hint)

    def day_start_reference(self, pool: list[Attraction], profile_cities: list[str]) -> LatLng | None:
        """Ponto de partida do dia: centróide das cidades do perfil ou das atrações disponíveis."""
        for city in profile_cities:
            point = coords_for_city_label(city)
            if point:
                return point
        points = [self.coords(a) for a in pool if self.coords(a)]
        if not points:
            return None
        lat = sum(p.lat for p in points) / len(points)
        lng = sum(p.lng for p in points) / len(points)
        return LatLng(lat, lng)

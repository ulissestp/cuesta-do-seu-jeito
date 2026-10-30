from __future__ import annotations

import io
import math
import os
import urllib.error
import urllib.parse
import urllib.request
from functools import lru_cache

from PIL import Image, ImageDraw

from src.branding import APP_VERSION_SHORT
from src.geo_distance import LatLng, coords_for_city_label, coords_for_maps_url

USER_AGENT = f"RoteiroTurismo/{APP_VERSION_SHORT} (PDF map thumbnail)"
_MAP_SIZE = "320x200"
_MAP_ZOOM = 14
_TILE_SIZE = 256
_OSM_TILE = "https://tile.openstreetmap.org/{z}/{x}/{y}.png"


def google_maps_static_api_key() -> str:
    env = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()
    if env:
        return env
    for path in (writable_root() / "google_maps_api_key.txt", data_dir() / "google_maps_api_key.txt"):
        if not path.is_file():
            continue
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                return line
    return ""


def coords_for_schedule_block(*, maps_url: str, title: str) -> LatLng | None:
    point = coords_for_maps_url(maps_url)
    if point:
        return point
    if "—" in title:
        city = title.split("—")[-1].strip()
        return coords_for_city_label(city)
    return None


def _google_static_map_url(lat: float, lng: float, *, api_key: str) -> str:
    params = urllib.parse.urlencode(
        {
            "center": f"{lat:.6f},{lng:.6f}",
            "zoom": str(_MAP_ZOOM),
            "size": _MAP_SIZE,
            "scale": "2",
            "maptype": "roadmap",
            "markers": f"color:red|size:mid|{lat:.6f},{lng:.6f}",
            "key": api_key,
        }
    )
    return f"https://maps.googleapis.com/maps/api/staticmap?{params}"


def _osm_static_map_url(lat: float, lng: float) -> str:
    params = urllib.parse.urlencode(
        {
            "center": f"{lat:.6f},{lng:.6f}",
            "zoom": str(_MAP_ZOOM),
            "size": "320x200",
            "maptype": "mapnik",
            "markers": f"{lat:.6f},{lng:.6f},red-pushpin",
        }
    )
    return f"https://staticmap.openstreetmap.de/staticmap.php?{params}"


def _fetch_bytes(url: str) -> bytes | None:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = resp.read()
        if len(data) < 500:
            return None
        if data[:8].startswith(b"{"):
            return None
        return data
    except (urllib.error.URLError, TimeoutError, OSError, ValueError):
        return None


def _latlng_to_world_pixel(lat: float, lng: float, zoom: int) -> tuple[float, float]:
    n = 2**zoom
    x = (lng + 180.0) / 360.0 * n * _TILE_SIZE
    lat_rad = math.radians(lat)
    y = (1.0 - math.asinh(math.tan(lat_rad)) / math.pi) / 2.0 * n * _TILE_SIZE
    return x, y


def _fetch_osm_tile(zoom: int, x: int, y: int) -> Image.Image | None:
    url = _OSM_TILE.format(z=zoom, x=x, y=y)
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return Image.open(io.BytesIO(resp.read())).convert("RGB")
    except (urllib.error.URLError, TimeoutError, OSError, ValueError):
        return None


def _render_osm_tile_thumbnail(lat: float, lng: float, *, width: int = 320, height: int = 200) -> bytes | None:
    zoom = _MAP_ZOOM
    wx, wy = _latlng_to_world_pixel(lat, lng, zoom)
    half_w, half_h = width / 2, height / 2
    min_wx, max_wx = wx - half_w, wx + half_w
    min_wy, max_wy = wy - half_h, wy + half_h
    tile_x0 = int(math.floor(min_wx / _TILE_SIZE))
    tile_x1 = int(math.floor(max_wx / _TILE_SIZE))
    tile_y0 = int(math.floor(min_wy / _TILE_SIZE))
    tile_y1 = int(math.floor(max_wy / _TILE_SIZE))

    map_w = (tile_x1 - tile_x0 + 1) * _TILE_SIZE
    map_h = (tile_y1 - tile_y0 + 1) * _TILE_SIZE
    if map_w <= 0 or map_h <= 0:
        return None

    canvas = Image.new("RGB", (map_w, map_h), (236, 240, 244))
    ok = 0
    total = (tile_x1 - tile_x0 + 1) * (tile_y1 - tile_y0 + 1)
    for tx in range(tile_x0, tile_x1 + 1):
        for ty in range(tile_y0, tile_y1 + 1):
            tile = _fetch_osm_tile(zoom, tx, ty)
            if tile is None:
                continue
            ok += 1
            canvas.paste(tile, ((tx - tile_x0) * _TILE_SIZE, (ty - tile_y0) * _TILE_SIZE))
    if ok < max(1, total // 3):
        return None

    px = wx - tile_x0 * _TILE_SIZE
    py = wy - tile_y0 * _TILE_SIZE
    left = max(0, int(px - half_w))
    top = max(0, int(py - half_h))
    crop = canvas.crop((left, top, left + width, top + height))
    draw = ImageDraw.Draw(crop)
    mx, my = int(px - left), int(py - top)
    draw.ellipse((mx - 7, my - 7, mx + 7, my + 7), fill=(229, 62, 62), outline=(255, 255, 255), width=2)
    buf = io.BytesIO()
    crop.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


@lru_cache(maxsize=256)
def fetch_map_thumbnail_bytes(lat: float, lng: float) -> bytes | None:
    """Miniatura de mapa: Google Static Maps (com chave) ou mapa OSM na coordenada."""
    key = google_maps_static_api_key()
    if key:
        data = _fetch_bytes(_google_static_map_url(lat, lng, api_key=key))
        if data:
            return data
    data = _fetch_bytes(_osm_static_map_url(lat, lng))
    if data:
        return data
    return _render_osm_tile_thumbnail(lat, lng)

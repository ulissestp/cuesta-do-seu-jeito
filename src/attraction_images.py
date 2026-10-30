from __future__ import annotations

import io

from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Image

from src.maps_static_thumbnail import coords_for_schedule_block, fetch_map_thumbnail_bytes


def attraction_thumbnail_flowable(
    *,
    maps_url: str = "",
    place_title: str = "",
    max_width: float | None = None,
    max_height: float | None = None,
) -> Image | None:
    """Miniatura de mapa (Google Static Maps ou OSM) abaixo do horário no PDF."""
    max_width = max_width or 2.35 * cm
    max_height = max_height or 1.75 * cm
    point = coords_for_schedule_block(maps_url=maps_url, title=place_title)
    if not point:
        return None
    thumb = fetch_map_thumbnail_bytes(round(point.lat, 5), round(point.lng, 5))
    if not thumb:
        return None
    try:
        reader = ImageReader(io.BytesIO(thumb))
        iw, ih = reader.getSize()
        if iw <= 0 or ih <= 0:
            return None
        scale = min(max_width / iw, max_height / ih)
        w, h = iw * scale, ih * scale
        return Image(io.BytesIO(thumb), width=w, height=h)
    except (OSError, ValueError, ZeroDivisionError):
        return None

from __future__ import annotations

import io
import urllib.error
import urllib.parse
import urllib.request
from functools import lru_cache

from PIL import Image as PILImage
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Image

from src.branding import APP_VERSION_SHORT
from src.maps_static_thumbnail import coords_for_schedule_block, fetch_map_thumbnail_bytes

USER_AGENT = f"RoteiroTurismo/{APP_VERSION_SHORT} (attraction photo)"
_MAX_IMAGE_BYTES = 8_000_000


def _larger_google_thumb(url: str) -> str:
    """Links colados da busca de imagens do Google costumam vir com s=10 (10 px)."""
    parsed = urllib.parse.urlparse(url)
    if "gstatic.com" not in parsed.netloc.lower():
        return url
    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    changed = False
    updated: list[tuple[str, str]] = []
    for key, value in query:
        if key == "s" and value.isdigit() and int(value) < 400:
            updated.append((key, "480"))
            changed = True
        else:
            updated.append((key, value))
    if not changed:
        return url
    return urllib.parse.urlunparse(parsed._replace(query=urllib.parse.urlencode(updated)))


@lru_cache(maxsize=128)
def _download_image_bytes(url: str) -> bytes | None:
    request = urllib.request.Request(
        _larger_google_thumb(url),
        headers={"User-Agent": USER_AGENT},
    )
    try:
        with urllib.request.urlopen(request, timeout=15) as response:
            data = response.read(_MAX_IMAGE_BYTES + 1)
    except (urllib.error.URLError, TimeoutError, OSError, ValueError):
        return None
    if not data or len(data) > _MAX_IMAGE_BYTES:
        return None
    return data


def _jpeg_bytes(raw: bytes) -> bytes | None:
    try:
        with PILImage.open(io.BytesIO(raw)) as photo:
            frame = photo.convert("RGB")
            buffer = io.BytesIO()
            frame.save(buffer, format="JPEG", quality=85)
            return buffer.getvalue()
    except (OSError, ValueError):
        return None


def _flowable_from_bytes(
    raw: bytes,
    *,
    max_width: float,
    max_height: float,
) -> Image | None:
    jpeg = _jpeg_bytes(raw)
    if not jpeg:
        return None
    try:
        reader = ImageReader(io.BytesIO(jpeg))
        width, height = reader.getSize()
        if width <= 0 or height <= 0:
            return None
        scale = min(max_width / width, max_height / height)
        return Image(io.BytesIO(jpeg), width=width * scale, height=height * scale)
    except (OSError, ValueError, ZeroDivisionError):
        return None


def _map_thumbnail_flowable(
    *,
    maps_url: str,
    place_title: str,
    max_width: float,
    max_height: float,
) -> Image | None:
    point = coords_for_schedule_block(maps_url=maps_url, title=place_title)
    if not point:
        return None
    thumb = fetch_map_thumbnail_bytes(round(point.lat, 5), round(point.lng, 5))
    if not thumb:
        return None
    return _flowable_from_bytes(thumb, max_width=max_width, max_height=max_height)


def attraction_thumbnail_flowable(
    *,
    maps_url: str = "",
    place_title: str = "",
    image_url: str = "",
    max_width: float | None = None,
    max_height: float | None = None,
) -> Image | None:
    """Foto da coluna Imagem abaixo do horário. Sem link, usa a miniatura de mapa."""
    max_width = max_width or 2.35 * cm
    max_height = max_height or 1.75 * cm
    url = (image_url or "").strip()
    if url.lower().startswith(("http://", "https://")):
        photo = _download_image_bytes(url)
        if photo:
            flowable = _flowable_from_bytes(photo, max_width=max_width, max_height=max_height)
            if flowable:
                return flowable
    return _map_thumbnail_flowable(
        maps_url=maps_url,
        place_title=place_title,
        max_width=max_width,
        max_height=max_height,
    )

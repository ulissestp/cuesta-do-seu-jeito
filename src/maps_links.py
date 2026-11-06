from __future__ import annotations

import json
import re
import threading
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from typing import Any
from xml.sax.saxutils import escape

from src.app_paths import output_dir
from src.branding import APP_VERSION_SHORT

_MAPS_URL_RE = re.compile(r"^https?://", re.I)
_MAPS_COORD_IN_URL = re.compile(
    r"@(-?\d+\.\d+),(-?\d+\.\d+)|!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)|[?&]q=(-?\d+\.\d+),(-?\d+\.\d+)"
)
_MAPS_EXPAND_USER_AGENT = f"RoteiroTurismo/{APP_VERSION_SHORT} (itinerary distance)"
_EXPAND_TIMEOUT = 5
_MAX_REDIRECTS = 4
_cache_lock = threading.Lock()
_expand_cache: dict[str, str] | None = None
_expand_cache_dirty = False
_expand_misses: set[str] = set()


def clean_maps_url(raw: Any) -> str:
    if raw is None:
        return ""
    text = str(raw).strip()
    if not text or text.lower() in ("none", "—", "-"):
        return ""
    if text.startswith("="):
        return ""
    if not _MAPS_URL_RE.match(text):
        if "google.com/maps" in text or "maps.app.goo.gl" in text or "goo.gl/" in text:
            if not text.startswith("http"):
                text = "https://" + text.lstrip("/")
        else:
            return ""
    return text


def _already_short(url: str) -> bool:
    lower = url.lower()
    return "maps.app.goo.gl" in lower or "goo.gl/" in lower


def maps_url_has_embedded_coords(url: str) -> bool:
    text = clean_maps_url(url)
    return bool(text and _MAPS_COORD_IN_URL.search(text))


def maps_url_needs_redirect_expand(url: str) -> bool:
    text = clean_maps_url(url)
    if not text or maps_url_has_embedded_coords(text):
        return False
    lower = text.lower()
    return (
        "maps.app.goo.gl" in lower
        or "goo.gl/" in lower
        or "google.com/maps" in lower
        or "maps.google.com" in lower
    )


def _expand_cache_path():
    return output_dir() / "maps_expand_cache.json"


def _load_expand_cache() -> dict[str, str]:
    global _expand_cache
    with _cache_lock:
        if _expand_cache is None:
            loaded: dict[str, str] = {}
            path = _expand_cache_path()
            if path.is_file():
                try:
                    raw = json.loads(path.read_text(encoding="utf-8"))
                    if isinstance(raw, dict):
                        loaded = {str(key): str(value) for key, value in raw.items() if value}
                except (OSError, json.JSONDecodeError, TypeError):
                    loaded = {}
            _expand_cache = loaded
        return _expand_cache


def _remember_expand(url: str, resolved: str) -> None:
    global _expand_cache_dirty
    if not resolved or resolved == url or not maps_url_has_embedded_coords(resolved):
        return
    cache = _load_expand_cache()
    with _cache_lock:
        if cache.get(url) == resolved:
            return
        cache[url] = resolved
        _expand_cache_dirty = True


def flush_maps_expand_cache() -> None:
    global _expand_cache_dirty
    cache = _load_expand_cache()
    with _cache_lock:
        if not _expand_cache_dirty:
            return
        path = _expand_cache_path()
        path.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
        _expand_cache_dirty = False


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def _follow_until_coords(url: str) -> str:
    """Para no primeiro redirecionamento que já traz latitude e longitude."""
    opener = urllib.request.build_opener(_NoRedirect)
    current = url
    for _ in range(_MAX_REDIRECTS):
        if maps_url_has_embedded_coords(current):
            return current
        request = urllib.request.Request(current, headers={"User-Agent": _MAPS_EXPAND_USER_AGENT})
        try:
            with opener.open(request, timeout=_EXPAND_TIMEOUT) as response:
                return response.geturl() or current
        except urllib.error.HTTPError as exc:
            if exc.code not in (301, 302, 303, 307, 308):
                return current
            location = exc.headers.get("Location")
            if not location:
                return current
            current = urllib.parse.urljoin(current, location.strip())
        except (urllib.error.URLError, TimeoutError, ValueError):
            return current
    return current


def expand_maps_url(url: str) -> str:
    """Segue redirecionamento de links curtos do Google Maps para obter URL com coordenadas."""
    cleaned = clean_maps_url(url)
    if not cleaned:
        return ""
    if maps_url_has_embedded_coords(cleaned) or not maps_url_needs_redirect_expand(cleaned):
        return cleaned
    if cleaned in _expand_misses:
        return cleaned
    cached = _load_expand_cache().get(cleaned)
    if cached:
        return cached
    resolved = _follow_until_coords(cleaned)
    if maps_url_has_embedded_coords(resolved):
        _remember_expand(cleaned, resolved)
    else:
        _expand_misses.add(cleaned)
    return resolved or cleaned


def prefetch_maps_urls(urls: list[str], *, workers: int = 8) -> None:
    """Resolve vários links curtos ao mesmo tempo e grava o cache em disco."""
    pending: list[str] = []
    seen: set[str] = set()
    cache = _load_expand_cache()
    for raw in urls:
        cleaned = clean_maps_url(raw)
        if (
            not cleaned
            or cleaned in seen
            or maps_url_has_embedded_coords(cleaned)
            or not maps_url_needs_redirect_expand(cleaned)
            or cleaned in cache
            or cleaned in _expand_misses
        ):
            continue
        seen.add(cleaned)
        pending.append(cleaned)
    if not pending:
        return
    with ThreadPoolExecutor(max_workers=min(workers, len(pending))) as pool:
        list(pool.map(expand_maps_url, pending, chunksize=1))
    flush_maps_expand_cache()


def try_shorten_maps_url(url: str, *, min_length: int = 55) -> str:
    """Encurta URLs longas do Maps (ex.: maps.google.com). Links goo.gl da planilha são mantidos."""
    url = clean_maps_url(url)
    if not url or len(url) <= min_length or _already_short(url):
        return url
    try:
        query = urllib.parse.urlencode({"format": "simple", "url": url})
        api = f"https://is.gd/create.php?{query}"
        with urllib.request.urlopen(api, timeout=8) as resp:
            short = resp.read().decode("utf-8", errors="replace").strip()
            if short.startswith("http"):
                return short
    except (urllib.error.URLError, TimeoutError, ValueError):
        pass
    return url


def maps_markdown_line(url: str) -> str:
    href = try_shorten_maps_url(url)
    if not href:
        return ""
    label = href if _already_short(href) else "Abrir no Google Maps"
    return f"Localização: [{label}]({href})"


def maps_pdf_html(url: str) -> str:
    href = try_shorten_maps_url(url)
    if not href:
        return ""
    label = escape(href if _already_short(href) else "Abrir no Google Maps")
    safe_href = escape(href, {"'": "&apos;", '"': "&quot;"})
    return (
        f'<br/><b>Google Maps:</b> '
        f'<font color="#2c5282"><u><a href="{safe_href}">{label}</a></u></font>'
    )

from __future__ import annotations

import re
import urllib.error
import urllib.parse
import urllib.request
from typing import Any
from xml.sax.saxutils import escape

_MAPS_URL_RE = re.compile(r"^https?://", re.I)


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

"""Verifica atualização do executável via GitHub Releases."""

from __future__ import annotations

import json
import re
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path

from src.app_paths import output_dir
from src.branding import APP_VERSION

GITHUB_REPO = "ulissestp/cuesta-do-seu-jeito"
RELEASE_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
_FETCH_TIMEOUT = 25
_USER_AGENT = "CUESTA-RT inventory-sync"
_EXE_DOWNLOAD_NAME = "RoteiroTurismo_atualizacao.exe"


@dataclass(frozen=True)
class AppUpdateResult:
    status: str
    message: str
    current_version: str
    latest_version: str | None = None
    release_url: str | None = None
    downloaded_exe: Path | None = None


def _http_get(url: str) -> bytes:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": _USER_AGENT,
            "Accept": "application/vnd.github+json",
        },
    )
    with urllib.request.urlopen(request, timeout=_FETCH_TIMEOUT) as response:
        return response.read()


def _parse_version(text: str) -> tuple[int, ...]:
    cleaned = re.sub(r"^[^0-9]*", "", (text or "").strip())
    parts: list[int] = []
    for piece in cleaned.split("."):
        digits = "".join(ch for ch in piece if ch.isdigit())
        if digits:
            parts.append(int(digits))
    return tuple(parts) if parts else (0,)


def _version_less(left: str, right: str) -> bool:
    a = _parse_version(left)
    b = _parse_version(right)
    return a < b


def _running_frozen() -> bool:
    return getattr(sys, "frozen", False)


def check_application_update(*, download_if_newer: bool = True) -> AppUpdateResult:
    current = APP_VERSION
    if not _running_frozen():
        return AppUpdateResult(
            status="dev",
            message=(
                f"Programa em modo desenvolvimento (versão {current}). "
                "Para atualizar o .exe, use a release mais recente no GitHub."
            ),
            current_version=current,
        )

    try:
        payload = json.loads(_http_get(RELEASE_API).decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, OSError, json.JSONDecodeError) as exc:
        return AppUpdateResult(
            status="offline",
            message=f"Não foi possível verificar atualização do programa: {exc}",
            current_version=current,
        )

    tag = str(payload.get("tag_name") or "").strip()
    latest = tag.lstrip("vV") or str(payload.get("name") or "").strip()
    release_url = str(payload.get("html_url") or "").strip() or None
    if not latest:
        return AppUpdateResult(
            status="error",
            message="Release no GitHub sem número de versão.",
            current_version=current,
            release_url=release_url,
        )

    if not _version_less(current, latest):
        return AppUpdateResult(
            status="up_to_date",
            message=f"Programa atualizado (versão {current}).",
            current_version=current,
            latest_version=latest,
            release_url=release_url,
        )

    downloaded: Path | None = None
    if download_if_newer:
        for asset in payload.get("assets") or []:
            name = str(asset.get("name") or "")
            if name.lower().endswith(".exe"):
                url = str(asset.get("browser_download_url") or "")
                if url:
                    dest = output_dir() / _EXE_DOWNLOAD_NAME
                    dest.write_bytes(_http_get(url))
                    downloaded = dest
                    break

    if downloaded is not None:
        msg = (
            f"Há uma versão nova do programa ({latest}; você usa {current}).\n"
            f"Arquivo baixado:\n{downloaded}\n\n"
            "Feche o CUESTA, substitua o .exe antigo por este arquivo e abra de novo."
        )
    else:
        msg = (
            f"Há uma versão nova do programa ({latest}; você usa {current}).\n"
            f"Baixe em: {release_url or RELEASE_API.replace('api.', '')}"
        )

    return AppUpdateResult(
        status="update_available",
        message=msg,
        current_version=current,
        latest_version=latest,
        release_url=release_url,
        downloaded_exe=downloaded,
    )

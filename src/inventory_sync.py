"""Verifica e aplica atualizações do inventário turístico publicado no GitHub."""

from __future__ import annotations

import json
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from src.app_paths import data_dir, output_dir
from src.excel_io import (
    InventoryMergeStats,
    merge_inventories_deduped,
    read_inventory,
    read_inventory_from_uploaded,
    save_inventory,
)

_SYNC_STATE = "inventario_sync.json"
_CATALOG_CACHE = "catalogo_inventario.xlsx"
_FETCH_TIMEOUT = 20
_USER_AGENT = "CUESTA-RT/1.1 inventory-sync"


@dataclass(frozen=True)
class InventoryUpdateResult:
    status: str
    message: str
    remote_version: str | None = None
    local_version: str | None = None
    stats: InventoryMergeStats | None = None


def _sync_state_path() -> Path:
    return output_dir() / _SYNC_STATE


def _catalog_cache_path() -> Path:
    return output_dir() / _CATALOG_CACHE


def load_local_inventory_version() -> str | None:
    path = _sync_state_path()
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        version = data.get("version")
        return str(version) if version else None
    except (json.JSONDecodeError, OSError, TypeError):
        return None


def _save_local_inventory_version(version: str) -> None:
    payload = {
        "version": version,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }
    _sync_state_path().write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _bundled_manifest_path() -> Path:
    return data_dir() / "inventario_manifest.json"


def _http_get(url: str) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": _USER_AGENT})
    with urllib.request.urlopen(request, timeout=_FETCH_TIMEOUT) as response:
        return response.read()


def _load_manifest_from_bytes(raw: bytes) -> dict:
    data = json.loads(raw.decode("utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Manifesto de inventário inválido.")
    return data


def fetch_remote_manifest(manifest_url: str | None = None) -> dict:
    bundled = _bundled_manifest_path()
    url = manifest_url
    if not url and bundled.is_file():
        url = json.loads(bundled.read_text(encoding="utf-8")).get("manifest_url")
    if not url:
        raise ValueError("URL do manifesto de inventário não configurada.")
    return _load_manifest_from_bytes(_http_get(url))


def _download_inventory(url: str, dest: Path) -> None:
    dest.write_bytes(_http_get(url))


def _format_stats(stats: InventoryMergeStats) -> str:
    removed = stats.removed_duplicate_ids + stats.removed_duplicate_names
    parts = [f"{stats.after} atrativos na planilha"]
    if removed:
        parts.append(f"{removed} duplicata(s) removida(s)")
    return " · ".join(parts)


def check_and_apply_inventory_update(workbook_path: Path) -> InventoryUpdateResult:
    """
    Compara a versão local com o manifesto no GitHub.
    Se houver versão nova, baixa o catálogo, mescla na sessão e remove duplicatas.
    """
    local_version = load_local_inventory_version()
    try:
        remote_manifest = fetch_remote_manifest()
    except (urllib.error.URLError, TimeoutError, OSError, ValueError) as exc:
        return InventoryUpdateResult(
            status="offline",
            message=f"Sem conexão ou manifesto indisponível: {exc}",
            local_version=local_version,
        )

    remote_version = str(remote_manifest.get("version") or "").strip()
    inventory_url = str(remote_manifest.get("inventory_url") or "").strip()
    if not remote_version or not inventory_url:
        return InventoryUpdateResult(
            status="error",
            message="Manifesto remoto incompleto (version ou inventory_url).",
            local_version=local_version,
        )

    if local_version == remote_version:
        return InventoryUpdateResult(
            status="up_to_date",
            message=f"Inventário já está na versão {remote_version}.",
            remote_version=remote_version,
            local_version=local_version,
        )

    cache_path = _catalog_cache_path()
    try:
        _download_inventory(inventory_url, cache_path)
        incoming = read_inventory_from_uploaded(cache_path)
        if not incoming:
            raise ValueError("Catálogo remoto não contém atrativos válidos.")
    except (urllib.error.URLError, TimeoutError, OSError, ValueError) as exc:
        return InventoryUpdateResult(
            status="error",
            message=f"Falha ao baixar inventário: {exc}",
            remote_version=remote_version,
            local_version=local_version,
        )

    existing = read_inventory(workbook_path)
    merged, stats = merge_inventories_deduped(existing, incoming, prefer_incoming=True)
    save_inventory(workbook_path, merged)
    _save_local_inventory_version(remote_version)

    removed = stats.removed_duplicate_ids + stats.removed_duplicate_names
    detail = f"{len(existing)} → {len(merged)} atrativos"
    if removed:
        detail += f" ({removed} duplicata(s) removida(s))"

    return InventoryUpdateResult(
        status="updated",
        message=f"Inventário atualizado para a versão {remote_version}. {detail}.",
        remote_version=remote_version,
        local_version=remote_version,
        stats=stats,
    )


def apply_inventory_update_from_file(workbook_path: Path, catalog_path: Path) -> InventoryUpdateResult:
    """Mescla um .xlsx local (mesma lógica da atualização remota)."""
    incoming = read_inventory_from_uploaded(catalog_path)
    if not incoming:
        raise ValueError("Nenhuma atração válida no arquivo.")
    existing = read_inventory(workbook_path)
    merged, stats = merge_inventories_deduped(existing, incoming, prefer_incoming=True)
    save_inventory(workbook_path, merged)
    return InventoryUpdateResult(
        status="updated",
        message=f"Inventário mesclado. {_format_stats(stats)}.",
        stats=stats,
    )

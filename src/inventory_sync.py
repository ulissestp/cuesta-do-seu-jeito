"""Verificação manual de atualizações (programa + inventário)."""

from __future__ import annotations

import json
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from src.app_paths import data_dir, output_dir
from src.app_update import AppUpdateResult, check_application_update
from src.excel_io import (
    InventoryMergeStats,
    merge_inventories_incremental,
    read_inventory,
    read_inventory_from_uploaded,
    save_inventory,
)

_SYNC_STATE = "inventario_sync.json"
_CATALOG_CACHE = "catalogo_inventario.xlsx"
_FETCH_TIMEOUT = 20
_USER_AGENT = "CUESTA-RT inventory-sync"


@dataclass(frozen=True)
class InventoryUpdateResult:
    status: str
    message: str
    remote_version: str | None = None
    local_version: str | None = None
    stats: InventoryMergeStats | None = None


@dataclass(frozen=True)
class VerifyUpdatesResult:
    app: AppUpdateResult
    inventory: InventoryUpdateResult

    @property
    def summary(self) -> str:
        return f"{self.app.message}\n\n{self.inventory.message}"


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


def fetch_remote_manifest() -> dict:
    bundled = _bundled_manifest_path()
    url = None
    if bundled.is_file():
        url = json.loads(bundled.read_text(encoding="utf-8")).get("manifest_url")
    if not url:
        raise ValueError("URL do manifesto de inventário não configurada.")
    data = json.loads(_http_get(url).decode("utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Manifesto de inventário inválido.")
    return data


def _inventory_change_message(stats: InventoryMergeStats, remote_version: str) -> str:
    skipped = stats.skipped_duplicate_names
    if stats.added:
        parts = [f"catálogo {remote_version}", f"{stats.added} novo(s)"]
        if skipped:
            parts.append(f"{skipped} nome(s) já existente(s) não importado(s)")
        return f"Inventário atualizado ({', '.join(parts)}). Total: {stats.after} atrativos."
    if skipped:
        return (
            f"Inventário já alinhado ao catálogo {remote_version} ({stats.after} atrativos). "
            f"{skipped} nome(s) já existente(s) não importado(s)."
        )
    return f"Inventário já alinhado ao catálogo {remote_version} ({stats.after} atrativos)."


def apply_incremental_inventory_update(workbook_path: Path) -> InventoryUpdateResult:
    """Baixa o catálogo remoto, compara com a sessão e aplica só diferenças."""
    local_version = load_local_inventory_version()
    try:
        remote_manifest = fetch_remote_manifest()
    except (urllib.error.URLError, TimeoutError, OSError, ValueError) as exc:
        return InventoryUpdateResult(
            status="offline",
            message=f"Inventário: sem conexão ou manifesto indisponível ({exc}).",
            local_version=local_version,
        )

    remote_version = str(remote_manifest.get("version") or "").strip()
    inventory_url = str(remote_manifest.get("inventory_url") or "").strip()
    if not remote_version or not inventory_url:
        return InventoryUpdateResult(
            status="error",
            message="Inventário: manifesto remoto incompleto.",
            local_version=local_version,
        )

    cache_path = _catalog_cache_path()
    try:
        cache_path.write_bytes(_http_get(inventory_url))
        incoming = read_inventory_from_uploaded(cache_path)
        if not incoming:
            raise ValueError("Catálogo remoto vazio.")
    except (urllib.error.URLError, TimeoutError, OSError, ValueError) as exc:
        return InventoryUpdateResult(
            status="error",
            message=f"Inventário: falha ao baixar catálogo ({exc}).",
            remote_version=remote_version,
            local_version=local_version,
        )

    existing = read_inventory(workbook_path)
    merged, stats = merge_inventories_incremental(existing, incoming)
    changed = stats.added or stats.updated or stats.removed_duplicate_names
    if changed or len(merged) != len(existing):
        save_inventory(workbook_path, merged)
        _save_local_inventory_version(remote_version)
        status = "updated"
    else:
        _save_local_inventory_version(remote_version)
        status = "up_to_date"

    return InventoryUpdateResult(
        status=status,
        message=_inventory_change_message(stats, remote_version),
        remote_version=remote_version,
        local_version=remote_version,
        stats=stats,
    )


def verify_updates(workbook_path: Path) -> VerifyUpdatesResult:
    """Verifica .exe (GitHub Releases) e inventário (planilha remota) sob demanda."""
    app = check_application_update(download_if_newer=True)
    inventory = apply_incremental_inventory_update(workbook_path)
    return VerifyUpdatesResult(app=app, inventory=inventory)


# Compatibilidade com imports antigos
check_and_apply_inventory_update = apply_incremental_inventory_update

from __future__ import annotations

from dataclasses import dataclass

from .excel_repository import WorkbookFingerprint
from .models import Asset, ChangeRecord, DictionaryItem, StorageMedium


@dataclass(frozen=True, slots=True)
class CacheStatus:
    loaded_at: str = ""
    stale: bool = True
    error: str = ""


@dataclass(frozen=True, slots=True)
class LedgerSnapshot:
    assets: tuple[Asset, ...]
    assets_by_id: dict[str, Asset]
    storage_by_asset_id: dict[str, tuple[StorageMedium, ...]]
    changes_by_asset_id: dict[str, tuple[ChangeRecord, ...]]
    storage_search_text: dict[str, str]
    dictionaries: dict[str, tuple[DictionaryItem, ...]]
    enum_values: dict[str, tuple[tuple[int, str, bool], ...]]
    config: dict[str, str]
    fingerprint: WorkbookFingerprint
    loaded_at: str

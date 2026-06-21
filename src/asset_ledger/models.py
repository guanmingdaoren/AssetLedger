from __future__ import annotations

from dataclasses import dataclass, fields
from datetime import date, datetime
from typing import Any


@dataclass(slots=True)
class Asset:
    asset_id: str = ""
    equipment_code: str = ""
    asset_identifier: str = ""
    name: str = ""
    primary_category: str = ""
    secondary_category: str = ""
    category_path: str = ""
    product_spec: str = ""
    model: str = ""
    serial_number: str = ""
    manufacturer: str = ""
    supplier: str = ""
    brand: str = ""
    source: str = ""
    purchase_date: str = ""
    start_date: str = ""
    valuation_method: str = ""
    price: float = 0.0
    department: str = ""
    owner: str = ""
    use_department: str = ""
    user: str = ""
    location: str = ""
    manufacture_date: str = ""
    grade: str = ""
    status: str = ""
    label_printed: bool = False
    notes: str = ""
    notes2: str = ""
    created_at: str = ""
    updated_at: str = ""

    @classmethod
    def from_mapping(cls, values: dict[str, Any]) -> "Asset":
        allowed = {field.name for field in fields(cls)}
        normalized = {key: values.get(key, "") for key in allowed}
        normalized["price"] = float(normalized.get("price") or 0)
        normalized["label_printed"] = bool(normalized.get("label_printed"))
        for key in ("purchase_date", "start_date", "manufacture_date"):
            value = normalized.get(key)
            if isinstance(value, (date, datetime)):
                normalized[key] = value.strftime("%Y-%m-%d")
        for key in ("created_at", "updated_at"):
            value = normalized.get(key)
            if isinstance(value, datetime):
                normalized[key] = value.strftime("%Y-%m-%d %H:%M:%S")
        return cls(**normalized)


@dataclass(slots=True)
class StorageMedium:
    storage_id: str = ""
    asset_id: str = ""
    medium_type: str = ""
    status: str = ""
    name: str = ""
    brand: str = ""
    capacity_value: float = 0.0
    capacity_unit: str = ""
    model: str = ""
    serial_number: str = ""
    notes: str = ""
    created_at: str = ""
    updated_at: str = ""

    @classmethod
    def from_mapping(cls, values: dict[str, Any]) -> "StorageMedium":
        allowed = {field.name for field in fields(cls)}
        normalized = {key: values.get(key, "") for key in allowed}
        normalized["capacity_value"] = float(normalized.get("capacity_value") or 0)
        for key in ("created_at", "updated_at"):
            value = normalized.get(key)
            if isinstance(value, datetime):
                normalized[key] = value.strftime("%Y-%m-%d %H:%M:%S")
        return cls(**normalized)


@dataclass(slots=True)
class ChangeRecord:
    change_id: str
    event_group_id: str
    asset_id: str
    event_type: str
    field_name: str
    old_value: str
    new_value: str
    changed_at: str
    operator: str
    note: str


@dataclass(slots=True)
class DictionaryItem:
    item_id: str
    parent_id: str
    name: str
    order: int
    enabled: bool


def timestamp_now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

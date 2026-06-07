from __future__ import annotations

from dataclasses import asdict
from datetime import date, datetime
from typing import Any
from uuid import uuid4

from openpyxl.utils import get_column_letter

from .excel_repository import (
    ASSET_HEADERS,
    ASSET_SHEET,
    CATEGORY_SHEET,
    CHANGE_SHEET,
    CONFIG_SHEET,
    ENUM_SHEET,
    LOCATION_SHEET,
    ExcelRepository,
)
from .models import Asset, ChangeRecord, DictionaryItem, timestamp_now


class AssetServiceError(ValueError):
    pass


class DuplicateEquipmentCodeError(AssetServiceError):
    pass


class AssetNotFoundError(AssetServiceError):
    pass


FIELD_TO_HEADER = {
    "asset_id": "资产ID",
    "equipment_code": "装备编码",
    "name": "设备名称",
    "primary_category": "一级类别",
    "secondary_category": "二级类别",
    "brand": "品牌",
    "model": "型号",
    "serial_number": "序列号",
    "source": "来源",
    "price": "价格",
    "purchase_date": "购入日期",
    "start_date": "启用日期",
    "status": "状态",
    "location": "位置",
    "department": "责任部门",
    "owner": "责任人",
    "notes": "备注",
    "created_at": "创建时间",
    "updated_at": "更新时间",
}
HEADER_TO_FIELD = {header: field for field, header in FIELD_TO_HEADER.items()}
EDITABLE_FIELDS = [
    field
    for field in FIELD_TO_HEADER
    if field not in {"asset_id", "created_at", "updated_at"}
]


class AssetService:
    def __init__(self, repository: ExcelRepository) -> None:
        self.repository = repository

    def list_assets(
        self, filters: dict[str, str | None] | None = None, search_text: str = ""
    ) -> list[Asset]:
        workbook = self.repository.load(data_only=True)
        try:
            assets = [
                self._asset_from_row(row)
                for row in workbook[ASSET_SHEET].iter_rows(min_row=2, values_only=True)
                if row[0]
            ]
        finally:
            workbook.close()

        filters = filters or {}
        filtered: list[Asset] = []
        needle = search_text.strip().lower()
        searchable_fields = ("asset_id", "equipment_code", "name", "brand", "model")
        for asset in assets:
            if needle and not any(
                needle in str(getattr(asset, field) or "").lower()
                for field in searchable_fields
            ):
                continue
            if not self._matches_filters(asset, filters):
                continue
            filtered.append(asset)
        return sorted(filtered, key=lambda item: item.updated_at, reverse=True)

    def get_asset(self, asset_id: str) -> Asset:
        for asset in self.list_assets():
            if asset.asset_id == asset_id:
                return asset
        raise AssetNotFoundError(f"未找到设备：{asset_id}")

    def create_asset(
        self, asset_data: dict[str, Any] | Asset, change_note: str = ""
    ) -> Asset:
        asset = self._normalize_asset(asset_data)
        self._validate_asset(asset)
        workbook = self.repository.load()
        try:
            asset_sheet = workbook[ASSET_SHEET]
            self._ensure_equipment_code_unique(asset_sheet, asset.equipment_code)
            asset.asset_id = self._next_asset_id(workbook)
            now = timestamp_now()
            asset.created_at = now
            asset.updated_at = now
            self._append_or_replace_blank(asset_sheet, self._asset_to_row(asset))
            self._format_asset_row(asset_sheet, asset_sheet.max_row)
            self._resize_table(asset_sheet)
            self._append_change(
                workbook,
                ChangeRecord(
                    change_id=self._new_id("CHG"),
                    event_group_id=self._new_id("EVT"),
                    asset_id=asset.asset_id,
                    event_type="新建设备",
                    field_name="*",
                    old_value="",
                    new_value=asset.name,
                    changed_at=now,
                    operator=self._get_config_value(workbook, "默认操作者", "管理员"),
                    note=change_note,
                ),
            )
            self.repository.save(workbook)
        finally:
            workbook.close()
        return asset

    def update_asset(
        self,
        asset_id: str,
        asset_data: dict[str, Any] | Asset,
        change_note: str = "",
    ) -> Asset:
        candidate = self._normalize_asset(asset_data)
        self._validate_asset(candidate)
        workbook = self.repository.load()
        try:
            sheet = workbook[ASSET_SHEET]
            row_number = self._find_asset_row(sheet, asset_id)
            if row_number is None:
                raise AssetNotFoundError(f"未找到设备：{asset_id}")
            current = self._asset_from_row(
                tuple(sheet.cell(row=row_number, column=index).value for index in range(1, 20))
            )
            self._ensure_equipment_code_unique(
                sheet, candidate.equipment_code, excluded_asset_id=asset_id
            )
            candidate.asset_id = current.asset_id
            candidate.created_at = current.created_at
            candidate.updated_at = current.updated_at
            changes = [
                field
                for field in EDITABLE_FIELDS
                if self._comparable(getattr(current, field))
                != self._comparable(getattr(candidate, field))
            ]
            if not changes:
                return current

            now = timestamp_now()
            candidate.updated_at = now
            for column, value in enumerate(self._asset_to_row(candidate), start=1):
                sheet.cell(row=row_number, column=column).value = value
            self._format_asset_row(sheet, row_number)

            event_group_id = self._new_id("EVT")
            operator = self._get_config_value(workbook, "默认操作者", "管理员")
            for field in changes:
                header = FIELD_TO_HEADER[field]
                event_type = (
                    "状态变更"
                    if field == "status"
                    else "位置变更"
                    if field == "location"
                    else "信息修改"
                )
                self._append_change(
                    workbook,
                    ChangeRecord(
                        change_id=self._new_id("CHG"),
                        event_group_id=event_group_id,
                        asset_id=asset_id,
                        event_type=event_type,
                        field_name=header,
                        old_value=self._display_value(getattr(current, field)),
                        new_value=self._display_value(getattr(candidate, field)),
                        changed_at=now,
                        operator=operator,
                        note=change_note,
                    ),
                )
            self.repository.save(workbook)
        finally:
            workbook.close()
        return candidate

    def list_changes(self, asset_id: str, field_name: str = "") -> list[ChangeRecord]:
        workbook = self.repository.load(data_only=True)
        try:
            records = [
                ChangeRecord(
                    change_id=str(row[0] or ""),
                    event_group_id=str(row[1] or ""),
                    asset_id=str(row[2] or ""),
                    event_type=str(row[3] or ""),
                    field_name=str(row[4] or ""),
                    old_value=self._display_value(row[5]),
                    new_value=self._display_value(row[6]),
                    changed_at=self._display_value(row[7]),
                    operator=str(row[8] or ""),
                    note=str(row[9] or ""),
                )
                for row in workbook[CHANGE_SHEET].iter_rows(min_row=2, values_only=True)
                if row[0]
                and str(row[2]) == asset_id
                and (not field_name or str(row[4]) == field_name)
            ]
        finally:
            workbook.close()
        return sorted(records, key=lambda change: change.changed_at, reverse=True)

    def list_departments(self) -> list[str]:
        return sorted({asset.department for asset in self.list_assets() if asset.department})

    def list_owners(self) -> list[str]:
        return sorted({asset.owner for asset in self.list_assets() if asset.owner})

    def get_categories(self, *, include_disabled: bool = False) -> list[DictionaryItem]:
        return self._get_dictionary_items(CATEGORY_SHEET, include_disabled)

    def get_locations(self, *, include_disabled: bool = False) -> list[DictionaryItem]:
        return self._get_dictionary_items(LOCATION_SHEET, include_disabled)

    def get_enum_values(self, enum_type: str, *, include_disabled: bool = False) -> list[str]:
        workbook = self.repository.load(data_only=True)
        try:
            values = [
                (int(row[2] or 0), str(row[1]))
                for row in workbook[ENUM_SHEET].iter_rows(min_row=2, values_only=True)
                if row[0] == enum_type and (include_disabled or bool(row[3]))
            ]
        finally:
            workbook.close()
        return [value for _, value in sorted(values)]

    def get_operator(self) -> str:
        workbook = self.repository.load(data_only=True)
        try:
            return self._get_config_value(workbook, "默认操作者", "管理员")
        finally:
            workbook.close()

    def set_operator(self, operator: str) -> None:
        workbook = self.repository.load()
        try:
            self._set_config_value(workbook, "默认操作者", operator.strip() or "管理员")
            self.repository.save(workbook)
        finally:
            workbook.close()

    def add_dictionary_item(
        self, sheet_name: str, name: str, *, parent_id: str = ""
    ) -> DictionaryItem:
        if sheet_name not in {CATEGORY_SHEET, LOCATION_SHEET}:
            raise AssetServiceError(f"不支持的字典：{sheet_name}")
        clean_name = name.strip()
        if not clean_name:
            raise AssetServiceError("名称不能为空")
        workbook = self.repository.load()
        try:
            sheet = workbook[sheet_name]
            existing_rows = [
                row
                for row in sheet.iter_rows(min_row=2, values_only=True)
                if row[0]
            ]
            if any(
                str(row[2]) == clean_name and str(row[1] or "") == parent_id
                for row in existing_rows
            ):
                raise AssetServiceError(f"字典项已存在：{clean_name}")
            order = max((int(row[3] or 0) for row in existing_rows), default=0) + 10
            prefix = "CAT" if sheet_name == CATEGORY_SHEET else "LOC"
            item = DictionaryItem(
                item_id=self._new_id(prefix),
                parent_id=parent_id,
                name=clean_name,
                order=order,
                enabled=True,
            )
            self._append_or_replace_blank(
                sheet, [item.item_id, item.parent_id, item.name, item.order, item.enabled]
            )
            self._resize_table(sheet)
            self.repository.save(workbook)
        finally:
            workbook.close()
        return item

    def set_dictionary_item_enabled(
        self, sheet_name: str, item_id: str, enabled: bool
    ) -> None:
        if sheet_name not in {CATEGORY_SHEET, LOCATION_SHEET}:
            raise AssetServiceError(f"不支持的字典：{sheet_name}")
        workbook = self.repository.load()
        try:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows(min_row=2):
                if row[0].value == item_id:
                    row[4].value = bool(enabled)
                    self.repository.save(workbook)
                    return
            raise AssetServiceError(f"未找到字典项：{item_id}")
        finally:
            workbook.close()

    def set_enum_values(self, enum_type: str, values: list[str]) -> None:
        cleaned = list(dict.fromkeys(value.strip() for value in values if value.strip()))
        if not cleaned:
            raise AssetServiceError(f"{enum_type}至少需要一个选项")
        workbook = self.repository.load()
        try:
            sheet = workbook[ENUM_SHEET]
            preserved = [
                list(row)
                for row in sheet.iter_rows(min_row=2, values_only=True)
                if row[0] and row[0] != enum_type
            ]
            if sheet.max_row > 1:
                sheet.delete_rows(2, sheet.max_row - 1)
            for row in preserved:
                sheet.append(row)
            for index, value in enumerate(cleaned, start=1):
                sheet.append([enum_type, value, index * 10, True])
            self._resize_table(sheet)
            self.repository.save(workbook)
        finally:
            workbook.close()

    def _get_dictionary_items(
        self, sheet_name: str, include_disabled: bool
    ) -> list[DictionaryItem]:
        workbook = self.repository.load(data_only=True)
        try:
            items = [
                DictionaryItem(
                    item_id=str(row[0] or ""),
                    parent_id=str(row[1] or ""),
                    name=str(row[2] or ""),
                    order=int(row[3] or 0),
                    enabled=bool(row[4]),
                )
                for row in workbook[sheet_name].iter_rows(min_row=2, values_only=True)
                if row[0] and (include_disabled or bool(row[4]))
            ]
        finally:
            workbook.close()
        return sorted(items, key=lambda item: item.order)

    @staticmethod
    def _matches_filters(asset: Asset, filters: dict[str, str | None]) -> bool:
        for key, value in filters.items():
            if value == "":
                continue
            asset_value = str(getattr(asset, key, "") or "")
            if value is None:
                if asset_value:
                    return False
            elif asset_value != str(value):
                return False
        return True

    @staticmethod
    def _normalize_asset(asset_data: dict[str, Any] | Asset) -> Asset:
        if isinstance(asset_data, Asset):
            return Asset.from_mapping(asdict(asset_data))
        return Asset.from_mapping(asset_data)

    @staticmethod
    def _validate_asset(asset: Asset) -> None:
        if not asset.equipment_code.strip():
            raise AssetServiceError("装备编码不能为空")
        if not asset.name.strip():
            raise AssetServiceError("设备名称不能为空")
        if asset.price < 0:
            raise AssetServiceError("价格不能为负数")

    @staticmethod
    def _ensure_equipment_code_unique(
        sheet, equipment_code: str, excluded_asset_id: str = ""
    ) -> None:
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0] and str(row[0]) != excluded_asset_id and str(row[1]) == equipment_code:
                raise DuplicateEquipmentCodeError(f"装备编码已存在：{equipment_code}")

    @staticmethod
    def _find_asset_row(sheet, asset_id: str) -> int | None:
        for row_number in range(2, sheet.max_row + 1):
            if sheet.cell(row=row_number, column=1).value == asset_id:
                return row_number
        return None

    @staticmethod
    def _next_asset_id(workbook) -> str:
        value = int(AssetService._get_config_value(workbook, "下一个资产序号", "1"))
        existing_ids = {
            str(row[0])
            for row in workbook[ASSET_SHEET].iter_rows(min_row=2, values_only=True)
            if row[0]
        }
        candidate = f"AST-{datetime.now().year}-{value:06d}"
        while candidate in existing_ids:
            value += 1
            candidate = f"AST-{datetime.now().year}-{value:06d}"
        AssetService._set_config_value(workbook, "下一个资产序号", str(value + 1))
        return candidate

    @staticmethod
    def _get_config_value(workbook, key: str, default: str = "") -> str:
        for row in workbook[CONFIG_SHEET].iter_rows(min_row=2):
            if row[0].value == key:
                return str(row[1].value or default)
        return default

    @staticmethod
    def _set_config_value(workbook, key: str, value: str) -> None:
        sheet = workbook[CONFIG_SHEET]
        for row in sheet.iter_rows(min_row=2):
            if row[0].value == key:
                row[1].value = value
                return
        sheet.append([key, value, ""])
        AssetService._resize_table(sheet)

    @staticmethod
    def _append_change(workbook, change: ChangeRecord) -> None:
        sheet = workbook[CHANGE_SHEET]
        AssetService._append_or_replace_blank(
            sheet,
            [
                change.change_id,
                change.event_group_id,
                change.asset_id,
                change.event_type,
                change.field_name,
                change.old_value,
                change.new_value,
                change.changed_at,
                change.operator,
                change.note,
            ],
        )
        AssetService._format_change_row(sheet, sheet.max_row)
        AssetService._resize_table(sheet)

    @staticmethod
    def _append_or_replace_blank(sheet, values: list[Any]) -> None:
        if sheet.max_row == 2 and not any(
            sheet.cell(row=2, column=column).value for column in range(1, sheet.max_column + 1)
        ):
            for column, value in enumerate(values, start=1):
                sheet.cell(row=2, column=column).value = value
        else:
            sheet.append(values)

    @staticmethod
    def _resize_table(sheet) -> None:
        if not sheet.tables:
            return
        table = next(iter(sheet.tables.values()))
        table.ref = f"A1:{get_column_letter(sheet.max_column)}{sheet.max_row}"
        sheet.auto_filter.ref = table.ref

    @staticmethod
    def _asset_to_row(asset: Asset) -> list[Any]:
        row: list[Any] = []
        for header in ASSET_HEADERS:
            field = HEADER_TO_FIELD[header]
            value = getattr(asset, field)
            if field in {"purchase_date", "start_date"} and value:
                try:
                    value = date.fromisoformat(str(value))
                except ValueError:
                    pass
            row.append(value)
        return row

    @staticmethod
    def _asset_from_row(row: tuple[Any, ...]) -> Asset:
        values = {
            HEADER_TO_FIELD[header]: row[index] if index < len(row) else ""
            for index, header in enumerate(ASSET_HEADERS)
        }
        for key in values:
            if values[key] is None:
                values[key] = ""
        return Asset.from_mapping(values)

    @staticmethod
    def _new_id(prefix: str) -> str:
        return f"{prefix}-{uuid4().hex.upper()}"

    @staticmethod
    def _display_value(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    @staticmethod
    def _comparable(value: Any) -> str:
        return AssetService._display_value(value).strip()

    @staticmethod
    def _format_asset_row(sheet, row_number: int) -> None:
        sheet.cell(row=row_number, column=10).number_format = '¥#,##0.00'
        for column in (11, 12):
            sheet.cell(row=row_number, column=column).number_format = "yyyy-mm-dd"
        for column in (18, 19):
            sheet.cell(row=row_number, column=column).number_format = "yyyy-mm-dd hh:mm:ss"

    @staticmethod
    def _format_change_row(sheet, row_number: int) -> None:
        sheet.cell(row=row_number, column=8).number_format = "yyyy-mm-dd hh:mm:ss"

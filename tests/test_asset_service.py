from pathlib import Path
from tempfile import TemporaryDirectory
from time import perf_counter
import unittest
from unittest.mock import patch

from asset_ledger.asset_service import (
    AssetService,
    AssetServiceError,
    CacheReloadAfterSaveError,
    CacheUnavailableError,
    DuplicateEquipmentCodeError,
)
from asset_ledger.excel_repository import (
    ASSET_HEADERS,
    ASSET_SHEET,
    CATEGORY_SHEET,
    CHANGE_SHEET,
    ExcelRepository,
    LOCATION_SHEET,
    STORAGE_SHEET,
    WorkbookChangedExternallyError,
)


def sample_asset(**overrides):
    asset = {
        "equipment_code": "EQ-001",
        "asset_identifier": "",
        "name": "测试电脑",
        "primary_category": "IT 与通信设备",
        "secondary_category": "电脑",
        "category_path": "IT 与通信设备 / 电脑",
        "brand": "联想",
        "model": "T14",
        "serial_number": "SN-001",
        "product_spec": "14英寸 / 32GB",
        "manufacturer": "联想（北京）有限公司",
        "supplier": "示例供应商",
        "source": "采购",
        "valuation_method": "原值",
        "price": 5999.0,
        "purchase_date": "2026-01-02",
        "start_date": "2026-01-05",
        "manufacture_date": "2025-12-01",
        "grade": "A",
        "status": "在用",
        "location": "办公室",
        "department": "技术部",
        "owner": "张三",
        "use_department": "研发部",
        "user": "李四",
        "label_printed": True,
        "notes": "初始设备",
        "notes2": "",
    }
    asset.update(overrides)
    if "category_path" not in overrides:
        asset["category_path"] = " / ".join(
            value for value in (asset["primary_category"], asset["secondary_category"]) if value
        )
    return asset


def sample_storage(**overrides):
    medium = {
        "medium_type": "SSD",
        "status": "",
        "name": "系统盘",
        "brand": "三星",
        "capacity_value": 2,
        "capacity_unit": "TB",
        "model": "990 PRO",
        "serial_number": "SSD-SN-001",
        "notes": "系统与软件",
    }
    medium.update(overrides)
    return medium


class AssetServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.path = Path(self.temporary_directory.name) / "assets.xlsx"
        self.repository = ExcelRepository(self.path)
        self.repository.initialize()
        self.service = AssetService(self.repository)

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def test_create_asset_generates_id_and_creation_history(self) -> None:
        created = self.service.create_asset(sample_asset(asset_identifier="ZC-001"))

        self.assertRegex(created.asset_id, r"^AST-\d{4}-000001$")
        self.assertEqual(created.equipment_code, "EQ-001")
        self.assertEqual(created.asset_identifier, "ZC-001")
        changes = self.service.list_changes(created.asset_id)
        self.assertEqual(len(changes), 1)
        self.assertEqual(changes[0].event_type, "新建设备")
        self.assertEqual(changes[0].field_name, "*")

    def test_queries_use_cache_without_reopening_workbook(self) -> None:
        created = self.service.create_asset(sample_asset(), [sample_storage()])

        with patch.object(
            self.repository, "load", side_effect=AssertionError("不应重新读取 Excel")
        ):
            self.assertEqual(self.service.get_asset(created.asset_id).name, "测试电脑")
            self.assertEqual(len(self.service.list_assets(search_text="联想")), 1)
            self.assertEqual(len(self.service.list_changes(created.asset_id)), 2)
            self.assertEqual(len(self.service.list_storage_media(created.asset_id)), 1)
            self.assertGreater(len(self.service.get_categories()), 0)
            self.assertGreater(len(self.service.get_locations()), 0)
            self.assertIn("采购", self.service.get_enum_values("取得方式"))
            self.assertEqual(self.service.get_operator(), "管理员")

    def test_service_initialization_loads_and_validates_workbook_once(self) -> None:
        with patch.object(
            self.repository,
            "load_with_fingerprint",
            wraps=self.repository.load_with_fingerprint,
        ) as load, patch.object(
            self.repository,
            "validate_workbook",
            side_effect=AssertionError("不应单独重复校验"),
        ):
            service = AssetService(self.repository)

        self.assertEqual(load.call_count, 1)
        self.assertFalse(service.cache_status.stale)

    def test_query_results_are_copies_and_cannot_mutate_cache(self) -> None:
        created = self.service.create_asset(sample_asset(), [sample_storage()])

        asset = self.service.get_asset(created.asset_id)
        medium = self.service.list_storage_media(created.asset_id)[0]
        change = self.service.list_changes(created.asset_id)[0]
        category = self.service.get_categories()[0]
        asset.name = "未保存修改"
        medium.serial_number = "未保存介质修改"
        change.new_value = "未保存历史修改"
        category.name = "未保存分类修改"

        self.assertEqual(self.service.get_asset(created.asset_id).name, "测试电脑")
        self.assertEqual(
            self.service.list_storage_media(created.asset_id)[0].serial_number,
            "SSD-SN-001",
        )
        self.assertNotEqual(
            self.service.list_changes(created.asset_id)[0].new_value,
            "未保存历史修改",
        )
        self.assertNotEqual(self.service.get_categories()[0].name, "未保存分类修改")

    def test_external_workbook_change_blocks_write_until_reload(self) -> None:
        self.service.create_asset(sample_asset())
        backups_before = list((self.path.parent / "backups").glob("assets-*.xlsx"))
        workbook = self.repository.load()
        workbook["系统配置"]["B4"] = "外部修改"
        self.repository.save(workbook, create_backup=False)
        workbook.close()

        with self.assertRaises(WorkbookChangedExternallyError):
            self.service.create_asset(
                sample_asset(equipment_code="EQ-002", serial_number="SN-002")
            )

        self.assertEqual(
            list((self.path.parent / "backups").glob("assets-*.xlsx")),
            backups_before,
        )
        self.assertEqual(len(self.service.list_assets()), 1)
        self.service.reload_cache()
        created = self.service.create_asset(
            sample_asset(equipment_code="EQ-002", serial_number="SN-002")
        )
        self.assertEqual(created.equipment_code, "EQ-002")

    def test_external_workbook_change_blocks_delete_until_reload(self) -> None:
        created = self.service.create_asset(sample_asset())
        backups_before = list((self.path.parent / "backups").glob("assets-*.xlsx"))
        workbook = self.repository.load()
        workbook["系统配置"]["B4"] = "外部修改"
        self.repository.save(workbook, create_backup=False)
        workbook.close()

        with self.assertRaises(WorkbookChangedExternallyError):
            self.service.delete_asset(created.asset_id)

        self.assertEqual(
            list((self.path.parent / "backups").glob("assets-*.xlsx")),
            backups_before,
        )
        persisted = self.repository.load(data_only=True)
        try:
            self.assertTrue(
                any(
                    row[0] == created.asset_id
                    for row in persisted[ASSET_SHEET].iter_rows(
                        min_row=2, values_only=True
                    )
                )
            )
        finally:
            persisted.close()

    def test_failed_manual_reload_keeps_old_cache_and_blocks_writes(self) -> None:
        created = self.service.create_asset(sample_asset())
        original_bytes = self.path.read_bytes()
        self.path.write_text("not an xlsx", encoding="utf-8")

        with self.assertRaises(Exception):
            self.service.reload_cache()

        self.assertTrue(self.service.cache_status.stale)
        self.assertEqual(self.service.get_asset(created.asset_id).name, "测试电脑")
        with self.assertRaises(CacheUnavailableError):
            self.service.create_asset(
                sample_asset(equipment_code="EQ-002", serial_number="SN-002")
            )
        self.path.write_bytes(original_bytes)
        self.service.reload_cache()
        self.assertFalse(self.service.cache_status.stale)

    def test_save_success_followed_by_cache_reload_failure_blocks_more_writes(self) -> None:
        original_reload = self.service.reload_cache
        with patch.object(
            self.service, "reload_cache", side_effect=RuntimeError("缓存读取失败")
        ):
            with self.assertRaises(CacheReloadAfterSaveError):
                self.service.create_asset(sample_asset())

        self.assertTrue(self.service.cache_status.stale)
        with self.assertRaises(CacheUnavailableError):
            self.service.create_asset(
                sample_asset(equipment_code="EQ-002", serial_number="SN-002")
            )
        original_reload()
        self.assertEqual(len(self.service.list_assets()), 1)

    def test_failed_save_keeps_workbook_and_cache_unchanged(self) -> None:
        first = self.service.create_asset(sample_asset())

        with patch.object(self.repository, "save", side_effect=RuntimeError("写入失败")):
            with self.assertRaisesRegex(RuntimeError, "写入失败"):
                self.service.create_asset(
                    sample_asset(equipment_code="EQ-002", serial_number="SN-002")
                )

        self.assertFalse(self.service.cache_status.stale)
        self.assertEqual([asset.asset_id for asset in self.service.list_assets()], [first.asset_id])
        workbook = self.repository.load(data_only=True)
        rows = [
            row
            for row in workbook["资产台账"].iter_rows(min_row=2, values_only=True)
            if row[0]
        ]
        workbook.close()
        self.assertEqual(len(rows), 1)

    def test_no_change_update_does_not_backup_or_reload_cache(self) -> None:
        created = self.service.create_asset(sample_asset())
        backups_before = list((self.path.parent / "backups").glob("assets-*.xlsx"))

        with patch.object(self.service, "reload_cache", wraps=self.service.reload_cache) as reload:
            self.service.update_asset(created.asset_id, sample_asset())

        self.assertEqual(reload.call_count, 0)
        self.assertEqual(
            list((self.path.parent / "backups").glob("assets-*.xlsx")),
            backups_before,
        )

    def test_create_asset_rejects_duplicate_equipment_code(self) -> None:
        self.service.create_asset(sample_asset())

        with self.assertRaises(DuplicateEquipmentCodeError):
            self.service.create_asset(sample_asset(name="另一台电脑"))

    def test_asset_identifier_saves_searches_and_requires_unique_when_present(self) -> None:
        first = self.service.create_asset(sample_asset(asset_identifier="ZC-UNIQUE-001"))

        searched = self.service.list_assets(search_text="ZC-UNIQUE-001")

        self.assertEqual(first.asset_identifier, "ZC-UNIQUE-001")
        self.assertEqual([asset.asset_id for asset in searched], [first.asset_id])
        with self.assertRaisesRegex(AssetServiceError, "资产唯一标识符已存在"):
            self.service.create_asset(
                sample_asset(
                    equipment_code="EQ-002",
                    serial_number="SN-002",
                    asset_identifier="ZC-UNIQUE-001",
                )
            )
        second = self.service.create_asset(
            sample_asset(
                equipment_code="EQ-002",
                serial_number="SN-002",
                name="未登记标识符设备",
                asset_identifier="",
            )
        )
        self.assertEqual(second.asset_identifier, "")

    def test_two_asset_notes_save_search_and_record_field_history(self) -> None:
        created = self.service.create_asset(
            sample_asset(notes="第一段备注", notes2="第二段备注")
        )

        self.assertEqual(created.notes, "第一段备注")
        self.assertEqual(created.notes2, "第二段备注")
        self.assertEqual(
            [asset.asset_id for asset in self.service.list_assets(search_text="第二段备注")],
            [created.asset_id],
        )

        self.service.update_asset(
            created.asset_id,
            sample_asset(notes="第一段备注已改", notes2="第二段备注已改"),
            "整理备注",
        )
        changes = [
            change
            for change in self.service.list_changes(created.asset_id)
            if change.field_name in {"备注1", "备注2"}
        ]

        self.assertEqual({change.field_name for change in changes}, {"备注1", "备注2"})
        self.assertEqual({change.note for change in changes}, {"整理备注"})

    def test_create_asset_allows_blank_code_but_requires_name(self) -> None:
        first = self.service.create_asset(sample_asset(equipment_code=""))
        second = self.service.create_asset(
            sample_asset(equipment_code="", name="第二台设备", serial_number="SN-002")
        )

        self.assertNotEqual(first.asset_id, second.asset_id)
        with self.assertRaisesRegex(AssetServiceError, "设备器材"):
            self.service.create_asset(sample_asset(equipment_code="", name=""))

    def test_create_asset_saves_multiple_storage_media(self) -> None:
        created = self.service.create_asset(
            sample_asset(),
            [
                sample_storage(),
                sample_storage(
                    medium_type="机械硬盘",
                    name="数据盘",
                    capacity_value=8,
                    model="ST8000",
                    serial_number="HDD-SN-001",
                ),
            ],
        )

        media = self.service.list_storage_media(created.asset_id)

        self.assertEqual(len(media), 2)
        self.assertTrue(all(item.storage_id.startswith("MED-") for item in media))
        self.assertEqual({item.asset_id for item in media}, {created.asset_id})

    def test_delete_asset_removes_asset_storage_and_history_and_creates_backup(self) -> None:
        created = self.service.create_asset(sample_asset(), [sample_storage()])
        self.service.update_asset(
            created.asset_id,
            sample_asset(status="维修", location="仓库"),
            "删除前变更",
        )
        self.assertGreater(len(self.service.list_changes(created.asset_id)), 0)

        self.service.delete_asset(created.asset_id)

        self.assertEqual(self.service.list_assets(), [])
        self.assertEqual(self.service.list_storage_media(created.asset_id), [])
        self.assertEqual(self.service.list_changes(created.asset_id), [])
        backups = list((self.path.parent / "backups").glob("*.xlsx"))
        self.assertGreaterEqual(len(backups), 1)
        with self.assertRaises(AssetServiceError):
            self.service.get_asset(created.asset_id)

        workbook = self.repository.load(data_only=True)
        try:
            self.assertFalse(
                any(
                    row[0] == created.asset_id
                    for row in workbook[ASSET_SHEET].iter_rows(
                        min_row=2, values_only=True
                    )
                )
            )
            self.assertFalse(
                any(
                    row[1] == created.asset_id
                    for row in workbook[STORAGE_SHEET].iter_rows(
                        min_row=2, values_only=True
                    )
                )
            )
            self.assertFalse(
                any(
                    row[2] == created.asset_id
                    for row in workbook[CHANGE_SHEET].iter_rows(
                        min_row=2, values_only=True
                    )
                )
            )
        finally:
            workbook.close()

    def test_delete_asset_rejects_missing_asset(self) -> None:
        with self.assertRaisesRegex(AssetServiceError, "未找到设备"):
            self.service.delete_asset("AST-2099-000001")

    def test_blank_storage_capacity_is_saved_as_empty_excel_cell(self) -> None:
        self.service.create_asset(
            sample_asset(),
            [sample_storage(capacity_value=0, capacity_unit="")],
        )

        workbook = self.repository.load(data_only=True)
        try:
            sheet = workbook[STORAGE_SHEET]
            headers = [cell.value for cell in sheet[1]]
            row = next(sheet.iter_rows(min_row=2, values_only=True))
        finally:
            workbook.close()

        self.assertIsNone(row[headers.index("容量数值")])

    def test_category_path_supports_multi_level_assets_and_prefix_filter(self) -> None:
        notebook = self.service.add_dictionary_item(
            CATEGORY_SHEET, "笔记本电脑", parent_id="CAT-IT-PC"
        )

        created = self.service.create_asset(
            sample_asset(
                category_path="IT 与通信设备 / 电脑 / 笔记本电脑",
                primary_category="IT 与通信设备",
                secondary_category="电脑",
            )
        )
        exact = self.service.list_assets(filters={"category_path": "IT 与通信设备 / 电脑 / 笔记本电脑"})
        parent = self.service.list_assets(filters={"category_path": "IT 与通信设备 / 电脑"})

        self.assertTrue(notebook.item_id.startswith("CAT-"))
        self.assertEqual(created.category_path, "IT 与通信设备 / 电脑 / 笔记本电脑")
        self.assertEqual([asset.asset_id for asset in exact], [created.asset_id])
        self.assertEqual([asset.asset_id for asset in parent], [created.asset_id])

    def test_delete_dictionary_item_allows_unused_leaf_and_rejects_used_or_parent(self) -> None:
        unused = self.service.add_dictionary_item(LOCATION_SHEET, "临时库房")
        used = self.service.add_dictionary_item(LOCATION_SHEET, "已用库房")
        parent = self.service.add_dictionary_item(CATEGORY_SHEET, "测试父类")
        self.service.add_dictionary_item(CATEGORY_SHEET, "测试子类", parent_id=parent.item_id)
        self.service.create_asset(sample_asset(location="已用库房"))

        self.service.delete_dictionary_item(LOCATION_SHEET, unused.item_id)

        self.assertNotIn(
            unused.item_id,
            [item.item_id for item in self.service.get_locations(include_disabled=True)],
        )
        with self.assertRaisesRegex(AssetServiceError, "子项"):
            self.service.delete_dictionary_item(CATEGORY_SHEET, parent.item_id)
        with self.assertRaisesRegex(AssetServiceError, "已被设备使用"):
            self.service.delete_dictionary_item(LOCATION_SHEET, used.item_id)

    def test_update_storage_media_records_add_modify_and_remove_history(self) -> None:
        created = self.service.create_asset(sample_asset(), [sample_storage()])
        original = self.service.list_storage_media(created.asset_id)[0]

        self.service.update_asset(
            created.asset_id,
            sample_asset(),
            [
                {
                    **sample_storage(
                        name="系统盘",
                        brand="西数",
                        serial_number=original.serial_number,
                    ),
                    "storage_id": original.storage_id,
                },
                sample_storage(
                    name="数据盘",
                    serial_number="SSD-SN-002",
                ),
            ],
            "调整存储",
        )
        media_after_add = self.service.list_storage_media(created.asset_id)
        added = next(item for item in media_after_add if item.storage_id != original.storage_id)
        self.service.update_asset(
            created.asset_id,
            sample_asset(),
            [
                {
                    **sample_storage(
                        name="数据盘",
                        serial_number=added.serial_number,
                    ),
                    "storage_id": added.storage_id,
                }
            ],
            "移除旧盘",
        )

        changes = self.service.list_changes(created.asset_id)
        event_types = {change.event_type for change in changes}
        self.assertIn("存储介质新增", event_types)
        self.assertIn("存储介质修改", event_types)
        self.assertIn("存储介质移除", event_types)

    def test_storage_status_saves_searches_and_records_history(self) -> None:
        created = self.service.create_asset(sample_asset(), [sample_storage(status="在用")])
        original = self.service.list_storage_media(created.asset_id)[0]

        self.service.update_asset(
            created.asset_id,
            sample_asset(),
            [
                {
                    **sample_storage(status="报废", serial_number=original.serial_number),
                    "storage_id": original.storage_id,
                }
            ],
            "介质报废",
        )

        updated = self.service.list_storage_media(created.asset_id)[0]
        searched = self.service.list_assets(search_text="报废")
        storage_status_changes = [
            change
            for change in self.service.list_changes(created.asset_id)
            if change.field_name == "存储介质-使用状态"
        ]

        self.assertEqual(updated.status, "报废")
        self.assertEqual([asset.asset_id for asset in searched], [created.asset_id])
        self.assertEqual(len(storage_status_changes), 1)
        self.assertEqual(storage_status_changes[0].old_value, "在用")
        self.assertEqual(storage_status_changes[0].new_value, "报废")

    def test_removing_all_storage_media_keeps_workbook_valid(self) -> None:
        created = self.service.create_asset(sample_asset(), [sample_storage()])

        self.service.update_asset(created.asset_id, sample_asset(), [], "移除全部存储介质")

        self.assertEqual(self.service.list_storage_media(created.asset_id), [])
        self.assertEqual(self.repository.validate_workbook(), [])

    def test_searches_storage_serial_and_filters_user(self) -> None:
        self.service.create_asset(sample_asset(), [sample_storage(serial_number="SECRET-SSD")])
        self.service.create_asset(
            sample_asset(
                equipment_code="EQ-002",
                name="打印机",
                serial_number="SN-002",
                user="王五",
            )
        )

        searched = self.service.list_assets(search_text="SECRET-SSD")
        filtered = self.service.list_assets(filters={"user": "王五"})

        self.assertEqual([asset.name for asset in searched], ["测试电脑"])
        self.assertEqual([asset.name for asset in filtered], ["打印机"])

    def test_asset_id_generation_skips_existing_id_when_counter_is_stale(self) -> None:
        first = self.service.create_asset(sample_asset())
        workbook = self.repository.load()
        for row in workbook["系统配置"].iter_rows(min_row=2):
            if row[0].value == "下一个资产序号":
                row[1].value = "1"
        self.repository.save(workbook, create_backup=False)
        workbook.close()
        self.service.reload_cache()

        second = self.service.create_asset(
            sample_asset(equipment_code="EQ-002", serial_number="SN-002")
        )

        self.assertNotEqual(first.asset_id, second.asset_id)
        self.assertTrue(second.asset_id.endswith("000002"))

    def test_update_asset_records_changed_fields_in_one_event_group(self) -> None:
        created = self.service.create_asset(sample_asset())

        updated_data = sample_asset(
            status="维修",
            location="仓库",
            price=5200.0,
            asset_identifier="ZC-UPDATED-001",
        )
        updated = self.service.update_asset(created.asset_id, updated_data, "送修并转入仓库")

        self.assertEqual(updated.status, "维修")
        self.assertEqual(updated.asset_identifier, "ZC-UPDATED-001")
        changes = self.service.list_changes(created.asset_id)
        updated_changes = [change for change in changes if change.event_type != "新建设备"]
        self.assertEqual(
            {change.field_name for change in updated_changes},
            {"资产唯一标识符", "金额", "使用状态", "存放地点"},
        )
        self.assertEqual(len({change.event_group_id for change in updated_changes}), 1)
        self.assertEqual({change.note for change in updated_changes}, {"送修并转入仓库"})

    def test_update_without_changes_does_not_append_history(self) -> None:
        created = self.service.create_asset(sample_asset())
        before = self.service.list_changes(created.asset_id)

        self.service.update_asset(created.asset_id, sample_asset(), "没有实际变化")

        after = self.service.list_changes(created.asset_id)
        self.assertEqual(len(after), len(before))

    def test_list_assets_supports_search_and_filters(self) -> None:
        self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(
                equipment_code="EQ-002",
                name="打印机",
                secondary_category="打印机",
                brand="惠普",
                serial_number="SN-002",
                status="封存",
                location="仓库",
            )
        )

        searched = self.service.list_assets(search_text="惠普")
        filtered = self.service.list_assets(filters={"status": "在用", "location": "办公室"})

        self.assertEqual([asset.equipment_code for asset in searched], ["EQ-002"])
        self.assertEqual([asset.equipment_code for asset in filtered], ["EQ-001"])

    def test_list_assets_can_filter_unassigned_owner(self) -> None:
        self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(
                equipment_code="EQ-002",
                serial_number="SN-002",
                owner="",
            )
        )

        unassigned = self.service.list_assets(filters={"owner": None})

        self.assertEqual([asset.equipment_code for asset in unassigned], ["EQ-002"])

    def test_lists_unique_departments_and_owners_from_assets(self) -> None:
        self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(
                equipment_code="EQ-002",
                serial_number="SN-002",
                department="行政部",
                owner="李四",
            )
        )
        self.service.create_asset(
            sample_asset(
                equipment_code="EQ-003",
                serial_number="SN-003",
                department="技术部",
                owner="",
            )
        )

        self.assertEqual(self.service.list_departments(), ["技术部", "行政部"])
        self.assertEqual(self.service.list_owners(), ["张三", "李四"])

    def test_list_changes_can_filter_by_field_name(self) -> None:
        created = self.service.create_asset(sample_asset())
        self.service.update_asset(
            created.asset_id,
            sample_asset(status="维修", location="仓库", owner="李四"),
            "调整归属",
        )

        status_changes = self.service.list_changes(created.asset_id, field_name="使用状态")

        self.assertEqual(len(status_changes), 1)
        self.assertEqual(status_changes[0].field_name, "使用状态")

    def test_write_operation_creates_backup(self) -> None:
        self.service.create_asset(sample_asset())

        backups = list((self.path.parent / "backups").glob("assets-*.xlsx"))

        self.assertEqual(len(backups), 1)

    def test_settings_can_add_and_disable_dictionary_items(self) -> None:
        self.service.add_dictionary_item(
            "位置字典", "测试室", parent_id="LOC-OFFICE"
        )
        added = next(item for item in self.service.get_locations() if item.name == "测试室")

        self.service.set_dictionary_item_enabled("位置字典", added.item_id, False)

        self.assertNotIn("测试室", [item.name for item in self.service.get_locations()])
        self.assertIn(
            "测试室",
            [item.name for item in self.service.get_locations(include_disabled=True)],
        )

    def test_settings_can_replace_enum_values(self) -> None:
        self.service.set_enum_values("取得方式", ["采购", "租赁"])

        self.assertEqual(self.service.get_enum_values("取得方式"), ["采购", "租赁"])

    def test_disabled_category_does_not_hide_existing_asset(self) -> None:
        created = self.service.create_asset(sample_asset())
        category = next(
            item for item in self.service.get_categories() if item.name == "IT 与通信设备"
        )

        self.service.set_dictionary_item_enabled("分类字典", category.item_id, False)

        self.assertEqual(self.service.get_asset(created.asset_id).name, "测试电脑")

    def test_lists_two_thousand_assets_within_five_seconds(self) -> None:
        workbook = self.repository.load()
        sheet = workbook["资产台账"]
        for index in range(2000):
            sheet.append(
                [
                    f"AST-2026-{index + 1:06d}",
                    f"EQ-{index + 1:05d}",
                    "",
                    f"设备 {index + 1}",
                    "IT 与通信设备",
                    "电脑",
                    "品牌",
                    "型号",
                    f"SN-{index + 1}",
                    "采购",
                    1000,
                    "",
                    "",
                    "在用",
                    "办公室",
                    "",
                    "",
                    "",
                    "2026-06-05 10:00:00",
                    "2026-06-05 10:00:00",
                ]
            )
        self.repository.save(workbook, create_backup=False)
        workbook.close()
        self.service.reload_cache()

        started = perf_counter()
        assets = self.service.list_assets()
        elapsed = perf_counter() - started

        self.assertEqual(len(assets), 2000)
        self.assertLess(elapsed, 5)

    def test_searches_five_thousand_cached_assets_within_half_second(self) -> None:
        workbook = self.repository.load()
        sheet = workbook["资产台账"]
        for index in range(5000):
            sheet.append(
                [
                    f"AST-2026-{index + 1:06d}",
                    f"EQ-{index + 1:05d}",
                    "",
                    f"设备 {index + 1}",
                    *([""] * (len(ASSET_HEADERS) - 4)),
                ]
            )
        self.repository.save(workbook, create_backup=False)
        workbook.close()
        self.service.reload_cache()

        started = perf_counter()
        assets = self.service.list_assets(search_text="设备 5000")
        elapsed = perf_counter() - started

        self.assertEqual(len(assets), 1)
        self.assertLess(elapsed, 0.5)


if __name__ == "__main__":
    unittest.main()

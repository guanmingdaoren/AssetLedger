from pathlib import Path
from tempfile import TemporaryDirectory
from time import perf_counter
import unittest

from asset_ledger.asset_service import AssetService, DuplicateEquipmentCodeError
from asset_ledger.excel_repository import ExcelRepository


def sample_asset(**overrides):
    asset = {
        "equipment_code": "EQ-001",
        "name": "测试电脑",
        "primary_category": "IT 与通信设备",
        "secondary_category": "电脑",
        "brand": "联想",
        "model": "T14",
        "serial_number": "SN-001",
        "source": "采购",
        "price": 5999.0,
        "purchase_date": "2026-01-02",
        "start_date": "2026-01-05",
        "status": "在用",
        "location": "办公室",
        "department": "技术部",
        "owner": "张三",
        "notes": "初始设备",
    }
    asset.update(overrides)
    return asset


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
        created = self.service.create_asset(sample_asset())

        self.assertRegex(created.asset_id, r"^AST-\d{4}-000001$")
        self.assertEqual(created.equipment_code, "EQ-001")
        changes = self.service.list_changes(created.asset_id)
        self.assertEqual(len(changes), 1)
        self.assertEqual(changes[0].event_type, "新建设备")
        self.assertEqual(changes[0].field_name, "*")

    def test_create_asset_rejects_duplicate_equipment_code(self) -> None:
        self.service.create_asset(sample_asset())

        with self.assertRaises(DuplicateEquipmentCodeError):
            self.service.create_asset(sample_asset(name="另一台电脑"))

    def test_asset_id_generation_skips_existing_id_when_counter_is_stale(self) -> None:
        first = self.service.create_asset(sample_asset())
        workbook = self.repository.load()
        for row in workbook["系统配置"].iter_rows(min_row=2):
            if row[0].value == "下一个资产序号":
                row[1].value = "1"
        self.repository.save(workbook, create_backup=False)
        workbook.close()

        second = self.service.create_asset(
            sample_asset(equipment_code="EQ-002", serial_number="SN-002")
        )

        self.assertNotEqual(first.asset_id, second.asset_id)
        self.assertTrue(second.asset_id.endswith("000002"))

    def test_update_asset_records_changed_fields_in_one_event_group(self) -> None:
        created = self.service.create_asset(sample_asset())

        updated_data = sample_asset(status="维修", location="仓库", price=5200.0)
        updated = self.service.update_asset(created.asset_id, updated_data, "送修并转入仓库")

        self.assertEqual(updated.status, "维修")
        changes = self.service.list_changes(created.asset_id)
        updated_changes = [change for change in changes if change.event_type != "新建设备"]
        self.assertEqual({change.field_name for change in updated_changes}, {"价格", "状态", "位置"})
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

        status_changes = self.service.list_changes(created.asset_id, field_name="状态")

        self.assertEqual(len(status_changes), 1)
        self.assertEqual(status_changes[0].field_name, "状态")

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
        self.service.set_enum_values("来源", ["采购", "租赁"])

        self.assertEqual(self.service.get_enum_values("来源"), ["采购", "租赁"])

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

        started = perf_counter()
        assets = self.service.list_assets()
        elapsed = perf_counter() - started

        self.assertEqual(len(assets), 2000)
        self.assertLess(elapsed, 5)


if __name__ == "__main__":
    unittest.main()

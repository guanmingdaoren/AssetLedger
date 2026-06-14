from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from asset_ledger.excel_repository import (
    CATEGORY_HEADERS,
    CHANGE_HEADERS,
    CONFIG_HEADERS,
    DEFAULT_CATEGORIES,
    DEFAULT_CONFIG,
    DEFAULT_ENUMS,
    DEFAULT_LOCATIONS,
    ENUM_HEADERS,
    ExcelRepository,
    LOCATION_HEADERS,
    STORAGE_HEADERS,
    WorkbookLockedError,
)


V1_ASSET_HEADERS = [
    "资产ID",
    "装备编码",
    "设备名称",
    "一级类别",
    "二级类别",
    "品牌",
    "型号",
    "序列号",
    "来源",
    "价格",
    "购入日期",
    "启用日期",
    "状态",
    "位置",
    "责任部门",
    "责任人",
    "备注",
    "创建时间",
    "更新时间",
]


class ExcelRepositoryTests(unittest.TestCase):
    def test_initializes_workbook_with_required_sheets_and_defaults(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "assets.xlsx"

            repository = ExcelRepository(path)
            repository.initialize()

            workbook = load_workbook(path, read_only=True, data_only=True)
            self.assertEqual(
                workbook.sheetnames,
                [
                    "资产台账",
                    "存储介质",
                    "变更历史",
                    "分类字典",
                    "位置字典",
                    "枚举字典",
                    "系统配置",
                ],
            )
            self.assertEqual(workbook["系统配置"]["B2"].value, "2")
            self.assertEqual(
                [cell.value for cell in next(workbook["存储介质"].iter_rows(max_row=1))],
                STORAGE_HEADERS,
            )
            self.assertGreater(workbook["分类字典"].max_row, 7)
            self.assertGreater(workbook["枚举字典"].max_row, 7)
            workbook.close()

    def test_initialize_migrates_v1_workbook_and_creates_backup(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "assets.xlsx"
            workbook = Workbook()
            workbook.remove(workbook.active)
            ExcelRepository._create_sheet(
                workbook,
                "资产台账",
                V1_ASSET_HEADERS,
                [
                    [
                        "AST-2026-000001",
                        "BM-001",
                        "旧电脑",
                        "IT 与通信设备",
                        "电脑",
                        "联想",
                        "T14",
                        "SN-001",
                        "采购",
                        5999,
                        "2026-01-02",
                        "2026-01-05",
                        "在用",
                        "办公室",
                        "技术部",
                        "张三",
                        "旧备注",
                        "2026-01-02 10:00:00",
                        "2026-01-02 10:00:00",
                    ]
                ],
                "AssetsTable",
            )
            ExcelRepository._create_sheet(
                workbook, "变更历史", CHANGE_HEADERS, [], "ChangesTable"
            )
            ExcelRepository._create_sheet(
                workbook,
                "分类字典",
                CATEGORY_HEADERS,
                DEFAULT_CATEGORIES,
                "CategoriesTable",
            )
            ExcelRepository._create_sheet(
                workbook,
                "位置字典",
                LOCATION_HEADERS,
                DEFAULT_LOCATIONS,
                "LocationsTable",
            )
            ExcelRepository._create_sheet(
                workbook, "枚举字典", ENUM_HEADERS, DEFAULT_ENUMS, "EnumsTable"
            )
            v1_config = [
                ("工作簿版本", "1", "用于识别工作簿结构"),
                *[row for row in DEFAULT_CONFIG if row[0] != "工作簿版本"],
            ]
            ExcelRepository._create_sheet(
                workbook, "系统配置", CONFIG_HEADERS, v1_config, "ConfigTable"
            )
            workbook.save(path)
            workbook.close()

            repository = ExcelRepository(path)
            repository.initialize()

            migrated = load_workbook(path, read_only=True, data_only=True)
            self.assertIn("存储介质", migrated.sheetnames)
            self.assertEqual(migrated["系统配置"]["B2"].value, "2")
            row = next(migrated["资产台账"].iter_rows(min_row=2, values_only=True))
            headers = [
                cell.value for cell in next(migrated["资产台账"].iter_rows(max_row=1))
            ]
            values = dict(zip(headers, row))
            self.assertEqual(values["bm编码"], "BM-001")
            self.assertEqual(values["设备器材"], "旧电脑")
            self.assertEqual(values["产品型号"], "T14")
            self.assertEqual(values["管理部门"], "技术部")
            self.assertEqual(values["管理人"], "张三")
            migrated.close()
            self.assertEqual(len(list((path.parent / "backups").glob("assets-*.xlsx"))), 1)
            self.assertEqual(repository.validate_workbook(), [])

    def test_validate_workbook_reports_missing_sheet(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "assets.xlsx"
            repository = ExcelRepository(path)
            repository.initialize()

            workbook = load_workbook(path)
            del workbook["变更历史"]
            workbook.save(path)

            errors = repository.validate_workbook()

            self.assertIn("缺少工作表：变更历史", errors)

    def test_validate_workbook_reports_wrong_headers_and_duplicate_ids(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "assets.xlsx"
            repository = ExcelRepository(path)
            repository.initialize()
            workbook = load_workbook(path)
            sheet = workbook["资产台账"]
            sheet["C1"] = "错误字段"
            sheet.append(["AST-2026-000001", "EQ-1"])
            sheet.append(["AST-2026-000001", "EQ-2"])
            sheet.append(["", "", "无编码设备"])
            workbook.save(path)
            workbook.close()

            errors = repository.validate_workbook()

            self.assertIn("工作表字段不正确：资产台账", errors)
            self.assertIn("资产唯一标识符重复：AST-2026-000001", errors)
            self.assertIn("资产唯一标识符为空：第4行", errors)
            self.assertFalse(any("bm编码为空" in error for error in errors))

    def test_failed_atomic_save_preserves_original_workbook(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "assets.xlsx"
            repository = ExcelRepository(path)
            repository.initialize()
            workbook = load_workbook(path)

            with patch.object(workbook, "save", side_effect=RuntimeError("写入失败")):
                with self.assertRaisesRegex(RuntimeError, "写入失败"):
                    repository.save(workbook, create_backup=False)
            workbook.close()

            self.assertEqual(repository.validate_workbook(), [])

    def test_locked_workbook_prevents_save(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "assets.xlsx"
            repository = ExcelRepository(path)
            repository.initialize()
            workbook = load_workbook(path)

            with patch.object(Path, "open", side_effect=PermissionError("locked")):
                with self.assertRaises(WorkbookLockedError):
                    repository.save(workbook, create_backup=False)
            workbook.close()


if __name__ == "__main__":
    unittest.main()

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from asset_ledger.excel_repository import ExcelRepository, WorkbookLockedError


class ExcelRepositoryTests(unittest.TestCase):
    def test_initializes_workbook_with_required_sheets_and_defaults(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "assets.xlsx"

            repository = ExcelRepository(path)
            repository.initialize()

            workbook = load_workbook(path, read_only=True, data_only=True)
            self.assertEqual(
                workbook.sheetnames,
                ["资产台账", "变更历史", "分类字典", "位置字典", "枚举字典", "系统配置"],
            )
            self.assertEqual(workbook["系统配置"]["B2"].value, "1")
            self.assertGreater(workbook["分类字典"].max_row, 7)
            self.assertGreater(workbook["枚举字典"].max_row, 7)
            workbook.close()

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
            self.assertIn("资产ID重复：AST-2026-000001", errors)
            self.assertIn("资产ID为空：第4行", errors)
            self.assertIn("装备编码为空：第4行", errors)

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

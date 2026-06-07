import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from asset_ledger.asset_dialog import AssetDialog
from asset_ledger.asset_service import AssetService
from asset_ledger.excel_repository import ExcelRepository
from asset_ledger.main_window import MainWindow
from asset_ledger.settings_dialog import SettingsDialog

from tests.test_asset_service import sample_asset


class GuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.application = QApplication.instance() or QApplication([])

    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.path = Path(self.temporary_directory.name) / "assets.xlsx"
        self.repository = ExcelRepository(self.path)
        self.repository.initialize()
        self.service = AssetService(self.repository)

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def test_asset_dialog_round_trips_asset_data(self) -> None:
        asset = self.service.create_asset(sample_asset())

        dialog = AssetDialog(self.service, asset)
        result = dialog.asset_data()

        self.assertEqual(result["equipment_code"], "EQ-001")
        self.assertEqual(result["name"], "测试电脑")
        self.assertEqual(result["price"], 5999.0)
        self.assertFalse(dialog.asset_id_edit.isEnabled())

    def test_asset_dialog_uses_editable_department_and_owner_suggestions(self) -> None:
        asset = self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(
                equipment_code="EQ-002",
                serial_number="SN-002",
                department="行政部",
                owner="李四",
            )
        )

        dialog = AssetDialog(self.service, asset)

        self.assertTrue(dialog.department_combo.isEditable())
        self.assertTrue(dialog.owner_combo.isEditable())
        self.assertIn(
            "行政部",
            [dialog.department_combo.itemText(index) for index in range(dialog.department_combo.count())],
        )
        self.assertIn(
            "李四",
            [dialog.owner_combo.itemText(index) for index in range(dialog.owner_combo.count())],
        )
        dialog.department_combo.setCurrentText("新部门")
        dialog.owner_combo.setCurrentText("王五")
        self.assertEqual(dialog.asset_data()["department"], "新部门")
        self.assertEqual(dialog.asset_data()["owner"], "王五")
        dialog.owner_combo.setCurrentText("")
        self.assertEqual(dialog.asset_data()["owner"], "")

    def test_main_window_loads_asset_table_and_detail_tabs(self) -> None:
        self.service.create_asset(sample_asset())

        window = MainWindow(self.service, self.path)
        window.refresh_assets()
        window.asset_table.selectRow(0)
        self.application.processEvents()

        self.assertEqual(window.asset_table.rowCount(), 1)
        self.assertEqual(window.detail_tabs.count(), 2)
        self.assertIn("AST-", window.detail_text.toPlainText())
        self.assertIn("最近刷新", window.result_label.text())
        window.close()

    def test_main_window_filters_unassigned_owner(self) -> None:
        self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(equipment_code="EQ-002", serial_number="SN-002", owner="")
        )
        window = MainWindow(self.service, self.path)

        window.owner_filter.setCurrentIndex(window.owner_filter.findText("未分配责任人"))
        self.application.processEvents()

        self.assertEqual(window.asset_table.rowCount(), 1)
        self.assertEqual(window.asset_table.item(0, 1).text(), "EQ-002")
        window.close()

    def test_history_can_filter_by_changed_field(self) -> None:
        asset = self.service.create_asset(sample_asset())
        self.service.update_asset(
            asset.asset_id, sample_asset(status="维修", location="仓库"), "送修"
        )
        window = MainWindow(self.service, self.path)
        window._select_asset(asset.asset_id)
        self.application.processEvents()

        fields = [
            window.history_filter.itemText(index)
            for index in range(window.history_filter.count())
        ]
        self.assertIn("状态", fields)
        self.assertIn("位置", fields)
        self.assertNotIn("*", fields)

        window.history_filter.setCurrentText("位置")
        self.application.processEvents()

        self.assertIn("位置变更", window.history_text.toPlainText())
        self.assertNotIn("状态变更", window.history_text.toPlainText())
        window.close()

    def test_refresh_preserves_selection_and_shows_refresh_time(self) -> None:
        first = self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(equipment_code="EQ-002", serial_number="SN-002")
        )
        window = MainWindow(self.service, self.path)
        window._select_asset(first.asset_id)

        window.refresh_all()
        self.application.processEvents()

        self.assertEqual(window._selected_asset().asset_id, first.asset_id)
        self.assertIn("最近刷新", window.result_label.text())
        window.close()

    def test_refresh_preserves_unassigned_owner_filter(self) -> None:
        self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(equipment_code="EQ-002", serial_number="SN-002", owner="")
        )
        window = MainWindow(self.service, self.path)
        window.owner_filter.setCurrentIndex(window.owner_filter.findText("未分配责任人"))

        window.refresh_all()
        self.application.processEvents()

        self.assertEqual(window.owner_filter.currentText(), "未分配责任人")
        self.assertEqual(window.asset_table.rowCount(), 1)
        window.close()

    def test_refresh_preserves_category_filter(self) -> None:
        self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(
                equipment_code="EQ-002",
                serial_number="SN-002",
                primary_category="办公设备",
                secondary_category="打印机",
            )
        )
        window = MainWindow(self.service, self.path)
        computer_item = window.category_tree.topLevelItem(0).child(0)
        window.category_tree.setCurrentItem(computer_item)
        window.refresh_assets()

        window.refresh_all()
        self.application.processEvents()

        self.assertEqual(
            window.category_tree.currentItem().data(0, Qt.ItemDataRole.UserRole), "电脑"
        )
        self.assertEqual(window.asset_table.rowCount(), 1)
        window.close()

    def test_refresh_after_save_selects_asset_and_opens_history(self) -> None:
        asset = self.service.create_asset(sample_asset())
        window = MainWindow(self.service, self.path)
        self.service.update_asset(asset.asset_id, sample_asset(status="维修"), "送修")

        window.refresh_after_save(asset.asset_id)
        self.application.processEvents()

        self.assertEqual(window._selected_asset().asset_id, asset.asset_id)
        self.assertEqual(window.detail_tabs.currentIndex(), 1)
        self.assertIn("状态变更", window.history_text.toPlainText())
        window.close()

    def test_refresh_after_save_reveals_asset_that_no_longer_matches_filter(self) -> None:
        asset = self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(equipment_code="EQ-002", serial_number="SN-002")
        )
        window = MainWindow(self.service, self.path)
        window.status_filter.setCurrentIndex(window.status_filter.findText("在用"))
        self.service.update_asset(asset.asset_id, sample_asset(status="维修"), "送修")

        window.refresh_after_save(asset.asset_id)
        self.application.processEvents()

        self.assertEqual(window._selected_asset().asset_id, asset.asset_id)
        self.assertEqual(window.status_filter.currentText(), "全部状态")
        window.close()

    def test_main_window_defines_readable_selected_row_style(self) -> None:
        window = MainWindow(self.service, self.path)

        style = window.styleSheet().upper()

        self.assertIn("QTABLEWIDGET::ITEM:SELECTED", style)
        self.assertIn("#EAF4FB", style)
        self.assertIn("#17212B", style)
        window.close()

    def test_settings_dialog_loads_operator_and_dictionary_tabs(self) -> None:
        dialog = SettingsDialog(self.service, self.path)

        self.assertEqual(dialog.operator_edit.text(), "管理员")
        self.assertEqual(dialog.tabs.count(), 4)
        self.assertGreater(dialog.category_tree.topLevelItemCount(), 0)


if __name__ == "__main__":
    unittest.main()

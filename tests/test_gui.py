import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QMessageBox

from asset_ledger.asset_dialog import AssetDialog
from asset_ledger.asset_service import AssetService, CacheReloadAfterSaveError
from asset_ledger.excel_repository import ExcelRepository
from asset_ledger.main_window import MainWindow
from asset_ledger.settings_dialog import SettingsDialog
from asset_ledger.widgets import NoWheelComboBox, NoWheelDoubleSpinBox

from tests.test_asset_service import sample_asset, sample_storage


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
        asset = self.service.create_asset(sample_asset(asset_identifier="ZC-GUI-001"))

        dialog = AssetDialog(self.service, asset)
        result = dialog.asset_data()

        self.assertEqual(result["equipment_code"], "EQ-001")
        self.assertEqual(result["asset_identifier"], "ZC-GUI-001")
        self.assertEqual(result["name"], "测试电脑")
        self.assertEqual(result["price"], 5999.0)
        self.assertFalse(dialog.asset_id_edit.isEnabled())
        self.assertEqual(dialog.asset_identifier_edit.text(), "ZC-GUI-001")

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

    def test_new_asset_dialog_defaults_optional_choices_to_blank(self) -> None:
        dialog = AssetDialog(self.service)

        self.assertEqual(dialog.primary_category_combo.currentText(), "")
        self.assertEqual(dialog.source_combo.currentText(), "")
        self.assertEqual(dialog.status_combo.currentText(), "")
        self.assertEqual(dialog.equipment_code_edit.text(), "")
        self.assertEqual(dialog.asset_identifier_edit.text(), "")

    def test_asset_dialog_loads_storage_media_table(self) -> None:
        asset = self.service.create_asset(sample_asset(), [sample_storage()])

        dialog = AssetDialog(self.service, asset)

        self.assertEqual(dialog.storage_table.rowCount(), 1)
        self.assertEqual(dialog.storage_media_data()[0]["serial_number"], "SSD-SN-001")

    def test_no_wheel_controls_ignore_wheel_changes(self) -> None:
        combo = NoWheelComboBox()
        combo.addItems(["", "A", "B"])
        combo.setCurrentIndex(1)
        spin = NoWheelDoubleSpinBox()
        spin.setValue(10)

        class WheelEvent:
            def ignore(self) -> None:
                pass

        combo.wheelEvent(WheelEvent())
        spin.wheelEvent(WheelEvent())

        self.assertEqual(combo.currentIndex(), 1)
        self.assertEqual(spin.value(), 10)

    def test_main_window_loads_asset_table_and_detail_tabs(self) -> None:
        self.service.create_asset(sample_asset(asset_identifier="ZC-MAIN-001"))

        window = MainWindow(self.service, self.path)
        window.refresh_assets()
        window.asset_table.selectRow(0)
        self.application.processEvents()

        self.assertEqual(window.asset_table.rowCount(), 1)
        self.assertEqual(window.asset_table.horizontalHeaderItem(0).text(), "资产UID")
        self.assertEqual(window.asset_table.horizontalHeaderItem(2).text(), "资产唯一标识符")
        self.assertEqual(window.detail_tabs.count(), 2)
        self.assertIn("AST-", window.detail_text.toPlainText())
        self.assertIn("ZC-MAIN-001", window.detail_text.toPlainText())
        self.assertIn("最近刷新", window.result_label.text())
        window.close()

    def test_main_window_filters_unassigned_owner(self) -> None:
        self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(equipment_code="EQ-002", serial_number="SN-002", owner="")
        )
        window = MainWindow(self.service, self.path)

        window.owner_filter.setCurrentIndex(window.owner_filter.findText("未分配管理人"))
        self.application.processEvents()

        self.assertEqual(window.asset_table.rowCount(), 1)
        self.assertEqual(window.asset_table.item(0, 1).text(), "EQ-002")
        window.close()

    def test_main_window_filters_user_and_shows_storage_in_detail(self) -> None:
        first = self.service.create_asset(sample_asset(), [sample_storage()])
        self.service.create_asset(
            sample_asset(
                equipment_code="EQ-002",
                serial_number="SN-002",
                user="王五",
            )
        )
        window = MainWindow(self.service, self.path)

        window.user_filter.setCurrentIndex(window.user_filter.findText("王五"))
        self.application.processEvents()

        self.assertEqual(window.asset_table.rowCount(), 1)
        window.clear_filters()
        window._select_asset(first.asset_id)
        window.show_selected_asset()
        self.assertIn("SSD-SN-001", window.detail_text.toPlainText())
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
        self.assertIn("使用状态", fields)
        self.assertIn("存放地点", fields)
        self.assertNotIn("*", fields)

        window.history_filter.setCurrentText("存放地点")
        self.application.processEvents()

        self.assertIn("存放地点变更", window.history_text.toPlainText())
        self.assertNotIn("使用状态变更", window.history_text.toPlainText())
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

    def test_manual_refresh_reloads_cache_but_save_refresh_only_redraws(self) -> None:
        asset = self.service.create_asset(sample_asset())
        window = MainWindow(self.service, self.path)

        with patch.object(
            self.service, "reload_cache", wraps=self.service.reload_cache
        ) as reload:
            window.refresh_all()
            window.refresh_after_save(asset.asset_id)

        self.assertEqual(reload.call_count, 1)
        window.close()

    def test_failed_manual_refresh_keeps_current_table_visible(self) -> None:
        self.service.create_asset(sample_asset())
        window = MainWindow(self.service, self.path)

        with patch.object(
            self.service, "reload_cache", side_effect=RuntimeError("刷新失败")
        ), patch.object(QMessageBox, "critical") as critical:
            window.refresh_all()

        self.assertEqual(window.asset_table.rowCount(), 1)
        critical.assert_called_once()
        window.close()

    def test_cache_reload_after_save_error_uses_non_retrying_warning_title(self) -> None:
        window = MainWindow(self.service, self.path)

        with patch.object(QMessageBox, "critical") as critical:
            window._show_write_error(
                "新增失败",
                CacheReloadAfterSaveError("Excel 已保存，但缓存刷新失败"),
            )

        self.assertEqual(critical.call_args.args[1], "Excel 已保存，请勿重复操作")
        window.close()

    def test_refresh_preserves_unassigned_owner_filter(self) -> None:
        self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(equipment_code="EQ-002", serial_number="SN-002", owner="")
        )
        window = MainWindow(self.service, self.path)
        window.owner_filter.setCurrentIndex(window.owner_filter.findText("未分配管理人"))

        window.refresh_all()
        self.application.processEvents()

        self.assertEqual(window.owner_filter.currentText(), "未分配管理人")
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

    def test_clear_filters_removes_category_filter(self) -> None:
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
        self.assertEqual(window.asset_table.rowCount(), 1)

        window.clear_filters()
        self.application.processEvents()

        self.assertIsNone(window.category_tree.currentItem())
        self.assertEqual(window.asset_table.rowCount(), 2)
        window.close()

    def test_quick_all_view_clears_category_filter(self) -> None:
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
        self.assertEqual(window.asset_table.rowCount(), 1)

        window._quick_filter("全部")
        self.application.processEvents()

        self.assertIsNone(window.category_tree.currentItem())
        self.assertEqual(window.asset_table.rowCount(), 2)
        window.close()

    def test_refresh_after_save_selects_asset_and_opens_history(self) -> None:
        asset = self.service.create_asset(sample_asset())
        window = MainWindow(self.service, self.path)
        self.service.update_asset(asset.asset_id, sample_asset(status="维修"), "送修")

        window.refresh_after_save(asset.asset_id)
        self.application.processEvents()

        self.assertEqual(window._selected_asset().asset_id, asset.asset_id)
        self.assertEqual(window.detail_tabs.currentIndex(), 1)
        self.assertIn("使用状态变更", window.history_text.toPlainText())
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
        self.assertEqual(window.status_filter.currentText(), "全部使用状态")
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

    def test_settings_dialog_does_not_write_operator_to_old_workbook_when_switching(self) -> None:
        dialog = SettingsDialog(self.service, self.path)
        dialog.selected_path = self.path.parent / "other.xlsx"
        dialog.operator_edit.setText("新操作者")

        with patch.object(self.service, "set_operator") as set_operator:
            dialog._save_general()

        set_operator.assert_not_called()

    def test_settings_dialog_keeps_open_when_operator_save_fails(self) -> None:
        dialog = SettingsDialog(self.service, self.path)

        with patch.object(
            self.service, "set_operator", side_effect=RuntimeError("外部修改冲突")
        ), patch.object(dialog, "accept") as accept, patch.object(
            QMessageBox, "warning"
        ) as warning:
            dialog._save_general()

        accept.assert_not_called()
        warning.assert_called_once()


if __name__ == "__main__":
    unittest.main()

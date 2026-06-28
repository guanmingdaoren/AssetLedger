import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QLabel, QMessageBox

from asset_ledger.asset_dialog import AssetDialog, StorageMediumDialog
from asset_ledger.asset_service import AssetService, CacheReloadAfterSaveError
from asset_ledger.excel_repository import CATEGORY_SHEET, ExcelRepository
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
        asset = self.service.create_asset(
            sample_asset(asset_identifier="ZC-GUI-001", notes2="第二段界面备注")
        )

        dialog = AssetDialog(self.service, asset)
        result = dialog.asset_data()

        self.assertEqual(result["equipment_code"], "EQ-001")
        self.assertEqual(result["asset_identifier"], "ZC-GUI-001")
        self.assertEqual(result["name"], "测试电脑")
        self.assertEqual(result["primary_category"], "IT 与通信设备")
        self.assertEqual(result["secondary_category"], "电脑")
        self.assertEqual(result["category_path"], "IT 与通信设备 / 电脑")
        self.assertEqual(result["price"], 5999.0)
        self.assertEqual(result["notes"], "初始设备")
        self.assertEqual(result["notes2"], "第二段界面备注")
        self.assertEqual(dialog.notes1_edit.toPlainText(), "初始设备")
        self.assertEqual(dialog.notes2_edit.toPlainText(), "第二段界面备注")
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

    def test_asset_dialog_uses_consistent_label_column(self) -> None:
        dialog = AssetDialog(self.service)

        form_labels = dialog.findChildren(QLabel, "formLabel")
        widths = {label.minimumWidth() for label in form_labels}

        self.assertGreater(len(form_labels), 20)
        self.assertEqual(widths, {132})
        self.assertTrue(all(label.alignment() & Qt.AlignmentFlag.AlignLeft for label in form_labels))
        self.assertTrue(all(not (label.alignment() & Qt.AlignmentFlag.AlignRight) for label in form_labels))

    def test_new_asset_dialog_defaults_optional_choices_to_blank(self) -> None:
        dialog = AssetDialog(self.service)

        self.assertEqual(dialog.primary_category_combo.currentText(), "")
        self.assertEqual(dialog.secondary_category_combo.currentText(), "")
        self.assertEqual(dialog.source_combo.currentText(), "")
        self.assertTrue(dialog.source_combo.isEditable())
        self.assertTrue(dialog.location_combo.isEditable())
        self.assertEqual(dialog.status_combo.currentText(), "")
        self.assertEqual(dialog.equipment_code_edit.text(), "")
        self.assertEqual(dialog.asset_identifier_edit.text(), "")
        self.assertIn("建议填写", dialog.grade_label.text())
        self.assertIn("未打印", dialog.label_status_label.text())
        self.assertIn("#FFF7D6", dialog.label_status_label.styleSheet())
        self.assertGreaterEqual(dialog.notes1_edit.minimumHeight(), 120)
        self.assertGreaterEqual(dialog.notes2_edit.minimumHeight(), 120)

    def test_asset_dialog_has_two_large_note_fields(self) -> None:
        dialog = AssetDialog(self.service)
        labels = [label.text() for label in dialog.findChildren(QLabel, "formLabel")]

        self.assertIn("备注1", labels)
        self.assertIn("备注2", labels)
        self.assertGreater(dialog.notes1_edit.minimumHeight(), dialog.change_note_edit.minimumHeight())
        self.assertGreater(dialog.notes2_edit.minimumHeight(), dialog.change_note_edit.minimumHeight())

    def test_asset_dialog_supports_manual_source_location_and_grade_warning(self) -> None:
        dialog = AssetDialog(self.service)
        dialog.name_edit.setText("手工录入设备")
        dialog.source_combo.setCurrentText("自制")
        dialog.location_combo.setCurrentText("临时地点A")

        with patch.object(
            QMessageBox,
            "question",
            return_value=QMessageBox.StandardButton.Yes,
        ) as question, patch.object(dialog, "accept") as accept:
            dialog._validate_and_accept()

        self.assertEqual(dialog.asset_data()["source"], "自制")
        self.assertEqual(dialog.asset_data()["location"], "临时地点A")
        question.assert_called_once()
        accept.assert_called_once()

    def test_asset_dialog_loads_storage_media_table(self) -> None:
        asset = self.service.create_asset(sample_asset(), [sample_storage(status="在用")])

        dialog = AssetDialog(self.service, asset)

        self.assertEqual(dialog.storage_table.rowCount(), 1)
        self.assertEqual(dialog.storage_media_data()[0]["status"], "在用")
        self.assertEqual(dialog.storage_media_data()[0]["serial_number"], "SSD-SN-001")

    def test_asset_dialog_copy_mode_clears_unique_fields_and_storage_serials(self) -> None:
        asset = self.service.create_asset(
            sample_asset(asset_identifier="ZC-COPY-001"),
            [sample_storage(status="在用", serial_number="SSD-COPY-001")],
        )

        dialog = AssetDialog(self.service, asset, copy_mode=True)
        media = dialog.storage_media_data()

        self.assertIn("复制", dialog.windowTitle())
        self.assertEqual(dialog.asset_id_edit.text(), "")
        self.assertEqual(dialog.equipment_code_edit.text(), "")
        self.assertEqual(dialog.asset_identifier_edit.text(), "")
        self.assertEqual(dialog.serial_number_edit.text(), "")
        self.assertEqual(dialog.name_edit.text(), "测试电脑")
        self.assertEqual(dialog.model_edit.text(), "T14")
        self.assertEqual(len(media), 1)
        self.assertEqual(media[0]["storage_id"], "")
        self.assertEqual(media[0]["asset_id"], "")
        self.assertEqual(media[0]["serial_number"], "")
        self.assertEqual(media[0]["model"], "990 PRO")

    def test_storage_medium_dialog_edits_status(self) -> None:
        dialog = StorageMediumDialog(
            sample_storage(status="在用"),
            status_values=["入库", "在用", "报废"],
        )

        self.assertIn(
            "报废",
            [dialog.status_combo.itemText(index) for index in range(dialog.status_combo.count())],
        )
        dialog.status_combo.setCurrentText("报废")

        self.assertEqual(dialog.medium_data()["status"], "报废")

    def test_storage_medium_capacity_starts_blank_and_can_be_left_empty(self) -> None:
        dialog = StorageMediumDialog(status_values=["在用"])

        self.assertEqual(dialog.capacity_edit.text(), "")

        dialog.type_combo.setCurrentText("SSD")
        data = dialog.medium_data()

        self.assertEqual(data["capacity_value"], 0.0)
        self.assertEqual(dialog.capacity_edit.text(), "")

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
        self.assertEqual(window.asset_table.horizontalHeaderItem(0).text(), "序号")
        self.assertEqual(window.asset_table.horizontalHeaderItem(1).text(), "资产UID")
        self.assertEqual(window.asset_table.horizontalHeaderItem(3).text(), "资产唯一标识符")
        self.assertEqual(window.asset_table.item(0, 0).text(), "1")
        self.assertEqual(window.detail_tabs.count(), 2)
        self.assertIn("身份信息", window.detail_text.toPlainText())
        self.assertIn("管理与使用信息", window.detail_text.toPlainText())
        self.assertIn("资产概览", window.detail_text.toPlainText())
        self.assertIn("当前状态", window.detail_text.toPlainText())
        self.assertIn("#f7fafc", window.detail_text.toHtml().lower())
        self.assertIn("#f8fbf7", window.detail_text.toHtml().lower())
        self.assertIn("border-collapse", window.detail_text.toHtml().lower())
        self.assertIn('width="100%"', window.detail_text.toHtml().lower())
        self.assertIn("AST-", window.detail_text.toPlainText())
        self.assertIn("ZC-MAIN-001", window.detail_text.toPlainText())
        self.assertIn("最近刷新", window.result_label.text())
        window.close()

    def test_asset_table_context_menu_contains_expected_device_actions(self) -> None:
        asset = self.service.create_asset(sample_asset())

        window = MainWindow(self.service, self.path)
        window._select_asset(asset.asset_id)
        menu = window._build_asset_row_context_menu()

        actions = [action.text() for action in menu.actions() if not action.isSeparator()]
        self.assertEqual(
            actions,
            ["查看详情", "查看变化记录", "编辑设备", "复制为新设备", "删除设备", "刷新列表"],
        )
        self.assertEqual(
            window.asset_table.contextMenuPolicy(),
            Qt.ContextMenuPolicy.CustomContextMenu,
        )
        window.close()

    def test_asset_table_blank_context_menu_contains_safe_actions(self) -> None:
        window = MainWindow(self.service, self.path)

        menu = window._build_asset_blank_context_menu()

        actions = [action.text() for action in menu.actions() if not action.isSeparator()]
        self.assertEqual(actions, ["新增设备", "刷新列表"])
        window.close()

    def test_asset_table_context_menu_selects_row_under_cursor(self) -> None:
        self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(equipment_code="EQ-002", serial_number="SN-002")
        )
        window = MainWindow(self.service, self.path)
        window.asset_table.selectRow(0)
        target_item = window.asset_table.item(1, 0)
        target_asset_id = target_item.data(Qt.ItemDataRole.UserRole)
        position = window.asset_table.visualItemRect(target_item).center()

        selected = window._select_asset_row_at_context_position(position)

        self.assertTrue(selected)
        self.assertEqual(window._selected_asset().asset_id, target_asset_id)
        window.close()

    def test_context_menu_view_actions_switch_detail_tabs(self) -> None:
        asset = self.service.create_asset(sample_asset())
        self.service.update_asset(asset.asset_id, sample_asset(status="维修"), "送修")
        window = MainWindow(self.service, self.path)
        window._select_asset(asset.asset_id)
        menu = window._build_asset_row_context_menu()

        window.detail_tabs.setCurrentIndex(1)
        menu.actions()[0].trigger()
        self.assertEqual(window.detail_tabs.currentIndex(), 0)

        menu.actions()[1].trigger()
        self.assertEqual(window.detail_tabs.currentIndex(), 1)
        self.assertIn("使用状态变更", window.history_text.toPlainText())
        window.close()

    def test_delete_selected_asset_cancel_keeps_asset_and_shows_summary(self) -> None:
        asset = self.service.create_asset(sample_asset(), [sample_storage()])
        self.service.update_asset(asset.asset_id, sample_asset(status="维修"), "送修")
        window = MainWindow(self.service, self.path)
        window._select_asset(asset.asset_id)

        with patch.object(
            QMessageBox,
            "question",
            return_value=QMessageBox.StandardButton.No,
        ) as question:
            window.delete_selected_asset()

        self.assertEqual([item.asset_id for item in self.service.list_assets()], [asset.asset_id])
        message = question.call_args.args[2]
        self.assertIn(asset.asset_id, message)
        self.assertIn("1 个存储介质", message)
        self.assertIn("变化记录", message)
        window.close()

    def test_delete_selected_asset_removes_asset_and_refreshes_table(self) -> None:
        asset = self.service.create_asset(sample_asset(), [sample_storage()])
        window = MainWindow(self.service, self.path)
        window._select_asset(asset.asset_id)

        with patch.object(
            QMessageBox,
            "question",
            return_value=QMessageBox.StandardButton.Yes,
        ):
            window.delete_selected_asset()
        self.application.processEvents()

        self.assertEqual(self.service.list_assets(), [])
        self.assertEqual(window.asset_table.rowCount(), 0)
        self.assertIn("未选择设备", window.detail_text.toPlainText())
        self.assertEqual(window.history_text.toPlainText(), "")
        window.close()

    def test_main_window_filter_conditions_hide_owner_and_source(self) -> None:
        window = MainWindow(self.service, self.path)

        filter_labels = [
            combo.itemText(0)
            for combo in window.findChildren(NoWheelComboBox)
        ]

        self.assertIn("全部使用状态", filter_labels)
        self.assertIn("全部位置", filter_labels)
        self.assertIn("全部品牌", filter_labels)
        self.assertIn("全部使用人", filter_labels)
        self.assertNotIn("全部管理人", filter_labels)
        self.assertNotIn("全部取得方式", filter_labels)
        self.assertFalse(hasattr(window, "owner_filter"))
        self.assertFalse(hasattr(window, "source_filter"))
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

    def test_copy_selected_asset_creates_new_asset_selects_and_highlights_row(self) -> None:
        original = self.service.create_asset(
            sample_asset(asset_identifier="ZC-COPY-001"),
            [sample_storage(serial_number="SSD-COPY-001")],
        )
        window = MainWindow(self.service, self.path)
        window._select_asset(original.asset_id)

        class FakeCopyDialog:
            DialogCode = AssetDialog.DialogCode

            def __init__(self, service, asset, parent=None, copy_mode=False):
                self.asset = asset
                self.copy_mode = copy_mode

            def exec(self):
                return AssetDialog.DialogCode.Accepted

            def asset_data(self):
                data = sample_asset(
                    equipment_code="",
                    asset_identifier="",
                    serial_number="",
                )
                data["asset_id"] = ""
                return data

            def storage_media_data(self):
                medium = sample_storage(serial_number="")
                medium["storage_id"] = ""
                medium["asset_id"] = ""
                return [medium]

            def change_note(self):
                return "复制测试"

        with patch("asset_ledger.main_window.AssetDialog", FakeCopyDialog):
            window.copy_selected_asset()
        self.application.processEvents()

        selected = window._selected_asset()
        self.assertIsNotNone(selected)
        self.assertNotEqual(selected.asset_id, original.asset_id)
        self.assertEqual(selected.equipment_code, "")
        self.assertEqual(selected.asset_identifier, "")
        self.assertEqual(selected.serial_number, "")
        self.assertEqual(window.recently_copied_asset_id, selected.asset_id)
        self.assertIn("#FFF4CC", window.asset_table.styleSheet())
        selected_row = window.asset_table.currentRow()
        copied_background = window.asset_table.item(selected_row, 0).background().color().name().upper()
        self.assertEqual(copied_background, "#FFF4CC")
        media = self.service.list_storage_media(selected.asset_id)
        self.assertEqual(len(media), 1)
        self.assertEqual(media[0].serial_number, "")
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

    def test_history_uses_alternating_cards_grouped_by_save_event(self) -> None:
        asset = self.service.create_asset(sample_asset())
        self.service.update_asset(
            asset.asset_id,
            sample_asset(status="维修", location="仓库"),
            "送修",
        )
        self.service.update_asset(
            asset.asset_id,
            sample_asset(status="维修", location="仓库", price=5200),
            "金额调整",
        )
        window = MainWindow(self.service, self.path)
        window._select_asset(asset.asset_id)
        self.application.processEvents()

        text = window.history_text.toPlainText()
        html = window.history_text.toHtml().lower()

        self.assertIn("2 项变化", text)
        self.assertIn("1 项变化", text)
        self.assertIn("金额调整", text)
        self.assertIn("送修", text)
        self.assertIn("#f7fafc", html)
        self.assertIn("#f8fbf7", html)
        self.assertIn("border-top", html)
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

    def test_refresh_preserves_unassigned_user_filter(self) -> None:
        self.service.create_asset(sample_asset())
        self.service.create_asset(
            sample_asset(equipment_code="EQ-002", serial_number="SN-002", user="")
        )
        window = MainWindow(self.service, self.path)
        window.user_filter.setCurrentIndex(window.user_filter.findText("未分配使用人"))

        window.refresh_all()
        self.application.processEvents()

        self.assertEqual(window.user_filter.currentText(), "未分配使用人")
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
            window.category_tree.currentItem().data(0, Qt.ItemDataRole.UserRole),
            "IT 与通信设备 / 电脑",
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
        self.assertIn("QTABLEWIDGET::ITEM:SELECTED:ACTIVE", style)
        self.assertIn("QABSTRACTITEMVIEW::ITEM:SELECTED", style)
        self.assertIn("#EAF4FB", style)
        self.assertIn("#17212B", style)
        self.assertNotIn("#8E24AA", style)
        window.close()

    def test_settings_dialog_loads_operator_and_dictionary_tabs(self) -> None:
        dialog = SettingsDialog(self.service, self.path)

        self.assertEqual(dialog.operator_edit.text(), "管理员")
        self.assertEqual(dialog.tabs.count(), 4)
        self.assertGreater(dialog.category_tree.topLevelItemCount(), 0)

    def test_settings_dialog_deletes_selected_dictionary_item(self) -> None:
        dialog = SettingsDialog(self.service, self.path)
        item = dialog.category_tree.topLevelItem(0)
        item_id = str(item.data(0, Qt.ItemDataRole.UserRole))
        dialog.category_tree.setCurrentItem(item)

        with patch.object(dialog.service, "delete_dictionary_item") as delete:
            dialog._delete_dictionary(CATEGORY_SHEET, dialog.category_tree)

        delete.assert_called_once_with(CATEGORY_SHEET, item_id)

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

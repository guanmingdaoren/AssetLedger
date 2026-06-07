from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFormLayout,
    QGroupBox,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from .asset_service import AssetService
from .models import Asset


class AssetDialog(QDialog):
    def __init__(
        self, service: AssetService, asset: Asset | None = None, parent=None
    ) -> None:
        super().__init__(parent)
        self.service = service
        self.asset = asset
        self.setWindowTitle("编辑设备" if asset else "新增设备")
        self.setMinimumSize(680, 720)

        self.asset_id_edit = QLineEdit()
        self.asset_id_edit.setEnabled(False)
        self.asset_id_edit.setPlaceholderText("保存后自动生成")
        self.equipment_code_edit = QLineEdit()
        self.name_edit = QLineEdit()
        self.primary_category_combo = QComboBox()
        self.secondary_category_combo = QComboBox()
        self.brand_edit = QLineEdit()
        self.model_edit = QLineEdit()
        self.serial_number_edit = QLineEdit()
        self.source_combo = QComboBox()
        self.price_spin = QDoubleSpinBox()
        self.price_spin.setRange(0, 999_999_999)
        self.price_spin.setDecimals(2)
        self.price_spin.setPrefix("¥ ")
        self.purchase_date_edit = QLineEdit()
        self.purchase_date_edit.setPlaceholderText("YYYY-MM-DD")
        self.start_date_edit = QLineEdit()
        self.start_date_edit.setPlaceholderText("YYYY-MM-DD")
        self.status_combo = QComboBox()
        self.location_combo = QComboBox()
        self.department_combo = QComboBox()
        self.department_combo.setEditable(True)
        self.department_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.owner_combo = QComboBox()
        self.owner_combo.setEditable(True)
        self.owner_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.notes_edit = QPlainTextEdit()
        self.notes_edit.setMaximumHeight(90)
        self.change_note_edit = QLineEdit()
        self.change_note_edit.setPlaceholderText("可选，用于说明本次修改原因")

        self._load_options()
        self.primary_category_combo.currentTextChanged.connect(self._load_secondary_categories)
        self._build_ui()
        if asset:
            self._populate(asset)
        else:
            self._load_secondary_categories(self.primary_category_combo.currentText())

    def _build_ui(self) -> None:
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setSpacing(12)

        identity = self._group(
            "身份信息",
            [
                ("资产 ID", self.asset_id_edit),
                ("装备编码 *", self.equipment_code_edit),
                ("设备名称 *", self.name_edit),
                ("一级类别", self.primary_category_combo),
                ("二级类别", self.secondary_category_combo),
            ],
        )
        product = self._group(
            "产品信息",
            [
                ("品牌", self.brand_edit),
                ("型号", self.model_edit),
                ("序列号", self.serial_number_edit),
                ("来源", self.source_combo),
            ],
        )
        asset_info = self._group(
            "资产信息",
            [
                ("价格", self.price_spin),
                ("购入日期", self.purchase_date_edit),
                ("启用日期", self.start_date_edit),
            ],
        )
        management = self._group(
            "管理信息",
            [
                ("状态", self.status_combo),
                ("位置", self.location_combo),
                ("责任部门", self.department_combo),
                ("责任人", self.owner_combo),
            ],
        )
        extra = self._group(
            "补充信息",
            [("备注", self.notes_edit), ("修改说明", self.change_note_edit)],
        )
        for group in (identity, product, asset_info, management, extra):
            content_layout.addWidget(group)
        content_layout.addStretch()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        scroll.setWidget(content)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("保存")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(scroll)
        layout.addWidget(buttons)

    @staticmethod
    def _group(title: str, rows: list[tuple[str, QWidget]]) -> QGroupBox:
        group = QGroupBox(title)
        form = QFormLayout(group)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form.setHorizontalSpacing(16)
        form.setVerticalSpacing(9)
        for label, widget in rows:
            form.addRow(QLabel(label), widget)
        return group

    def _load_options(self) -> None:
        categories = self.service.get_categories()
        primary = [item for item in categories if not item.parent_id]
        self.primary_category_combo.addItems([item.name for item in primary])
        self.source_combo.addItems(self.service.get_enum_values("来源"))
        self.status_combo.addItems(self.service.get_enum_values("状态"))
        self.location_combo.addItem("")
        self.location_combo.addItems([item.name for item in self.service.get_locations()])
        self.department_combo.addItem("")
        self.department_combo.addItems(self.service.list_departments())
        self.owner_combo.addItem("")
        self.owner_combo.addItems(self.service.list_owners())

    def _load_secondary_categories(self, primary_name: str) -> None:
        current = self.secondary_category_combo.currentText()
        categories = self.service.get_categories(include_disabled=True)
        parent_id = next(
            (item.item_id for item in categories if item.name == primary_name and not item.parent_id),
            "",
        )
        children = [item.name for item in categories if item.parent_id == parent_id and item.enabled]
        self.secondary_category_combo.blockSignals(True)
        self.secondary_category_combo.clear()
        self.secondary_category_combo.addItem("")
        self.secondary_category_combo.addItems(children)
        self.secondary_category_combo.setCurrentText(current)
        self.secondary_category_combo.blockSignals(False)

    def _populate(self, asset: Asset) -> None:
        self.asset_id_edit.setText(asset.asset_id)
        self.equipment_code_edit.setText(asset.equipment_code)
        self.name_edit.setText(asset.name)
        self.primary_category_combo.setCurrentText(asset.primary_category)
        self._load_secondary_categories(asset.primary_category)
        self.secondary_category_combo.setCurrentText(asset.secondary_category)
        self.brand_edit.setText(asset.brand)
        self.model_edit.setText(asset.model)
        self.serial_number_edit.setText(asset.serial_number)
        self.source_combo.setCurrentText(asset.source)
        self.price_spin.setValue(asset.price)
        self.purchase_date_edit.setText(asset.purchase_date)
        self.start_date_edit.setText(asset.start_date)
        self.status_combo.setCurrentText(asset.status)
        self.location_combo.setCurrentText(asset.location)
        self.department_combo.setCurrentText(asset.department)
        self.owner_combo.setCurrentText(asset.owner)
        self.notes_edit.setPlainText(asset.notes)

    def asset_data(self) -> dict[str, object]:
        return {
            "asset_id": self.asset_id_edit.text().strip(),
            "equipment_code": self.equipment_code_edit.text().strip(),
            "name": self.name_edit.text().strip(),
            "primary_category": self.primary_category_combo.currentText(),
            "secondary_category": self.secondary_category_combo.currentText(),
            "brand": self.brand_edit.text().strip(),
            "model": self.model_edit.text().strip(),
            "serial_number": self.serial_number_edit.text().strip(),
            "source": self.source_combo.currentText(),
            "price": self.price_spin.value(),
            "purchase_date": self.purchase_date_edit.text().strip(),
            "start_date": self.start_date_edit.text().strip(),
            "status": self.status_combo.currentText(),
            "location": self.location_combo.currentText(),
            "department": self.department_combo.currentText().strip(),
            "owner": self.owner_combo.currentText().strip(),
            "notes": self.notes_edit.toPlainText().strip(),
        }

    def change_note(self) -> str:
        return self.change_note_edit.text().strip()

    def _validate_and_accept(self) -> None:
        if not self.equipment_code_edit.text().strip():
            QMessageBox.warning(self, "无法保存", "装备编码不能为空。")
            self.equipment_code_edit.setFocus()
            return
        if not self.name_edit.text().strip():
            QMessageBox.warning(self, "无法保存", "设备名称不能为空。")
            self.name_edit.setFocus()
            return
        self.accept()

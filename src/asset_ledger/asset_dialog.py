from __future__ import annotations

from dataclasses import asdict

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .asset_service import AssetService
from .models import Asset, StorageMedium
from .widgets import NoWheelComboBox, NoWheelDoubleSpinBox


STORAGE_COLUMNS = [
    ("medium_type", "介质类型"),
    ("name", "名称/编号"),
    ("brand", "品牌"),
    ("capacity", "容量"),
    ("model", "型号"),
    ("serial_number", "序列号"),
]


class StorageMediumDialog(QDialog):
    def __init__(self, medium: dict[str, object] | StorageMedium | None = None, parent=None):
        super().__init__(parent)
        self.original = (
            asdict(medium)
            if isinstance(medium, StorageMedium)
            else dict(medium or {})
        )
        self.setWindowTitle("编辑存储介质" if medium else "添加存储介质")
        self.setMinimumWidth(520)

        self.type_combo = NoWheelComboBox()
        self.type_combo.setEditable(True)
        self.type_combo.addItems(["", "SSD", "机械硬盘", "其他"])
        self.name_edit = QLineEdit()
        self.brand_edit = QLineEdit()
        self.capacity_spin = NoWheelDoubleSpinBox()
        self.capacity_spin.setRange(0, 999_999)
        self.capacity_spin.setDecimals(2)
        self.capacity_unit_combo = NoWheelComboBox()
        self.capacity_unit_combo.addItems(["", "GB", "TB"])
        self.model_edit = QLineEdit()
        self.serial_edit = QLineEdit()
        self.notes_edit = QPlainTextEdit()
        self.notes_edit.setMaximumHeight(80)

        form = QFormLayout()
        for label, widget in (
            ("介质类型", self.type_combo),
            ("名称/编号", self.name_edit),
            ("品牌", self.brand_edit),
            ("容量数值", self.capacity_spin),
            ("容量单位", self.capacity_unit_combo),
            ("型号", self.model_edit),
            ("序列号", self.serial_edit),
            ("备注", self.notes_edit),
        ):
            form.addRow(label, widget)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("保存")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)
        if medium:
            self._populate()

    def _populate(self) -> None:
        self.type_combo.setCurrentText(str(self.original.get("medium_type", "")))
        self.name_edit.setText(str(self.original.get("name", "")))
        self.brand_edit.setText(str(self.original.get("brand", "")))
        self.capacity_spin.setValue(float(self.original.get("capacity_value", 0) or 0))
        self.capacity_unit_combo.setCurrentText(
            str(self.original.get("capacity_unit", ""))
        )
        self.model_edit.setText(str(self.original.get("model", "")))
        self.serial_edit.setText(str(self.original.get("serial_number", "")))
        self.notes_edit.setPlainText(str(self.original.get("notes", "")))

    def medium_data(self) -> dict[str, object]:
        return {
            **self.original,
            "medium_type": self.type_combo.currentText().strip(),
            "name": self.name_edit.text().strip(),
            "brand": self.brand_edit.text().strip(),
            "capacity_value": self.capacity_spin.value(),
            "capacity_unit": self.capacity_unit_combo.currentText(),
            "model": self.model_edit.text().strip(),
            "serial_number": self.serial_edit.text().strip(),
            "notes": self.notes_edit.toPlainText().strip(),
        }

    def _validate_and_accept(self) -> None:
        data = self.medium_data()
        if not any(
            str(data.get(key, "") or "").strip()
            for key in (
                "medium_type",
                "name",
                "brand",
                "capacity_value",
                "capacity_unit",
                "model",
                "serial_number",
                "notes",
            )
        ):
            QMessageBox.warning(self, "无法保存", "存储介质内容不能全部为空。")
            return
        self.accept()


class AssetDialog(QDialog):
    def __init__(
        self, service: AssetService, asset: Asset | None = None, parent=None
    ) -> None:
        super().__init__(parent)
        self.service = service
        self.asset = asset
        self.storage_media: list[dict[str, object]] = []
        self.setWindowTitle("编辑设备" if asset else "新增设备")
        self.setMinimumSize(760, 780)

        self.asset_id_edit = QLineEdit()
        self.asset_id_edit.setEnabled(False)
        self.asset_id_edit.setPlaceholderText("保存后自动生成")
        self.equipment_code_edit = QLineEdit()
        self.asset_identifier_edit = QLineEdit()
        self.name_edit = QLineEdit()
        self.primary_category_combo = self._choice_combo()
        self.secondary_category_combo = self._choice_combo()
        self.product_spec_edit = QLineEdit()
        self.model_edit = QLineEdit()
        self.serial_number_edit = QLineEdit()
        self.manufacturer_edit = QLineEdit()
        self.supplier_edit = QLineEdit()
        self.brand_edit = QLineEdit()
        self.source_combo = self._choice_combo()
        self.purchase_date_edit = self._date_edit()
        self.start_date_edit = self._date_edit()
        self.valuation_method_combo = self._editable_combo()
        self.price_spin = NoWheelDoubleSpinBox()
        self.price_spin.setRange(0, 999_999_999)
        self.price_spin.setDecimals(2)
        self.price_spin.setPrefix("¥ ")
        self.department_combo = self._editable_combo()
        self.owner_combo = self._editable_combo()
        self.use_department_combo = self._editable_combo()
        self.user_combo = self._editable_combo()
        self.location_combo = self._choice_combo()
        self.manufacture_date_edit = self._date_edit()
        self.grade_combo = self._editable_combo()
        self.status_combo = self._choice_combo()
        self.label_printed_check = QCheckBox("已打印")
        self.notes_edit = QPlainTextEdit()
        self.notes_edit.setMaximumHeight(90)
        self.change_note_edit = QLineEdit()
        self.change_note_edit.setPlaceholderText("可选，用于说明本次修改原因")
        self.storage_table = QTableWidget(0, len(STORAGE_COLUMNS))
        self.storage_table.setHorizontalHeaderLabels(
            [label for _, label in STORAGE_COLUMNS]
        )
        self.storage_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.storage_table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self.storage_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.storage_table.verticalHeader().setVisible(False)
        self.storage_table.setMinimumHeight(150)

        self._load_options()
        self.primary_category_combo.currentTextChanged.connect(
            self._load_secondary_categories
        )
        self._build_ui()
        if asset:
            self._populate(asset)
        else:
            self._load_secondary_categories("")

    @staticmethod
    def _choice_combo() -> NoWheelComboBox:
        combo = NoWheelComboBox()
        combo.addItem("")
        return combo

    @staticmethod
    def _editable_combo() -> NoWheelComboBox:
        combo = NoWheelComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(NoWheelComboBox.InsertPolicy.NoInsert)
        combo.addItem("")
        return combo

    @staticmethod
    def _date_edit() -> QLineEdit:
        edit = QLineEdit()
        edit.setPlaceholderText("YYYY-MM-DD")
        return edit

    def _build_ui(self) -> None:
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setSpacing(12)
        groups = [
            self._group(
                "身份与分类信息",
                [
                    ("资产UID", self.asset_id_edit),
                    ("bm编码", self.equipment_code_edit),
                    ("资产唯一标识符", self.asset_identifier_edit),
                    ("设备器材 *", self.name_edit),
                    ("一级类别", self.primary_category_combo),
                    ("二级类别", self.secondary_category_combo),
                ],
            ),
            self._group(
                "产品信息",
                [
                    ("产品规格", self.product_spec_edit),
                    ("产品型号", self.model_edit),
                    ("设备序列号", self.serial_number_edit),
                    ("生产厂家", self.manufacturer_edit),
                    ("供应商", self.supplier_edit),
                    ("品牌", self.brand_edit),
                ],
            ),
            self._group(
                "取得与资产信息",
                [
                    ("取得方式", self.source_combo),
                    ("取得日期", self.purchase_date_edit),
                    ("启用日期", self.start_date_edit),
                    ("计价方式", self.valuation_method_combo),
                    ("金额", self.price_spin),
                    ("出厂日期", self.manufacture_date_edit),
                    ("等级", self.grade_combo),
                    ("是否已打印标签", self.label_printed_check),
                ],
            ),
            self._group(
                "管理与使用信息",
                [
                    ("管理部门", self.department_combo),
                    ("管理人", self.owner_combo),
                    ("使用部门", self.use_department_combo),
                    ("使用人", self.user_combo),
                    ("存放地点", self.location_combo),
                    ("使用状态", self.status_combo),
                ],
            ),
            self._storage_group(),
            self._group(
                "备注与修改说明",
                [("备注", self.notes_edit), ("修改说明", self.change_note_edit)],
            ),
        ]
        for group in groups:
            content_layout.addWidget(group)
        content_layout.addStretch()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
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

    def _storage_group(self) -> QGroupBox:
        group = QGroupBox("存储介质")
        layout = QVBoxLayout(group)
        layout.addWidget(self.storage_table)
        controls = QHBoxLayout()
        for text, slot in (
            ("添加存储介质", self._add_storage),
            ("编辑", self._edit_storage),
            ("移除", self._remove_storage),
        ):
            button = QPushButton(text)
            button.clicked.connect(slot)
            controls.addWidget(button)
        controls.addStretch()
        layout.addLayout(controls)
        return group

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
        self.primary_category_combo.addItems(
            [item.name for item in categories if not item.parent_id]
        )
        self.source_combo.addItems(self.service.get_enum_values("取得方式"))
        self.status_combo.addItems(self.service.get_enum_values("使用状态"))
        self.location_combo.addItems([item.name for item in self.service.get_locations()])
        self.department_combo.addItems(self.service.list_departments())
        self.owner_combo.addItems(self.service.list_owners())
        self.use_department_combo.addItems(self.service.list_use_departments())
        self.user_combo.addItems(self.service.list_users())
        self.grade_combo.addItems(self.service.list_grades())
        self.valuation_method_combo.addItems(self.service.list_valuation_methods())

    def _load_secondary_categories(self, primary_name: str) -> None:
        current = self.secondary_category_combo.currentText()
        categories = self.service.get_categories(include_disabled=True)
        parent_id = next(
            (
                item.item_id
                for item in categories
                if item.name == primary_name and not item.parent_id
            ),
            "",
        )
        children = [
            item.name
            for item in categories
            if parent_id and item.parent_id == parent_id and item.enabled
        ]
        self.secondary_category_combo.blockSignals(True)
        self.secondary_category_combo.clear()
        self.secondary_category_combo.addItem("")
        self.secondary_category_combo.addItems(children)
        self.secondary_category_combo.setCurrentText(current)
        self.secondary_category_combo.blockSignals(False)

    def _populate(self, asset: Asset) -> None:
        for widget, value in (
            (self.asset_id_edit, asset.asset_id),
            (self.equipment_code_edit, asset.equipment_code),
            (self.asset_identifier_edit, asset.asset_identifier),
            (self.name_edit, asset.name),
            (self.product_spec_edit, asset.product_spec),
            (self.model_edit, asset.model),
            (self.serial_number_edit, asset.serial_number),
            (self.manufacturer_edit, asset.manufacturer),
            (self.supplier_edit, asset.supplier),
            (self.brand_edit, asset.brand),
            (self.purchase_date_edit, asset.purchase_date),
            (self.start_date_edit, asset.start_date),
            (self.manufacture_date_edit, asset.manufacture_date),
        ):
            widget.setText(value)
        for combo, value in (
            (self.primary_category_combo, asset.primary_category),
            (self.source_combo, asset.source),
            (self.valuation_method_combo, asset.valuation_method),
            (self.department_combo, asset.department),
            (self.owner_combo, asset.owner),
            (self.use_department_combo, asset.use_department),
            (self.user_combo, asset.user),
            (self.location_combo, asset.location),
            (self.grade_combo, asset.grade),
            (self.status_combo, asset.status),
        ):
            combo.setCurrentText(value)
        self._load_secondary_categories(asset.primary_category)
        self.secondary_category_combo.setCurrentText(asset.secondary_category)
        self.price_spin.setValue(asset.price)
        self.label_printed_check.setChecked(asset.label_printed)
        self.notes_edit.setPlainText(asset.notes)
        self.storage_media = [
            asdict(medium) for medium in self.service.list_storage_media(asset.asset_id)
        ]
        self._refresh_storage_table()

    def _add_storage(self) -> None:
        dialog = StorageMediumDialog(parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.storage_media.append(dialog.medium_data())
            self._refresh_storage_table()

    def _edit_storage(self) -> None:
        row = self.storage_table.currentRow()
        if row < 0:
            return
        dialog = StorageMediumDialog(self.storage_media[row], self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.storage_media[row] = dialog.medium_data()
            self._refresh_storage_table()

    def _remove_storage(self) -> None:
        row = self.storage_table.currentRow()
        if row >= 0:
            self.storage_media.pop(row)
            self._refresh_storage_table()

    def _refresh_storage_table(self) -> None:
        self.storage_table.setRowCount(len(self.storage_media))
        for row, medium in enumerate(self.storage_media):
            capacity = ""
            if medium.get("capacity_value"):
                capacity = (
                    f"{float(medium['capacity_value']):g}"
                    f" {medium.get('capacity_unit', '')}"
                ).strip()
            values = [
                medium.get("medium_type", ""),
                medium.get("name", ""),
                medium.get("brand", ""),
                capacity,
                medium.get("model", ""),
                medium.get("serial_number", ""),
            ]
            for column, value in enumerate(values):
                self.storage_table.setItem(row, column, QTableWidgetItem(str(value)))

    def storage_media_data(self) -> list[dict[str, object]]:
        return [dict(item) for item in self.storage_media]

    def asset_data(self) -> dict[str, object]:
        return {
            "asset_id": self.asset_id_edit.text().strip(),
            "equipment_code": self.equipment_code_edit.text().strip(),
            "asset_identifier": self.asset_identifier_edit.text().strip(),
            "name": self.name_edit.text().strip(),
            "primary_category": self.primary_category_combo.currentText(),
            "secondary_category": self.secondary_category_combo.currentText(),
            "product_spec": self.product_spec_edit.text().strip(),
            "model": self.model_edit.text().strip(),
            "serial_number": self.serial_number_edit.text().strip(),
            "manufacturer": self.manufacturer_edit.text().strip(),
            "supplier": self.supplier_edit.text().strip(),
            "brand": self.brand_edit.text().strip(),
            "source": self.source_combo.currentText(),
            "purchase_date": self.purchase_date_edit.text().strip(),
            "start_date": self.start_date_edit.text().strip(),
            "valuation_method": self.valuation_method_combo.currentText().strip(),
            "price": self.price_spin.value(),
            "department": self.department_combo.currentText().strip(),
            "owner": self.owner_combo.currentText().strip(),
            "use_department": self.use_department_combo.currentText().strip(),
            "user": self.user_combo.currentText().strip(),
            "location": self.location_combo.currentText(),
            "manufacture_date": self.manufacture_date_edit.text().strip(),
            "grade": self.grade_combo.currentText().strip(),
            "status": self.status_combo.currentText(),
            "label_printed": self.label_printed_check.isChecked(),
            "notes": self.notes_edit.toPlainText().strip(),
        }

    def change_note(self) -> str:
        return self.change_note_edit.text().strip()

    def _validate_and_accept(self) -> None:
        if not self.name_edit.text().strip():
            QMessageBox.warning(self, "无法保存", "设备器材不能为空。")
            self.name_edit.setFocus()
            return
        self.accept()

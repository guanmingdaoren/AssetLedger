from __future__ import annotations

from dataclasses import asdict

from PySide6.QtCore import Qt
from PySide6.QtGui import QDoubleValidator
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
    QSizePolicy,
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
    ("status", "使用状态"),
    ("name", "名称/编号"),
    ("brand", "品牌"),
    ("capacity", "容量"),
    ("model", "型号"),
    ("serial_number", "序列号"),
]


class StorageMediumDialog(QDialog):
    def __init__(
        self,
        medium: dict[str, object] | StorageMedium | None = None,
        parent=None,
        status_values: list[str] | None = None,
    ):
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
        self.status_combo = NoWheelComboBox()
        self.status_combo.addItem("")
        self.status_combo.addItems(status_values or [])
        self.name_edit = QLineEdit()
        self.brand_edit = QLineEdit()
        self.capacity_edit = QLineEdit()
        self.capacity_edit.setPlaceholderText("可留空")
        capacity_validator = QDoubleValidator(0, 999_999, 2, self.capacity_edit)
        capacity_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
        self.capacity_edit.setValidator(capacity_validator)
        self.capacity_unit_combo = NoWheelComboBox()
        self.capacity_unit_combo.addItems(["", "GB", "TB"])
        self.model_edit = QLineEdit()
        self.serial_edit = QLineEdit()
        self.notes_edit = QPlainTextEdit()
        self.notes_edit.setMaximumHeight(80)

        form = QFormLayout()
        for label, widget in (
            ("介质类型", self.type_combo),
            ("使用状态", self.status_combo),
            ("名称/编号", self.name_edit),
            ("品牌", self.brand_edit),
            ("容量数值", self.capacity_edit),
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
        self.status_combo.setCurrentText(str(self.original.get("status", "")))
        self.name_edit.setText(str(self.original.get("name", "")))
        self.brand_edit.setText(str(self.original.get("brand", "")))
        capacity_value = float(self.original.get("capacity_value", 0) or 0)
        self.capacity_edit.setText(f"{capacity_value:g}" if capacity_value else "")
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
            "status": self.status_combo.currentText().strip(),
            "name": self.name_edit.text().strip(),
            "brand": self.brand_edit.text().strip(),
            "capacity_value": self._capacity_value(),
            "capacity_unit": self.capacity_unit_combo.currentText(),
            "model": self.model_edit.text().strip(),
            "serial_number": self.serial_edit.text().strip(),
            "notes": self.notes_edit.toPlainText().strip(),
        }

    def _validate_and_accept(self) -> None:
        if self.capacity_edit.text().strip() and not self._capacity_text_is_valid():
            QMessageBox.warning(self, "无法保存", "容量数值必须是非负数字。")
            self.capacity_edit.setFocus()
            return
        data = self.medium_data()
        if not any(
            str(data.get(key, "") or "").strip()
            for key in (
                "medium_type",
                "status",
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

    def _capacity_value(self) -> float:
        text = self.capacity_edit.text().strip()
        return float(text) if text else 0.0

    def _capacity_text_is_valid(self) -> bool:
        text = self.capacity_edit.text().strip()
        if not text:
            return True
        state, _, _ = self.capacity_edit.validator().validate(text, 0)
        return state == QDoubleValidator.State.Acceptable


class AssetDialog(QDialog):
    def __init__(
        self,
        service: AssetService,
        asset: Asset | None = None,
        parent=None,
        copy_mode: bool = False,
    ) -> None:
        super().__init__(parent)
        self.service = service
        self.asset = asset
        self.copy_mode = copy_mode
        self.storage_media: list[dict[str, object]] = []
        self.setWindowTitle("复制设备" if copy_mode else "编辑设备" if asset else "新增设备")
        self.setMinimumSize(760, 780)

        self.asset_id_edit = QLineEdit()
        self.asset_id_edit.setEnabled(False)
        self.asset_id_edit.setPlaceholderText("保存后自动生成")
        self.equipment_code_edit = QLineEdit()
        self.asset_identifier_edit = QLineEdit()
        self.name_edit = QLineEdit()
        self.category_path_combo = self._choice_combo()
        self.primary_category_combo = self._choice_combo()
        self.secondary_category_combo = self._choice_combo()
        self.product_spec_edit = QLineEdit()
        self.model_edit = QLineEdit()
        self.serial_number_edit = QLineEdit()
        self.manufacturer_edit = QLineEdit()
        self.supplier_edit = QLineEdit()
        self.brand_edit = QLineEdit()
        self.source_combo = self._editable_combo()
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
        self.location_combo = self._editable_combo()
        self.manufacture_date_edit = self._date_edit()
        self.grade_combo = self._editable_combo()
        self.grade_label = QLabel("等级（建议填写）")
        self.status_combo = self._choice_combo()
        self.label_printed_check = QCheckBox("已打印标签")
        self.label_status_label = QLabel()
        self.notes1_edit = self._large_notes_edit()
        self.notes2_edit = self._large_notes_edit()
        self.notes_edit = self.notes1_edit
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
        self.grade_combo.currentTextChanged.connect(self._update_grade_hint)
        self.label_printed_check.toggled.connect(self._update_label_status_hint)
        self._build_ui()
        self._apply_style()
        if asset and copy_mode:
            self._populate_copy(asset)
        elif asset:
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

    @staticmethod
    def _large_notes_edit() -> QPlainTextEdit:
        edit = QPlainTextEdit()
        edit.setObjectName("largeNotesEdit")
        edit.setMinimumHeight(130)
        edit.setMaximumHeight(190)
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
                    ("一级分类", self.primary_category_combo),
                    ("二级分类", self.secondary_category_combo),
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
                    (self.grade_label, self.grade_combo),
                    ("标签状态", self._label_printed_widget()),
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
                [
                    ("备注1", self.notes1_edit),
                    ("备注2", self.notes2_edit),
                    ("修改说明", self.change_note_edit),
                ],
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
        group.setObjectName("formGroup")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(20, 22, 20, 16)
        layout.setSpacing(10)
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

    def _label_printed_widget(self) -> QWidget:
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        layout.addWidget(self.label_printed_check)
        layout.addWidget(self.label_status_label)
        layout.addStretch()
        self._update_label_status_hint()
        return widget

    @staticmethod
    def _group(title: str, rows: list[tuple[str | QLabel, QWidget]]) -> QGroupBox:
        group = QGroupBox(title)
        group.setObjectName("formGroup")
        form = QFormLayout(group)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        form.setContentsMargins(20, 22, 20, 16)
        form.setHorizontalSpacing(18)
        form.setVerticalSpacing(11)
        for label, widget in rows:
            if isinstance(label, QLabel):
                label_widget = label
            else:
                label_widget = QLabel(label)
            label_widget.setObjectName("formLabel")
            label_widget.setMinimumWidth(132)
            label_widget.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            widget.setSizePolicy(QSizePolicy.Policy.Expanding, widget.sizePolicy().verticalPolicy())
            form.addRow(label_widget, widget)
        return group

    def _load_options(self) -> None:
        categories = self.service.get_categories()
        self.category_path_combo.addItems(self._category_paths(categories))
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
        self._update_grade_hint()
        self._update_label_status_hint()

    @staticmethod
    def _category_paths(categories) -> list[str]:
        by_id = {item.item_id: item for item in categories}
        paths: list[str] = []
        for item in categories:
            names: list[str] = []
            current = item
            seen: set[str] = set()
            while current and current.item_id not in seen:
                seen.add(current.item_id)
                names.append(current.name)
                current = by_id.get(current.parent_id)
            paths.append(" / ".join(reversed(names)))
        return sorted(dict.fromkeys(path for path in paths if path))

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
            (self.category_path_combo, asset.category_path),
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
        self.notes1_edit.setPlainText(asset.notes)
        self.notes2_edit.setPlainText(asset.notes2)
        self._update_grade_hint()
        self.storage_media = [
            asdict(medium) for medium in self.service.list_storage_media(asset.asset_id)
        ]
        self._refresh_storage_table()

    def _populate_copy(self, asset: Asset) -> None:
        self._populate(asset)
        self.asset_id_edit.clear()
        self.equipment_code_edit.clear()
        self.asset_identifier_edit.clear()
        self.serial_number_edit.clear()
        for medium in self.storage_media:
            medium["storage_id"] = ""
            medium["asset_id"] = ""
            medium["serial_number"] = ""
            medium["created_at"] = ""
            medium["updated_at"] = ""
        self.change_note_edit.setText(f"复制自 {asset.asset_id}")
        self._refresh_storage_table()

    def _add_storage(self) -> None:
        dialog = StorageMediumDialog(
            parent=self, status_values=self.service.get_enum_values("使用状态")
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.storage_media.append(dialog.medium_data())
            self._refresh_storage_table()

    def _edit_storage(self) -> None:
        row = self.storage_table.currentRow()
        if row < 0:
            return
        dialog = StorageMediumDialog(
            self.storage_media[row], self, status_values=self.service.get_enum_values("使用状态")
        )
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
                medium.get("status", ""),
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
        primary_category = self.primary_category_combo.currentText().strip()
        secondary_category = self.secondary_category_combo.currentText().strip()
        category_parts = [
            part for part in (primary_category, secondary_category) if part
        ]
        return {
            "asset_id": self.asset_id_edit.text().strip(),
            "equipment_code": self.equipment_code_edit.text().strip(),
            "asset_identifier": self.asset_identifier_edit.text().strip(),
            "name": self.name_edit.text().strip(),
            "primary_category": primary_category,
            "secondary_category": secondary_category,
            "category_path": " / ".join(category_parts),
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
            "notes": self.notes1_edit.toPlainText().strip(),
            "notes2": self.notes2_edit.toPlainText().strip(),
        }

    def change_note(self) -> str:
        return self.change_note_edit.text().strip()

    def _validate_and_accept(self) -> None:
        if not self.name_edit.text().strip():
            QMessageBox.warning(self, "无法保存", "设备器材不能为空。")
            self.name_edit.setFocus()
            return
        if not self.grade_combo.currentText().strip():
            answer = QMessageBox.question(
                self,
                "等级未填写",
                "等级尚未填写。该字段建议填写，是否继续保存？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if answer != QMessageBox.StandardButton.Yes:
                self.grade_combo.setFocus()
                return
        self.accept()

    def _update_grade_hint(self) -> None:
        if self.grade_combo.currentText().strip():
            self.grade_combo.setStyleSheet("")
        else:
            self.grade_combo.setStyleSheet("background: #FFF7D6; border-color: #E0A800;")

    def _update_label_status_hint(self) -> None:
        if self.label_printed_check.isChecked():
            self.label_status_label.setText("已打印")
            self.label_status_label.setStyleSheet(
                "background: #DDF4E7; color: #116149; border-radius: 8px; padding: 3px 8px;"
            )
        else:
            self.label_status_label.setText("未打印标签")
            self.label_status_label.setStyleSheet(
                "background: #FFF7D6; color: #8A6100; border-radius: 8px; padding: 3px 8px;"
            )

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QDialog {
                background: #F4F7F8;
                color: #17212B;
                font-size: 13px;
                font-family: "Microsoft YaHei UI", "Segoe UI";
            }
            QGroupBox#formGroup {
                background: #FFFFFF;
                border: 1px solid #D6E1E5;
                border-radius: 8px;
                margin-top: 16px;
                font-weight: 700;
            }
            QGroupBox#formGroup::title {
                subcontrol-origin: margin;
                left: 16px;
                padding: 0 6px;
                color: #0F3D3E;
                background: #F4F7F8;
            }
            QLabel#formLabel {
                background: transparent;
                color: #344054;
                font-weight: 500;
            }
            QLineEdit, QComboBox, QDoubleSpinBox, QPlainTextEdit {
                min-height: 32px;
                border: 1px solid #BFD0D6;
                border-radius: 6px;
                background: #FFFFFF;
                padding: 0 8px;
            }
            QPlainTextEdit#largeNotesEdit {
                min-height: 130px;
                padding: 8px;
            }
            QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus, QPlainTextEdit:focus {
                border-color: #0F766E;
            }
            QTableWidget {
                background: #FFFFFF;
                alternate-background-color: #F8FAFB;
                border: 1px solid #D6E1E5;
                gridline-color: #E8EEF1;
                selection-background-color: #EAF4FB;
                selection-color: #17212B;
            }
            QTableWidget::item:selected {
                background: #EAF4FB;
                color: #17212B;
            }
            QPushButton {
                min-height: 30px;
                padding: 0 10px;
                border: 1px solid #BFD0D6;
                background: #FFFFFF;
                border-radius: 5px;
            }
            QPushButton:hover {
                background: #EEF8F6;
                border-color: #7EBDB7;
            }
            """
        )

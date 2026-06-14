from __future__ import annotations

from datetime import datetime
from html import escape
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStyle,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QTreeWidget,
    QTreeWidgetItem,
    QTreeWidgetItemIterator,
    QVBoxLayout,
    QWidget,
)

from .asset_dialog import AssetDialog
from .asset_service import AssetService
from .excel_repository import ExcelRepository
from .models import Asset
from .settings_dialog import SettingsDialog
from .widgets import NoWheelComboBox


STATUS_COLORS = {
    "在用": "#DDF4E7",
    "入库": "#E8F0FE",
    "闲置": "#F1F3F4",
    "维修": "#FFF1CC",
    "封存": "#EAE4F7",
    "报废": "#FDE2E1",
    "丢失": "#FFD7D7",
}

TABLE_COLUMNS = [
    ("asset_id", "资产唯一标识符"),
    ("equipment_code", "bm编码"),
    ("name", "设备器材"),
    ("secondary_category", "类别"),
    ("brand_model", "品牌 / 产品型号"),
    ("status", "使用状态"),
    ("location", "存放地点"),
    ("owner", "管理人"),
    ("user", "使用人"),
    ("price", "金额"),
    ("updated_at", "更新时间"),
]


class MainWindow(QMainWindow):
    workbook_changed = Signal(object)

    def __init__(self, service: AssetService, workbook_path: Path, parent=None) -> None:
        super().__init__(parent)
        self.service = service
        self.workbook_path = Path(workbook_path)
        self.assets: list[Asset] = []
        self.last_refresh_time = ""
        self.setWindowTitle("设备资产台账")
        self.resize(1500, 900)
        self.setMinimumSize(1100, 680)
        self._build_ui()
        self._apply_style()
        self.refresh_filters()
        self.refresh_assets()
        self._update_refresh_label()

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        header = QFrame()
        header.setObjectName("appHeader")
        header_layout = QHBoxLayout(header)
        title = QLabel("设备资产台账")
        title.setObjectName("appTitle")
        self.path_label = QLabel(str(self.workbook_path))
        self.path_label.setObjectName("pathLabel")
        settings_button = QPushButton("设置")
        settings_button.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView)
        )
        settings_button.clicked.connect(self.open_settings)
        header_layout.addWidget(title)
        header_layout.addWidget(self.path_label, 1)
        header_layout.addWidget(settings_button)
        root_layout.addWidget(header)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self._build_filter_panel())
        splitter.addWidget(self._build_table_panel())
        splitter.addWidget(self._build_detail_panel())
        splitter.setSizes([250, 850, 380])
        splitter.setCollapsible(0, False)
        splitter.setCollapsible(1, False)
        root_layout.addWidget(splitter, 1)
        self.setCentralWidget(root)
        self.statusBar().showMessage("就绪")

    def _build_filter_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("sidePanel")
        panel.setMinimumWidth(220)
        layout = QVBoxLayout(panel)
        layout.addWidget(self._section_label("快捷视图"))
        quick_row = QHBoxLayout()
        for text in ("全部", "在用", "维修", "封存"):
            button = QPushButton(text)
            button.setCheckable(True)
            button.clicked.connect(lambda checked=False, value=text: self._quick_filter(value))
            quick_row.addWidget(button)
        layout.addLayout(quick_row)

        layout.addWidget(self._section_label("筛选条件"))
        self.status_filter = self._filter_combo("全部使用状态")
        self.location_filter = self._filter_combo("全部位置")
        self.brand_filter = self._filter_combo("全部品牌")
        self.owner_filter = self._filter_combo("全部管理人")
        self.user_filter = self._filter_combo("全部使用人")
        self.source_filter = self._filter_combo("全部取得方式")
        for combo in (
            self.status_filter,
            self.location_filter,
            self.brand_filter,
            self.owner_filter,
            self.user_filter,
            self.source_filter,
        ):
            combo.currentTextChanged.connect(self.refresh_assets)
            layout.addWidget(combo)

        layout.addWidget(self._section_label("设备类别"))
        self.category_tree = QTreeWidget()
        self.category_tree.setHeaderHidden(True)
        self.category_tree.itemClicked.connect(self._category_clicked)
        layout.addWidget(self.category_tree, 1)
        clear_button = QPushButton("清除全部筛选")
        clear_button.clicked.connect(self.clear_filters)
        layout.addWidget(clear_button)
        return panel

    def _build_table_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        toolbar = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText(
            "搜索资产ID、bm编码、设备器材、品牌型号、厂家、人员或存储介质序列号"
        )
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.textChanged.connect(self.refresh_assets)
        new_button = QPushButton("新增设备")
        new_button.setObjectName("primaryButton")
        new_button.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogNewFolder)
        )
        new_button.clicked.connect(self.create_asset)
        edit_button = QPushButton("编辑")
        edit_button.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView)
        )
        edit_button.clicked.connect(self.edit_selected_asset)
        refresh_button = QPushButton("刷新")
        refresh_button.setToolTip("重新读取 Excel 工作簿")
        refresh_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        refresh_button.clicked.connect(self.refresh_all)
        toolbar.addWidget(self.search_edit, 1)
        toolbar.addWidget(new_button)
        toolbar.addWidget(edit_button)
        toolbar.addWidget(refresh_button)
        layout.addLayout(toolbar)

        self.result_label = QLabel()
        self.result_label.setObjectName("mutedLabel")
        layout.addWidget(self.result_label)
        self.asset_table = QTableWidget(0, len(TABLE_COLUMNS))
        self.asset_table.setHorizontalHeaderLabels([label for _, label in TABLE_COLUMNS])
        self.asset_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.asset_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.asset_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.asset_table.setAlternatingRowColors(True)
        self.asset_table.setSortingEnabled(True)
        self.asset_table.verticalHeader().setVisible(False)
        self.asset_table.itemSelectionChanged.connect(self.show_selected_asset)
        self.asset_table.itemDoubleClicked.connect(self.edit_selected_asset)
        widths = [155, 125, 150, 110, 150, 90, 110, 90, 90, 105, 145]
        for index, width in enumerate(widths):
            self.asset_table.setColumnWidth(index, width)
        layout.addWidget(self.asset_table, 1)
        return panel

    def _build_detail_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("detailPanel")
        panel.setMinimumWidth(320)
        layout = QVBoxLayout(panel)
        self.detail_tabs = QTabWidget()
        self.detail_text = QTextBrowser()
        history_panel = QWidget()
        history_layout = QVBoxLayout(history_panel)
        history_layout.setContentsMargins(0, 0, 0, 0)
        history_toolbar = QHBoxLayout()
        self.history_filter = NoWheelComboBox()
        self.history_filter.addItem("全部变化", "")
        self.history_filter.currentTextChanged.connect(self.refresh_history)
        history_refresh_button = QPushButton("刷新")
        history_refresh_button.setToolTip("重新读取当前设备的变化记录")
        history_refresh_button.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload)
        )
        history_refresh_button.clicked.connect(self.refresh_history)
        history_toolbar.addWidget(self.history_filter, 1)
        history_toolbar.addWidget(history_refresh_button)
        self.history_refresh_label = QLabel("尚未刷新")
        self.history_refresh_label.setObjectName("mutedLabel")
        self.history_text = QTextBrowser()
        history_layout.addLayout(history_toolbar)
        history_layout.addWidget(self.history_refresh_label)
        history_layout.addWidget(self.history_text, 1)
        self.detail_tabs.addTab(self.detail_text, "设备详情")
        self.detail_tabs.addTab(history_panel, "变化记录")
        layout.addWidget(self.detail_tabs)
        return panel

    @staticmethod
    def _section_label(text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("sectionLabel")
        return label

    @staticmethod
    def _filter_combo(all_text: str) -> NoWheelComboBox:
        combo = NoWheelComboBox()
        combo.addItem(all_text, "")
        return combo

    def refresh_all(self) -> None:
        selected_asset_id = self._selected_asset_id()
        self.refresh_filters()
        self.refresh_assets()
        if selected_asset_id:
            self._select_asset(selected_asset_id)
        self.show_selected_asset()
        self._update_refresh_label()

    def refresh_filters(self) -> None:
        assets = self.service.list_assets()
        selected_category = self._selected_category()
        current_values = {
            self.status_filter: self.status_filter.currentData(),
            self.location_filter: self.location_filter.currentData(),
            self.brand_filter: self.brand_filter.currentData(),
            self.owner_filter: self.owner_filter.currentData(),
            self.user_filter: self.user_filter.currentData(),
            self.source_filter: self.source_filter.currentData(),
        }
        options = {
            self.status_filter: sorted({asset.status for asset in assets if asset.status}),
            self.location_filter: sorted({asset.location for asset in assets if asset.location}),
            self.brand_filter: sorted({asset.brand for asset in assets if asset.brand}),
            self.owner_filter: self.service.list_owners(),
            self.user_filter: self.service.list_users(),
            self.source_filter: sorted({asset.source for asset in assets if asset.source}),
        }
        for combo, values in options.items():
            label = combo.itemText(0)
            combo.blockSignals(True)
            combo.clear()
            combo.addItem(label, "")
            if combo is self.owner_filter:
                combo.addItem("未分配管理人", None)
            if combo is self.user_filter:
                combo.addItem("未分配使用人", None)
            for value in values:
                combo.addItem(value, value)
            index = combo.findData(current_values[combo])
            combo.setCurrentIndex(max(index, 0))
            combo.blockSignals(False)
        self._load_category_tree(assets)
        self._restore_category_selection(selected_category)

    def _load_category_tree(self, assets: list[Asset]) -> None:
        self.category_tree.clear()
        categories = self.service.get_categories(include_disabled=True)
        primary_items: dict[str, QTreeWidgetItem] = {}
        for category in categories:
            count = sum(
                1
                for asset in assets
                if asset.primary_category == category.name
                or asset.secondary_category == category.name
            )
            node = QTreeWidgetItem([f"{category.name}  ({count})"])
            node.setData(0, Qt.ItemDataRole.UserRole, category.name)
            node.setData(0, Qt.ItemDataRole.UserRole + 1, bool(category.parent_id))
            if category.parent_id and category.parent_id in primary_items:
                primary_items[category.parent_id].addChild(node)
            else:
                self.category_tree.addTopLevelItem(node)
                primary_items[category.item_id] = node
        self.category_tree.expandAll()

    def refresh_assets(self, *_args) -> None:
        selected_asset_id = self._selected_asset_id()
        filters: dict[str, str | None] = {}
        for field, combo in (
            ("status", self.status_filter),
            ("location", self.location_filter),
            ("brand", self.brand_filter),
            ("owner", self.owner_filter),
            ("user", self.user_filter),
            ("source", self.source_filter),
        ):
            value = combo.currentData()
            if value != "":
                filters[field] = value
        selected_category = self.category_tree.currentItem()
        if selected_category:
            category = selected_category.data(0, Qt.ItemDataRole.UserRole)
            is_secondary = selected_category.data(0, Qt.ItemDataRole.UserRole + 1)
            filters["secondary_category" if is_secondary else "primary_category"] = category
        self.assets = self.service.list_assets(filters, self.search_edit.text())
        self.asset_table.setSortingEnabled(False)
        self.asset_table.setRowCount(len(self.assets))
        for row, asset in enumerate(self.assets):
            values = [
                asset.asset_id,
                asset.equipment_code,
                asset.name,
                asset.secondary_category or asset.primary_category,
                " / ".join(value for value in (asset.brand, asset.model) if value),
                asset.status,
                asset.location,
                asset.owner,
                asset.user,
                f"¥ {asset.price:,.2f}",
                asset.updated_at,
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setData(Qt.ItemDataRole.UserRole, asset.asset_id)
                if column == 5:
                    item.setBackground(QColor(STATUS_COLORS.get(asset.status, "#FFFFFF")))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if column == 9:
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                self.asset_table.setItem(row, column, item)
        self.asset_table.setSortingEnabled(True)
        self._show_result_label()
        if self.assets:
            if selected_asset_id:
                self._select_asset(selected_asset_id)
            if not self.asset_table.selectedItems():
                self.asset_table.selectRow(0)
        elif not self.assets:
            self.detail_text.setText("未选择设备")
            self.history_text.clear()

    def show_selected_asset(self) -> None:
        asset = self._selected_asset()
        if not asset:
            return
        fields = [
            ("资产唯一标识符", asset.asset_id),
            ("bm编码", asset.equipment_code),
            ("设备器材", asset.name),
            (
                "设备分类",
                " / ".join(v for v in (asset.primary_category, asset.secondary_category) if v),
            ),
            ("产品规格", asset.product_spec),
            ("品牌 / 产品型号", " / ".join(v for v in (asset.brand, asset.model) if v)),
            ("设备序列号", asset.serial_number),
            ("生产厂家", asset.manufacturer),
            ("供应商", asset.supplier),
            ("取得方式", asset.source),
            ("计价方式 / 金额", " / ".join(v for v in (asset.valuation_method, f"¥ {asset.price:,.2f}") if v)),
            (
                "取得 / 启用 / 出厂日期",
                " / ".join(
                    v
                    for v in (
                        asset.purchase_date,
                        asset.start_date,
                        asset.manufacture_date,
                    )
                    if v
                ),
            ),
            ("等级", asset.grade),
            ("使用状态", asset.status),
            ("存放地点", asset.location),
            (
                "管理部门 / 管理人",
                " / ".join(v for v in (asset.department, asset.owner) if v),
            ),
            (
                "使用部门 / 使用人",
                " / ".join(v for v in (asset.use_department, asset.user) if v),
            ),
            ("是否已打印标签", "是" if asset.label_printed else "否"),
            ("备注", asset.notes),
            ("创建时间", asset.created_at),
            ("更新时间", asset.updated_at),
        ]
        media = self.service.list_storage_media(asset.asset_id)
        storage_html = (
            "<h3>存储介质</h3>"
            + "".join(
                (
                    "<p><b>{}</b><br>{}</p>".format(
                        escape(medium.name or medium.medium_type or "未命名介质"),
                        escape(
                            " / ".join(
                                value
                                for value in (
                                    medium.medium_type,
                                    medium.brand,
                                    (
                                        f"{self.service._display_value(medium.capacity_value)}"
                                        f"{medium.capacity_unit}"
                                        if medium.capacity_value
                                        else ""
                                    ),
                                    medium.model,
                                    medium.serial_number,
                                    medium.notes,
                                )
                                if value
                            )
                        ),
                    )
                )
                for medium in media
            )
            if media
            else "<h3>存储介质</h3><p>无存储介质</p>"
        )
        self.detail_text.setHtml(
            "<h2>{}</h2>{}{}".format(
                escape(asset.name),
                "".join(
                    f"<p><b>{escape(label)}</b><br>{escape(str(value or '—'))}</p>"
                    for label, value in fields
                ),
                storage_html,
            )
        )
        self._load_history_filter(asset.asset_id)
        self.refresh_history()

    def refresh_history(self, *_args) -> None:
        asset = self._selected_asset()
        if not asset:
            self.history_text.clear()
            return
        changes = self.service.list_changes(
            asset.asset_id, field_name=str(self.history_filter.currentData() or "")
        )
        self.history_text.setHtml(
            "".join(
                (
                    f"<h3>{escape(change.event_type)} · {escape(change.field_name)}</h3>"
                    f"<p>{escape(change.old_value or '—')} → "
                    f"<b>{escape(change.new_value or '—')}</b></p>"
                    f"<p style='color:#667085'>{escape(change.changed_at)} · "
                    f"{escape(change.operator)}"
                    f"{' · ' + escape(change.note) if change.note else ''}</p>"
                )
                for change in changes
            )
            or "暂无变化记录"
        )
        self.history_refresh_label.setText(
            f"最近刷新 {datetime.now().strftime('%H:%M:%S')}"
        )

    def create_asset(self) -> None:
        dialog = AssetDialog(self.service, parent=self)
        if dialog.exec() != AssetDialog.DialogCode.Accepted:
            return
        try:
            created = self.service.create_asset(
                dialog.asset_data(), dialog.storage_media_data(), dialog.change_note()
            )
            self.refresh_all()
            self._select_asset(created.asset_id)
            self.statusBar().showMessage(f"已新增 {created.name}", 5000)
        except Exception as exc:
            QMessageBox.critical(self, "新增失败", str(exc))

    def edit_selected_asset(self, *_args) -> None:
        asset = self._selected_asset()
        if not asset:
            QMessageBox.information(self, "请选择设备", "请先选择需要编辑的设备。")
            return
        dialog = AssetDialog(self.service, asset, self)
        if dialog.exec() != AssetDialog.DialogCode.Accepted:
            return
        try:
            updated = self.service.update_asset(
                asset.asset_id,
                dialog.asset_data(),
                dialog.storage_media_data(),
                dialog.change_note(),
            )
            self.refresh_after_save(updated.asset_id)
            self.statusBar().showMessage(f"已更新 {updated.name}", 5000)
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", str(exc))

    def open_settings(self) -> None:
        dialog = SettingsDialog(self.service, self.workbook_path, self)
        if dialog.exec() != SettingsDialog.DialogCode.Accepted:
            return
        if dialog.selected_path != self.workbook_path:
            repository = ExcelRepository(dialog.selected_path)
            repository.initialize()
            errors = repository.validate_workbook()
            if errors:
                QMessageBox.critical(self, "工作簿无效", "\n".join(errors))
                return
            self.workbook_path = dialog.selected_path
            self.service = AssetService(repository)
            self.service.set_operator(dialog.operator_edit.text())
            self.path_label.setText(str(self.workbook_path))
            self.workbook_changed.emit(self.workbook_path)
        self.refresh_all()

    def clear_filters(self) -> None:
        for combo in (
            self.status_filter,
            self.location_filter,
            self.brand_filter,
            self.owner_filter,
            self.user_filter,
            self.source_filter,
        ):
            combo.setCurrentIndex(0)
        self.category_tree.setCurrentItem(None)
        self.category_tree.clearSelection()
        self.search_edit.clear()
        self.refresh_assets()

    def _quick_filter(self, value: str) -> None:
        if value == "全部":
            self.clear_filters()
            return
        self.status_filter.setCurrentIndex(
            max(self.status_filter.findData(value), 0)
        )
        self.refresh_assets()

    def _category_clicked(self, *_args) -> None:
        self.refresh_assets()

    def _selected_asset(self) -> Asset | None:
        items = self.asset_table.selectedItems()
        if not items:
            return None
        asset_id = items[0].data(Qt.ItemDataRole.UserRole)
        return next((asset for asset in self.assets if asset.asset_id == asset_id), None)

    def _selected_asset_id(self) -> str:
        asset = self._selected_asset()
        return asset.asset_id if asset else ""

    def _select_asset(self, asset_id: str) -> None:
        for row in range(self.asset_table.rowCount()):
            if self.asset_table.item(row, 0).data(Qt.ItemDataRole.UserRole) == asset_id:
                self.asset_table.selectRow(row)
                break

    def refresh_after_save(self, asset_id: str) -> None:
        self.refresh_filters()
        self.refresh_assets()
        if not any(asset.asset_id == asset_id for asset in self.assets):
            self.clear_filters()
        self._select_asset(asset_id)
        self.show_selected_asset()
        self.detail_tabs.setCurrentIndex(1)
        self._update_refresh_label()

    def _load_history_filter(self, asset_id: str) -> None:
        current = self.history_filter.currentData()
        fields = sorted(
            {
                change.field_name
                for change in self.service.list_changes(asset_id)
                if change.field_name != "*"
            }
        )
        self.history_filter.blockSignals(True)
        self.history_filter.clear()
        self.history_filter.addItem("全部变化", "")
        for field in fields:
            self.history_filter.addItem(field, field)
        index = self.history_filter.findData(current)
        self.history_filter.setCurrentIndex(max(index, 0))
        self.history_filter.blockSignals(False)

    def _selected_category(self) -> tuple[str, bool] | None:
        item = self.category_tree.currentItem()
        if not item:
            return None
        return (
            str(item.data(0, Qt.ItemDataRole.UserRole)),
            bool(item.data(0, Qt.ItemDataRole.UserRole + 1)),
        )

    def _restore_category_selection(self, selected: tuple[str, bool] | None) -> None:
        if not selected:
            return
        iterator = QTreeWidgetItemIterator(self.category_tree)
        while iterator.value():
            item = iterator.value()
            key = (
                str(item.data(0, Qt.ItemDataRole.UserRole)),
                bool(item.data(0, Qt.ItemDataRole.UserRole + 1)),
            )
            if key == selected:
                self.category_tree.setCurrentItem(item)
                return
            iterator += 1

    def _update_refresh_label(self) -> None:
        self.last_refresh_time = datetime.now().strftime("%H:%M:%S")
        self._show_result_label()

    def _show_result_label(self) -> None:
        refresh_text = (
            f" · 最近刷新 {self.last_refresh_time}" if self.last_refresh_time else ""
        )
        self.result_label.setText(f"共 {len(self.assets)} 台设备{refresh_text}")

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow, QWidget { background: #F6F8FA; color: #17212B; font-size: 13px;
                                   font-family: "Microsoft YaHei UI", "Segoe UI"; }
            #appHeader { background: #FFFFFF; border-bottom: 1px solid #D9E0E7; }
            #appTitle { font-size: 20px; font-weight: 700; padding: 10px 4px; }
            #pathLabel, #mutedLabel { color: #667085; }
            #sidePanel, #detailPanel { background: #FFFFFF; }
            #sidePanel { border-right: 1px solid #D9E0E7; }
            #detailPanel { border-left: 1px solid #D9E0E7; }
            #sectionLabel { color: #344054; font-weight: 700; margin-top: 8px; }
            QPushButton { min-height: 30px; padding: 0 10px; border: 1px solid #C9D2DC;
                          background: #FFFFFF; border-radius: 4px; }
            QPushButton:hover { background: #F0F5FA; }
            QPushButton:pressed, QPushButton:checked { background: #E1ECF7; border-color: #6A94B8; }
            #primaryButton { background: #1769AA; color: white; border-color: #1769AA; font-weight: 600; }
            QLineEdit, QComboBox, QDoubleSpinBox, QPlainTextEdit {
                min-height: 30px; border: 1px solid #C9D2DC; border-radius: 4px; background: white;
                padding: 0 7px;
            }
            QTableWidget, QTreeWidget, QTextBrowser, QListWidget {
                background: white; border: 1px solid #D9E0E7; gridline-color: #E8EDF2;
            }
            QTableWidget::item:selected { background: #EAF4FB; color: #17212B; }
            QHeaderView::section { background: #EDF2F6; padding: 8px; border: 0;
                                   border-right: 1px solid #D9E0E7; font-weight: 600; }
            QTabBar::tab { padding: 9px 14px; background: #EDF2F6; }
            QTabBar::tab:selected { background: white; border-bottom: 2px solid #1769AA; }
            QGroupBox { font-weight: 700; border: 1px solid #D9E0E7; margin-top: 12px;
                        padding-top: 14px; background: #FFFFFF; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
            """
        )

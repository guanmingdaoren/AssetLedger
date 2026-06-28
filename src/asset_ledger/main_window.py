from __future__ import annotations

from datetime import datetime
from html import escape
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QAction, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
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
from .asset_service import AssetService, CacheReloadAfterSaveError, CacheUnavailableError
from .excel_repository import ExcelRepository, WorkbookChangedExternallyError
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
    ("row_number", "序号"),
    ("asset_id", "资产UID"),
    ("equipment_code", "bm编码"),
    ("asset_identifier", "资产唯一标识符"),
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
        self.recently_copied_asset_id = ""
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
        self.user_filter = self._filter_combo("全部使用人")
        for combo in (
            self.status_filter,
            self.location_filter,
            self.brand_filter,
            self.user_filter,
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
            "搜索资产UID、bm编码、资产唯一标识符、设备器材、品牌型号、厂家、人员或存储介质序列号"
        )
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.textChanged.connect(self.refresh_assets)
        new_button = QPushButton("新增设备")
        new_button.setObjectName("primaryButton")
        new_button.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogNewFolder)
        )
        new_button.clicked.connect(self.create_asset)
        copy_button = QPushButton("复制")
        copy_button.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogToParent)
        )
        copy_button.clicked.connect(self.copy_selected_asset)
        edit_button = QPushButton("编辑")
        edit_button.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView)
        )
        edit_button.clicked.connect(self.edit_selected_asset)
        delete_button = QPushButton("删除")
        delete_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon))
        delete_button.clicked.connect(self.delete_selected_asset)
        refresh_button = QPushButton("刷新")
        refresh_button.setToolTip("重新读取 Excel 工作簿")
        refresh_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        refresh_button.clicked.connect(self.refresh_all)
        toolbar.addWidget(self.search_edit, 1)
        toolbar.addWidget(new_button)
        toolbar.addWidget(copy_button)
        toolbar.addWidget(edit_button)
        toolbar.addWidget(delete_button)
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
        self.asset_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.asset_table.customContextMenuRequested.connect(self._open_asset_context_menu)
        self.asset_table.itemSelectionChanged.connect(self.show_selected_asset)
        self.asset_table.itemDoubleClicked.connect(self.edit_selected_asset)
        widths = [58, 155, 125, 155, 150, 150, 150, 90, 110, 90, 90, 105, 145]
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

    def refresh_all(self, *_args, reload_cache: bool = True) -> None:
        selected_asset_id = self._selected_asset_id()
        if reload_cache:
            try:
                self.service.reload_cache()
            except Exception as exc:
                self.statusBar().showMessage("刷新失败，当前显示仍为上次有效缓存")
                QMessageBox.critical(
                    self,
                    "刷新失败",
                    f"{exc}\n\n当前画面仍显示上次成功加载的数据，保存功能已被禁用。",
                )
                return
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
            self.user_filter: self.user_filter.currentData(),
        }
        options = {
            self.status_filter: sorted({asset.status for asset in assets if asset.status}),
            self.location_filter: sorted({asset.location for asset in assets if asset.location}),
            self.brand_filter: sorted({asset.brand for asset in assets if asset.brand}),
            self.user_filter: self.service.list_users(),
        }
        for combo, values in options.items():
            label = combo.itemText(0)
            combo.blockSignals(True)
            combo.clear()
            combo.addItem(label, "")
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
        by_parent: dict[str, list] = {}
        for category in categories:
            by_parent.setdefault(category.parent_id, []).append(category)

        def add_nodes(parent_item: QTreeWidgetItem | None, parent_id: str, parent_path: str) -> None:
            for category in by_parent.get(parent_id, []):
                path = " / ".join(value for value in (parent_path, category.name) if value)
                count = sum(
                    1
                    for asset in assets
                    if asset.category_path == path
                    or asset.category_path.startswith(f"{path} / ")
                    or (not asset.category_path and category.name in {asset.primary_category, asset.secondary_category})
                )
                node = QTreeWidgetItem([f"{category.name}  ({count})"])
                node.setData(0, Qt.ItemDataRole.UserRole, path)
                node.setData(0, Qt.ItemDataRole.UserRole + 1, bool(category.parent_id))
                if parent_item is None:
                    self.category_tree.addTopLevelItem(node)
                else:
                    parent_item.addChild(node)
                add_nodes(node, category.item_id, path)

        add_nodes(None, "", "")
        self.category_tree.expandAll()

    def refresh_assets(self, *_args) -> None:
        selected_asset_id = self._selected_asset_id()
        filters: dict[str, str | None] = {}
        for field, combo in (
            ("status", self.status_filter),
            ("location", self.location_filter),
            ("brand", self.brand_filter),
            ("user", self.user_filter),
        ):
            value = combo.currentData()
            if value != "":
                filters[field] = value
        selected_category = self.category_tree.currentItem()
        if selected_category:
            filters["category_path"] = selected_category.data(0, Qt.ItemDataRole.UserRole)
        self.assets = self.service.list_assets(filters, self.search_edit.text())
        self.asset_table.setSortingEnabled(False)
        self.asset_table.setRowCount(len(self.assets))
        for row, asset in enumerate(self.assets):
            values = [
                row + 1,
                asset.asset_id,
                asset.equipment_code,
                asset.asset_identifier,
                asset.name,
                asset.category_path or asset.secondary_category or asset.primary_category,
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
                if column == 0:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if column == 7:
                    item.setBackground(QColor(STATUS_COLORS.get(asset.status, "#FFFFFF")))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if column == 11:
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                if asset.asset_id == self.recently_copied_asset_id:
                    item.setBackground(QColor("#FFF4CC"))
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
            self._update_copied_selection_style("")
            return
        self._update_copied_selection_style(asset.asset_id)
        groups = [
            (
                "身份信息",
                [
                    ("资产UID", asset.asset_id),
                    ("bm编码", asset.equipment_code),
                    ("资产唯一标识符", asset.asset_identifier),
                    ("设备器材", asset.name),
                    ("设备分类", asset.category_path or " / ".join(v for v in (asset.primary_category, asset.secondary_category) if v)),
                ],
            ),
            (
                "产品信息",
                [
                    ("产品规格", asset.product_spec),
                    ("品牌 / 产品型号", " / ".join(v for v in (asset.brand, asset.model) if v)),
                    ("设备序列号", asset.serial_number),
                    ("生产厂家", asset.manufacturer),
                    ("供应商", asset.supplier),
                ],
            ),
            (
                "取得与资产信息",
                [
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
                    ("是否已打印标签", "是" if asset.label_printed else "否"),
                ],
            ),
            (
                "管理与使用信息",
                [
                    ("使用状态", asset.status),
                    ("存放地点", asset.location),
                    ("管理部门 / 管理人", " / ".join(v for v in (asset.department, asset.owner) if v)),
                    ("使用部门 / 使用人", " / ".join(v for v in (asset.use_department, asset.user) if v)),
                ],
            ),
            (
                "备注信息",
                [
                    ("备注1", asset.notes),
                    ("备注2", asset.notes2),
                ],
            ),
            (
                "系统信息",
                [("创建时间", asset.created_at), ("更新时间", asset.updated_at)],
            ),
        ]
        system_group = groups.pop()
        media = self.service.list_storage_media(asset.asset_id)
        card_colors = ["#F7FAFC", "#F8FBF7", "#FBFAF6", "#F7F8FC"]

        def fields_html(fields) -> str:
            return (
                "<table width='100%' cellspacing='0' cellpadding='0' style='border-collapse:collapse;'>"
                + "".join(
                    (
                        "<tr>"
                        "<td width='42%' style='color:#667085; padding:6px 0; vertical-align:top;'>"
                        f"{escape(label)}</td>"
                        "<td width='58%' style='color:#17212B; font-weight:600; padding:6px 0; vertical-align:top;'>"
                        f"{escape(str(value or '—'))}</td>"
                        "</tr>"
                    )
                    for label, value in fields
                )
                + "</table>"
            )

        def card_html(title: str, fields, color: str) -> str:
            return (
                "<table width='100%' cellspacing='0' cellpadding='10' "
                f"style='background:{color}; border:1px solid #E3EBEF; "
                "border-collapse:separate; margin:0 0 12px 0;'>"
                "<tr><td>"
                f"<h3 style='margin:0 0 8px 0; color:#0F3D3E; font-size:15px;'>{escape(title)}</h3>"
                f"{fields_html(fields)}"
                "</td></tr></table>"
            )

        storage_html = (
            "<table width='100%' cellspacing='0' cellpadding='10' "
            "style='background:#F7FAFC; border:1px solid #E3EBEF; "
            "border-collapse:separate; margin:0 0 12px 0;'><tr><td>"
            "<h3 style='margin:0 0 8px 0; color:#0F3D3E; font-size:15px;'>存储介质</h3>"
            + "".join(
                (
                    "<div style='padding:8px 0; border-top:1px solid #E8EEF1;'>"
                    "<b style='color:#17212B'>{}</b><br>"
                    "<span style='color:#667085'>{}</span></div>".format(
                        escape(medium.name or medium.medium_type or "未命名介质"),
                        escape(
                            " / ".join(
                                value
                                for value in (
                                    medium.medium_type,
                                    medium.status,
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
            else (
                "<table width='100%' cellspacing='0' cellpadding='10' "
                "style='background:#F7FAFC; border:1px solid #E3EBEF; "
                "border-collapse:separate; margin:0 0 12px 0;'><tr><td>"
                "<h3 style='margin:0 0 8px 0; color:#0F3D3E; font-size:15px;'>存储介质</h3>"
                "<p style='color:#667085; margin:0;'>无存储介质</p>"
            )
        ) + "</td></tr></table>"
        body = "".join(
            card_html(title, fields, card_colors[index % len(card_colors)])
            for index, (title, fields) in enumerate(groups)
        )
        self.detail_text.setHtml(
            (
                "<div style='background:#FFFFFF; color:#17212B;'>"
                "{}{}{}{}"
                "</div>"
            ).format(
                self._asset_summary_html(asset, media),
                body,
                storage_html,
                card_html(system_group[0], system_group[1], "#F7FAFC"),
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
        self.history_text.setHtml(self._history_html(changes))
        self.history_refresh_label.setText(
            f"最近刷新 {datetime.now().strftime('%H:%M:%S')}"
        )

    @staticmethod
    def _history_html(changes) -> str:
        if not changes:
            return (
                "<div style='background:#F7FAFC; border:1px solid #E3EBEF; "
                "padding:14px; color:#667085;'>暂无变化记录</div>"
            )
        grouped: dict[str, list] = {}
        for change in changes:
            key = change.event_group_id or change.change_id
            grouped.setdefault(key, []).append(change)

        colors = ["#F7FAFC", "#F8FBF7", "#FFF8E7", "#F7F8FC"]
        accents = ["#2F80ED", "#0F766E", "#D89E00", "#6B7280"]

        def unique(values: list[str]) -> list[str]:
            return list(dict.fromkeys(value for value in values if value))

        cards = []
        for index, group in enumerate(grouped.values()):
            first = group[0]
            event_title = " / ".join(unique([change.event_type for change in group]))
            note = f" · {escape(first.note)}" if first.note else ""
            rows = "".join(
                (
                    "<tr>"
                    "<td style='border-top:1px solid #E3EBEF; padding:8px 0;'>"
                    f"<div style='color:#0F3D3E; font-weight:700;'>"
                    f"{escape(change.event_type)}"
                    f"{'' if change.field_name == '*' else ' · ' + escape(change.field_name)}"
                    "</div>"
                    f"<div style='color:#17212B; margin-top:4px;'>"
                    f"{escape(change.old_value or '—')} → "
                    f"<b>{escape(change.new_value or '—')}</b>"
                    "</div>"
                    "</td>"
                    "</tr>"
                )
                for change in group
            )
            cards.append(
                (
                    "<table width='100%' cellspacing='0' cellpadding='10' "
                    f"style='background:{colors[index % len(colors)]}; "
                    "border:1px solid #DDE8EE; border-collapse:separate; "
                    "margin:0 0 12px 0;'>"
                    "<tr><td>"
                    "<table width='100%' cellspacing='0' cellpadding='0'>"
                    "<tr>"
                    f"<td style='border-left:4px solid {accents[index % len(accents)]}; "
                    "padding-left:8px;'>"
                    f"<div style='font-size:15px; color:#0F3D3E; font-weight:700;'>"
                    f"{escape(event_title)}</div>"
                    f"<div style='color:#667085; margin-top:4px;'>"
                    f"{escape(first.changed_at)} · {escape(first.operator)}{note}</div>"
                    "</td>"
                    "<td align='right' style='color:#344054; font-weight:700;'>"
                    f"{len(group)} 项变化</td>"
                    "</tr>"
                    "</table>"
                    "<table width='100%' cellspacing='0' cellpadding='0' "
                    "style='margin-top:8px; border-collapse:collapse;'>"
                    f"{rows}"
                    "</table>"
                    "</td></tr>"
                    "</table>"
                )
            )
        return "<div style='background:#FFFFFF; color:#17212B;'>" + "".join(cards) + "</div>"

    def create_asset(self) -> None:
        dialog = AssetDialog(self.service, parent=self)
        if dialog.exec() != AssetDialog.DialogCode.Accepted:
            return
        try:
            created = self.service.create_asset(
                dialog.asset_data(), dialog.storage_media_data(), dialog.change_note()
            )
            self.refresh_all(reload_cache=False)
            self._select_asset(created.asset_id)
            self.statusBar().showMessage(f"已新增 {created.name}", 5000)
        except Exception as exc:
            self._show_write_error("新增失败", exc)

    def copy_selected_asset(self, *_args) -> None:
        asset = self._selected_asset()
        if not asset:
            QMessageBox.information(self, "请选择设备", "请先选择需要复制的设备。")
            return
        dialog = AssetDialog(self.service, asset, self, copy_mode=True)
        if dialog.exec() != AssetDialog.DialogCode.Accepted:
            return
        try:
            created = self.service.create_asset(
                dialog.asset_data(),
                dialog.storage_media_data(),
                dialog.change_note() or f"复制自 {asset.asset_id}",
            )
            self.recently_copied_asset_id = created.asset_id
            self.refresh_after_copy(created.asset_id)
            self.statusBar().showMessage(f"已复制生成 {created.name}", 5000)
        except Exception as exc:
            self._show_write_error("复制失败", exc)

    def show_selected_asset_detail(self) -> None:
        if not self._selected_asset():
            return
        self.show_selected_asset()
        self.detail_tabs.setCurrentIndex(0)

    def show_selected_asset_history(self) -> None:
        if not self._selected_asset():
            return
        self.show_selected_asset()
        self.detail_tabs.setCurrentIndex(1)

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
            self._show_write_error("保存失败", exc)

    def delete_selected_asset(self, *_args) -> None:
        asset = self._selected_asset()
        if not asset:
            QMessageBox.information(self, "请选择设备", "请先选择需要删除的设备。")
            return
        media_count = len(self.service.list_storage_media(asset.asset_id))
        change_count = len(self.service.list_changes(asset.asset_id))
        message = (
            "确认删除这台设备及其关联信息吗？\n\n"
            f"设备器材：{asset.name or '—'}\n"
            f"资产UID：{asset.asset_id}\n"
            f"bm编码：{asset.equipment_code or '—'}\n"
            f"资产唯一标识符：{asset.asset_identifier or '—'}\n"
            f"关联数据：{media_count} 个存储介质，{change_count} 条变化记录\n\n"
            "删除后，该设备、存储介质和变化记录会从当前台账中移除；"
            "如需找回，只能从删除前自动备份中恢复。"
        )
        result = QMessageBox.question(
            self,
            "确认删除设备",
            message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if result != QMessageBox.StandardButton.Yes:
            return
        try:
            deleted_name = asset.name
            deleted_asset_id = asset.asset_id
            self.service.delete_asset(asset.asset_id)
            if self.recently_copied_asset_id == deleted_asset_id:
                self.recently_copied_asset_id = ""
            self.refresh_all(reload_cache=False)
            self.statusBar().showMessage(f"已删除 {deleted_name}", 5000)
        except Exception as exc:
            self._show_write_error("删除失败", exc)

    def open_settings(self) -> None:
        dialog = SettingsDialog(self.service, self.workbook_path, self)
        if dialog.exec() != SettingsDialog.DialogCode.Accepted:
            return
        if dialog.selected_path != self.workbook_path:
            try:
                repository = ExcelRepository(dialog.selected_path)
                repository.initialize()
                new_service = AssetService(repository)
                new_service.set_operator(dialog.operator_edit.text())
            except Exception as exc:
                QMessageBox.critical(self, "无法切换工作簿", str(exc))
                return
            self.workbook_path = dialog.selected_path
            self.service = new_service
            self.path_label.setText(str(self.workbook_path))
            self.workbook_changed.emit(self.workbook_path)
        self.refresh_all(reload_cache=False)

    def clear_filters(self) -> None:
        for combo in (
            self.status_filter,
            self.location_filter,
            self.brand_filter,
            self.user_filter,
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

    def _open_asset_context_menu(self, position) -> None:
        if self._select_asset_row_at_context_position(position):
            menu = self._build_asset_row_context_menu()
        else:
            menu = self._build_asset_blank_context_menu()
        menu.exec(self.asset_table.viewport().mapToGlobal(position))

    def _select_asset_row_at_context_position(self, position) -> bool:
        item = self.asset_table.itemAt(position)
        if not item:
            return False
        self.asset_table.selectRow(item.row())
        return True

    def _build_asset_row_context_menu(self) -> QMenu:
        menu = QMenu(self)
        self._add_context_action(menu, "查看详情", self.show_selected_asset_detail)
        self._add_context_action(menu, "查看变化记录", self.show_selected_asset_history)
        self._add_context_action(menu, "编辑设备", self.edit_selected_asset)
        self._add_context_action(menu, "复制为新设备", self.copy_selected_asset)
        self._add_context_action(menu, "删除设备", self.delete_selected_asset)
        menu.addSeparator()
        self._add_context_action(menu, "刷新列表", self.refresh_all)
        return menu

    def _build_asset_blank_context_menu(self) -> QMenu:
        menu = QMenu(self)
        self._add_context_action(menu, "新增设备", self.create_asset)
        self._add_context_action(menu, "刷新列表", self.refresh_all)
        return menu

    @staticmethod
    def _add_context_action(menu: QMenu, text: str, callback) -> QAction:
        action = QAction(text, menu)
        action.triggered.connect(lambda _checked=False, callback=callback: callback())
        menu.addAction(action)
        return action

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

    def refresh_after_copy(self, asset_id: str) -> None:
        self.refresh_filters()
        self.refresh_assets()
        if not any(asset.asset_id == asset_id for asset in self.assets):
            self.clear_filters()
        self._select_asset(asset_id)
        self.show_selected_asset()
        self.detail_tabs.setCurrentIndex(0)
        self._update_refresh_label()

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

    def _update_copied_selection_style(self, selected_asset_id: str) -> None:
        if self.recently_copied_asset_id and selected_asset_id == self.recently_copied_asset_id:
            self.asset_table.setStyleSheet(
                "QTableWidget::item:selected { background: #FFF4CC; color: #17212B; }"
            )
        else:
            self.asset_table.setStyleSheet("")

    @staticmethod
    def _asset_summary_html(asset: Asset, media) -> str:
        status_color = STATUS_COLORS.get(asset.status, "#F1F3F4")
        media_text = f"{len(media)} 个存储介质" if media else "无存储介质"
        tag_text = "已打印标签" if asset.label_printed else "未打印标签"
        owner_text = " / ".join(value for value in (asset.department, asset.owner) if value) or "未分配"
        user_text = " / ".join(value for value in (asset.use_department, asset.user) if value) or "未分配"
        return (
            "<table width='100%' cellspacing='0' cellpadding='12' "
            "style='background:#F7FBFC; border:1px solid #D6E1E5; "
            "border-collapse:separate; margin:0 0 12px 0;'>"
            "<tr>"
            "<td colspan='2'>"
            "<div style='color:#667085; font-size:12px; font-weight:700;'>资产概览</div>"
            f"<div style='color:#0F3D3E; font-size:22px; font-weight:800; margin-top:4px;'>{escape(asset.name or '未命名设备')}</div>"
            f"<div style='color:#667085; margin-top:4px;'>{escape(asset.asset_id)}"
            f"{' · ' + escape(asset.equipment_code) if asset.equipment_code else ''}</div>"
            "</td>"
            "<td align='right'>"
            f"<span style='background:{status_color}; color:#17212B; padding:5px 10px; "
            "font-weight:700;'>当前状态："
            f"{escape(asset.status or '未设置')}</span>"
            "</td>"
            "</tr>"
            "<tr>"
            f"<td style='color:#344054;'><b>存放地点</b><br>{escape(asset.location or '未设置')}</td>"
            f"<td style='color:#344054;'><b>管理归属</b><br>{escape(owner_text)}</td>"
            f"<td style='color:#344054;'><b>使用归属</b><br>{escape(user_text)}</td>"
            "</tr>"
            "<tr>"
            f"<td style='color:#344054;'><b>金额</b><br>¥ {asset.price:,.2f}</td>"
            f"<td style='color:#344054;'><b>标签状态</b><br>{escape(tag_text)}</td>"
            f"<td style='color:#344054;'><b>存储介质</b><br>{escape(media_text)}</td>"
            "</tr>"
            "</table>"
        )

    def _update_refresh_label(self) -> None:
        loaded_at = self.service.cache_status.loaded_at
        self.last_refresh_time = loaded_at.split(" ")[-1] if loaded_at else ""
        self._show_result_label()
        if self.service.cache_status.stale:
            self.statusBar().showMessage("缓存已过期，保存功能已禁用，请点击刷新")
        else:
            self.statusBar().showMessage(f"缓存已加载：{loaded_at}")

    def _show_write_error(self, default_title: str, error: Exception) -> None:
        if isinstance(error, CacheReloadAfterSaveError):
            title = "Excel 已保存，请勿重复操作"
        elif isinstance(error, (WorkbookChangedExternallyError, CacheUnavailableError)):
            title = "数据已变化，无法保存"
        else:
            title = default_title
        QMessageBox.critical(self, title, str(error))

    def _show_result_label(self) -> None:
        refresh_text = (
            f" · 最近刷新 {self.last_refresh_time}" if self.last_refresh_time else ""
        )
        self.result_label.setText(f"共 {len(self.assets)} 台设备{refresh_text}")

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow, QWidget {
                background: #F4F7F8;
                color: #17212B;
                font-size: 13px;
                font-family: "Microsoft YaHei UI", "Segoe UI";
            }
            #appHeader {
                background: #FFFFFF;
                border-bottom: 1px solid #D6E1E5;
            }
            #appTitle {
                color: #0F3D3E;
                font-size: 20px;
                font-weight: 700;
                padding: 10px 4px;
            }
            #pathLabel, #mutedLabel { color: #667085; }
            #sidePanel, #detailPanel {
                background: #FFFFFF;
            }
            #sidePanel {
                border-right: 1px solid #D6E1E5;
            }
            #detailPanel {
                border-left: 1px solid #D6E1E5;
            }
            #sectionLabel {
                color: #0F3D3E;
                font-weight: 700;
                margin-top: 8px;
                padding: 4px 0;
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
            QPushButton:pressed, QPushButton:checked {
                background: #DDF2EF;
                border-color: #0F766E;
                color: #0F3D3E;
            }
            #primaryButton {
                background: #0F766E;
                color: white;
                border-color: #0F766E;
                font-weight: 600;
            }
            #primaryButton:hover { background: #11887F; }
            QLineEdit, QComboBox, QDoubleSpinBox, QPlainTextEdit {
                min-height: 30px; border: 1px solid #BFD0D6; border-radius: 5px; background: white;
                padding: 0 7px;
            }
            QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus, QPlainTextEdit:focus {
                border-color: #0F766E;
            }
            QTableWidget, QTreeWidget, QTextBrowser, QListWidget {
                background: #FFFFFF;
                alternate-background-color: #F8FAFB;
                border: 1px solid #D6E1E5;
                gridline-color: #E8EEF1;
                selection-background-color: #EAF4FB;
                selection-color: #17212B;
            }
            QTreeWidget::item { min-height: 24px; padding: 2px 4px; }
            QTreeWidget::item:selected { background: #EAF4FB; color: #17212B; }
            QTreeWidget::item:selected:active,
            QTreeWidget::item:selected:!active,
            QListWidget::item:selected,
            QAbstractItemView::item:selected {
                background: #EAF4FB;
                color: #17212B;
            }
            QTableWidget::item:selected { background: #EAF4FB; color: #17212B; }
            QTableWidget::item:selected:active,
            QTableWidget::item:selected:!active {
                background: #EAF4FB;
                color: #17212B;
            }
            QComboBox QAbstractItemView::item:selected {
                background: #EAF4FB;
                color: #17212B;
            }
            QHeaderView::section {
                background: #E6F0F1;
                color: #0F3D3E;
                padding: 8px;
                border: 0;
                border-right: 1px solid #D6E1E5;
                font-weight: 600;
            }
            QTabWidget::pane { border: 1px solid #D6E1E5; background: #FFFFFF; }
            QTabBar::tab { padding: 9px 14px; background: #E7EEF0; }
            QTabBar::tab:selected {
                background: #FFFFFF;
                border-bottom: 2px solid #0F766E;
                color: #0F3D3E;
            }
            QGroupBox { font-weight: 700; border: 1px solid #D6E1E5; margin-top: 12px;
                        padding-top: 14px; background: #FFFFFF; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
            """
        )

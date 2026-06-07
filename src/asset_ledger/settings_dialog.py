from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .asset_service import AssetService, AssetServiceError
from .excel_repository import CATEGORY_SHEET, LOCATION_SHEET


class SettingsDialog(QDialog):
    def __init__(self, service: AssetService, workbook_path: Path, parent=None) -> None:
        super().__init__(parent)
        self.service = service
        self.workbook_path = Path(workbook_path)
        self.selected_path = self.workbook_path
        self.setWindowTitle("设置")
        self.setMinimumSize(720, 560)

        self.tabs = QTabWidget()
        self.operator_edit = QLineEdit()
        self.path_edit = QLineEdit(str(self.workbook_path))
        self.path_edit.setReadOnly(True)
        self.category_tree = self._new_tree()
        self.location_tree = self._new_tree()
        self.status_list = QListWidget()
        self.source_list = QListWidget()

        self.tabs.addTab(self._build_general_tab(), "常规")
        self.tabs.addTab(
            self._build_dictionary_tab(self.category_tree, CATEGORY_SHEET), "设备分类"
        )
        self.tabs.addTab(
            self._build_dictionary_tab(self.location_tree, LOCATION_SHEET), "位置"
        )
        self.tabs.addTab(self._build_enum_tab(), "状态与来源")

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("保存并关闭")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        buttons.accepted.connect(self._save_general)
        buttons.rejected.connect(self.reject)
        layout = QVBoxLayout(self)
        layout.addWidget(self.tabs)
        layout.addWidget(buttons)
        self._load_data()

    @staticmethod
    def _new_tree() -> QTreeWidget:
        tree = QTreeWidget()
        tree.setHeaderLabels(["名称", "状态"])
        tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        tree.setColumnWidth(0, 360)
        return tree

    def _build_general_tab(self) -> QWidget:
        tab = QWidget()
        form = QFormLayout(tab)
        choose_path_button = QPushButton("选择工作簿")
        choose_path_button.clicked.connect(self._choose_path)
        path_row = QWidget()
        path_layout = QHBoxLayout(path_row)
        path_layout.setContentsMargins(0, 0, 0, 0)
        path_layout.addWidget(self.path_edit, 1)
        path_layout.addWidget(choose_path_button)
        form.addRow("当前操作者", self.operator_edit)
        form.addRow("Excel 工作簿", path_row)
        form.addRow("", QLabel("更换工作簿后主窗口会自动重新加载。"))
        return tab

    def _build_dictionary_tab(self, tree: QTreeWidget, sheet_name: str) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.addWidget(tree)
        controls = QHBoxLayout()
        add_root = QPushButton("新增一级项")
        add_child = QPushButton("新增子项")
        toggle = QPushButton("启用 / 停用")
        add_root.clicked.connect(lambda: self._add_dictionary(sheet_name, tree, False))
        add_child.clicked.connect(lambda: self._add_dictionary(sheet_name, tree, True))
        toggle.clicked.connect(lambda: self._toggle_dictionary(sheet_name, tree))
        controls.addWidget(add_root)
        controls.addWidget(add_child)
        controls.addWidget(toggle)
        controls.addStretch()
        layout.addLayout(controls)
        return tab

    def _build_enum_tab(self) -> QWidget:
        tab = QWidget()
        layout = QHBoxLayout(tab)
        for title, widget, enum_type in (
            ("设备状态", self.status_list, "状态"),
            ("设备来源", self.source_list, "来源"),
        ):
            panel = QWidget()
            panel_layout = QVBoxLayout(panel)
            panel_layout.addWidget(QLabel(title))
            panel_layout.addWidget(widget)
            buttons = QHBoxLayout()
            add = QPushButton("新增")
            remove = QPushButton("移除")
            add.clicked.connect(lambda _=False, t=enum_type, w=widget: self._add_enum(t, w))
            remove.clicked.connect(lambda _=False, t=enum_type, w=widget: self._remove_enum(t, w))
            buttons.addWidget(add)
            buttons.addWidget(remove)
            panel_layout.addLayout(buttons)
            layout.addWidget(panel)
        return tab

    def _load_data(self) -> None:
        self.operator_edit.setText(self.service.get_operator())
        self._populate_tree(self.category_tree, self.service.get_categories(include_disabled=True))
        self._populate_tree(self.location_tree, self.service.get_locations(include_disabled=True))
        self.status_list.addItems(self.service.get_enum_values("状态", include_disabled=True))
        self.source_list.addItems(self.service.get_enum_values("来源", include_disabled=True))

    @staticmethod
    def _populate_tree(tree: QTreeWidget, items) -> None:
        tree.clear()
        by_id: dict[str, QTreeWidgetItem] = {}
        for item in items:
            node = QTreeWidgetItem([item.name, "启用" if item.enabled else "停用"])
            node.setData(0, Qt.ItemDataRole.UserRole, item.item_id)
            node.setData(0, Qt.ItemDataRole.UserRole + 1, item.enabled)
            if not item.enabled:
                node.setForeground(0, Qt.GlobalColor.gray)
            by_id[item.item_id] = node
            if item.parent_id and item.parent_id in by_id:
                by_id[item.parent_id].addChild(node)
            else:
                tree.addTopLevelItem(node)
        tree.expandAll()

    def _add_dictionary(self, sheet_name: str, tree: QTreeWidget, child: bool) -> None:
        parent_id = ""
        if child:
            selected = tree.currentItem()
            if not selected:
                QMessageBox.information(self, "请选择父项", "请先选择一个父级项目。")
                return
            parent_id = str(selected.data(0, Qt.ItemDataRole.UserRole))
        name, accepted = QInputDialog.getText(self, "新增项目", "名称")
        if not accepted:
            return
        try:
            self.service.add_dictionary_item(sheet_name, name, parent_id=parent_id)
            self._refresh_tree(sheet_name, tree)
        except AssetServiceError as exc:
            QMessageBox.warning(self, "无法新增", str(exc))

    def _toggle_dictionary(self, sheet_name: str, tree: QTreeWidget) -> None:
        selected = tree.currentItem()
        if not selected:
            return
        item_id = str(selected.data(0, Qt.ItemDataRole.UserRole))
        enabled = bool(selected.data(0, Qt.ItemDataRole.UserRole + 1))
        self.service.set_dictionary_item_enabled(sheet_name, item_id, not enabled)
        self._refresh_tree(sheet_name, tree)

    def _refresh_tree(self, sheet_name: str, tree: QTreeWidget) -> None:
        items = (
            self.service.get_categories(include_disabled=True)
            if sheet_name == CATEGORY_SHEET
            else self.service.get_locations(include_disabled=True)
        )
        self._populate_tree(tree, items)

    def _add_enum(self, enum_type: str, widget: QListWidget) -> None:
        value, accepted = QInputDialog.getText(self, f"新增{enum_type}", "名称")
        if accepted and value.strip():
            values = [widget.item(index).text() for index in range(widget.count())]
            values.append(value.strip())
            self.service.set_enum_values(enum_type, values)
            widget.clear()
            widget.addItems(self.service.get_enum_values(enum_type))

    def _remove_enum(self, enum_type: str, widget: QListWidget) -> None:
        row = widget.currentRow()
        if row < 0:
            return
        values = [
            widget.item(index).text() for index in range(widget.count()) if index != row
        ]
        try:
            self.service.set_enum_values(enum_type, values)
            widget.clear()
            widget.addItems(values)
        except AssetServiceError as exc:
            QMessageBox.warning(self, "无法移除", str(exc))

    def _choose_path(self) -> None:
        selected, _ = QFileDialog.getSaveFileName(
            self,
            "选择设备资产工作簿",
            str(self.workbook_path),
            "Excel 工作簿 (*.xlsx)",
        )
        if selected:
            path = Path(selected)
            self.selected_path = path if path.suffix.lower() == ".xlsx" else path.with_suffix(".xlsx")
            self.path_edit.setText(str(self.selected_path))

    def _save_general(self) -> None:
        self.service.set_operator(self.operator_edit.text())
        self.accept()

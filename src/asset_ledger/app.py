from __future__ import annotations

from pathlib import Path
import sys

from PySide6.QtCore import QSettings
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QApplication, QMessageBox

from .asset_service import AssetService
from .excel_repository import ExcelRepository
from .main_window import MainWindow


def main() -> int:
    application = QApplication(sys.argv)
    application.setApplicationName("设备资产台账")
    application.setOrganizationName("AssetLedger")
    application.setFont(QFont("Microsoft YaHei UI", 9))
    settings = QSettings()
    default_path = Path.cwd() / "data" / "设备资产台账.xlsx"
    workbook_path = Path(settings.value("workbook_path", str(default_path)))
    repository = ExcelRepository(workbook_path)
    try:
        repository.initialize()
        errors = repository.validate_workbook()
        if errors:
            QMessageBox.critical(None, "无法打开工作簿", "\n".join(errors))
            return 1
    except Exception as exc:
        QMessageBox.critical(None, "启动失败", str(exc))
        return 1

    window = MainWindow(AssetService(repository), workbook_path)
    window.workbook_changed.connect(lambda path: settings.setValue("workbook_path", str(path)))
    window.show()
    return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

from pathlib import Path
import sys

from PySide6.QtCore import QSettings
from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import QApplication, QMessageBox

from .asset_service import AssetService
from .excel_repository import ExcelRepository
from .main_window import MainWindow


def application_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd()


def default_workbook_path() -> Path:
    return application_base_dir() / "data" / "设备资产台账.xlsx"


def resource_path(relative_path: str | Path) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path
    return Path.cwd() / relative_path


def main() -> int:
    application = QApplication(sys.argv)
    application.setApplicationName("设备资产台账")
    application.setOrganizationName("AssetLedger")
    application.setFont(QFont("Microsoft YaHei UI", 9))
    icon_path = resource_path("assets/app_icon.ico")
    if icon_path.exists():
        application.setWindowIcon(QIcon(str(icon_path)))
    settings = QSettings()
    default_path = default_workbook_path()
    workbook_path = Path(settings.value("workbook_path", str(default_path)))
    repository = ExcelRepository(workbook_path)
    try:
        repository.initialize()
        service = AssetService(repository)
    except Exception as exc:
        QMessageBox.critical(None, "启动失败", str(exc))
        return 1

    window = MainWindow(service, workbook_path)
    window.workbook_changed.connect(lambda path: settings.setValue("workbook_path", str(path)))
    window.show()
    return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

from PySide6.QtWidgets import QComboBox, QDoubleSpinBox


class NoWheelComboBox(QComboBox):
    def wheelEvent(self, event) -> None:
        event.ignore()


class NoWheelDoubleSpinBox(QDoubleSpinBox):
    def wheelEvent(self, event) -> None:
        event.ignore()

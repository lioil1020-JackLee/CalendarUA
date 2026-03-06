from __future__ import annotations

from PySide6.QtCore import QItemSelectionModel, Qt
from PySide6.QtWidgets import QListWidget


class WheelSelectListWidget(QListWidget):
    """List widget for picking one/multiple items with predictable wheel behavior."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSelectionMode(QListWidget.SingleSelection)
        self._hover_select_enabled = False
        self.setMouseTracking(True)

    def set_hover_select_enabled(self, enabled: bool):
        self._hover_select_enabled = bool(enabled)

    def mouseMoveEvent(self, event):
        if self._hover_select_enabled:
            item = self.itemAt(event.pos())
            if item is not None:
                self.setCurrentItem(item)
        super().mouseMoveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.selectionMode() == QListWidget.ExtendedSelection:
            item = self.itemAt(event.pos())
            if item is not None and not (event.modifiers() & (Qt.ControlModifier | Qt.ShiftModifier)):
                idx = self.indexFromItem(item)
                self.selectionModel().setCurrentIndex(idx, QItemSelectionModel.ClearAndSelect)
                event.accept()
                return
        super().mousePressEvent(event)

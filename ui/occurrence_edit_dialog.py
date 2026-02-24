from __future__ import annotations

from datetime import datetime, timedelta

from PySide6.QtWidgets import (
    QDateTimeEdit,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)


class OccurrenceEditDialog(QDialog):
    """編輯單次 occurrence 的例外資料"""

    def __init__(
        self,
        parent=None,
        title: str = "",
        target_value: str = "",
        start_dt: datetime | None = None,
        end_dt: datetime | None = None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Edit Occurrence")
        self.setMinimumWidth(420)
        self.setModal(True)

        if start_dt is None:
            start_dt = datetime.now().replace(minute=0, second=0, microsecond=0)
        if end_dt is None:
            end_dt = start_dt + timedelta(hours=1)

        layout = QVBoxLayout(self)

        hint = QLabel("修改本次 occurrence（不影響整個 series）")
        layout.addWidget(hint)

        form = QFormLayout()
        self.title_edit = QLineEdit(title)
        self.target_value_edit = QLineEdit(target_value)

        self.start_edit = QDateTimeEdit(start_dt)
        self.start_edit.setDisplayFormat("yyyy/MM/dd HH:mm:ss")
        self.start_edit.setCalendarPopup(True)

        self.end_edit = QDateTimeEdit(end_dt)
        self.end_edit.setDisplayFormat("yyyy/MM/dd HH:mm:ss")
        self.end_edit.setCalendarPopup(True)

        form.addRow("Subject", self.title_edit)
        form.addRow("ValueSet Value", self.target_value_edit)
        form.addRow("Start", self.start_edit)
        form.addRow("End", self.end_edit)

        layout.addLayout(form)

        buttons = QHBoxLayout()
        buttons.addStretch()
        ok_btn = QPushButton("Ok")
        ok_btn.clicked.connect(self._on_ok)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        buttons.addWidget(ok_btn)
        buttons.addWidget(cancel_btn)
        layout.addLayout(buttons)

    def _on_ok(self):
        start_dt = self.start_edit.dateTime().toPython()
        end_dt = self.end_edit.dateTime().toPython()
        if end_dt <= start_dt:
            QMessageBox.warning(self, "時間錯誤", "End 時間必須晚於 Start 時間。")
            return
        self.accept()

    def get_data(self) -> dict:
        return {
            "title": self.title_edit.text().strip(),
            "target_value": self.target_value_edit.text().strip(),
            "start": self.start_edit.dateTime().toPython(),
            "end": self.end_edit.dateTime().toPython(),
        }

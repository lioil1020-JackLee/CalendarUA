from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QLabel,
    QRadioButton,
    QButtonGroup,
    QHBoxLayout,
    QPushButton,
)
from PySide6.QtCore import Qt


class OccurrenceChoiceDialog(QDialog):
    """Recurring 事件開啟方式選擇對話框"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Appointment Recurrence")
        self.setModal(True)
        self.setMinimumWidth(320)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        title = QLabel("Open Recurring Item?")
        title.setAlignment(Qt.AlignLeft)
        layout.addWidget(title)

        self.radio_occurrence = QRadioButton("Open this occurrence.")
        self.radio_series = QRadioButton("Open the series.")
        self.radio_occurrence.setChecked(True)

        self.group = QButtonGroup(self)
        self.group.addButton(self.radio_occurrence)
        self.group.addButton(self.radio_series)

        layout.addWidget(self.radio_occurrence)
        layout.addWidget(self.radio_series)

        buttons = QHBoxLayout()
        buttons.addStretch()

        ok_btn = QPushButton("Ok")
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)

        buttons.addWidget(ok_btn)
        buttons.addWidget(cancel_btn)

        layout.addLayout(buttons)

    def selected_mode(self) -> str:
        return "occurrence" if self.radio_occurrence.isChecked() else "series"

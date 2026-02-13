from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QTabWidget,
    QWidget,
    QLabel,
    QComboBox,
    QSpinBox,
    QCheckBox,
    QPushButton,
    QDateEdit,
    QTimeEdit,
    QGroupBox,
    QRadioButton,
    QButtonGroup,
    QGridLayout,
    QFrame,
)
from PySide6.QtCore import Qt, QDate, QTime, Signal, QTimer
from PySide6.QtGui import QFont, QColor
import sys


class RecurrenceDialog(QDialog):
    """週期性設定對話框 - 模仿 Office/Outlook 風格"""

    rrule_created = Signal(str)

    def __init__(self, parent=None, current_rrule: str = ""):
        super().__init__(parent)
        self.current_rrule = current_rrule
        self.setup_ui()
        self.apply_modern_style()

        if current_rrule:
            self.parse_existing_rrule(current_rrule)

    def setup_ui(self):
        self.setWindowTitle("週期性設定")
        self.setMinimumWidth(450)
        self.setMinimumHeight(400)
        self.setModal(True)

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(20, 20, 20, 20)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)

        self.daily_tab = self.create_daily_tab()
        self.weekly_tab = self.create_weekly_tab()
        self.monthly_tab = self.create_monthly_tab()

        self.tabs.addTab(self.daily_tab, "每天")
        self.tabs.addTab(self.weekly_tab, "每週")
        self.tabs.addTab(self.monthly_tab, "每月")

        main_layout.addWidget(self.tabs)

        main_layout.addWidget(self.create_time_group())

        main_layout.addWidget(self.create_range_group())

        main_layout.addWidget(self.create_button_group())

        self.setLayout(main_layout)

    def create_daily_tab(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(15)

        options_group = QGroupBox("每天重複")
        options_layout = QVBoxLayout(options_group)

        self.daily_every_radio = QRadioButton("每天")
        self.daily_every_radio.setChecked(True)
        self.daily_weekday_radio = QRadioButton("每個工作天 (星期一至星期五)")

        options_layout.addWidget(self.daily_every_radio)
        options_layout.addWidget(self.daily_weekday_radio)

        layout.addWidget(options_group)
        layout.addStretch()

        return widget

    def create_weekly_tab(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(15)

        pattern_group = QGroupBox("每週重複")
        pattern_layout = QVBoxLayout(pattern_group)

        every_week_layout = QHBoxLayout()
        every_week_layout.addWidget(QLabel("每"))
        self.weekly_interval = QSpinBox()
        self.weekly_interval.setMinimum(1)
        self.weekly_interval.setMaximum(52)
        self.weekly_interval.setValue(1)
        self.weekly_interval.setFixedWidth(60)
        every_week_layout.addWidget(self.weekly_interval)
        every_week_layout.addWidget(QLabel("週"))
        every_week_layout.addStretch()
        pattern_layout.addLayout(every_week_layout)

        days_layout = QGridLayout()
        days_layout.setSpacing(8)

        self.day_checkboxes = {}
        days = [
            ("星期一", "MO"),
            ("星期二", "TU"),
            ("星期三", "WE"),
            ("星期四", "TH"),
            ("星期五", "FR"),
            ("星期六", "SA"),
            ("星期日", "SU"),
        ]

        for i, (day_name, day_code) in enumerate(days):
            checkbox = QCheckBox(day_name)
            self.day_checkboxes[day_code] = checkbox
            row = i // 4
            col = i % 4
            days_layout.addWidget(checkbox, row, col)

        pattern_layout.addLayout(days_layout)

        layout.addWidget(pattern_group)
        layout.addStretch()

        return widget

    def create_monthly_tab(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(15)

        pattern_group = QGroupBox("每月重複")
        pattern_layout = QVBoxLayout(pattern_group)

        self.monthly_day_radio = QRadioButton("在第")
        self.monthly_day_radio.setChecked(True)

        day_pattern_layout = QHBoxLayout()
        day_pattern_layout.addWidget(self.monthly_day_radio)

        self.monthly_day_number = QSpinBox()
        self.monthly_day_number.setMinimum(1)
        self.monthly_day_number.setMaximum(31)
        self.monthly_day_number.setValue(1)
        self.monthly_day_number.setFixedWidth(60)
        day_pattern_layout.addWidget(self.monthly_day_number)

        day_pattern_layout.addWidget(QLabel("天"))
        day_pattern_layout.addStretch()

        pattern_layout.addLayout(day_pattern_layout)

        self.monthly_week_radio = QRadioButton("在第")

        week_pattern_layout = QHBoxLayout()
        week_pattern_layout.addWidget(self.monthly_week_radio)

        self.monthly_week_num = QComboBox()
        self.monthly_week_num.addItems(["1", "2", "3", "4", "5"])
        self.monthly_week_num.setFixedWidth(60)
        week_pattern_layout.addWidget(self.monthly_week_num)

        self.monthly_week_day = QComboBox()
        self.monthly_week_day.addItems(
            ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"]
        )
        week_pattern_layout.addWidget(self.monthly_week_day)

        week_pattern_layout.addWidget(QLabel("天"))
        week_pattern_layout.addStretch()

        pattern_layout.addLayout(week_pattern_layout)

        every_month_layout = QHBoxLayout()
        every_month_layout.addWidget(QLabel("每"))
        self.monthly_interval = QSpinBox()
        self.monthly_interval.setMinimum(1)
        self.monthly_interval.setMaximum(12)
        self.monthly_interval.setValue(1)
        self.monthly_interval.setFixedWidth(60)
        every_month_layout.addWidget(self.monthly_interval)
        every_month_layout.addWidget(QLabel("月的第"))
        every_month_layout.addWidget(self.monthly_day_number)
        every_month_layout.addWidget(QLabel("天"))
        every_month_layout.addStretch()

        pattern_layout.addLayout(every_month_layout)

        layout.addWidget(pattern_group)
        layout.addStretch()

        return widget

    def create_time_group(self) -> QGroupBox:
        group = QGroupBox("時間")
        layout = QHBoxLayout(group)

        layout.addWidget(QLabel("於"))

        self.time_edit = QTimeEdit()
        self.time_edit.setDisplayFormat("HH:mm")
        self.time_edit.setTime(QTime(8, 0))
        self.time_edit.setFixedWidth(100)
        layout.addWidget(self.time_edit)

        layout.addStretch()

        return group

    def create_range_group(self) -> QGroupBox:
        group = QGroupBox("範圍")
        layout = QVBoxLayout(group)

        start_layout = QHBoxLayout()
        start_layout.addWidget(QLabel("開始於:"))

        self.start_date = QDateEdit()
        self.start_date.setDisplayFormat("yyyy/MM/dd")
        self.start_date.setDate(QDate.currentDate())
        self.start_date.setFixedWidth(120)
        start_layout.addWidget(self.start_date)

        self.start_time = QTimeEdit()
        self.start_time.setDisplayFormat("HH:mm")
        self.start_time.setTime(QTime(8, 0))
        self.start_time.setFixedWidth(80)
        start_layout.addWidget(self.start_time)

        start_layout.addStretch()
        layout.addLayout(start_layout)

        end_layout = QHBoxLayout()

        self.end_never_radio = QRadioButton("無結束日期")
        self.end_never_radio.setChecked(True)
        end_layout.addWidget(self.end_never_radio)

        end_layout.addSpacing(20)

        self.end_after_radio = QRadioButton("在")
        end_layout.addWidget(self.end_after_radio)

        self.end_occurrences = QSpinBox()
        self.end_occurrences.setMinimum(1)
        self.end_occurrences.setMaximum(999)
        self.end_occurrences.setValue(10)
        self.end_occurrences.setFixedWidth(60)
        end_layout.addWidget(self.end_occurrences)

        end_layout.addWidget(QLabel("次後結束"))

        end_layout.addStretch()
        layout.addLayout(end_layout)

        end_date_layout = QHBoxLayout()

        self.end_on_radio = QRadioButton("結束於:")
        end_date_layout.addWidget(self.end_on_radio)

        self.end_date = QDateEdit()
        self.end_date.setDisplayFormat("yyyy/MM/dd")
        self.end_date.setDate(QDate.currentDate().addMonths(3))
        self.end_date.setFixedWidth(120)
        end_date_layout.addWidget(self.end_date)

        end_date_layout.addStretch()
        layout.addLayout(end_date_layout)

        return group

    def create_button_group(self) -> QWidget:
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 10, 0, 0)

        layout.addStretch()

        cancel_btn = QPushButton("取消")
        cancel_btn.setFixedWidth(80)
        cancel_btn.clicked.connect(self.reject)
        layout.addWidget(cancel_btn)

        ok_btn = QPushButton("確定")
        ok_btn.setFixedWidth(80)
        ok_btn.setDefault(True)
        ok_btn.clicked.connect(self.on_ok_clicked)
        layout.addWidget(ok_btn)

        return widget

    def apply_modern_style(self):
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #d0d0d0;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 12px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 8px;
                color: #2c3e50;
            }
            QTabWidget::pane {
                border: 1px solid #d0d0d0;
                border-radius: 6px;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #e8e8e8;
                border: 1px solid #d0d0d0;
                border-bottom: none;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 8px 20px;
                margin-right: 2px;
                color: #555;
            }
            QTabBar::tab:selected {
                background-color: white;
                color: #0078d4;
                font-weight: bold;
            }
            QTabBar::tab:hover:!selected {
                background-color: #d8d8d8;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 6px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QPushButton[default="true"] {
                background-color: #0078d4;
            }
            QPushButton[default="true"]:hover {
                background-color: #106ebe;
            }
            QCheckBox {
                spacing: 8px;
                color: #333;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 3px;
                border: 2px solid #a0a0a0;
            }
            QCheckBox::indicator:checked {
                background-color: #0078d4;
                border-color: #0078d4;
            }
            QRadioButton {
                spacing: 8px;
                color: #333;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
            }
            QRadioButton::indicator:checked {
                background-color: #0078d4;
            }
            QSpinBox, QComboBox, QDateEdit, QTimeEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 4px 8px;
                background-color: white;
            }
            QSpinBox:focus, QComboBox:focus, QDateEdit:focus, QTimeEdit:focus {
                border: 2px solid #0078d4;
            }
            QLabel {
                color: #333;
            }
        """)

    def build_rrule(self) -> str:
        freq = ""
        byday = ""
        bymonthday = ""
        bysetpos = ""
        interval = 1
        until = ""
        count = 0

        time = self.time_edit.time()
        hour = time.hour()
        minute = time.minute()

        start_date = self.start_date.date()
        dtstart = f"{start_date.year()}{start_date.month():02d}{start_date.day():02d}T{hour:02d}{minute:02d}00"

        tab_index = self.tabs.currentIndex()

        if tab_index == 0:
            freq = "DAILY"
            if self.daily_weekday_radio.isChecked():
                byday = "MO,TU,WE,TH,FR"
                interval = 1
        elif tab_index == 1:
            freq = "WEEKLY"
            interval = self.weekly_interval.value()

            selected_days = []
            for day_code, checkbox in self.day_checkboxes.items():
                if checkbox.isChecked():
                    selected_days.append(day_code)

            if selected_days:
                byday = ",".join(selected_days)
        elif tab_index == 2:
            freq = "MONTHLY"
            interval = self.monthly_interval.value()

            if self.monthly_day_radio.isChecked():
                bymonthday = str(self.monthly_day_number.value())
            else:
                week_num = int(self.monthly_week_num.currentText())
                day_map = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"]
                byday = day_map[self.monthly_week_day.currentIndex()]
                bysetpos = str(week_num)

        if self.end_never_radio.isChecked():
            pass
        elif self.end_after_radio.isChecked():
            count = self.end_occurrences.value()
        elif self.end_on_radio.isChecked():
            end_date = self.end_date.date()
            until = (
                f"{end_date.year()}{end_date.month():02d}{end_date.day():02d}T235959"
            )

        parts = [f"FREQ={freq}"]

        if interval > 1:
            parts.append(f"INTERVAL={interval}")

        if byday:
            parts.append(f"BYDAY={byday}")

        if bymonthday:
            parts.append(f"BYMONTHDAY={bymonthday}")

        if bysetpos:
            parts.append(f"BYSETPOS={bysetpos}")

        parts.append(f"BYHOUR={hour}")
        parts.append(f"BYMINUTE={minute}")

        if count > 0:
            parts.append(f"COUNT={count}")

        if until:
            parts.append(f"UNTIL={until}")

        parts.append(f"DTSTART:{dtstart}")

        return ";".join(parts)

    def parse_existing_rrule(self, rrule_str: str):
        try:
            import re

            parts = rrule_str.split(";")
            freq = ""
            interval = 1
            byday = ""
            bymonthday = ""
            bysetpos = ""
            byhour = 8
            byminute = 0

            dtstart_match = re.search(
                r"DTSTART:(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})", rrule_str
            )
            if dtstart_match:
                year, month, day = (
                    int(dtstart_match.group(1)),
                    int(dtstart_match.group(2)),
                    int(dtstart_match.group(3)),
                )
                hour, minute = int(dtstart_match.group(4)), int(dtstart_match.group(5))
                self.start_date.setDate(QDate(year, month, day))
                self.start_time.setTime(QTime(hour, minute))
                self.time_edit.setTime(QTime(hour, minute))

            for part in parts:
                if part.startswith("FREQ="):
                    freq = part.split("=")[1]
                elif part.startswith("INTERVAL="):
                    interval = int(part.split("=")[1])
                elif part.startswith("BYDAY="):
                    byday = part.split("=")[1]
                elif part.startswith("BYMONTHDAY="):
                    bymonthday = part.split("=")[1]
                elif part.startswith("BYSETPOS="):
                    bysetpos = part.split("=")[1]
                elif part.startswith("BYHOUR="):
                    byhour = int(part.split("=")[1])
                elif part.startswith("BYMINUTE="):
                    byminute = int(part.split("=")[1])
                elif part.startswith("COUNT="):
                    count = int(part.split("=")[1])
                    self.end_after_radio.setChecked(True)
                    self.end_occurrences.setValue(count)
                elif part.startswith("UNTIL="):
                    until_str = part.split("=")[1][:8]
                    year, month, day = (
                        int(until_str[:4]),
                        int(until_str[4:6]),
                        int(until_str[6:8]),
                    )
                    self.end_on_radio.setChecked(True)
                    self.end_date.setDate(QDate(year, month, day))

            self.time_edit.setTime(QTime(byhour, byminute))

            if freq == "DAILY":
                self.tabs.setCurrentIndex(0)
                if byday == "MO,TU,WE,TH,FR":
                    self.daily_weekday_radio.setChecked(True)
                else:
                    self.daily_every_radio.setChecked(True)
            elif freq == "WEEKLY":
                self.tabs.setCurrentIndex(1)
                self.weekly_interval.setValue(interval)
                if byday:
                    for day_code in byday.split(","):
                        if day_code in self.day_checkboxes:
                            self.day_checkboxes[day_code].setChecked(True)
            elif freq == "MONTHLY":
                self.tabs.setCurrentIndex(2)
                self.monthly_interval.setValue(interval)
                if bymonthday:
                    self.monthly_day_radio.setChecked(True)
                    self.monthly_day_number.setValue(int(bymonthday))
                elif bysetpos and byday:
                    self.monthly_week_radio.setChecked(True)
                    self.monthly_week_num.setCurrentText(bysetpos)
                    day_map = {
                        "SU": 0,
                        "MO": 1,
                        "TU": 2,
                        "WE": 3,
                        "TH": 4,
                        "FR": 5,
                        "SA": 6,
                    }
                    if byday in day_map:
                        self.monthly_week_day.setCurrentIndex(day_map[byday])

        except Exception as e:
            print(f"解析現有 RRULE 失敗: {e}")

    def on_ok_clicked(self):
        rrule_str = self.build_rrule()
        self.rrule_created.emit(rrule_str)
        self.accept()

    def get_rrule(self) -> str:
        return self.build_rrule()


def show_recurrence_dialog(parent=None, current_rrule: str = "") -> str:
    """
    顯示週期性設定對話框並返回 RRULE 字串

    Args:
        parent: 父視窗
        current_rrule: 現有的 RRULE 字串（可選，用於編輯）

    Returns:
        str: RRULE 字串，使用者取消則返回空字串
    """
    dialog = RecurrenceDialog(parent, current_rrule)
    if dialog.exec() == QDialog.Accepted:
        return dialog.get_rrule()
    return ""


if __name__ == "__main__":
    from PySide6.QtWidgets import QApplication

    app = QApplication(sys.argv)

    rrule = show_recurrence_dialog()
    print(f"生成的 RRULE: {rrule}")

    app.exit()

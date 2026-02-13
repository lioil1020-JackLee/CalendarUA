"""
週期性設定對話框 - Outlook 風格
提供每天、每週、每月、每年的循環設定
"""

from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
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
from PySide6.QtCore import Qt, QDate, QTime, Signal
from PySide6.QtGui import QFont
import sys


class RecurrenceDialog(QDialog):
    """週期性設定對話框 - Outlook 風格"""

    rrule_created = Signal(str)

    def __init__(self, parent=None, current_rrule: str = ""):
        super().__init__(parent)
        self.current_rrule = current_rrule
        self.setWindowTitle("週期性約會")
        self.setMinimumWidth(520)
        self.setMinimumHeight(480)
        self.setModal(True)

        self.setup_ui()
        self.apply_modern_style()
        self.connect_signals()

        if current_rrule:
            self.parse_existing_rrule(current_rrule)

    def setup_ui(self):
        """設定主介面"""
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # 約會時間區塊
        main_layout.addWidget(self.create_time_group())

        # 循環模式區塊
        main_layout.addWidget(self.create_recurrence_pattern_group())

        # 循環範圍區塊
        main_layout.addWidget(self.create_range_group())

        # 按鈕
        main_layout.addWidget(self.create_button_group())

    def create_time_group(self) -> QGroupBox:
        """建立約會時間區塊"""
        group = QGroupBox("約會時間")
        layout = QGridLayout(group)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        # 開始時間
        start_label = QLabel("開始(T):")
        start_label.setObjectName("fieldLabel")
        layout.addWidget(start_label, 0, 0)
        self.start_time_edit = QTimeEdit()
        self.start_time_edit.setDisplayFormat("hh:mm")
        self.start_time_edit.setTime(QTime(9, 0))
        self.start_time_edit.setFixedWidth(80)
        layout.addWidget(self.start_time_edit, 0, 1)

        # 結束時間
        end_label = QLabel("結束(N):")
        end_label.setObjectName("fieldLabel")
        layout.addWidget(end_label, 1, 0)
        self.end_time_edit = QTimeEdit()
        self.end_time_edit.setDisplayFormat("hh:mm")
        self.end_time_edit.setTime(QTime(9, 30))
        self.end_time_edit.setFixedWidth(80)
        layout.addWidget(self.end_time_edit, 1, 1)

        # 期間
        duration_label = QLabel("期間(U):")
        duration_label.setObjectName("fieldLabel")
        layout.addWidget(duration_label, 2, 0)
        self.duration_combo = QComboBox()
        self.duration_combo.setFixedWidth(130)
        self.update_duration_combo()
        layout.addWidget(self.duration_combo, 2, 1)

        layout.setColumnStretch(2, 1)
        return group

    def update_duration_combo(self):
        """更新期間下拉選單"""
        self.duration_combo.clear()
        durations = [
            ("0 分", 0),
            ("15 分", 15),
            ("30 分", 30),
            ("45 分", 45),
            ("1 小時", 60),
            ("1.5 小時", 90),
            ("2 小時", 120),
            ("3 小時", 180),
            ("4 小時", 240),
            ("5 小時", 300),
            ("6 小時", 360),
            ("7 小時", 420),
            ("8 小時", 480),
        ]
        for text, minutes in durations:
            self.duration_combo.addItem(text, minutes)
        self.duration_combo.setCurrentIndex(2)  # 預設 30 分

    def create_recurrence_pattern_group(self) -> QGroupBox:
        """建立循環模式區塊"""
        group = QGroupBox("循環模式")
        layout = QHBoxLayout(group)
        layout.setSpacing(15)
        layout.setContentsMargins(12, 12, 12, 12)

        # 左側：頻率選擇
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(8)
        left_layout.setContentsMargins(0, 0, 0, 0)

        self.freq_button_group = QButtonGroup(self)

        self.radio_daily = QRadioButton("每天(D)")
        self.radio_daily.setChecked(True)
        self.freq_button_group.addButton(self.radio_daily)
        left_layout.addWidget(self.radio_daily)

        self.radio_weekly = QRadioButton("每週(W)")
        self.freq_button_group.addButton(self.radio_weekly)
        left_layout.addWidget(self.radio_weekly)

        self.radio_monthly = QRadioButton("每月(M)")
        self.freq_button_group.addButton(self.radio_monthly)
        left_layout.addWidget(self.radio_monthly)

        self.radio_yearly = QRadioButton("每年(Y)")
        self.freq_button_group.addButton(self.radio_yearly)
        left_layout.addWidget(self.radio_yearly)

        left_layout.addStretch()
        layout.addWidget(left_widget)

        # 分隔線
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.VLine)
        separator.setStyleSheet("color: #d0d0d0;")
        layout.addWidget(separator)

        # 右側：詳細設定
        self.detail_widget = QWidget()
        self.detail_layout = QVBoxLayout(self.detail_widget)
        self.detail_layout.setSpacing(8)
        self.detail_layout.setContentsMargins(0, 0, 0, 0)

        # 建立各頻率的詳細設定面板
        self.create_daily_detail()
        self.create_weekly_detail()
        self.create_monthly_detail()
        self.create_yearly_detail()

        layout.addWidget(self.detail_widget, 1)
        return group

    def create_daily_detail(self):
        """建立每天選項的詳細設定"""
        self.daily_widget = QWidget()
        layout = QHBoxLayout(self.daily_widget)
        layout.setSpacing(5)
        layout.setContentsMargins(0, 0, 0, 0)

        self.radio_daily_every = QRadioButton("每(V)")
        self.radio_daily_every.setChecked(True)
        layout.addWidget(self.radio_daily_every)

        self.daily_interval = QSpinBox()
        self.daily_interval.setMinimum(1)
        self.daily_interval.setMaximum(999)
        self.daily_interval.setValue(1)
        self.daily_interval.setFixedWidth(50)
        layout.addWidget(self.daily_interval)

        layout.addWidget(QLabel("天"))
        layout.addStretch()

        self.daily_weekday_radio = QRadioButton("每個工作日(K)")
        layout.addWidget(self.daily_weekday_radio)
        layout.addStretch()

        self.detail_layout.addWidget(self.daily_widget)

    def create_weekly_detail(self):
        """建立每週選項的詳細設定"""
        self.weekly_widget = QWidget()
        layout = QVBoxLayout(self.weekly_widget)
        layout.setSpacing(8)
        layout.setContentsMargins(0, 0, 0, 0)

        # 每幾週
        top_layout = QHBoxLayout()
        repeat_label = QLabel("重複於每(C)")
        repeat_label.setObjectName("fieldLabel")
        top_layout.addWidget(repeat_label)
        self.weekly_interval = QSpinBox()
        self.weekly_interval.setMinimum(1)
        self.weekly_interval.setMaximum(52)
        self.weekly_interval.setValue(1)
        self.weekly_interval.setFixedWidth(50)
        top_layout.addWidget(self.weekly_interval)
        week_label = QLabel("週的:")
        week_label.setObjectName("fieldLabel")
        top_layout.addWidget(week_label)
        top_layout.addStretch()
        layout.addLayout(top_layout)

        # 星期選擇
        days_layout = QGridLayout()
        days_layout.setSpacing(8)

        self.day_checkboxes = {}
        days = [
            ("星期日", "SU"),
            ("星期一", "MO"),
            ("星期二", "TU"),
            ("星期三", "WE"),
            ("星期四", "TH"),
            ("星期五", "FR"),
            ("星期六", "SA"),
        ]

        for i, (day_name, day_code) in enumerate(days):
            checkbox = QCheckBox(day_name)
            self.day_checkboxes[day_code] = checkbox
            row = i // 4
            col = i % 4
            days_layout.addWidget(checkbox, row, col)

        layout.addLayout(days_layout)
        self.weekly_widget.hide()
        self.detail_layout.addWidget(self.weekly_widget)

    def create_monthly_detail(self):
        """建立每月選項的詳細設定"""
        self.monthly_widget = QWidget()
        layout = QVBoxLayout(self.monthly_widget)
        layout.setSpacing(8)
        layout.setContentsMargins(0, 0, 0, 0)

        # 選項 1: 每月第 X 天
        day_layout = QHBoxLayout()
        self.radio_monthly_day = QRadioButton("每(A)")
        self.radio_monthly_day.setChecked(True)
        day_layout.addWidget(self.radio_monthly_day)

        self.monthly_interval = QSpinBox()
        self.monthly_interval.setMinimum(1)
        self.monthly_interval.setMaximum(12)
        self.monthly_interval.setValue(1)
        self.monthly_interval.setFixedWidth(50)
        day_layout.addWidget(self.monthly_interval)

        month_label = QLabel("個月的第")
        month_label.setObjectName("fieldLabel")
        day_layout.addWidget(month_label)

        self.monthly_day = QSpinBox()
        self.monthly_day.setMinimum(1)
        self.monthly_day.setMaximum(31)
        self.monthly_day.setValue(1)
        self.monthly_day.setFixedWidth(50)
        day_layout.addWidget(self.monthly_day)

        day_label = QLabel("天")
        day_label.setObjectName("fieldLabel")
        day_layout.addWidget(day_label)
        day_layout.addStretch()
        layout.addLayout(day_layout)

        # 選項 2: 每月第 X 個星期 Y
        week_layout = QHBoxLayout()
        self.radio_monthly_week = QRadioButton("每(E)")
        week_layout.addWidget(self.radio_monthly_week)

        self.monthly_week_interval = QSpinBox()
        self.monthly_week_interval.setMinimum(1)
        self.monthly_week_interval.setMaximum(12)
        self.monthly_week_interval.setValue(1)
        self.monthly_week_interval.setFixedWidth(50)
        week_layout.addWidget(self.monthly_week_interval)

        month_of_label = QLabel("個月的")
        month_of_label.setObjectName("fieldLabel")
        week_layout.addWidget(month_of_label)

        self.monthly_week_num = QComboBox()
        self.monthly_week_num.addItems(
            ["第 1 個", "第 2 個", "第 3 個", "第 4 個", "最後 1 個"]
        )
        self.monthly_week_num.setFixedWidth(80)
        week_layout.addWidget(self.monthly_week_num)

        self.monthly_week_day = QComboBox()
        self.monthly_week_day.addItems(
            ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"]
        )
        self.monthly_week_day.setFixedWidth(80)
        week_layout.addWidget(self.monthly_week_day)

        week_layout.addStretch()
        layout.addLayout(week_layout)

        self.monthly_widget.hide()
        self.detail_layout.addWidget(self.monthly_widget)

    def create_yearly_detail(self):
        """建立每年選項的詳細設定"""
        self.yearly_widget = QWidget()
        layout = QVBoxLayout(self.yearly_widget)
        layout.setSpacing(8)
        layout.setContentsMargins(0, 0, 0, 0)

        # 每幾年
        top_layout = QHBoxLayout()
        year_repeat_label = QLabel("重複於每(C)")
        year_repeat_label.setObjectName("fieldLabel")
        top_layout.addWidget(year_repeat_label)
        self.yearly_interval = QSpinBox()
        self.yearly_interval.setMinimum(1)
        self.yearly_interval.setMaximum(10)
        self.yearly_interval.setValue(1)
        self.yearly_interval.setFixedWidth(50)
        top_layout.addWidget(self.yearly_interval)
        year_label = QLabel("年的")
        year_label.setObjectName("fieldLabel")
        top_layout.addWidget(year_label)
        top_layout.addStretch()
        layout.addLayout(top_layout)

        # 選項 1: 於 X 月 Y 日
        date_layout = QHBoxLayout()
        self.radio_yearly_date = QRadioButton("於:")
        self.radio_yearly_date.setChecked(True)
        date_layout.addWidget(self.radio_yearly_date)

        self.yearly_month = QComboBox()
        self.yearly_month.addItems(
            [
                "一月",
                "二月",
                "三月",
                "四月",
                "五月",
                "六月",
                "七月",
                "八月",
                "九月",
                "十月",
                "十一月",
                "十二月",
            ]
        )
        self.yearly_month.setFixedWidth(80)
        date_layout.addWidget(self.yearly_month)

        self.yearly_day = QSpinBox()
        self.yearly_day.setMinimum(1)
        self.yearly_day.setMaximum(31)
        self.yearly_day.setValue(1)
        self.yearly_day.setFixedWidth(50)
        date_layout.addWidget(self.yearly_day)

        day_label2 = QLabel("日")
        day_label2.setObjectName("fieldLabel")
        date_layout.addWidget(day_label2)
        date_layout.addStretch()
        layout.addLayout(date_layout)

        # 選項 2: 於 X 月第 Y 個星期 Z
        week_layout = QHBoxLayout()
        self.radio_yearly_week = QRadioButton("於(E):")
        week_layout.addWidget(self.radio_yearly_week)

        self.yearly_week_month = QComboBox()
        self.yearly_week_month.addItems(
            [
                "一月",
                "二月",
                "三月",
                "四月",
                "五月",
                "六月",
                "七月",
                "八月",
                "九月",
                "十月",
                "十一月",
                "十二月",
            ]
        )
        self.yearly_week_month.setFixedWidth(80)
        week_layout.addWidget(self.yearly_week_month)

        of_label = QLabel("的")
        of_label.setObjectName("fieldLabel")
        week_layout.addWidget(of_label)

        self.yearly_week_num = QComboBox()
        self.yearly_week_num.addItems(
            ["第 1 個", "第 2 個", "第 3 個", "第 4 個", "最後 1 個"]
        )
        self.yearly_week_num.setFixedWidth(80)
        week_layout.addWidget(self.yearly_week_num)

        self.yearly_week_day = QComboBox()
        self.yearly_week_day.addItems(
            ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"]
        )
        self.yearly_week_day.setFixedWidth(80)
        week_layout.addWidget(self.yearly_week_day)

        week_layout.addStretch()
        layout.addLayout(week_layout)

        self.yearly_widget.hide()
        self.detail_layout.addWidget(self.yearly_widget)

    def create_range_group(self) -> QGroupBox:
        """建立循環範圍區塊"""
        group = QGroupBox("循環範圍")
        layout = QGridLayout(group)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        # 開始日期
        start_date_label = QLabel("開始(S):")
        start_date_label.setObjectName("fieldLabel")
        layout.addWidget(start_date_label, 0, 0)
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setDisplayFormat("yyyy/M/d (ddd)")
        self.start_date_edit.setDate(QDate.currentDate())
        self.start_date_edit.setFixedWidth(150)
        layout.addWidget(self.start_date_edit, 0, 1)

        # 結束選項
        self.end_button_group = QButtonGroup(self)

        # 結束於日期
        self.radio_end_by = QRadioButton("結束於(B):")
        self.radio_end_by.setChecked(True)
        self.end_button_group.addButton(self.radio_end_by)
        layout.addWidget(self.radio_end_by, 0, 2)

        self.end_date_edit = QDateEdit()
        self.end_date_edit.setDisplayFormat("yyyy/M/d (ddd)")
        self.end_date_edit.setDate(QDate.currentDate().addMonths(3))
        self.end_date_edit.setFixedWidth(150)
        layout.addWidget(self.end_date_edit, 0, 3)

        # 重複次數
        self.radio_end_after = QRadioButton("在反覆(F):")
        self.end_button_group.addButton(self.radio_end_after)
        layout.addWidget(self.radio_end_after, 1, 2)

        # SpinBox 和標籤放在同一個水平布局
        count_layout = QHBoxLayout()
        count_layout.setSpacing(5)
        self.end_count = QSpinBox()
        self.end_count.setMinimum(1)
        self.end_count.setMaximum(999)
        self.end_count.setValue(10)
        self.end_count.setFixedWidth(50)
        count_layout.addWidget(self.end_count)

        count_label = QLabel("次之後結束")
        count_label.setObjectName("fieldLabel")
        count_layout.addWidget(count_label)
        count_layout.addStretch()

        layout.addLayout(count_layout, 1, 3)

        # 沒有結束日期
        self.radio_end_never = QRadioButton("沒有結束日期(O)")
        self.end_button_group.addButton(self.radio_end_never)
        layout.addWidget(self.radio_end_never, 2, 2, 1, 2)

        layout.setColumnStretch(4, 1)
        return group

    def create_button_group(self) -> QWidget:
        """建立按區塊"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setSpacing(10)
        layout.setContentsMargins(0, 5, 0, 0)

        layout.addStretch()

        self.btn_ok = QPushButton("確定")
        self.btn_ok.setFixedWidth(80)
        self.btn_ok.setDefault(True)
        self.btn_ok.clicked.connect(self.on_ok_clicked)
        layout.addWidget(self.btn_ok)

        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setFixedWidth(80)
        self.btn_cancel.clicked.connect(self.reject)
        layout.addWidget(self.btn_cancel)

        self.btn_remove = QPushButton("移除循環(R)")
        self.btn_remove.setFixedWidth(100)
        self.btn_remove.clicked.connect(self.on_remove_clicked)
        layout.addWidget(self.btn_remove)

        return widget

    def connect_signals(self):
        """連接信號"""
        # 頻率選擇變更
        self.radio_daily.toggled.connect(self.on_frequency_changed)
        self.radio_weekly.toggled.connect(self.on_frequency_changed)
        self.radio_monthly.toggled.connect(self.on_frequency_changed)
        self.radio_yearly.toggled.connect(self.on_frequency_changed)

        # 時間變更時自動計算
        self.start_time_edit.timeChanged.connect(self.on_start_time_changed)
        self.end_time_edit.timeChanged.connect(self.on_end_time_changed)
        self.duration_combo.currentIndexChanged.connect(self.on_duration_changed)

        # 初始化時更新期間顯示
        self.update_duration_from_times()

    def on_frequency_changed(self):
        """頻率選擇變更時顯示對應的詳細設定"""
        self.daily_widget.setVisible(self.radio_daily.isChecked())
        self.weekly_widget.setVisible(self.radio_weekly.isChecked())
        self.monthly_widget.setVisible(self.radio_monthly.isChecked())
        self.yearly_widget.setVisible(self.radio_yearly.isChecked())

    def on_start_time_changed(self):
        """開始時間變更時，保持期間不變，更新結束時間"""
        # 暫時阻止 duration_combo 信號以避免遞迴
        self.duration_combo.blockSignals(True)
        
        duration = self.duration_combo.currentData()
        if duration is not None:
            start = self.start_time_edit.time()
            start_minutes = start.hour() * 60 + start.minute()
            end_minutes = start_minutes + duration

            end_hour = (end_minutes // 60) % 24
            end_minute = end_minutes % 60

            self.end_time_edit.blockSignals(True)
            self.end_time_edit.setTime(QTime(end_hour, end_minute))
            self.end_time_edit.blockSignals(False)
        
        self.duration_combo.blockSignals(False)

    def on_end_time_changed(self):
        """結束時間變更時，反推期間"""
        self.update_duration_from_times()

    def update_duration_from_times(self):
        """根據開始和結束時間計算期間"""
        start = self.start_time_edit.time()
        end = self.end_time_edit.time()

        start_minutes = start.hour() * 60 + start.minute()
        end_minutes = end.hour() * 60 + end.minute()

        if end_minutes < start_minutes:
            end_minutes += 24 * 60  # 跨日

        duration = end_minutes - start_minutes

        # 更新期間下拉選單而不觸發信號
        self.duration_combo.blockSignals(True)
        index = self.duration_combo.findData(duration)
        if index >= 0:
            self.duration_combo.setCurrentIndex(index)
        else:
            # 如果期間不在預設列表中，添加自訂期間
            self.duration_combo.addItem(f"{duration} 分", duration)
            self.duration_combo.setCurrentIndex(self.duration_combo.count() - 1)
        self.duration_combo.blockSignals(False)

    def on_duration_changed(self):
        """期間變更時更新結束時間"""
        duration = self.duration_combo.currentData()
        if duration is not None:
            start = self.start_time_edit.time()
            start_minutes = start.hour() * 60 + start.minute()
            end_minutes = start_minutes + duration

            end_hour = (end_minutes // 60) % 24
            end_minute = end_minutes % 60

            self.end_time_edit.blockSignals(True)
            self.end_time_edit.setTime(QTime(end_hour, end_minute))
            self.end_time_edit.blockSignals(False)

    def on_ok_clicked(self):
        """確定按點擊"""
        rrule_str = self.build_rrule()
        self.rrule_created.emit(rrule_str)
        self.accept()

    def on_remove_clicked(self):
        """移除循環按點擊"""
        self.rrule_created.emit("")
        self.accept()

    def build_rrule(self) -> str:
        """建立 RRULE 字串"""
        freq = ""
        byday = ""
        bymonthday = ""
        bymonth = ""
        bysetpos = ""
        interval = 1
        until = ""
        count = 0

        # 取得時間
        time = self.start_time_edit.time()
        hour = time.hour()
        minute = time.minute()

        # 開始日期
        start_date = self.start_date_edit.date()
        dtstart = f"{start_date.year()}{start_date.month():02d}{start_date.day():02d}T{hour:02d}{minute:02d}00"

        # 根據頻率設定
        if self.radio_daily.isChecked():
            freq = "DAILY"
            if self.daily_weekday_radio.isChecked():
                byday = "MO,TU,WE,TH,FR"
                interval = 1
            else:
                interval = self.daily_interval.value()

        elif self.radio_weekly.isChecked():
            freq = "WEEKLY"
            interval = self.weekly_interval.value()

            selected_days = []
            for day_code, checkbox in self.day_checkboxes.items():
                if checkbox.isChecked():
                    selected_days.append(day_code)

            if selected_days:
                byday = ",".join(selected_days)

        elif self.radio_monthly.isChecked():
            freq = "MONTHLY"

            if self.radio_monthly_day.isChecked():
                interval = self.monthly_interval.value()
                bymonthday = str(self.monthly_day.value())
            else:
                interval = self.monthly_week_interval.value()
                week_num = self.monthly_week_num.currentIndex() + 1
                if self.monthly_week_num.currentIndex() == 4:  # 最後一個
                    week_num = -1
                day_map = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"]
                byday = day_map[self.monthly_week_day.currentIndex()]
                bysetpos = str(week_num)

        elif self.radio_yearly.isChecked():
            freq = "YEARLY"
            interval = self.yearly_interval.value()

            if self.radio_yearly_date.isChecked():
                bymonth = str(self.yearly_month.currentIndex() + 1)
                bymonthday = str(self.yearly_day.value())
            else:
                bymonth = str(self.yearly_week_month.currentIndex() + 1)
                week_num = self.yearly_week_num.currentIndex() + 1
                if self.yearly_week_num.currentIndex() == 4:  # 最後一個
                    week_num = -1
                day_map = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"]
                byday = day_map[self.yearly_week_day.currentIndex()]
                bysetpos = str(week_num)

        # 結束條件
        if self.radio_end_never.isChecked():
            pass
        elif self.radio_end_after.isChecked():
            count = self.end_count.value()
        elif self.radio_end_by.isChecked():
            end_date = self.end_date_edit.date()
            until = (
                f"{end_date.year()}{end_date.month():02d}{end_date.day():02d}T235959"
            )

        # 組合 RRULE
        parts = [f"FREQ={freq}"]

        if interval > 1:
            parts.append(f"INTERVAL={interval}")

        if bymonth:
            parts.append(f"BYMONTH={bymonth}")

        if bymonthday:
            parts.append(f"BYMONTHDAY={bymonthday}")

        if byday:
            parts.append(f"BYDAY={byday}")

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
        """解析現有的 RRULE 字串"""
        try:
            import re

            parts = rrule_str.split(";")
            freq = ""
            interval = 1
            byday = ""
            bymonthday = ""
            bymonth = ""
            bysetpos = ""
            byhour = 9
            byminute = 0

            # 解析 DTSTART
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
                self.start_date_edit.setDate(QDate(year, month, day))
                self.start_time_edit.setTime(QTime(hour, minute))
                byhour = hour
                byminute = minute

            # 解析各個部分
            for part in parts:
                if part.startswith("FREQ="):
                    freq = part.split("=")[1]
                elif part.startswith("INTERVAL="):
                    interval = int(part.split("=")[1])
                elif part.startswith("BYDAY="):
                    byday = part.split("=")[1]
                elif part.startswith("BYMONTHDAY="):
                    bymonthday = part.split("=")[1]
                elif part.startswith("BYMONTH="):
                    bymonth = part.split("=")[1]
                elif part.startswith("BYSETPOS="):
                    bysetpos = part.split("=")[1]
                elif part.startswith("BYHOUR="):
                    byhour = int(part.split("=")[1])
                elif part.startswith("BYMINUTE="):
                    byminute = int(part.split("=")[1])
                elif part.startswith("COUNT="):
                    count = int(part.split("=")[1])
                    self.radio_end_after.setChecked(True)
                    self.end_count.setValue(count)
                elif part.startswith("UNTIL="):
                    until_str = part.split("=")[1][:8]
                    year, month, day = (
                        int(until_str[:4]),
                        int(until_str[4:6]),
                        int(until_str[6:8]),
                    )
                    self.radio_end_by.setChecked(True)
                    self.end_date_edit.setDate(QDate(year, month, day))

            # 設定時間
            self.start_time_edit.setTime(QTime(byhour, byminute))
            self.update_duration()

            # 設定頻率和詳細選項
            if freq == "DAILY":
                self.radio_daily.setChecked(True)
                if byday == "MO,TU,WE,TH,FR":
                    self.daily_weekday_radio.setChecked(True)
                else:
                    self.radio_daily_every.setChecked(True)
                    self.daily_interval.setValue(interval)

            elif freq == "WEEKLY":
                self.radio_weekly.setChecked(True)
                self.weekly_interval.setValue(interval)
                if byday:
                    for day_code in byday.split(","):
                        if day_code in self.day_checkboxes:
                            self.day_checkboxes[day_code].setChecked(True)

            elif freq == "MONTHLY":
                self.radio_monthly.setChecked(True)
                if bymonthday:
                    self.radio_monthly_day.setChecked(True)
                    self.monthly_interval.setValue(interval)
                    self.monthly_day.setValue(int(bymonthday))
                elif bysetpos and byday:
                    self.radio_monthly_week.setChecked(True)
                    self.monthly_week_interval.setValue(interval)
                    setpos = int(bysetpos)
                    if setpos == -1:
                        self.monthly_week_num.setCurrentIndex(4)
                    else:
                        self.monthly_week_num.setCurrentIndex(min(setpos - 1, 4))
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

            elif freq == "YEARLY":
                self.radio_yearly.setChecked(True)
                self.yearly_interval.setValue(interval)
                if bymonthday:
                    self.radio_yearly_date.setChecked(True)
                    if bymonth:
                        self.yearly_month.setCurrentIndex(int(bymonth) - 1)
                    self.yearly_day.setValue(int(bymonthday))
                elif bysetpos and byday and bymonth:
                    self.radio_yearly_week.setChecked(True)
                    if bymonth:
                        self.yearly_week_month.setCurrentIndex(int(bymonth) - 1)
                    setpos = int(bysetpos)
                    if setpos == -1:
                        self.yearly_week_num.setCurrentIndex(4)
                    else:
                        self.yearly_week_num.setCurrentIndex(min(setpos - 1, 4))
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
                        self.yearly_week_day.setCurrentIndex(day_map[byday])

        except Exception as e:
            print(f"解析現有 RRULE 失敗: {e}")

    def is_dark_mode(self) -> bool:
        """檢查是否使用暗色模式"""
        # 遍历父窗口链查找主题设置
        parent = self.parent()
        while parent:
            if hasattr(parent, "current_theme"):
                if parent.current_theme == "dark":
                    return True
                elif parent.current_theme == "system":
                    if hasattr(parent, "is_system_dark_mode"):
                        return parent.is_system_dark_mode()
                return False
            parent = parent.parent() if hasattr(parent, "parent") else None
        return False

    def apply_modern_style(self):
        """套用現代化樣式，支援主題切換"""
        is_dark = self.is_dark_mode()

        if is_dark:
            # 暗色主题樣式
            self.setStyleSheet("""
                QDialog {
                    background-color: #2b2b2b;
                }
                QGroupBox {
                    font-weight: bold;
                    border: 1px solid #3d3d3d;
                    border-radius: 6px;
                    margin-top: 12px;
                    padding-top: 12px;
                    background-color: #363636;
                    color: #cccccc;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top left;
                    padding: 0 8px;
                    color: #ffffff;
                }
                QPushButton {
                    background-color: #0e639c;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 6px 16px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #1177bb;
                }
                QPushButton:pressed {
                    background-color: #094771;
                }
                QPushButton:disabled {
                    background-color: #4a4a4a;
                    color: #808080;
                }
                QCheckBox {
                    spacing: 8px;
                    color: #cccccc;
                    outline: none;
                }
                QCheckBox::indicator {
                    width: 18px;
                    height: 18px;
                    border-radius: 3px;
                    border: 2px solid #606060;
                    background-color: #1e1e1e;
                }
                QCheckBox::indicator:checked {
                    background-color: #0e639c;
                    border-color: #0e639c;
                }
                QRadioButton {
                    spacing: 8px;
                    color: #cccccc;
                    outline: none;
                }
                QRadioButton::indicator {
                    width: 18px;
                    height: 18px;
                    border: 2px solid #606060;
                    border-radius: 9px;
                    background-color: #1e1e1e;
                }
                QRadioButton::indicator:checked {
                    background-color: #0e639c;
                    border-color: #0e639c;
                }
                QSpinBox, QComboBox, QDateEdit, QTimeEdit {
                    border: 1px solid #3d3d3d;
                    border-radius: 4px;
                    padding: 4px 8px;
                    background-color: #1e1e1e;
                    color: #cccccc;
                }
                QSpinBox:focus, QComboBox:focus, QDateEdit:focus, QTimeEdit:focus {
                    border: 2px solid #0e639c;
                }
                QLabel {
                    color: #cccccc;
                }
                QLabel#fieldLabel {
                    color: #ffffff;
                    font-weight: bold;
                }
                QFrame {
                    color: #3d3d3d;
                }
            """)
        else:
            # 亮色主题樣式
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
                QPushButton:disabled {
                    background-color: #cccccc;
                    color: #888888;
                }
                QCheckBox {
                    spacing: 8px;
                    color: #333;
                    outline: none;
                }
                QCheckBox::indicator {
                    width: 18px;
                    height: 18px;
                    border-radius: 3px;
                    border: 2px solid #a0a0a0;
                    background-color: white;
                }
                QCheckBox::indicator:checked {
                    background-color: #0078d4;
                    border-color: #0078d4;
                }
                QRadioButton {
                    spacing: 8px;
                    color: #333;
                    outline: none;
                }
                QRadioButton::indicator {
                    width: 18px;
                    height: 18px;
                    border: 2px solid #a0a0a0;
                    border-radius: 9px;
                    background-color: white;
                }
                QRadioButton::indicator:checked {
                    background-color: #0078d4;
                    border-color: #0078d4;
                }
                QSpinBox, QComboBox, QDateEdit, QTimeEdit {
                    border: 1px solid #d0d0d0;
                    border-radius: 4px;
                    padding: 4px 8px;
                    background-color: white;
                    color: #333;
                }
                QSpinBox:focus, QComboBox:focus, QDateEdit:focus, QTimeEdit:focus {
                    border: 2px solid #0078d4;
                }
                QLabel {
                    color: #333;
                }
                QLabel#fieldLabel {
                    color: #2c3e50;
                    font-weight: bold;
                }
            """)

    def get_rrule(self) -> str:
        """取得 RRULE 字串"""
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

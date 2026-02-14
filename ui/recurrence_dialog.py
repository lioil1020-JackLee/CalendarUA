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
    QGroupBox,
    QRadioButton,
    QButtonGroup,
    QGridLayout,
    QFrame,
    QMessageBox,
)
from PySide6.QtCore import Qt, QDate, QTime, Signal
from PySide6.QtGui import QFont, QIcon
import sys


class RecurrenceDialog(QDialog):
    """週期性設定對話框 - Outlook 風格"""

    rrule_created = Signal(str)

    def __init__(self, parent=None, current_rrule: str = ""):
        super().__init__(parent)
        self.current_rrule = current_rrule
        self.setWindowTitle("週期性約會")
        self.setWindowIcon(QIcon('lioil.ico'))
        self.setMinimumWidth(570)
        self.setMinimumHeight(480)
        self.setModal(True)

        self.setup_ui()
        self.apply_modern_style()
        self.connect_signals()

        # 初始化結束條件控制項狀態
        self.on_end_condition_changed(self.radio_end_never, True)

        # 初始化頻率選擇的顯示狀態
        self.on_frequency_changed()

        # 解析現有的 RRULE（如果有的話）
        if self.current_rrule:
            self.parse_existing_rrule()

        # 初始化時間同步：確保結束時間根據期間正確計算
        self.on_start_time_changed(None)

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
        self.start_time_combo = QComboBox()
        self.start_time_combo.setEditable(True)
        self.start_time_combo.setObjectName("startTimeCombo")
        self.populate_time_combo(self.start_time_combo)
        self.start_time_combo.setFixedWidth(120)
        layout.addWidget(self.start_time_combo, 0, 1)

        # 結束時間
        end_label = QLabel("結束(N):")
        end_label.setObjectName("fieldLabel")
        layout.addWidget(end_label, 1, 0)
        self.end_time_combo = QComboBox()
        self.end_time_combo.setEditable(True)
        self.end_time_combo.setObjectName("endTimeCombo")
        self.populate_time_combo(self.end_time_combo)
        self.end_time_combo.setFixedWidth(120)
        layout.addWidget(self.end_time_combo, 1, 1)

        # 期間
        duration_label = QLabel("期間(U):")
        duration_label.setObjectName("fieldLabel")
        layout.addWidget(duration_label, 2, 0)
        self.duration_combo = QComboBox()
        # 允許自訂輸入（可編輯），但不要自動插入新項目
        self.duration_combo.setEditable(True)
        self.duration_combo.setInsertPolicy(QComboBox.NoInsert)
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
            ("5 分", 5),
            ("10 分", 10),
            ("15 分", 15),
            ("30 分", 30),
            ("1 時", 60),
            ("2 時", 120),
            ("3 時", 180),
            ("4 時", 240),
            ("5 時", 300),
            ("6 時", 360),
            ("7 時", 420),
            ("8 時", 480),
            ("9 時", 540),
            ("10 時", 600),
            ("11 時", 660),
            ("0.5 日", 720),
            ("18 時", 1080),
            ("1 日", 1440),
            ("2 日", 2880),
            ("3 日", 4320),
            ("4 日", 5760),
            ("1 週", 10080),
            ("2 週", 20160),
        ]
        for text, minutes in durations:
            self.duration_combo.addItem(text, minutes)
        self.duration_combo.setCurrentIndex(1)  # 預設改為 5 分

    def populate_time_combo(self, combo: QComboBox, default_time: QTime = None):
        """填充時間下拉選單"""
        combo.clear()
        # 生成從上午 12:00:00 (00:00:00) 到下午 11:30:00 (23:30:00) 的選項，每30分鐘一個
        for hour in range(24):
            for minute in range(0, 60, 30):  # 每30分鐘
                time_str = "上午" if hour < 12 else "下午"
                display_hour = hour % 12
                if display_hour == 0:
                    display_hour = 12
                time_text = f"{time_str} {display_hour:02d}:{minute:02d}:00"
                combo.addItem(time_text, QTime(hour, minute, 0))

        # 如果提供了預設時間，使用它；否則設置為離目前系統時間最接近的
        if default_time is not None:
            # 首先嘗試找到完全匹配的時間選項
            exact_match_i = -1
            for i in range(combo.count()):
                time = combo.itemData(i)
                if time.hour() == default_time.hour() and time.minute() == default_time.minute():
                    exact_match_i = i
                    break
            
            if exact_match_i >= 0:
                # 找到完全匹配，設置為該項目
                combo.setCurrentIndex(exact_match_i)
            else:
                # 沒有完全匹配，格式化時間並設置為可編輯文字
                time_str = "上午" if default_time.hour() < 12 else "下午"
                display_hour = default_time.hour() % 12
                if display_hour == 0:
                    display_hour = 12
                time_text = f"{time_str} {display_hour:02d}:{default_time.minute():02d}:00"
                combo.setCurrentText(time_text)
        elif not self.current_rrule:
            # 只有在沒有現有 RRULE 時才設置預設值為離目前系統時間最接近的，但要大於目前系統時間
            current_time = QTime.currentTime()
            best_i = -1
            best_time = None
            for i in range(combo.count()):
                time = combo.itemData(i)
                if time > current_time and (best_time is None or time < best_time):
                    best_time = time
                    best_i = i
            if best_i >= 0:
                combo.setCurrentIndex(best_i)
            else:
                # 如果沒有找到合適的時間，設置為第一個項目
                combo.setCurrentIndex(0)
        # 在編輯模式下，如果沒有提供 default_time，不設置任何預設值

    def connect_signals(self):
        """連接信號"""
        # 連接時間和期間的互動
        self.start_time_combo.activated.connect(self.on_start_time_changed)
        self.start_time_combo.editTextChanged.connect(self.on_start_time_changed)
        self.end_time_combo.activated.connect(self.on_end_time_changed)
        self.end_time_combo.editTextChanged.connect(self.on_end_time_changed)
        self.duration_combo.currentIndexChanged.connect(self.on_duration_changed)
        # 支援使用者直接在可編輯的 combo 中輸入自訂期間
        if self.duration_combo.isEditable() and self.duration_combo.lineEdit() is not None:
            self.duration_combo.lineEdit().editingFinished.connect(self.on_duration_text_edited)
            self.duration_combo.lineEdit().textChanged.connect(self.on_duration_text_changed)

        # 頻率選擇變更
        self.radio_daily.toggled.connect(self.on_frequency_changed)
        self.radio_weekly.toggled.connect(self.on_frequency_changed)
        self.radio_monthly.toggled.connect(self.on_frequency_changed)
        self.radio_yearly.toggled.connect(self.on_frequency_changed)

        # 結束條件變更
        self.end_button_group.buttonToggled.connect(self.on_end_condition_changed)

    def parse_existing_rrule(self):
        """解析現有的 RRULE 字串並設置控制項"""
        if not self.current_rrule:
            return

        try:
            # 解析 RRULE 參數
            params = {}
            parts = self.current_rrule.split(";")
            for part in parts:
                if "=" in part:
                    key, value = part.split("=", 1)
                    params[key] = value

            # 設置頻率
            freq = params.get("FREQ", "DAILY")
            if freq == "DAILY":
                self.radio_daily.setChecked(True)
            elif freq == "WEEKLY":
                self.radio_weekly.setChecked(True)
            elif freq == "MONTHLY":
                self.radio_monthly.setChecked(True)
            elif freq == "YEARLY":
                self.radio_yearly.setChecked(True)

            # 設置間隔
            interval = int(params.get("INTERVAL", "1"))
            if freq == "DAILY":
                self.daily_interval.setValue(interval)
            elif freq == "WEEKLY":
                self.weekly_interval.setValue(interval)
            elif freq == "MONTHLY":
                self.monthly_interval.setValue(interval)
            elif freq == "YEARLY":
                self.yearly_interval.setValue(interval)

            # 設置開始時間
            byhour = params.get("BYHOUR")
            byminute = params.get("BYMINUTE", "0")
            if byhour:
                hour = int(byhour)
                minute = int(byminute)
                start_time = QTime(hour, minute, 0)
            else:
                # 如果沒有 BYHOUR，使用預設時間 (上午9:00)
                start_time = QTime(9, 0, 0)
            # 重新填充開始時間下拉選單，使用現有時間作為預設值
            self.populate_time_combo(self.start_time_combo, start_time)

            # 設置結束條件
            if "COUNT" in params:
                self.radio_end_after.setChecked(True)
                self.end_count.setValue(int(params["COUNT"]))
            elif "UNTIL" in params:
                self.radio_end_by.setChecked(True)
                until_str = params["UNTIL"]
                try:
                    # 解析 UNTIL 日期 (格式: YYYYMMDD)
                    year = int(until_str[:4])
                    month = int(until_str[4:6])
                    day = int(until_str[6:8])
                    self.end_date_edit.setDate(QDate(year, month, day))
                except (ValueError, IndexError):
                    pass  # 使用預設值
            else:
                self.radio_end_never.setChecked(True)

            # 設置頻率特定的參數
            self._parse_frequency_specific_params(params)

        except Exception as e:
            print(f"解析 RRULE 失敗: {e}")
            # 解析失敗時使用預設值

    def _parse_frequency_specific_params(self, params):
        """解析頻率特定的參數"""
        freq = params.get("FREQ", "DAILY")

        if freq == "WEEKLY":
            # 解析星期幾
            byday = params.get("BYDAY", "")
            if byday:
                days = byday.split(",")
                # 清除所有勾選
                for checkbox in self.day_checkboxes.values():
                    checkbox.setChecked(False)
                # 設置對應的勾選
                for day in days:
                    if day in self.day_checkboxes:
                        self.day_checkboxes[day].setChecked(True)

        elif freq == "MONTHLY":
            bymonthday = params.get("BYMONTHDAY")
            byday = params.get("BYDAY")
            bysetpos = params.get("BYSETPOS")

            if bymonthday:
                # 每月第幾天
                self.radio_monthly_day.setChecked(True)
                self.monthly_day.setValue(int(bymonthday))
            elif byday and bysetpos:
                # 每月第幾個星期幾
                self.radio_monthly_week.setChecked(True)
                self.monthly_week_interval.setValue(int(bysetpos))
                # 設置星期幾
                day_map = {
                    "MO": "星期一", "TU": "星期二", "WE": "星期三", "TH": "星期四",
                    "FR": "星期五", "SA": "星期六", "SU": "星期日"
                }
                if byday in day_map:
                    day_text = day_map[byday]
                    for i in range(self.monthly_week_day.count()):
                        if day_text in self.monthly_week_day.itemText(i):
                            self.monthly_week_day.setCurrentIndex(i)
                            break

        elif freq == "YEARLY":
            bymonth = params.get("BYMONTH")
            bymonthday = params.get("BYMONTHDAY")
            byday = params.get("BYDAY")
            bysetpos = params.get("BYSETPOS")

            if bymonth and bymonthday:
                # 每年第幾月第幾天
                self.radio_yearly_date.setChecked(True)
                self.yearly_month.setCurrentIndex(int(bymonth) - 1)  # 月份從0開始
                self.yearly_day.setValue(int(bymonthday))
            elif bymonth and byday and bysetpos:
                # 每年第幾月第幾個星期幾
                self.radio_yearly_week.setChecked(True)
                self.yearly_week_month.setCurrentIndex(int(bymonth) - 1)
                self.yearly_week_num.setValue(int(bysetpos))
                # 設置星期幾
                day_map = {
                    "MO": "星期一", "TU": "星期二", "WE": "星期三", "TH": "星期四",
                    "FR": "星期五", "SA": "星期六", "SU": "星期日"
                }
                if byday in day_map:
                    day_text = day_map[byday]
                    for i in range(self.yearly_week_day.count()):
                        if day_text in self.yearly_week_day.itemText(i):
                            self.yearly_week_day.setCurrentIndex(i)
                            break

    def on_start_time_changed(self, value):
        """開始時間改變時更新結束時間"""
        if not hasattr(self, "_updating_times") or not self._updating_times:
            self._updating_times = True
            try:
                start_time = self.get_time_from_combo(self.start_time_combo)
                duration_minutes = self.get_duration_minutes()
                if start_time and duration_minutes is not None:
                    # 重新格式化開始時間顯示，確保上午/下午格式正確
                    self.set_combo_to_time(self.start_time_combo, start_time)
                    end_time = start_time.addSecs(duration_minutes * 60)
                    self.set_combo_to_time(self.end_time_combo, end_time)
            finally:
                self._updating_times = False

    def on_end_time_changed(self, value):
        """結束時間改變時更新期間"""
        if not hasattr(self, "_updating_times") or not self._updating_times:
            self._updating_times = True
            try:
                start_time = self.get_time_from_combo(self.start_time_combo)
                end_time = self.get_time_from_combo(self.end_time_combo)
                if start_time and end_time:
                    # 重新格式化結束時間顯示，確保上午/下午格式正確
                    self.set_combo_to_time(self.end_time_combo, end_time)
                    duration_seconds = start_time.secsTo(end_time)
                    if duration_seconds < 0:
                        duration_seconds += 24 * 3600  # 跨日
                    duration_minutes = duration_seconds // 60
                    self.set_duration_to_minutes(duration_minutes)
            finally:
                self._updating_times = False

    def get_time_from_combo(self, combo: QComboBox) -> QTime:
        """從 ComboBox 獲取時間"""
        text = combo.currentText()
        if ":" in text:
            # 自訂時間格式 或 標準格式
            parts = text.split()
            if len(parts) >= 2 and (parts[0] == "上午" or parts[0] == "下午"):
                # 標準格式: "上午 12:00:00"
                period = parts[0]
                time_part = parts[-1]
                time_parts = time_part.split(":")
                if len(time_parts) == 3:
                    # 包含秒數: "12:00:30"
                    hour, minute, second = map(int, time_parts)
                else:
                    # 不包含秒數: "12:00"
                    hour, minute = map(int, time_parts)
                    second = 0
                if period == "上午":
                    if hour == 12:
                        hour = 0
                elif period == "下午":
                    if hour != 12:
                        hour += 12
                return QTime(hour, minute, second)
            else:
                # 自訂格式: "12:00:00" 或 "12:00"
                time_part = text.split()[-1]
                try:
                    time_parts = time_part.split(":")
                    if len(time_parts) == 3:
                        hour, minute, second = map(int, time_parts)
                    else:
                        hour, minute = map(int, time_parts)
                        second = 0
                    return QTime(hour, minute, second)
                except:
                    pass
        # 從項目數據獲取
        return combo.currentData()

    def set_combo_to_time(self, combo: QComboBox, time: QTime):
        """設置 ComboBox 到指定時間"""
        combo.blockSignals(True)
        try:
            # 首先嘗試找到匹配的項目
            for i in range(combo.count()):
                item_time = combo.itemData(i)
                if item_time == time:
                    combo.setCurrentIndex(i)
                    return
            
            # 如果沒有找到，設置為自訂文本，包含上午/下午和秒數
            hour = time.hour()
            minute = time.minute()
            second = time.second()
            time_str = "上午" if hour < 12 else "下午"
            display_hour = hour % 12
            if display_hour == 0:
                display_hour = 12
            time_text = f"{time_str} {display_hour:02d}:{minute:02d}:{second:02d}"
            combo.setCurrentText(time_text)
        finally:
            combo.blockSignals(False)

    def set_duration_to_minutes(self, minutes: int):
        """設置期間到最接近的分鐘數"""
        self.duration_combo.blockSignals(True)
        try:
            best_index = 0
            min_diff = abs(self.duration_combo.itemData(0) - minutes)

            for i in range(1, self.duration_combo.count()):
                diff = abs(self.duration_combo.itemData(i) - minutes)
                if diff < min_diff:
                    min_diff = diff
                    best_index = i

            self.duration_combo.setCurrentIndex(best_index)
        finally:
            self.duration_combo.blockSignals(False)

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
        self.radio_daily.setChecked(True)  # 預設改為每天
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
        layout.setSpacing(8)  # 增加間距
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

        # 為"天"標籤設置物件名稱與最小寬度，確保套用 fieldLabel 樣式並可見
        day_label = QLabel("天")
        day_label.setObjectName("fieldLabel")
        day_label.setMinimumWidth(20)
        layout.addWidget(day_label)

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
        self.monthly_week_num.setFixedWidth(100)
        week_layout.addWidget(self.monthly_week_num)

        self.monthly_week_day = QComboBox()
        self.monthly_week_day.addItems(
            ["週一到週五", "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"]
        )
        self.monthly_week_day.setFixedWidth(100)
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
        self.yearly_week_num.setFixedWidth(100)
        week_layout.addWidget(self.yearly_week_num)

        self.yearly_week_day = QComboBox()
        self.yearly_week_day.addItems(
            ["週一到週五", "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"]
        )
        self.yearly_week_day.setFixedWidth(100)
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
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setFixedWidth(150)
        layout.addWidget(self.start_date_edit, 0, 1)

        # 結束選項
        self.end_button_group = QButtonGroup(self)

        # 結束於日期
        self.radio_end_by = QRadioButton("結束於(B):")
        self.end_button_group.addButton(self.radio_end_by)
        layout.addWidget(self.radio_end_by, 0, 2)

        self.end_date_edit = QDateEdit()
        self.end_date_edit.setDisplayFormat("yyyy/M/d (ddd)")
        self.end_date_edit.setDate(QDate.currentDate().addMonths(3))
        self.end_date_edit.setCalendarPopup(True)
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
        self.radio_end_never.setChecked(True)  # 預設改為沒有結束日期
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

        return widget

    def on_frequency_changed(self):
        """頻率選擇變更時顯示對應的詳細設定"""
        self.daily_widget.setVisible(self.radio_daily.isChecked())
        self.weekly_widget.setVisible(self.radio_weekly.isChecked())
        self.monthly_widget.setVisible(self.radio_monthly.isChecked())
        self.yearly_widget.setVisible(self.radio_yearly.isChecked())

    def on_end_condition_changed(self, button, checked):
        """結束條件變更時啟用/禁用相關控制項"""
        if not checked:
            return

        # 根據選擇的結束條件啟用/禁用控制項
        if button == self.radio_end_never:
            # 沒有結束日期：禁用所有結束條件控制項
            self.end_date_edit.setEnabled(False)
            self.end_count.setEnabled(False)
        elif button == self.radio_end_by:
            # 結束於日期：只啟用日期選擇器
            self.end_date_edit.setEnabled(True)
            self.end_count.setEnabled(False)
        elif button == self.radio_end_after:
            # 重複次數：只啟用次數輸入框
            self.end_date_edit.setEnabled(False)
            self.end_count.setEnabled(True)

    def on_duration_changed(self, index):
        """期間變更時更新結束時間"""
        if not hasattr(self, "_updating_times") or not self._updating_times:
            self._updating_times = True
            try:
                start_time = self.get_time_from_combo(self.start_time_combo)
                duration_minutes = self.get_duration_minutes()
                # 選取內建項目時，取消自訂旗標
                if self.duration_combo.currentIndex() >= 0:
                    self._using_custom_duration = False

                if start_time and duration_minutes is not None:
                    end_time = start_time.addSecs(duration_minutes * 60)
                    self.set_combo_to_time(self.end_time_combo, end_time)
            finally:
                self._updating_times = False

    def on_duration_text_changed(self, text: str):
        """在使用者輸入期間文字時，提供即時的輸入驗證（不立即套用）"""
        # 目前不強制更新結束時間，等 editingFinished 再處理
        return

    def on_duration_text_edited(self):
        """使用者在可編輯的 combo 完成輸入後，解析並套用期間"""
        text = self.duration_combo.currentText()
        minutes = self.parse_duration_text(text)
        if minutes is None:
            return
        # 如果輸入可以解析，設定為自訂期間並更新結束時間
        self.set_custom_duration(minutes)
        # 觸發與選擇改變相同的行為
        self.on_duration_changed(self.duration_combo.currentIndex())

    def parse_duration_text(self, text: str):
        """解析使用者輸入的期間文字，回傳分鐘數或 None。支援簡單的單位：分/時/日 或純數字（視為分鐘）。"""
        if not text:
            return None
        s = text.strip()
        try:
            # 純數字視為分鐘
            if s.isdigit():
                return int(s)
            # 結尾包含單位
            if s.endswith('分'):
                num = s[:-1].strip()
                return int(float(num))
            if s.endswith('時'):
                num = s[:-1].strip()
                return int(float(num) * 60)
            if s.endswith('日'):
                num = s[:-1].strip()
                return int(float(num) * 1440)
            # 允許 '1.5 小時' 或 '1.5時' 之類的點號
            for unit, factor in [('分', 1), ('時', 60), ('日', 1440)]:
                if unit in s:
                    try:
                        num = float(s.replace(unit, '').strip())
                        return int(num * factor)
                    except:
                        return None
        except Exception:
            return None
        return None

    def get_duration_minutes(self):
        """取得目前期間的分鐘數：優先取選單項目的 data，否則取自訂儲存值。"""
        # 先檢查目前文字是否與選單中某個項目完全相符
        text = self.duration_combo.currentText()
        for i in range(self.duration_combo.count()):
            if self.duration_combo.itemText(i) == text:
                data = self.duration_combo.itemData(i)
                if isinstance(data, int):
                    return data

        # 若未匹配任何項目，嘗試解析目前文字為分鐘數
        minutes = self.parse_duration_text(text)
        if minutes is not None:
            return minutes

        # 如果都失敗，回傳 None
        return None

    def set_custom_duration(self, minutes: int):
        """把自訂分鐘設為 combo 的顯示文字（不新增到選單項目）並記錄。"""
        self._custom_duration_minutes = minutes
        # 顯示用文字（以分為單位）
        self.duration_combo.setCurrentText(f"{minutes} 分")

    def on_ok_clicked(self):
        """確定按點擊"""
        rrule_str = self.build_rrule()
        self.rrule_created.emit(rrule_str)
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
        time = self.get_time_from_combo(self.start_time_combo)
        if not time:
            time = QTime(9, 0)  # 預設值
        hour = time.hour()
        minute = time.minute()

        # 開始日期
        start_date = self.start_date_edit.date()
        dtstart = f"{start_date.year()}{start_date.month():02d}{start_date.day():02d}T{hour:02d}{minute:02d}00"

        # 期間
        duration_minutes = self.get_duration_minutes() or 30
        duration_str = f"DURATION=PT{duration_minutes}M"

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
                
                day_index = self.monthly_week_day.currentIndex()
                if day_index == 0:  # 週一到週五
                    byday = "MO,TU,WE,TH,FR"
                else:
                    day_map = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"]
                    byday = day_map[day_index - 1]  # 減1因為第一個選項是週一到週五
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
                
                day_index = self.yearly_week_day.currentIndex()
                if day_index == 0:  # 週一到週五
                    byday = "MO,TU,WE,TH,FR"
                else:
                    day_map = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"]
                    byday = day_map[day_index - 1]  # 減1因為第一個選項是週一到週五
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
        parts.append(duration_str)

        return ";".join(parts)

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
                QComboBox QListView::item {
                    background-color: #1e1e1e;
                    color: #cccccc;
                }
                QComboBox QListView::item:selected {
                    background-color: #094771;
                    color: white;
                }
                QComboBox#startTimeCombo, QComboBox#endTimeCombo {
                    color: white;
                }
                QComboBox#startTimeCombo QListView::item, QComboBox#endTimeCombo QListView::item {
                    color: white;
                }
                QCalendarWidget QWidget {
                    background-color: #2b2b2b;
                    color: #cccccc;
                }
                QCalendarWidget QAbstractItemView:enabled {
                    background-color: #363636;
                    color: #cccccc;
                    selection-background-color: #0e639c;
                    selection-color: white;
                }
                QCalendarWidget QAbstractItemView:disabled {
                    color: #666666;
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
                QComboBox::item {
                    background-color: white;
                    color: #333;
                }
                QComboBox::item:selected {
                    background-color: #0078d4;
                    color: white;
                }
                QCalendarWidget QWidget {
                    background-color: #f5f5f5;
                    color: #333;
                }
                QCalendarWidget QAbstractItemView:enabled {
                    background-color: white;
                    color: #333;
                    selection-background-color: #0078d4;
                    selection-color: white;
                }
                QCalendarWidget QAbstractItemView:disabled {
                    color: #cccccc;
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

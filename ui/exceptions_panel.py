"""
Exceptions Panel - 例外記錄管理面板
用於查看和管理 schedule exceptions (occurrence overrides and cancellations)
"""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QDialog, QFormLayout, QComboBox, QDateEdit, QLineEdit,
    QDateTimeEdit, QGroupBox, QRadioButton, QButtonGroup, QTabBar,
    QStackedWidget, QToolButton, QMenu, QCalendarWidget, QWidgetAction, QStyle
)
from PySide6.QtCore import Qt, Signal, QDate, QDateTime
from PySide6.QtGui import QColor, QBrush, QIcon, QPainter, QPen, QPixmap
from datetime import datetime, date, time, timedelta
from typing import List, Dict, Any, Optional

from core.schedule_resolver import ResolvedOccurrence
from ui.schedule_canvas import DayViewWidget, WeekViewWidget
from ui.month_grid import MonthViewWidget


class ExceptionEditDialog(QDialog):
    """例外記錄編輯對話框"""
    
    def __init__(self, parent=None, schedules: List[Dict[str, Any]] = None, 
                 initial_data: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.schedules = schedules or []
        self.initial_data = initial_data or {}
        self.setWindowTitle("編輯例外記錄 - Edit Exception")
        self.resize(550, 500)
        self._init_ui()
        self._load_data()
    
    def _init_ui(self):
        layout = QVBoxLayout(self)
        
        # Schedule 與日期選擇
        group_basic = QGroupBox("基本設定 (Basic Settings)")
        form_basic = QFormLayout()
        
        # Schedule 下拉選單
        self.combo_schedule = QComboBox()
        for sch in self.schedules:
            schedule_id = sch.get("id")
            task_name = sch.get("task_name", "未命名")
            rrule = sch.get("rrule_str", "")
            display = f"[{schedule_id}] {task_name} ({rrule[:50]}...)" if len(rrule) > 50 else f"[{schedule_id}] {task_name} ({rrule})"
            self.combo_schedule.addItem(display, schedule_id)
        form_basic.addRow("排程系列 (Schedule):", self.combo_schedule)
        
        # Occurrence 日期
        self.date_occurrence = QDateEdit()
        self.date_occurrence.setCalendarPopup(True)
        self.date_occurrence.setDate(QDate.currentDate())
        self.date_occurrence.setDisplayFormat("yyyy-MM-dd")
        form_basic.addRow("發生日期 (Occurrence Date):", self.date_occurrence)
        
        group_basic.setLayout(form_basic)
        layout.addWidget(group_basic)
        
        # Action 選擇
        group_action = QGroupBox("操作類型 (Action)")
        action_layout = QVBoxLayout()
        
        self.action_group = QButtonGroup(self)
        self.radio_cancel = QRadioButton("取消此次 (Cancel this occurrence)")
        self.radio_override = QRadioButton("覆寫此次 (Override this occurrence)")
        self.action_group.addButton(self.radio_cancel, 0)
        self.action_group.addButton(self.radio_override, 1)
        self.radio_override.setChecked(True)
        
        action_layout.addWidget(self.radio_cancel)
        action_layout.addWidget(self.radio_override)
        group_action.setLayout(action_layout)
        layout.addWidget(group_action)
        
        # Override 設定（僅當選 override 時啟用）
        self.group_override = QGroupBox("覆寫設定 (Override Settings)")
        form_override = QFormLayout()
        
        self.edit_override_title = QLineEdit()
        self.edit_override_title.setPlaceholderText("留空則使用原排程標題")
        form_override.addRow("標題 (Subject):", self.edit_override_title)
        
        self.edit_override_value = QLineEdit()
        self.edit_override_value.setPlaceholderText("留空則使用原排程數值")
        form_override.addRow("目標數值 (Target Value):", self.edit_override_value)
        
        self.datetime_start = QDateTimeEdit()
        self.datetime_start.setCalendarPopup(True)
        self.datetime_start.setDisplayFormat("yyyy-MM-dd HH:mm")
        self.datetime_start.setDateTime(QDateTime.currentDateTime())
        form_override.addRow("開始時間 (Start):", self.datetime_start)
        
        self.datetime_end = QDateTimeEdit()
        self.datetime_end.setCalendarPopup(True)
        self.datetime_end.setDisplayFormat("yyyy-MM-dd HH:mm")
        self.datetime_end.setDateTime(QDateTime.currentDateTime().addSecs(3600))
        form_override.addRow("結束時間 (End):", self.datetime_end)
        
        self.group_override.setLayout(form_override)
        layout.addWidget(self.group_override)
        
        # 按鈕
        button_layout = QHBoxLayout()
        self.btn_ok = QPushButton("確定 (OK)")
        self.btn_cancel = QPushButton("取消 (Cancel)")
        
        button_layout.addStretch()
        button_layout.addWidget(self.btn_ok)
        button_layout.addWidget(self.btn_cancel)
        
        layout.addLayout(button_layout)
        
        # 訊號
        self.radio_cancel.toggled.connect(self._on_action_changed)
        self.radio_override.toggled.connect(self._on_action_changed)
        self.btn_ok.clicked.connect(self._on_ok)
        self.btn_cancel.clicked.connect(self.reject)
        
        self._on_action_changed()
    
    def _on_action_changed(self):
        """Action 變更時啟用/禁用 override 設定"""
        is_override = self.radio_override.isChecked()
        self.group_override.setEnabled(is_override)
    
    def _load_data(self):
        """載入初始資料（編輯模式）"""
        if not self.initial_data:
            return
        
        # 選擇 schedule
        schedule_id = self.initial_data.get("schedule_id")
        for i in range(self.combo_schedule.count()):
            if self.combo_schedule.itemData(i) == schedule_id:
                self.combo_schedule.setCurrentIndex(i)
                break
        
        # 日期
        occurrence_date_str = self.initial_data.get("occurrence_date")
        if occurrence_date_str:
            try:
                dt = datetime.strptime(occurrence_date_str, "%Y-%m-%d")
                self.date_occurrence.setDate(QDate(dt.year, dt.month, dt.day))
            except:
                pass
        
        # Action
        action = self.initial_data.get("action", "override")
        if action == "cancel":
            self.radio_cancel.setChecked(True)
        else:
            self.radio_override.setChecked(True)
        
        # Override 值
        self.edit_override_title.setText(self.initial_data.get("override_task_name") or "")
        self.edit_override_value.setText(self.initial_data.get("override_target_value") or "")
        
        override_start_str = self.initial_data.get("override_start")
        if override_start_str:
            try:
                dt = datetime.fromisoformat(override_start_str)
                self.datetime_start.setDateTime(QDateTime(dt.year, dt.month, dt.day, dt.hour, dt.minute))
            except:
                pass
        
        override_end_str = self.initial_data.get("override_end")
        if override_end_str:
            try:
                dt = datetime.fromisoformat(override_end_str)
                self.datetime_end.setDateTime(QDateTime(dt.year, dt.month, dt.day, dt.hour, dt.minute))
            except:
                pass
    
    def _on_ok(self):
        """驗證並確定"""
        # 驗證 schedule 已選擇
        if self.combo_schedule.currentIndex() < 0:
            QMessageBox.warning(self, "驗證錯誤", "請選擇排程系列")
            return
        
        # 如果是 override，驗證時間
        if self.radio_override.isChecked():
            start_dt = self.datetime_start.dateTime().toPython()
            end_dt = self.datetime_end.dateTime().toPython()
            if end_dt <= start_dt:
                QMessageBox.warning(self, "驗證錯誤", "結束時間必須大於開始時間")
                return
        
        self.accept()
    
    def get_data(self) -> Dict[str, Any]:
        """取得編輯結果"""
        action = "cancel" if self.radio_cancel.isChecked() else "override"
        occurrence_date = self.date_occurrence.date().toPython()
        
        result = {
            "exception_id": self.initial_data.get("id"),  # None for new
            "schedule_id": self.combo_schedule.currentData(),
            "occurrence_date": occurrence_date,
            "action": action,
        }
        
        if action == "override":
            start_dt = self.datetime_start.dateTime().toPython()
            end_dt = self.datetime_end.dateTime().toPython()
            
            result.update({
                "override_start": start_dt,
                "override_end": end_dt,
                "override_task_name": self.edit_override_title.text().strip() or None,
                "override_target_value": self.edit_override_value.text().strip() or None,
            })
        
        return result


class ExceptionsPanel(QWidget):
    """Exceptions 面板 - 管理例外記錄"""
    
    exception_changed = Signal()  # 通知主視窗重新載入
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db_manager = None
        self.schedules = []
        self.exceptions = []
        self.filtered_exceptions = []
        self.current_view_mode = "day"
        self.reference_date = QDate.currentDate()
        self._init_ui()
    
    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)

        # Day / Week / Month 與日期控制同列（右上角日期）
        view_row = QHBoxLayout()
        view_row.setContentsMargins(4, 2, 4, 0)
        view_row.setSpacing(4)
        self.view_tabs = QTabBar()
        self.view_tabs.addTab("Day")
        self.view_tabs.addTab("Week")
        self.view_tabs.addTab("Month")
        self.view_tabs.setCurrentIndex(0)
        self.view_tabs.currentChanged.connect(self._on_view_changed)
        self.label_month_title = QLabel("")
        self.label_month_title.setStyleSheet("font-family: 'Times New Roman'; font-size: 18px; font-weight: 700; padding-left: 2px;")
        view_row.addWidget(self.label_month_title)
        view_row.addStretch()
        view_row.addWidget(self.view_tabs)
        view_row.addStretch()

        self.btn_prev_period = QPushButton("<")
        self.btn_prev_period.setFixedSize(22, 22)
        self.btn_next_period = QPushButton(">")
        self.btn_next_period.setFixedSize(22, 22)
        self.btn_current_period = QPushButton("")
        self.btn_current_period.setMinimumWidth(140)
        self.btn_current_period.setFixedHeight(24)
        self.btn_calendar_popup = QToolButton()
        self.btn_calendar_popup.setIcon(self._create_calendar_icon())
        self.btn_calendar_popup.setToolTip("開啟日曆")
        self.btn_calendar_popup.setFixedSize(24, 22)

        view_row.addWidget(self.btn_current_period)
        view_row.addWidget(self.btn_prev_period)
        view_row.addWidget(self.btn_calendar_popup)
        view_row.addWidget(self.btn_next_period)
        layout.addLayout(view_row)

        # 日曆視圖區
        self.calendar_stack = QStackedWidget()
        self.day_view = DayViewWidget()
        self.week_view = WeekViewWidget()
        self.month_view = MonthViewWidget()
        self.calendar_stack.addWidget(self.day_view)
        self.calendar_stack.addWidget(self.week_view)
        self.calendar_stack.addWidget(self.month_view)
        layout.addWidget(self.calendar_stack)

        # 表格
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "ID", "排程名稱", "發生日期", "操作", 
            "覆寫標題", "覆寫數值", "覆寫時間", "建立時間"
        ])
        
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.hide()
        
        layout.addWidget(self.table)
        
        # 底部說明
        info_label = QLabel(
            "提示：此面板管理所有排程的例外記錄。"
            "「取消」會隱藏該次 occurrence；「覆寫」會替換時間/標題/數值。"
            "雙擊列可編輯。"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: gray; font-size: 10pt; padding: 5px;")
        info_label.hide()
        layout.addWidget(info_label)
        
        # 訊號
        self.table.doubleClicked.connect(self._edit_selected_exception)
        self.table.itemSelectionChanged.connect(self._update_button_states)
        self.btn_prev_period.clicked.connect(self._go_previous_period)
        self.btn_next_period.clicked.connect(self._go_next_period)
        self.btn_calendar_popup.clicked.connect(self._show_calendar_popup)
        self.btn_current_period.clicked.connect(self._go_current_period)

        self._update_range_label()
        self._apply_header_style()

    def _apply_header_style(self):
        self.view_tabs.setStyleSheet(
            "QTabBar::tab {"
            "background: transparent;"
            "border: none;"
            "min-width: 62px;"
            "padding: 1px 8px;"
            "font-family: 'Times New Roman';"
            "font-size: 18px;"
            "font-weight: 600;"
            "}"
            "QTabBar::tab:selected {"
            "border-bottom: 2px solid #d0d0d0;"
            "}"
        )
        button_style = "padding: 0 6px;"
        self.btn_current_period.setStyleSheet(
            "padding: 0 8px;"
            "font-family: 'Times New Roman';"
            "font-size: 14px;"
            "font-weight: 700;"
        )
        self.btn_prev_period.setStyleSheet(button_style)
        self.btn_next_period.setStyleSheet(button_style)
        self.btn_calendar_popup.setStyleSheet(button_style)

    def _create_calendar_icon(self) -> QIcon:
        pixmap = QPixmap(16, 16)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing, False)

        painter.setPen(QPen(QColor("#8a8f98"), 1))
        painter.setBrush(QColor("#f4f6f8"))
        painter.drawRect(1, 2, 14, 13)

        painter.setBrush(QColor("#2f73d9"))
        painter.drawRect(1, 2, 14, 4)

        painter.setPen(QPen(QColor("#d7e6ff"), 1))
        painter.drawLine(4, 3, 4, 5)
        painter.drawLine(11, 3, 11, 5)

        painter.setPen(QPen(QColor("#6d7580"), 1))
        painter.drawLine(4, 8, 12, 8)
        painter.drawLine(4, 11, 12, 11)
        painter.drawLine(6, 7, 6, 13)
        painter.drawLine(10, 7, 10, 13)

        painter.end()
        return QIcon(pixmap)
    
    def set_db_manager(self, db_manager):
        """設定資料庫管理器"""
        self.db_manager = db_manager
    
    def load_data(self, schedules: List[Dict[str, Any]], exceptions: List[Dict[str, Any]]):
        """載入排程與例外資料"""
        self.schedules = schedules
        self.exceptions = exceptions
        self._apply_time_filter()
        self._refresh_calendar_views()
        self._populate_table()
        self._update_button_states()
    
    def refresh(self):
        """重新整理"""
        if self.db_manager:
            self.schedules = self.db_manager.get_all_schedules()
            self.exceptions = self.db_manager.get_all_schedule_exceptions()
            self._apply_time_filter()
            self._refresh_calendar_views()
            self._populate_table()
            self._update_button_states()

    def _on_view_changed(self, index: int):
        mode_map = {0: "day", 1: "week", 2: "month"}
        self.current_view_mode = mode_map.get(index, "day")
        self.calendar_stack.setCurrentIndex(index if index in (0, 1, 2) else 0)
        self._update_range_label()
        self._apply_time_filter()
        self._refresh_calendar_views()
        self._populate_table()
        self._update_button_states()

    def _on_reference_date_changed(self, qdate: QDate):
        self.reference_date = qdate
        self._update_range_label()
        self._apply_time_filter()
        self._refresh_calendar_views()
        self._populate_table()
        self._update_button_states()

    def _apply_reference_date(self, qdate: QDate):
        self._on_reference_date_changed(qdate)

    def _go_previous_period(self):
        current = self.reference_date
        if self.current_view_mode == "day":
            self._apply_reference_date(current.addDays(-1))
        elif self.current_view_mode == "week":
            self._apply_reference_date(current.addDays(-7))
        else:
            self._apply_reference_date(current.addMonths(-1))

    def _go_next_period(self):
        current = self.reference_date
        if self.current_view_mode == "day":
            self._apply_reference_date(current.addDays(1))
        elif self.current_view_mode == "week":
            self._apply_reference_date(current.addDays(7))
        else:
            self._apply_reference_date(current.addMonths(1))

    def _go_current_period(self):
        current_date = QDate.currentDate()
        if self.current_view_mode == "month":
            self._apply_reference_date(QDate(current_date.year(), current_date.month(), 1))
        else:
            self._apply_reference_date(current_date)

    def _week_start(self, qdate: QDate) -> QDate:
        # 以週一為一週起始
        return qdate.addDays(-(qdate.dayOfWeek() - 1))

    def _update_range_label(self):
        self.label_month_title.setText(f"{self._month_zh(self.reference_date.month())} {self.reference_date.year()}")
        if self.current_view_mode == "day":
            self.btn_current_period.setText("Current Day")
        elif self.current_view_mode == "week":
            self.btn_current_period.setText("Current Week")
        else:
            self.btn_current_period.setText("Current Month")

    def _month_zh(self, month: int) -> str:
        months = ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月"]
        if 1 <= month <= 12:
            return months[month - 1]
        return str(month)

    def _show_calendar_popup(self):
        popup_menu = QMenu(self)

        container = QWidget(popup_menu)
        root_layout = QVBoxLayout(container)
        root_layout.setContentsMargins(6, 6, 6, 6)
        root_layout.setSpacing(4)

        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)
        header.setSpacing(6)

        btn_prev = QToolButton(container)
        btn_prev.setText("◀")
        btn_prev.setAutoRaise(True)

        btn_month = QToolButton(container)
        btn_month.setPopupMode(QToolButton.InstantPopup)
        btn_month.setToolButtonStyle(Qt.ToolButtonTextOnly)

        btn_year = QToolButton(container)
        btn_year.setPopupMode(QToolButton.InstantPopup)
        btn_year.setToolButtonStyle(Qt.ToolButtonTextOnly)

        btn_next = QToolButton(container)
        btn_next.setText("▶")
        btn_next.setAutoRaise(True)

        header.addWidget(btn_prev)
        header.addStretch()
        header.addWidget(btn_month)
        header.addWidget(btn_year)
        header.addStretch()
        header.addWidget(btn_next)
        root_layout.addLayout(header)

        mini_calendar = QCalendarWidget(container)
        mini_calendar.setGridVisible(True)
        mini_calendar.setSelectedDate(self.reference_date)
        mini_calendar.setFirstDayOfWeek(Qt.Monday)
        mini_calendar.setNavigationBarVisible(False)
        mini_calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        mini_calendar.setCurrentPage(self.reference_date.year(), self.reference_date.month())
        root_layout.addWidget(mini_calendar)

        def update_header(year: int, month: int):
            btn_month.setText(self._month_zh(month))
            btn_year.setText(str(year))

        def shift_month(delta: int):
            year = mini_calendar.yearShown()
            month = mini_calendar.monthShown()
            new_month = month + delta
            new_year = year
            if new_month < 1:
                new_month = 12
                new_year -= 1
            elif new_month > 12:
                new_month = 1
                new_year += 1
            mini_calendar.setCurrentPage(new_year, new_month)

        def show_month_picker():
            menu = QMenu(btn_month)
            for month in range(1, 13):
                action = menu.addAction(self._month_zh(month))
                action.triggered.connect(
                    lambda checked=False, m=month: mini_calendar.setCurrentPage(mini_calendar.yearShown(), m)
                )
            menu.exec(btn_month.mapToGlobal(btn_month.rect().bottomLeft()))

        def show_year_picker():
            menu = QMenu(btn_year)
            current_year = mini_calendar.yearShown()
            start_year = current_year - (current_year % 12)
            years = list(range(start_year, start_year + 12))

            grid_container = QWidget(menu)
            grid_layout = QVBoxLayout(grid_container)
            grid_layout.setContentsMargins(4, 4, 4, 4)
            grid_layout.setSpacing(2)

            for row_idx in range(3):
                row = QHBoxLayout()
                row.setContentsMargins(0, 0, 0, 0)
                row.setSpacing(2)
                for col_idx in range(4):
                    year = years[row_idx * 4 + col_idx]
                    year_btn = QPushButton(str(year), grid_container)
                    year_btn.setFixedSize(52, 24)
                    if year == current_year:
                        year_btn.setStyleSheet("font-weight: bold;")
                    year_btn.clicked.connect(
                        lambda checked=False, y=year: (
                            mini_calendar.setCurrentPage(y, mini_calendar.monthShown()),
                            menu.close(),
                        )
                    )
                    row.addWidget(year_btn)
                grid_layout.addLayout(row)

            action = QWidgetAction(menu)
            action.setDefaultWidget(grid_container)
            menu.addAction(action)
            menu.exec(btn_year.mapToGlobal(btn_year.rect().bottomLeft()))

        btn_prev.clicked.connect(lambda: shift_month(-1))
        btn_next.clicked.connect(lambda: shift_month(1))
        btn_month.clicked.connect(show_month_picker)
        btn_year.clicked.connect(show_year_picker)
        mini_calendar.currentPageChanged.connect(update_header)
        mini_calendar.clicked.connect(lambda qdate: (self._apply_reference_date(qdate), popup_menu.close()))

        update_header(mini_calendar.yearShown(), mini_calendar.monthShown())

        action = QWidgetAction(popup_menu)
        action.setDefaultWidget(container)
        popup_menu.addAction(action)
        popup_menu.exec(self.btn_calendar_popup.mapToGlobal(self.btn_calendar_popup.rect().bottomLeft()))

    def _apply_time_filter(self):
        """依 Day/Week/Month 篩選例外清單。"""
        filtered = []

        week_start = self._week_start(self.reference_date)
        week_end = week_start.addDays(6)

        for exc in self.exceptions:
            occ_str = str(exc.get("occurrence_date", "")).strip()
            try:
                occ = datetime.strptime(occ_str, "%Y-%m-%d").date()
            except ValueError:
                continue

            occ_qdate = QDate(occ.year, occ.month, occ.day)

            if self.current_view_mode == "day":
                if occ_qdate == self.reference_date:
                    filtered.append(exc)
            elif self.current_view_mode == "week":
                if week_start <= occ_qdate <= week_end:
                    filtered.append(exc)
            else:
                if occ_qdate.year() == self.reference_date.year() and occ_qdate.month() == self.reference_date.month():
                    filtered.append(exc)

        self.filtered_exceptions = filtered

    def _build_calendar_occurrences(self) -> List[ResolvedOccurrence]:
        """把目前篩選後的例外記錄轉為日曆可顯示的 occurrence。"""
        schedule_map = {s.get("id"): s.get("task_name", "未命名排程") for s in self.schedules}
        result: List[ResolvedOccurrence] = []

        for exc in self.filtered_exceptions:
            exc_id = int(exc.get("id", 0))
            schedule_id = int(exc.get("schedule_id", 0))
            action = str(exc.get("action", "")).strip().lower()
            occ_date_str = str(exc.get("occurrence_date", "")).strip()
            if not occ_date_str:
                continue

            try:
                occ_date = datetime.strptime(occ_date_str, "%Y-%m-%d").date()
            except ValueError:
                continue

            title_base = str(schedule_map.get(schedule_id, f"ID={schedule_id}"))
            if action == "cancel":
                title = f"{title_base} (取消)"
                start_dt = datetime.combine(occ_date, time(8, 0))
                end_dt = start_dt + timedelta(hours=1)
                bg = "#c0392b"
            else:
                title = str(exc.get("override_task_name") or f"{title_base} (覆寫)")
                start_raw = exc.get("override_start")
                end_raw = exc.get("override_end")
                try:
                    start_dt = datetime.fromisoformat(str(start_raw)) if start_raw else datetime.combine(occ_date, time(8, 0))
                except Exception:
                    start_dt = datetime.combine(occ_date, time(8, 0))
                try:
                    end_dt = datetime.fromisoformat(str(end_raw)) if end_raw else (start_dt + timedelta(hours=1))
                except Exception:
                    end_dt = start_dt + timedelta(hours=1)
                if end_dt <= start_dt:
                    end_dt = start_dt + timedelta(hours=1)
                bg = "#1f6fd6"

            target_value = str(exc.get("override_target_value") or action or "exception")
            result.append(
                ResolvedOccurrence(
                    schedule_id=schedule_id,
                    source="exception",
                    title=title,
                    start=start_dt,
                    end=end_dt,
                    category_bg=bg,
                    category_fg="#ffffff",
                    target_value=target_value,
                    is_exception=True,
                    is_holiday=False,
                    occurrence_key=f"exception-{exc_id}",
                )
            )

        return result

    def _refresh_calendar_views(self):
        """同步更新 Day/Week/Month 日曆內容。"""
        occurrences = self._build_calendar_occurrences()

        self.day_view.set_reference_date(self.reference_date)
        self.day_view.set_occurrences(occurrences)

        self.week_view.set_reference_date(self.reference_date)
        self.week_view.set_occurrences(occurrences)

        self.month_view.set_reference_date(self.reference_date)
        self.month_view.set_selected_date(self.reference_date)
        self.month_view.set_occurrences(occurrences)
    
    def _populate_table(self):
        """填充表格"""
        self.table.setRowCount(0)
        
        # 建立 schedule_id -> task_name 映射
        schedule_map = {s.get("id"): s.get("task_name", "未知") for s in self.schedules}
        
        for exc in self.filtered_exceptions:
            row = self.table.rowCount()
            self.table.insertRow(row)
            
            # ID
            item_id = QTableWidgetItem(str(exc.get("id", "")))
            item_id.setData(Qt.UserRole, exc)  # 儲存完整資料
            self.table.setItem(row, 0, item_id)
            
            # 排程名稱
            schedule_id = exc.get("schedule_id")
            schedule_name = schedule_map.get(schedule_id, f"ID={schedule_id}")
            self.table.setItem(row, 1, QTableWidgetItem(schedule_name))
            
            # 發生日期
            self.table.setItem(row, 2, QTableWidgetItem(exc.get("occurrence_date", "")))
            
            # 操作
            action = exc.get("action", "")
            action_text = "取消 (Cancel)" if action == "cancel" else "覆寫 (Override)"
            action_item = QTableWidgetItem(action_text)
            if action == "cancel":
                action_item.setBackground(QBrush(QColor(255, 200, 200)))  # 淡紅色
            else:
                action_item.setBackground(QBrush(QColor(200, 220, 255)))  # 淡藍色
            self.table.setItem(row, 3, action_item)
            
            # 覆寫標題
            override_title = exc.get("override_task_name") or "-"
            self.table.setItem(row, 4, QTableWidgetItem(override_title))
            
            # 覆寫數值
            override_value = exc.get("override_target_value") or "-"
            self.table.setItem(row, 5, QTableWidgetItem(override_value))
            
            # 覆寫時間
            override_start = exc.get("override_start")
            override_end = exc.get("override_end")
            if override_start and override_end:
                time_text = f"{override_start} ~ {override_end}"
            else:
                time_text = "-"
            self.table.setItem(row, 6, QTableWidgetItem(time_text))
            
            # 建立時間
            created_at = exc.get("created_at", "")
            self.table.setItem(row, 7, QTableWidgetItem(created_at))
        
        self.table.resizeColumnsToContents()

    def _update_button_states(self):
        """依據資料與選取狀態更新按鈕可用性。"""
        return
    
    def _create_exception(self):
        """建立新例外"""
        if not self.db_manager:
            QMessageBox.warning(self, "錯誤", "資料庫未連線")
            return
        
        if not self.schedules:
            QMessageBox.warning(self, "錯誤", "沒有可用的排程系列，請先建立排程")
            return
        
        dialog = ExceptionEditDialog(self, self.schedules, None)
        if dialog.exec() != QDialog.Accepted:
            return
        
        data = dialog.get_data()
        self._save_exception(data)
    
    def _edit_selected_exception(self):
        """編輯選取的例外"""
        selected_items = self.table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "提示", "請先選擇一筆例外記錄")
            return
        
        row = self.table.currentRow()
        exc_data = self.table.item(row, 0).data(Qt.UserRole)
        
        dialog = ExceptionEditDialog(self, self.schedules, exc_data)
        if dialog.exec() != QDialog.Accepted:
            return
        
        data = dialog.get_data()
        self._save_exception(data)
    
    def _delete_selected_exception(self):
        """刪除選取的例外"""
        selected_items = self.table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "提示", "請先選擇一筆例外記錄")
            return
        
        row = self.table.currentRow()
        exc_data = self.table.item(row, 0).data(Qt.UserRole)
        exc_id = exc_data.get("id")
        
        reply = QMessageBox.question(
            self,
            "確認刪除",
            f"確定要刪除此例外記錄嗎？\nID: {exc_id}",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes and self.db_manager:
            # 從資料庫刪除（需要添加此方法到 db_manager）
            self._delete_exception_from_db(exc_id)
            self.exception_changed.emit()
            self.refresh()
    
    def _save_exception(self, data: Dict[str, Any]):
        """儲存例外記錄"""
        if not self.db_manager:
            return
        
        action = data.get("action")
        schedule_id = data.get("schedule_id")
        occurrence_date = data.get("occurrence_date")
        
        if action == "cancel":
            self.db_manager.add_schedule_exception_cancel(schedule_id, occurrence_date)
        else:
            self.db_manager.add_schedule_exception_override(
                schedule_id=schedule_id,
                occurrence_date=occurrence_date,
                override_start=data.get("override_start"),
                override_end=data.get("override_end"),
                override_task_name=data.get("override_task_name"),
                override_target_value=data.get("override_target_value"),
            )
        
        self.exception_changed.emit()
        self.refresh()
    
    def _delete_exception_from_db(self, exception_id: int):
        """從資料庫刪除例外記錄"""
        if not self.db_manager:
            return
        
        success = self.db_manager.delete_schedule_exception(exception_id)
        if not success:
            QMessageBox.warning(self, "刪除失敗", f"無法刪除例外記錄 ID: {exception_id}")

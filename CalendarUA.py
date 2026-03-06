#!/usr/bin/env python3
"""
CalendarUA - 工業自動化排程管理系統主程式
採用 PySide6 開發，結合 Office/Outlook 風格行事曆介面
"""

import sys
import asyncio
import os
from datetime import datetime, timedelta
from typing import Optional, Dict, Any, List
import logging

# 設定統一的日誌記錄級別
logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QCalendarWidget,
    QPushButton,
    QLabel,
    QLineEdit,
    QInputDialog,
    QGroupBox,
    QMessageBox,
    QMenu,
    QSystemTrayIcon,
    QStyle,
    QDialog,
    QComboBox,
    QStatusBar,
    QToolBar,
    QTreeWidget,
    QRadioButton,
    QCheckBox,
    QFileDialog,
    QTreeWidgetItem,
    QStackedWidget,
    QToolButton,
)
from PySide6.QtCore import Qt, QTimer, Signal, Slot, QThread, QDate, QSize, QLocale, QTime, QEvent
from PySide6.QtGui import QAction, QFont, QColor, QGuiApplication
import qasync
import re
from datetime import date as dt_date

from database.sqlite_manager import SQLiteManager
from core.opc_handler import OPCHandler
from core.rrule_parser import RRuleParser
from core.schedule_resolver import resolve_occurrences_for_range
from ui.recurrence_dialog import RecurrenceDialog
from ui.database_settings_dialog import DatabaseSettingsDialog
from ui.holiday_settings_dialog import HolidaySettingsDialog
from ui.combo_wheel_helper import attach_combo_wheel_behavior
from ui.schedule_canvas import DayViewWidget, WeekViewWidget
from ui.month_grid import MonthViewWidget
from core.lunar_calendar import to_lunar, format_lunar_day_text
from ui.app_icon import get_app_icon


class NavCalendarWidget(QCalendarWidget):
    """
    導覽用小月曆（完全自繪 + 自行管理選取）。

    - role = "top"：顯示目前月份，隱藏「下個月」銜接格，並管理選取
    - role = "bottom"：顯示下一月份，隱藏「上個月」銜接格（純預覽，不選取）
    """

    date_clicked = Signal(QDate)  # 對外統一用這個訊號回報點擊日期

    def __init__(self, role: str, parent=None):
        super().__init__(parent)
        self.role = role  # "top" or "bottom"
        self._forced_selected_date: QDate | None = None  # 自行管理的選取日期
        self._holiday_checker = None
        self._is_dark_theme = False

        # 使用 Qt 內建 clicked(QDate) 訊號，不再自己做 hit-test
        self.clicked.connect(self._on_qt_clicked)

        # 延遲安裝事件過濾器，阻擋內部子元件（表格）滾輪切月
        QTimer.singleShot(0, self._install_wheel_blockers)

    def _install_wheel_blockers(self):
        self.installEventFilter(self)
        for child in self.findChildren(QWidget):
            child.installEventFilter(self)

    def set_forced_selected_date(self, date: QDate):
        """由外部設定目前選取日期（兩個小月曆共用同一個選取日期）"""
        self._forced_selected_date = date
        self.update()

    def set_holiday_checker(self, checker):
        """設定假日判斷函式，簽章為 checker(QDate) -> bool。"""
        self._holiday_checker = checker
        self.update()

    def set_theme_dark(self, is_dark: bool):
        self._is_dark_theme = bool(is_dark)
        self.update()

    def _is_holiday(self, date: QDate) -> bool:
        if date.dayOfWeek() in (6, 7):
            return True
        if self._holiday_checker is None:
            return False
        try:
            return bool(self._holiday_checker(date))
        except Exception:
            return False

    def _on_qt_clicked(self, clicked: QDate):
        """由 Qt 傳入正確的 clicked 日期，再套用我們的隱藏/跨月規則。"""
        shown_year = self.yearShown()
        shown_month = self.monthShown()
        first = QDate(shown_year, shown_month, 1)
        prev_month = first.addMonths(-1)
        next_month = first.addMonths(1)

        is_prev = (clicked.year() == prev_month.year() and clicked.month() == prev_month.month())
        is_next = (clicked.year() == next_month.year() and clicked.month() == next_month.month())

        # 對應 paintCell 的「隱藏交界格」規則：這些格子應該完全無效
        hide = False
        if self.role == "top" and is_next:
            hide = True      # 上方隱藏「下個月」交界格
        if self.role == "bottom" and is_prev:
            hide = True      # 下方隱藏「上個月」交界格

        if hide:
            return

        # 其餘日期（本月白字 + 合法淺灰）都視為有效點擊
        self._forced_selected_date = clicked
        self.update()
        self.date_clicked.emit(clicked)

    def wheelEvent(self, event):
        """停用滑鼠滾輪切換月份。"""
        event.accept()
        return

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Wheel:
            event.accept()
            return True
        return super().eventFilter(watched, event)

    def paintCell(self, painter, rect, date):
        shown_year = self.yearShown()
        shown_month = self.monthShown()
        first = QDate(shown_year, shown_month, 1)
        prev_month = first.addMonths(-1)
        next_month = first.addMonths(1)

        is_prev = (date.year() == prev_month.year() and date.month() == prev_month.month())
        is_this = (date.year() == shown_year and date.month() == shown_month)
        is_next = (date.year() == next_month.year() and date.month() == next_month.month())

        painter.save()
        # 底色
        cell_bg = QColor("#2b2b2b") if self._is_dark_theme else QColor("#ffffff")
        painter.fillRect(rect, cell_bg)

        # 決定這格是否完全不顯示（空白）
        hide = False
        if self.role == "top" and is_next:
            hide = True      # 上方隱藏「下個月」交界格
        if self.role == "bottom" and is_prev:
            hide = True      # 下方隱藏「上個月」交界格

        is_dark_palette = self._is_dark_theme

        lunar_text = ""
        try:
            lunar_info = to_lunar(dt_date(date.year(), date.month(), date.day()))
            if lunar_info:
                lunar_text = format_lunar_day_text(lunar_info)
        except Exception:
            lunar_text = ""

        if not hide:
            # 本月白字、前後月灰字
            is_holiday = self._is_holiday(date)
            if is_this:
                if is_holiday:
                    day_fg = QColor("#c62828")
                else:
                    day_fg = QColor("#f0f0f0") if is_dark_palette else QColor("#202020")
            else:
                day_fg = QColor("#b36b6b") if is_holiday else QColor("#808080")

            # 上方：國曆（較大）
            painter.setPen(day_fg)
            solar_font = QFont(painter.font())
            solar_font.setFamily("Segoe UI")
            solar_font.setBold(True)
            solar_font.setPointSize(13)
            painter.setFont(solar_font)
            top_rect = rect.adjusted(0, 1, 0, -rect.height() // 2)
            painter.drawText(top_rect, Qt.AlignHCenter | Qt.AlignVCenter, str(date.day()))

            # 下方：農曆（國字）
            if lunar_text:
                lunar_font = QFont(painter.font())
                lunar_font.setFamily("Microsoft JhengHei")
                lunar_font.setBold(False)
                lunar_font.setPointSize(9)
                painter.setFont(lunar_font)
                bottom_rect = rect.adjusted(0, rect.height() // 2 - 1, 0, 0)
                painter.drawText(bottom_rect, Qt.AlignHCenter | Qt.AlignVCenter, lunar_text)

            # 今日標記（左側小月曆永久保留一個 today 高亮）
            if date == QDate.currentDate():
                today_pen = QColor("#ff8f00")
                painter.setPen(today_pen)
                painter.setBrush(Qt.NoBrush)
                today_rect = rect.adjusted(3, 3, -3, -3)
                painter.drawRect(today_rect)

        # 畫選取高亮（兩個小月曆共用同一個選取日期；但隱藏格不畫）
        if (not hide) and self._forced_selected_date and date == self._forced_selected_date:
            sel = QColor("#2e7d32") if is_dark_palette else QColor("#66bb6a")
            painter.setPen(Qt.NoPen)
            painter.setBrush(sel)
            r = rect.adjusted(2, 2, -2, -2)
            painter.drawRect(r)
            is_holiday = self._is_holiday(date)
            if is_this:
                if is_holiday:
                    selected_day_fg = QColor("#c62828")
                else:
                    selected_day_fg = QColor("#f0f0f0") if is_dark_palette else QColor("#202020")
            else:
                selected_day_fg = QColor("#b36b6b") if is_holiday else QColor("#808080")
            painter.setPen(selected_day_fg)

            solar_font = QFont(painter.font())
            solar_font.setFamily("Segoe UI")
            solar_font.setBold(True)
            solar_font.setPointSize(13)
            painter.setFont(solar_font)
            top_rect = rect.adjusted(0, 1, 0, -rect.height() // 2)
            painter.drawText(top_rect, Qt.AlignHCenter | Qt.AlignVCenter, str(date.day()))

            if lunar_text:
                lunar_font = QFont(painter.font())
                lunar_font.setFamily("Microsoft JhengHei")
                lunar_font.setBold(False)
                lunar_font.setPointSize(9)
                painter.setFont(lunar_font)
                bottom_rect = rect.adjusted(0, rect.height() // 2 - 1, 0, 0)
                painter.drawText(bottom_rect, Qt.AlignHCenter | Qt.AlignVCenter, lunar_text)

        painter.restore()


class YearNavComboBox(QComboBox):
    """年份下拉：支援滑鼠滾輪觸發年份遞增/遞減。"""

    year_step_requested = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.installEventFilter(self)
        self.view().installEventFilter(self)
        self.view().viewport().installEventFilter(self)

    def _steps_from_wheel(self, event) -> int:
        delta = event.angleDelta().y()
        if delta == 0:
            return 0
        steps = int(delta / 120)
        if steps == 0:
            steps = 1 if delta > 0 else -1
        return -steps

    def event(self, event):
        if event.type() == QEvent.Wheel:
            self.wheelEvent(event)
            return True
        return super().event(event)

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Wheel and watched in (self, self.view(), self.view().viewport()):
            steps = self._steps_from_wheel(event)
            if steps != 0:
                self.year_step_requested.emit(steps)
            event.accept()
            return True
        return super().eventFilter(watched, event)

    def wheelEvent(self, event):
        steps = self._steps_from_wheel(event)
        if steps != 0:
            self.year_step_requested.emit(steps)
        event.accept()


def _combo_steps_from_wheel(event) -> int:
    delta = event.angleDelta().y()
    if delta == 0:
        return 0
    steps = int(delta / 120)
    if steps == 0:
        steps = 1 if delta > 0 else -1
    return steps


class SchedulerWorker(QThread):
    """背景排程工作執行緒"""

    trigger_task = Signal(dict)

    def __init__(self, db_manager: SQLiteManager, check_interval: int = 1):
        super().__init__()
        self.db_manager = db_manager
        self.check_interval = check_interval
        self.running = True
        self.last_check = datetime.now()
        # 記錄每個排程的上次觸發時間，防止重複觸發
        self.last_trigger_times: Dict[int, datetime] = {}

    def run(self):
        """持續檢查排程"""
        while self.running:
            try:
                current_time = datetime.now()

                # 取得所有啟用的排程
                schedules = self.db_manager.get_all_schedules(enabled_only=True)

                for schedule in schedules:
                    trigger_time = self.should_trigger(schedule, current_time)
                    if trigger_time is not None:
                        schedule_payload = dict(schedule)
                        schedule_payload["_trigger_time"] = trigger_time
                        self.trigger_task.emit(schedule_payload)

                self.last_check = current_time

                # 休眠指定秒數
                for _ in range(self.check_interval):
                    if not self.running:
                        break
                    self.msleep(1000)

            except Exception as e:
                print(f"排程檢查錯誤: {e}")
                self.msleep(5000)

    def _parse_duration_minutes(self, rrule_str: str) -> int:
        """從 RRULE 解析 DURATION（分鐘）。"""
        if not rrule_str:
            return 0

        try:
            match = re.search(r"DURATION=PT(?:(\d+)H)?(?:(\d+)M)?", rrule_str.upper())
            if not match:
                return 0
            hours = int(match.group(1) or 0)
            minutes = int(match.group(2) or 0)
            return max(0, hours * 60 + minutes)
        except Exception:
            return 0

    def should_trigger(self, schedule: Dict[str, Any], current_time: datetime) -> Optional[datetime]:
        """檢查是否應該觸發排程，回傳本次 occurrence 開始時間。"""
        schedule_id = schedule.get("id")
        rrule_str = schedule.get("rrule_str", "")
        if not rrule_str:
            return None

        tolerance_seconds = 30
        trigger_anchor: Optional[datetime] = None

        # 正常情況：檢查上次輪詢到現在之間是否有新觸發。
        check_start = max(self.last_check, current_time - timedelta(seconds=tolerance_seconds))
        near_triggers = RRuleParser.get_trigger_between(
            rrule_str,
            check_start,
            current_time,
        )
        if near_triggers:
            trigger_anchor = max(near_triggers)

        # 啟動回補：若第一次看到此排程且有 DURATION，回補目前期間內最近一次觸發。
        if trigger_anchor is None and schedule_id not in self.last_trigger_times:
            duration_minutes = self._parse_duration_minutes(rrule_str)
            if duration_minutes > 0:
                window_start = current_time - timedelta(minutes=duration_minutes)
                candidate_triggers = RRuleParser.get_trigger_between(rrule_str, window_start, current_time)
                if candidate_triggers:
                    latest_trigger = max(candidate_triggers)
                    if latest_trigger <= current_time < latest_trigger + timedelta(minutes=duration_minutes):
                        trigger_anchor = latest_trigger

        if trigger_anchor is None:
            return None

        # 以 occurrence 開始時間防止重複觸發（不是用「現在時間」）
        last_trigger = self.last_trigger_times.get(schedule_id)
        if isinstance(last_trigger, datetime) and last_trigger == trigger_anchor:
            return None

        self.last_trigger_times[schedule_id] = trigger_anchor
        return trigger_anchor

    def stop(self):
        """停止工作執行緒"""
        self.running = False
        self.wait(2000)


class CalendarUA(QMainWindow):
    """CalendarUA 主視窗"""

    def __init__(self):
        super().__init__()

        self.db_manager: Optional[SQLiteManager] = None
        self.scheduler_worker: Optional[SchedulerWorker] = None
        self.schedules: List[Dict[str, Any]] = []
        self.schedule_exceptions: List[Dict[str, Any]] = []
        self.holiday_entries: List[Dict[str, Any]] = []
        # 主行事曆視圖狀態
        self.current_view_mode: str = "month"
        self.reference_date: QDate = QDate.currentDate()
        
        # 執行計數器：schedule_id -> 已執行次數
        self.execution_counts: Dict[int, int] = {}
        
        # 正在執行的任務ID集合，防止重複執行
        self.running_tasks: set[int] = set()

        # 目前選取的排程 ID (Ribbon Edit/Delete 使用)
        self.selected_schedule_id: Optional[int] = None
        self._copied_schedule_ids: List[int] = []

        # 主題模式: "light", "dark", "system"
        self.current_theme = "system"
        self._allow_minimize_to_tray = False

        self.setup_ui()
        attach_combo_wheel_behavior(self)
        self.apply_modern_style()
        self.setup_connections()
        self.setup_system_tray()

        # 初始化資料庫連線
        self.init_database()

        # 設定系統主題監聽
        self.setup_theme_listener()
        QTimer.singleShot(1200, self._enable_minimize_to_tray)

    def _enable_minimize_to_tray(self):
        self._allow_minimize_to_tray = True

    def setup_ui(self):
        """設定使用者介面"""
        self.setWindowTitle("CalendarUA")
        self.setWindowIcon(get_app_icon())
        self.setMinimumSize(1100, 760)
        self.resize(1320, 840)

        # 建立中央widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # 主要內容面板 (全寬顯示)
        main_panel = self.create_main_panel()
        main_layout.addWidget(main_panel)

        # 建立選單列
        self.create_menu_bar()

        # 建立狀態列
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就緒")

    def create_main_panel(self) -> QWidget:
        """建立主要內容面板（模仿 Outlook 行事曆版面）"""
        panel = QWidget()
        root_layout = QHBoxLayout(panel)
        root_layout.setContentsMargins(0, 0, 0, 0)
        # 減少左右區塊間距，避免中間有明顯空白
        root_layout.setSpacing(4)

        # 左側：導覽月曆 + 行事曆清單
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        # 讓「上一月/月份/年份/下一月」這一排和月曆本體緊貼在一起
        left_layout.setSpacing(0)

        # 導覽月份標頭：上一月 / 月份下拉 / 年份下拉 / 下一月
        header_layout = QGridLayout()
        # 略留左右 2px，維持整體緊湊
        header_layout.setContentsMargins(2, 0, 2, 0)
        header_layout.setHorizontalSpacing(2)
        header_layout.setVerticalSpacing(0)

        # 使用與 QCalendarWidget 導覽列相似的 QToolButton 圖示，但放大圖示尺寸，增加點擊區
        self.btn_nav_prev = QToolButton()
        self.btn_nav_prev.setIcon(self.style().standardIcon(QStyle.SP_ArrowLeft))
        self.btn_nav_prev.setAutoRaise(True)
        self.btn_nav_prev.setIconSize(QSize(22, 22))
        self.btn_nav_prev.setFixedSize(34, 30)

        self.btn_nav_next = QToolButton()
        self.btn_nav_next.setIcon(self.style().standardIcon(QStyle.SP_ArrowRight))
        self.btn_nav_next.setAutoRaise(True)
        self.btn_nav_next.setIconSize(QSize(22, 22))
        self.btn_nav_next.setFixedSize(34, 30)

        self.combo_nav_month = QComboBox()
        self.combo_nav_year = YearNavComboBox()
        self.combo_nav_month.setEditable(True)
        self.combo_nav_year.setEditable(True)
        if self.combo_nav_month.lineEdit() is not None:
            self.combo_nav_month.lineEdit().setReadOnly(True)
            self.combo_nav_month.lineEdit().setAlignment(Qt.AlignCenter)
            self.combo_nav_month.lineEdit().setCursor(Qt.PointingHandCursor)
            self.combo_nav_month.lineEdit().installEventFilter(self)
        if self.combo_nav_year.lineEdit() is not None:
            self.combo_nav_year.lineEdit().setReadOnly(True)
            self.combo_nav_year.lineEdit().setAlignment(Qt.AlignCenter)
            self.combo_nav_year.lineEdit().setCursor(Qt.PointingHandCursor)
            self.combo_nav_year.lineEdit().installEventFilter(self)
        self.combo_nav_month.installEventFilter(self)
        self.combo_nav_year.installEventFilter(self)
        self.combo_nav_month.setMaxVisibleItems(12)
        self.combo_nav_year.setMaxVisibleItems(11)
        self.combo_nav_month.view().setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.combo_nav_month.view().setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.combo_nav_year.view().setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.combo_nav_year.view().setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        # 下拉箭頭隱藏後，可用較緊湊寬度並避免與右箭頭重疊
        self.combo_nav_month.setFixedWidth(64)
        # 年度需要完整顯示四位數
        self.combo_nav_year.setFixedWidth(60)
        common_combo_style = """
            QComboBox {
                font-family: 'Segoe UI';
                font-size: 16px;
                padding-right: 2px;
            }
            QComboBox QAbstractItemView {
                text-align: center;
                outline: 0;
            }
            QComboBox QAbstractItemView::item {
                min-height: 24px;
            }
            QComboBox QAbstractItemView QScrollBar:vertical {
                width: 0px;
            }
            QComboBox QAbstractItemView QScrollBar:horizontal {
                height: 0px;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 0px;
                border: none;
                padding: 0px;
                margin: 0px;
            }
            QComboBox::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
            }
        """
        self.combo_nav_month.setStyleSheet(common_combo_style)
        self.combo_nav_year.setStyleSheet(common_combo_style)

        # 今日按鈕（使用符號，縮小尺寸，避免套用全域 QPushButton 樣式變成大藍塊）
        self.btn_nav_today = QToolButton()
        self.btn_nav_today.setText("●")
        self.btn_nav_today.setAutoRaise(True)
        self.btn_nav_today.setFixedSize(26, 26)
        self.btn_nav_today.setToolTip("跳到今天")

        left_cluster = QWidget()
        left_cluster_layout = QHBoxLayout(left_cluster)
        left_cluster_layout.setContentsMargins(0, 0, 0, 0)
        left_cluster_layout.setSpacing(2)
        left_cluster_layout.addWidget(self.btn_nav_prev)
        left_cluster_layout.addWidget(self.combo_nav_year)

        right_cluster = QWidget()
        right_cluster_layout = QHBoxLayout(right_cluster)
        right_cluster_layout.setContentsMargins(0, 0, 0, 0)
        right_cluster_layout.setSpacing(2)
        right_cluster_layout.addWidget(self.combo_nav_month)
        right_cluster_layout.addWidget(self.btn_nav_next)

        # 以 3 欄對稱配置，讓「今日」固定在月曆正中央
        header_layout.setColumnStretch(0, 1)
        header_layout.setColumnStretch(1, 0)
        header_layout.setColumnStretch(2, 1)
        header_layout.addWidget(left_cluster, 0, 0, alignment=Qt.AlignRight | Qt.AlignVCenter)
        header_layout.addWidget(self.btn_nav_today, 0, 1, alignment=Qt.AlignCenter)
        header_layout.addWidget(right_cluster, 0, 2, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        left_layout.addLayout(header_layout)

        # 左上：目前月份導覽月曆
        self.nav_calendar = NavCalendarWidget(role="top")
        self.nav_calendar.setGridVisible(True)
        self.nav_calendar.setFirstDayOfWeek(Qt.Sunday)
        self.nav_calendar.setLocale(QLocale(QLocale.Chinese, QLocale.Taiwan))
        self.nav_calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.nav_calendar.setHorizontalHeaderFormat(QCalendarWidget.ShortDayNames)
        self.nav_calendar.setFixedWidth(236)
        self.nav_calendar.setNavigationBarVisible(False)
        left_layout.addWidget(self.nav_calendar)

        # 左下標題：顯示下一個月的「年 / 月」文字（不使用下拉）
        self.label_nav_next_header = QLabel("")
        self.label_nav_next_header.setAlignment(Qt.AlignCenter)
        self.label_nav_next_header.setStyleSheet("font-family: 'Segoe UI'; font-size: 16px; font-weight: 600;")
        left_layout.addWidget(self.label_nav_next_header)

        # 左下：下一個月預覽（僅顯示）
        self.nav_calendar_next = NavCalendarWidget(role="bottom")
        self.nav_calendar_next.setGridVisible(True)
        self.nav_calendar_next.setFirstDayOfWeek(Qt.Sunday)
        self.nav_calendar_next.setLocale(QLocale(QLocale.Chinese, QLocale.Taiwan))
        self.nav_calendar_next.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.nav_calendar_next.setHorizontalHeaderFormat(QCalendarWidget.ShortDayNames)
        self.nav_calendar_next.setFixedWidth(236)
        self.nav_calendar_next.setNavigationBarVisible(False)
        left_layout.addWidget(self.nav_calendar_next)

        # 讓左側整個區塊寬度與月曆一致，不再比行事曆寬
        left_widget.setFixedWidth(236)

        root_layout.addWidget(left_widget, 0)

        # 右側：視圖切換列 + Day/Week/Month 主視圖
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        # 讓標題列與行事曆視圖之間緊貼，消除中間大塊空白
        right_layout.setSpacing(0)

        # 視圖切換工具列（左：主視圖上一段/下一段；中：目前範圍；右：日/週/月）
        view_toolbar = QHBoxLayout()
        view_toolbar.setContentsMargins(4, 0, 4, 0)
        view_toolbar.setSpacing(4)

        # 主視圖上一段 / 下一段按鈕
        self.btn_main_prev = QToolButton()
        self.btn_main_prev.setIcon(self.style().standardIcon(QStyle.SP_ArrowLeft))
        self.btn_main_prev.setAutoRaise(True)
        self.btn_main_prev.setIconSize(QSize(18, 18))

        self.btn_main_next = QToolButton()
        self.btn_main_next.setIcon(self.style().standardIcon(QStyle.SP_ArrowRight))
        self.btn_main_next.setAutoRaise(True)
        self.btn_main_next.setIconSize(QSize(18, 18))

        self.label_current_range = QLabel("")
        self.label_current_range.setStyleSheet(
            "font-family: 'Segoe UI'; font-size: 16px; font-weight: 600;"
        )
        # 讓「分類 / 上一段 / 標題 / 下一段」這組元件整體置中
        view_toolbar.addStretch()
        view_toolbar.addWidget(self.btn_main_prev)
        view_toolbar.addWidget(self.label_current_range)
        view_toolbar.addWidget(self.btn_main_next)
        view_toolbar.addStretch()

        # 右側按鈕區：日/週/月/假日（可高亮）
        self.btn_view_schedule_list = QPushButton("排程清單")
        self.btn_view_day = QPushButton("日")
        self.btn_view_week = QPushButton("週")
        self.btn_view_month = QPushButton("月")
        self.btn_holiday_settings = QPushButton("假日")
        for btn in (self.btn_view_schedule_list, self.btn_view_day, self.btn_view_week, self.btn_view_month):
            btn.setCheckable(True)
            btn.setMinimumWidth(48 if btn != self.btn_view_schedule_list else 88)
            view_toolbar.addWidget(btn)
        self.btn_holiday_settings.setCheckable(False)
        self.btn_holiday_settings.setMinimumWidth(48)
        view_toolbar.addWidget(self.btn_holiday_settings)

        right_layout.addLayout(view_toolbar)

        # 主行事曆視圖堆疊
        self.calendar_stack = QStackedWidget()
        self.day_view = DayViewWidget()
        self.week_view = WeekViewWidget()
        self.month_view = MonthViewWidget()
        self.month_view.set_holiday_checker(self._is_holiday_qdate)
        self.schedule_list_view = QTreeWidget()
        self.schedule_list_view.setColumnCount(2)
        self.schedule_list_view.setHeaderLabels(["欄位", "內容"])
        self.schedule_list_view.setRootIsDecorated(True)
        self.schedule_list_view.setAlternatingRowColors(True)
        self.schedule_list_view.setUniformRowHeights(False)
        self.calendar_stack.addWidget(self.day_view)
        self.calendar_stack.addWidget(self.week_view)
        self.calendar_stack.addWidget(self.month_view)
        self.calendar_stack.addWidget(self.schedule_list_view)
        right_layout.addWidget(self.calendar_stack)

        # 資料庫狀態列
        status_layout = QHBoxLayout()
        self.db_status_label = QLabel("資料庫: 未連線")
        status_layout.addWidget(self.db_status_label)
        status_layout.addStretch()
        right_layout.addLayout(status_layout)

        root_layout.addWidget(right_widget, 1)

        # 訊號連接
        self.nav_calendar.date_clicked.connect(self.on_nav_calendar_date_clicked)
        self.nav_calendar_next.date_clicked.connect(self.on_nav_calendar_date_clicked)
        self.nav_calendar.set_holiday_checker(self._is_holiday_qdate)
        self.nav_calendar_next.set_holiday_checker(self._is_holiday_qdate)
        self.btn_nav_prev.clicked.connect(lambda: self._shift_nav_month(-1))
        self.btn_nav_next.clicked.connect(lambda: self._shift_nav_month(1))
        self.combo_nav_month.currentIndexChanged.connect(self._on_nav_combo_changed)
        self.combo_nav_year.currentIndexChanged.connect(self._on_nav_combo_changed)
        self.combo_nav_year.year_step_requested.connect(self._on_nav_year_wheel_step)
        self.btn_nav_today.clicked.connect(self._go_to_today_from_nav)
        self.btn_main_prev.clicked.connect(lambda: self._shift_main_range(-1))
        self.btn_main_next.clicked.connect(lambda: self._shift_main_range(1))
        self.btn_view_schedule_list.clicked.connect(lambda: self._set_view_mode("list"))
        self.btn_view_day.clicked.connect(lambda: self._set_view_mode("day"))
        self.btn_view_week.clicked.connect(lambda: self._set_view_mode("week"))
        self.btn_view_month.clicked.connect(lambda: self._set_view_mode("month"))
        self.btn_holiday_settings.clicked.connect(self.show_holiday_settings)

        # 初始化導覽月份下拉選單
        self._init_nav_month_year()

        # 預設為月視圖
        self._set_view_mode("month", initial=True)

        # 將 Day / Week / Month 視圖的右鍵動作導向主視窗邏輯
        self.day_view.context_action_requested.connect(self._on_calendar_context_action)
        self.week_view.context_action_requested.connect(self._on_calendar_context_action)
        self.day_view.time_scale_changed.connect(self._on_time_scale_changed)
        self.week_view.time_scale_changed.connect(self._on_time_scale_changed)
        self.month_view.context_action_requested.connect(self._on_calendar_context_action)
        self.month_view.date_selected.connect(self._on_main_calendar_date_selected)

        return panel

    def create_menu_bar(self):
        """建立選單列"""
        menubar = self.menuBar()
        
        # File 選單
        file_menu = menubar.addMenu("&File")
        
        # New Project：建立新的資料庫檔（新專案）
        self.action_new_project = QAction("&New Project...", self)
        self.action_new_project.setStatusTip("建立新的專案資料庫")
        self.action_new_project.triggered.connect(self.new_project)
        file_menu.addAction(self.action_new_project)

        file_menu.addSeparator()
        
        # Load Project：開啟另一個資料庫檔
        self.action_load_project = QAction("&Load Project...", self)
        self.action_load_project.setStatusTip("開啟既有專案資料庫")
        self.action_load_project.triggered.connect(self.load_project_database)
        file_menu.addAction(self.action_load_project)

        file_menu.addSeparator()

        self.action_db_settings = QAction("&Database Settings...", self)
        self.action_db_settings.setStatusTip("資料庫連線設定")
        self.action_db_settings.triggered.connect(self.show_database_settings)
        file_menu.addAction(self.action_db_settings)
        
        file_menu.addSeparator()
        
        self.action_exit = QAction("E&xit", self)
        self.action_exit.setShortcut("Ctrl+Q")
        self.action_exit.setStatusTip("離開程式")
        self.action_exit.triggered.connect(self.close)
        file_menu.addAction(self.action_exit)
        
        # Help 選單
        help_menu = menubar.addMenu("&Help")
        
        self.action_about = QAction("&About CalendarUA", self)
        self.action_about.triggered.connect(self.show_about)
        help_menu.addAction(self.action_about)

    def _on_theme_menu_triggered(self, action):
        """處理主題選單點擊，確保只有一個選項被選中"""
        for theme, act in self.theme_action_group.items():
            if act != action:
                act.setChecked(False)

    def create_tool_bar(self):
        """建立工具列"""
        toolbar = QToolBar("主工具列")
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(24, 24))
        self.addToolBar(toolbar)
        
        # Create 按鈕
        self.btn_toolbar_new = QPushButton("New")
        self.btn_toolbar_new.setToolTip("新增排程 (Ctrl+N)")
        self.btn_toolbar_new.setFixedWidth(80)
        self.btn_toolbar_new.clicked.connect(self.add_schedule)
        toolbar.addWidget(self.btn_toolbar_new)
        
        toolbar.addSeparator()
        
        # Edit 按鈕
        self.btn_toolbar_edit = QPushButton("Edit")
        self.btn_toolbar_edit.setToolTip("編輯排程 (Ctrl+E)")
        self.btn_toolbar_edit.setFixedWidth(80)
        self.btn_toolbar_edit.setEnabled(True)
        self.btn_toolbar_edit.clicked.connect(self.edit_selected_schedule)
        toolbar.addWidget(self.btn_toolbar_edit)
        
        # Delete 按鈕
        self.btn_toolbar_delete = QPushButton("Delete")
        self.btn_toolbar_delete.setToolTip("刪除排程 (Del)")
        self.btn_toolbar_delete.setFixedWidth(80)
        self.btn_toolbar_delete.setEnabled(True)
        self.btn_toolbar_delete.clicked.connect(self.delete_selected_schedule)
        toolbar.addWidget(self.btn_toolbar_delete)
        
        toolbar.addSeparator()
        
        # Refresh 按鈕
        self.btn_toolbar_refresh = QPushButton("Refresh")
        self.btn_toolbar_refresh.setToolTip("重新載入排程資料 (F5)")
        self.btn_toolbar_refresh.setFixedWidth(80)
        self.btn_toolbar_refresh.clicked.connect(self.refresh_schedules)
        toolbar.addWidget(self.btn_toolbar_refresh)

        # Apply 按鈕
        self.btn_toolbar_apply = QPushButton("Apply")
        self.btn_toolbar_apply.setToolTip("套用排程變更 (Ctrl+Shift+A)")
        self.btn_toolbar_apply.setFixedWidth(80)
        self.btn_toolbar_apply.clicked.connect(self.apply_schedules)
        toolbar.addWidget(self.btn_toolbar_apply)
        
        toolbar.addSeparator()
        
        toolbar.addWidget(QLabel(""))  # Spacer
        toolbar.addSeparator()
        
        # Scheduler 控制
        scheduler_label = QLabel("Scheduler:")
        toolbar.addWidget(scheduler_label)
        
        self.scheduler_status_label = QLabel("Running")
        self.scheduler_status_label.setStyleSheet("color: green; font-weight: bold;")
        toolbar.addWidget(self.scheduler_status_label)

    def setup_connections(self):
        """設定信號連接"""
        return

    def setup_system_tray(self):
        """設定系統托盤"""
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(get_app_icon())  # 使用應用圖標
        self.tray_icon.setToolTip("CalendarUA - 工業自動化排程管理系統")

        tray_menu = QMenu()
        show_action = QAction("顯示", self)
        show_action.triggered.connect(self.show_window)
        tray_menu.addAction(show_action)

        tray_menu.addSeparator()

        quit_action = QAction("結束", self)
        quit_action.triggered.connect(self.close)
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()  # 總是顯示托盤圖標

    def on_tray_activated(self, reason):
        """處理托盤圖示點擊"""
        if reason == QSystemTrayIcon.DoubleClick:
            self.setWindowState(self.windowState() & ~Qt.WindowMinimized)
            self.show()
            self.raise_()
            QTimer.singleShot(100, self.activateWindow)

    def show_window(self):
        """顯示視窗並置於最上層"""
        self.setWindowState(self.windowState() & ~Qt.WindowMinimized)
        self.show()
        self.raise_()
        QTimer.singleShot(100, self.activateWindow)

    def apply_modern_style(self):
        """套用現代化樣式，根據主題模式選擇亮色或暗色主題"""
        # 判斷是否使用暗色模式
        is_dark = False
        if self.current_theme == "dark":
            is_dark = True
        elif self.current_theme == "system":
            is_dark = self.is_system_dark_mode()

        if is_dark:
            self._apply_dark_theme()
        else:
            self._apply_light_theme()

        self._apply_nav_calendar_theme(is_dark)
        self._apply_view_toolbar_button_style(is_dark)
        if hasattr(self, "day_view"):
            self.day_view.apply_theme_style(is_dark)
        if hasattr(self, "week_view"):
            self.week_view.apply_theme_style(is_dark)
        self._refresh_theme_sensitive_views()

    def _refresh_theme_sensitive_views(self):
        """主題切換後，立即刷新自繪與依 palette 上色的內容。"""
        if hasattr(self, "nav_calendar"):
            self.nav_calendar.update()
        if hasattr(self, "nav_calendar_next"):
            self.nav_calendar_next.update()
        if hasattr(self, "day_view"):
            self.day_view.update()
            self.day_view.table.viewport().update()
        if hasattr(self, "week_view"):
            self.week_view.update()
            self.week_view.table.viewport().update()
        if hasattr(self, "month_view"):
            self.month_view.update()

        QTimer.singleShot(0, self._refresh_main_calendar_views)

    def _apply_nav_calendar_theme(self, is_dark: bool):
        """套用左側上下小月曆樣式（含週標字色）。"""
        weekday_color = "#ffffff" if is_dark else "#000000"
        body_color = "#f0f0f0" if is_dark else "#111111"
        header_bg = "#363636" if is_dark else "#e6e6e6"
        calendar_bg = "#2b2b2b" if is_dark else "#ffffff"
        selected_bg = "#0078d7" if is_dark else "#9ec6f3"
        selected_fg = "#ffffff" if is_dark else "#0f1f33"
        calendar_border = "#3d3d3d" if is_dark else "transparent"

        calendar_style = f"""
            QCalendarWidget {{
                border: 1px solid {calendar_border};
            }}
            QCalendarWidget QWidget {{
                background-color: {calendar_bg};
                color: {body_color};
            }}
            QCalendarWidget QAbstractItemView:enabled {{
                selection-background-color: {selected_bg};
                selection-color: {selected_fg};
                background-color: {calendar_bg};
                color: {body_color};
                border: 1px solid {calendar_border};
            }}
            QCalendarWidget QTableView {{
                border: 1px solid {calendar_border};
            }}
            QCalendarWidget QToolButton {{
                color: {body_color};
                background-color: transparent;
                border: none;
            }}
            QCalendarWidget QTableView QHeaderView::section {{
                background-color: {header_bg};
                color: {weekday_color};
                font-weight: 600;
            }}
        """

        self.nav_calendar.setStyleSheet(calendar_style)
        self.nav_calendar_next.setStyleSheet(calendar_style)
        if isinstance(self.nav_calendar, NavCalendarWidget):
            self.nav_calendar.set_theme_dark(is_dark)
        if isinstance(self.nav_calendar_next, NavCalendarWidget):
            self.nav_calendar_next.set_theme_dark(is_dark)
        weekday_fmt = self.nav_calendar.weekdayTextFormat(Qt.Monday)
        weekday_fmt.setForeground(QColor(weekday_color))
        for day in (
            Qt.Sunday,
            Qt.Monday,
            Qt.Tuesday,
            Qt.Wednesday,
            Qt.Thursday,
            Qt.Friday,
            Qt.Saturday,
        ):
            self.nav_calendar.setWeekdayTextFormat(day, weekday_fmt)
            self.nav_calendar_next.setWeekdayTextFormat(day, weekday_fmt)
        self.label_nav_next_header.setStyleSheet(
            f"font-family: 'Segoe UI'; font-size: 16px; font-weight: 600; color: {weekday_color};"
        )

    def _apply_view_toolbar_button_style(self, is_dark: bool):
        """套用右上角工具按鈕樣式：日/週/月可高亮。"""
        if not hasattr(self, "btn_view_day"):
            return

        if is_dark:
            toggle_style = """
                QPushButton {
                    background-color: #2f3540;
                    color: #d6d6d6;
                    border: 1px solid #3d3d3d;
                    border-radius: 4px;
                    padding: 6px 14px;
                    font-weight: 600;
                    min-width: 48px;
                }
                QPushButton:hover {
                    background-color: #3a4250;
                }
                QPushButton:checked {
                    background-color: #2e7d32;
                    color: #ffffff;
                    border: 1px solid #3f9a45;
                }
            """
        else:
            toggle_style = """
                QPushButton {
                    background-color: #e9ecef;
                    color: #2f2f2f;
                    border: 1px solid #cfd3d7;
                    border-radius: 4px;
                    padding: 6px 14px;
                    font-weight: 600;
                    min-width: 48px;
                }
                QPushButton:hover {
                    background-color: #c7d4e2;
                }
                QPushButton:checked {
                    background-color: #66bb6a;
                    color: #0e2b12;
                    border: 1px solid #3f9a45;
                }
            """
        self.btn_view_schedule_list.setStyleSheet(toggle_style)
        self.btn_view_day.setStyleSheet(toggle_style)
        self.btn_view_week.setStyleSheet(toggle_style)
        self.btn_view_month.setStyleSheet(toggle_style)

        if is_dark:
            holiday_style = """
                QPushButton {
                    background-color: #0e639c;
                    color: #ffffff;
                    border: 1px solid #2a8ccd;
                    border-radius: 4px;
                    padding: 6px 14px;
                    font-weight: 600;
                    min-width: 48px;
                }
                QPushButton:hover {
                    background-color: #1f89cd;
                }
                QPushButton:pressed {
                    background-color: #094771;
                }
            """
        else:
            holiday_style = """
                QPushButton {
                    background-color: #e9ecef;
                    color: #2f2f2f;
                    border: 1px solid #cfd3d7;
                    border-radius: 4px;
                    padding: 6px 14px;
                    font-weight: 600;
                    min-width: 48px;
                }
                QPushButton:hover {
                    background-color: #c7d4e2;
                }
                QPushButton:pressed {
                    background-color: #cfd6dd;
                }
            """
        self.btn_holiday_settings.setStyleSheet(holiday_style)

    def _apply_light_theme(self):
        """套用亮色主題"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QWidget {
                color: #222222;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #d0d0d0;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 12px;
                background-color: white;
                color: #333;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 8px;
                color: #2c3e50;
            }
            QPushButton {
                background-color: #e9ecef;
                color: #111111;
                border: 1px solid #9aa4ad;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #c7d4e2;
            }
            QPushButton:pressed {
                background-color: #cfd6dd;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
            QTableWidget {
                background-color: white;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                gridline-color: #e0e0e0;
            }
            QTableWidget::item:selected {
                background-color: #9ec6f3;
                color: #0f1f33;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 8px;
                border: none;
                border-bottom: 2px solid #0078d4;
                font-weight: bold;
                color: #333;
            }
            QTextEdit {
                background-color: white;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
                color: #333;
            }
            QLabel {
                color: #333;
            }
            QMenuBar {
                background-color: #f0f0f0;
                border-bottom: 1px solid #d0d0d0;
            }
            QMenuBar::item:selected {
                background-color: #9ec6f3;
                color: #0f1f33;
            }
            QStatusBar {
                background-color: #f0f0f0;
                border-top: 1px solid #d0d0d0;
            }
            QRadioButton {
                color: #333;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #999999;
                border-radius: 9px;
                background-color: white;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0078d4;
                background-color: #f0f0f0;
            }
            QRadioButton::indicator:checked {
                background-color: #0078d4;
                border: 2px solid #0078d4;
            }
            QCheckBox {
                color: #333;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #999999;
                border-radius: 2px;
                background-color: white;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #f0f0f0;
            }
            QCheckBox::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
                image: url(:/checkbox_check);
            }
            QComboBox {
                background-color: white;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 6px;
                color: #333;
            }
            QComboBox QAbstractItemView {
                background-color: white;
                color: #333;
                selection-background-color: #9ec6f3;
                selection-color: #0f1f33;
            }
            QComboBox::drop-down {
                width: 0px;
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
            }
        """)

    def _apply_dark_theme(self):
        """套用暗色主題"""
        self.setStyleSheet("""
            QMainWindow {
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
                border: 1px solid #2a8ccd;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #1f89cd;
            }
            QPushButton:pressed {
                background-color: #094771;
            }
            QPushButton:disabled {
                background-color: #4a4a4a;
                color: #808080;
            }
            QTableWidget {
                background-color: #1e1e1e;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                gridline-color: #3d3d3d;
                color: #cccccc;
            }
            QTableWidget::item:selected {
                background-color: #094771;
                color: white;
            }
            QHeaderView::section {
                background-color: #252526;
                padding: 8px;
                border: none;
                border-bottom: 2px solid #0e639c;
                font-weight: bold;
                color: #cccccc;
            }
            QTextEdit {
                background-color: #1e1e1e;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 8px;
                color: #cccccc;
            }
            QLabel {
                color: #cccccc;
            }
            QMenuBar {
                background-color: #2b2b2b;
                border-bottom: 1px solid #3d3d3d;
                color: #cccccc;
            }
            QMenuBar::item {
                color: #cccccc;
            }
            QMenuBar::item:selected {
                background-color: #094771;
                color: white;
            }
            QMenu {
                background-color: #2b2b2b;
                border: 1px solid #3d3d3d;
                color: #cccccc;
            }
            QMenu::item:selected {
                background-color: #094771;
                color: white;
            }
            QStatusBar {
                background-color: #2b2b2b;
                border-top: 1px solid #3d3d3d;
                color: #cccccc;
            }
            QLineEdit {
                background-color: #1e1e1e;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                color: #cccccc;
            }
            QLineEdit:focus {
                border: 2px solid #0e639c;
            }
            QComboBox {
                background-color: #1e1e1e;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                color: #cccccc;
            }
            QComboBox::drop-down {
                width: 0px;
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
            }
            QSpinBox {
                background-color: #1e1e1e;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                color: #cccccc;
            }
            QCalendarWidget {
                background-color: #2b2b2b;
            }
            QCalendarWidget QTableView {
                selection-background-color: #094771;
                selection-color: white;
                background-color: #1e1e1e;
                color: #cccccc;
            }
            QCalendarWidget QWidget#qt_calendar_navigationbar {
                background-color: #0e639c;
            }
            QCalendarWidget QToolButton {
                color: white;
                background-color: transparent;
                border: none;
                font-weight: bold;
            }
            QRadioButton {
                color: #cccccc;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #666666;
                border-radius: 9px;
                background-color: #1e1e1e;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #252526;
            }
            QRadioButton::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
            }
            QCheckBox {
                color: #cccccc;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #666666;
                border-radius: 2px;
                background-color: #1e1e1e;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #252526;
            }
            QCheckBox::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
                image: url(:/checkbox_check);
            }
        """)

    def setup_theme_listener(self):
        """設定系統主題監聽"""
        # 在 Windows 上定時檢查系統主題變化
        try:
            self._theme_timer = QTimer(self)
            self._theme_timer.timeout.connect(self.check_system_theme)
            self._theme_timer.start(2000)  # 每2秒檢查一次
            self._last_theme = self.is_system_dark_mode()
        except Exception:
            pass

    def is_system_dark_mode(self) -> bool:
        """檢查系統是否使用暗色模式"""
        try:
            import winreg

            registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            key = winreg.OpenKey(registry, key_path)
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            winreg.CloseKey(key)
            return value == 0
        except Exception:
            return False

    def check_system_theme(self):
        """檢查系統主題是否變化"""
        if self.current_theme == "system":
            current_system_theme = self.is_system_dark_mode()
            if current_system_theme != self._last_theme:
                self._last_theme = current_system_theme
                self.apply_modern_style()

    def set_theme(self, theme: str):
        """設定主題模式

        Args:
            theme: "light", "dark", 或 "system"
        """
        if theme in ["light", "dark", "system"]:
            self.current_theme = theme
            self.apply_modern_style()

    def on_theme_changed(self):
        """處理主題選擇改變"""
        theme_data = self.theme_combo.currentData()
        self.set_theme(theme_data)

    def eventFilter(self, watched, event):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            month_line = self.combo_nav_month.lineEdit() if hasattr(self, "combo_nav_month") else None
            year_line = self.combo_nav_year.lineEdit() if hasattr(self, "combo_nav_year") else None
            if watched in (self.combo_nav_month, month_line):
                QTimer.singleShot(0, self.combo_nav_month.showPopup)
                event.accept()
                return True
            if watched in (self.combo_nav_year, year_line):
                QTimer.singleShot(0, self.combo_nav_year.showPopup)
                event.accept()
                return True

        if event.type() == QEvent.Wheel:
            month_line = self.combo_nav_month.lineEdit() if hasattr(self, "combo_nav_month") else None
            if watched in (self.combo_nav_month, month_line):
                steps = _combo_steps_from_wheel(event)
                if steps != 0 and self.combo_nav_month.count() > 0:
                    current_index = self.combo_nav_month.currentIndex()
                    if current_index < 0:
                        current_index = 0
                    target_index = current_index - steps
                    if target_index < 0:
                        target_index = 0
                    elif target_index >= self.combo_nav_month.count():
                        target_index = self.combo_nav_month.count() - 1
                    if target_index != self.combo_nav_month.currentIndex():
                        self.combo_nav_month.setCurrentIndex(target_index)
                event.accept()
                return True
        return super().eventFilter(watched, event)

    # ========= 主行事曆視圖相關 =========

    def _init_nav_month_year(self):
        """初始化左側導覽月曆的月份與年份下拉選單"""
        months = ["1 月", "2 月", "3 月", "4 月", "5 月", "6 月",
                  "7 月", "8 月", "9 月", "10 月", "11 月", "12 月"]
        self.combo_nav_month.clear()
        self.combo_nav_month.addItems(months)

        current_year = QDate.currentDate().year()
        self._set_nav_year_window(current_year, current_year)

        # 設定當前年/月與預設 reference_date（高亮落在今日）
        today = QDate.currentDate()
        self.reference_date = today
        # 讓 Combo 代表「左上顯示月份」= 今天所在的月份
        self.combo_nav_month.blockSignals(True)
        self.combo_nav_year.blockSignals(True)
        self.combo_nav_month.setCurrentIndex(today.month() - 1)
        year_index = self.combo_nav_year.findData(today.year())
        self.combo_nav_year.setCurrentIndex(year_index)
        self.combo_nav_month.blockSignals(False)
        self.combo_nav_year.blockSignals(False)
        # 依 reference_date 同步兩個小月曆與主行事曆
        self._update_nav_calendars(today.year(), today.month())

    def _set_nav_year_window(self, center_year: int, selected_year: Optional[int] = None):
        """設定年份下拉視窗（固定 11 個選項，中心可滑動）。"""
        start_year = center_year - 5
        years = list(range(start_year, start_year + 11))

        target_year = selected_year if isinstance(selected_year, int) else center_year
        if target_year < years[0]:
            target_year = years[0]
        elif target_year > years[-1]:
            target_year = years[-1]

        self.combo_nav_year.blockSignals(True)
        self.combo_nav_year.clear()
        for y in years:
            self.combo_nav_year.addItem(str(y), y)

        target_index = self.combo_nav_year.findData(target_year)
        if target_index >= 0:
            self.combo_nav_year.setCurrentIndex(target_index)
        self.combo_nav_year.blockSignals(False)

    def _ensure_nav_year_available(self, year: int) -> int:
        """確保年份下拉內含指定年份。"""
        index = self.combo_nav_year.findData(year)
        if index >= 0:
            return index

        self._set_nav_year_window(year, year)
        return self.combo_nav_year.findData(year)

    def _on_nav_year_wheel_step(self, steps: int):
        """滑鼠滾輪平移年份選項視窗，不直接切換行事曆。"""
        if steps == 0:
            return

        center_idx = min(5, max(0, self.combo_nav_year.count() - 1))
        center_year = self.combo_nav_year.itemData(center_idx)
        if not isinstance(center_year, int):
            center_year = QDate.currentDate().year()

        selected_year = self.combo_nav_year.currentData()
        if not isinstance(selected_year, int):
            selected_year = center_year

        self._set_nav_year_window(center_year + steps, selected_year)

    def _update_nav_calendars(self, year: int, month: int):
        """同步更新兩個導覽月曆的顯示月份"""
        self.nav_calendar.setCurrentPage(year, month)

        next_qdate = QDate(year, month, 1).addMonths(1)
        self.nav_calendar_next.setCurrentPage(next_qdate.year(), next_qdate.month())

        # 更新左下方標題顯示的「下一個月 年/月」文字
        if hasattr(self, "label_nav_next_header") and self.label_nav_next_header is not None:
            self.label_nav_next_header.setText(next_qdate.toString("yyyy年 M月"))

        # 更新兩個小月曆的高亮日期（共用同一個 reference_date）
        if isinstance(self.nav_calendar, NavCalendarWidget):
            self.nav_calendar.set_forced_selected_date(self.reference_date)
        if isinstance(self.nav_calendar_next, NavCalendarWidget):
            self.nav_calendar_next.set_forced_selected_date(self.reference_date)

    def _shift_nav_month(self, delta: int):
        """上一月 / 下一月"""
        current_year = self.combo_nav_year.currentData()
        current_month = self.combo_nav_month.currentIndex() + 1
        if not current_year:
            current_year = QDate.currentDate().year()

        qd = QDate(current_year, current_month, 1).addMonths(delta)
        self.reference_date = qd
        # 更新下拉與月曆
        self.combo_nav_month.setCurrentIndex(qd.month() - 1)
        year_index = self._ensure_nav_year_available(qd.year())
        if year_index >= 0:
            self.combo_nav_year.setCurrentIndex(year_index)
        self._update_nav_calendars(qd.year(), qd.month())
        self._refresh_main_calendar_views()

    def _on_nav_combo_changed(self):
        """月份或年份下拉改變"""
        year = self.combo_nav_year.currentData()
        month = self.combo_nav_month.currentIndex() + 1
        if not year or month <= 0:
            return
        qd = QDate(year, month, 1)
        self.reference_date = qd
        self._update_nav_calendars(year, month)
        self._refresh_main_calendar_views()

    def on_nav_calendar_date_clicked(self, qdate: QDate):
        """左側小月曆點擊某天時（Outlook 行為）：
        - 點上方白字（本月）：只換 reference_date，不改左側顯示月份
        - 點下方白字（次月）：只換 reference_date，高亮移到下方，不改左側顯示月份
        - 點上方淺灰（上個月）：左側顯示月份往回一個月（上=上個月，下=本月）
        - 點下方淺灰（下下月）：左側顯示月份往前一個月（上=本月，下=下個月）
        """
        sender = self.sender()

        base_year = self.combo_nav_year.currentData()
        base_month = self.combo_nav_month.currentIndex() + 1
        if not base_year or base_month <= 0:
            base_year = QDate.currentDate().year()
            base_month = QDate.currentDate().month()

        top_first = QDate(base_year, base_month, 1)          # 上方顯示月份
        bottom_first = top_first.addMonths(1)                 # 下方顯示月份
        after_bottom_first = top_first.addMonths(2)           # 下方再下一個月
        prev_first = top_first.addMonths(-1)                  # 上方前一個月

        # 先更新 reference_date（右側主視圖需要跟著動）
        self.reference_date = qdate

        # 依點擊來源與點到的月份決定是否要推動左側兩個月曆的顯示月份
        new_base_first = None

        if sender is self.nav_calendar:
            # 上方：
            #   本月白字：qdate.month == top_first.month -> 不動
            #   上方淺灰（上個月任何一天）：qdate.month == prev_first.month -> 往回一個月
            if qdate.year() == prev_first.year() and qdate.month() == prev_first.month():
                new_base_first = prev_first
        elif sender is self.nav_calendar_next:
            # 下方：
            #   次月白字：qdate.month == bottom_first.month -> 不動
            #   下方淺灰（下下月任何一天）：qdate.month == after_bottom_first.month -> 往前一個月（base + 1）
            if qdate.year() == after_bottom_first.year() and qdate.month() == after_bottom_first.month():
                new_base_first = bottom_first

        if new_base_first is not None:
            # 同步下拉（代表左側顯示的「上方月份」）
            self.combo_nav_month.blockSignals(True)
            self.combo_nav_year.blockSignals(True)
            self.combo_nav_month.setCurrentIndex(new_base_first.month() - 1)
            year_index = self._ensure_nav_year_available(new_base_first.year())
            if year_index >= 0:
                self.combo_nav_year.setCurrentIndex(year_index)
            self.combo_nav_month.blockSignals(False)
            self.combo_nav_year.blockSignals(False)
            self._update_nav_calendars(new_base_first.year(), new_base_first.month())
        else:
            # 不改左側月份，只更新兩個月曆高亮
            if isinstance(self.nav_calendar, NavCalendarWidget):
                self.nav_calendar.set_forced_selected_date(self.reference_date)
            if isinstance(self.nav_calendar_next, NavCalendarWidget):
                self.nav_calendar_next.set_forced_selected_date(self.reference_date)

        self._refresh_main_calendar_views()

    def _on_main_calendar_date_selected(self, qdate: QDate):
        """右側主月曆點擊日期後，同步左側導覽月曆與年月下拉。"""
        if not isinstance(qdate, QDate) or not qdate.isValid():
            return

        self.reference_date = qdate

        self.combo_nav_month.blockSignals(True)
        self.combo_nav_year.blockSignals(True)
        self.combo_nav_month.setCurrentIndex(qdate.month() - 1)
        year_index = self._ensure_nav_year_available(qdate.year())
        if year_index >= 0:
            self.combo_nav_year.setCurrentIndex(year_index)
        self.combo_nav_month.blockSignals(False)
        self.combo_nav_year.blockSignals(False)

        self._update_nav_calendars(qdate.year(), qdate.month())
        self._refresh_main_calendar_views()

    def _go_to_today_from_nav(self):
        """點擊左側『今日』按鈕時，回到今天"""
        today = QDate.currentDate()
        self.reference_date = today
        # 同步下拉（暫時關閉 signals，避免 _on_nav_combo_changed 把日期改成該月 1 號）
        self.combo_nav_month.blockSignals(True)
        self.combo_nav_year.blockSignals(True)
        self.combo_nav_month.setCurrentIndex(today.month() - 1)
        year_index = self._ensure_nav_year_available(today.year())
        if year_index >= 0:
            self.combo_nav_year.setCurrentIndex(year_index)
        self.combo_nav_month.blockSignals(False)
        self.combo_nav_year.blockSignals(False)

        # 同步導覽月曆與主行事曆（此時 reference_date 已是今天）
        self._update_nav_calendars(today.year(), today.month())
        self._refresh_main_calendar_views()

    def _on_nav_date_changed(self):
        """左側導覽月曆日期變更"""
        # 若使用 NavCalendarWidget，自訂的點擊邏輯會呼叫 on_nav_calendar_date_clicked，
        # 這裡僅作保險同步目前 Qt 的 selectedDate。
        self.reference_date = self.nav_calendar.selectedDate()
        self._update_nav_calendars(self.reference_date.year(), self.reference_date.month())
        self._refresh_main_calendar_views()

    def _set_view_mode(self, mode: str, initial: bool = False):
        """切換 Day / Week / Month / Year 視圖"""
        self.current_view_mode = mode

        self.btn_view_schedule_list.setChecked(mode == "list")
        self.btn_view_day.setChecked(mode == "day")
        self.btn_view_week.setChecked(mode == "week")
        self.btn_view_month.setChecked(mode == "month")

        if mode == "list":
            self.calendar_stack.setCurrentIndex(3)
        elif mode == "day":
            self.calendar_stack.setCurrentIndex(0)
        elif mode == "week":
            self.calendar_stack.setCurrentIndex(1)
        else:
            # "month" 視圖使用月格顯示
            self.calendar_stack.setCurrentIndex(2)

        if not initial:
            self._refresh_main_calendar_views()

        if mode == "day" and hasattr(self, "day_view"):
            self.day_view.relayout_to_viewport()
            QTimer.singleShot(0, self.day_view.relayout_to_viewport)
        elif mode == "week" and hasattr(self, "week_view"):
            self.week_view.relayout_to_viewport()
            QTimer.singleShot(0, self.week_view.relayout_to_viewport)

    def _shift_main_range(self, delta: int):
        """
        主視窗上一段 / 下一段：
        - 日視圖：delta 天
        - 週視圖：delta 週（7 天）
        - 月視圖：delta 月
        """
        if not isinstance(self.reference_date, QDate) or not self.reference_date.isValid():
            self.reference_date = QDate.currentDate()

        if self.current_view_mode == "list":
            return
        elif self.current_view_mode == "day":
            self.reference_date = self.reference_date.addDays(delta)
        elif self.current_view_mode == "week":
            self.reference_date = self.reference_date.addDays(delta * 7)
        else:
            # month 視圖
            self.reference_date = self.reference_date.addMonths(delta)

        # 同步左側導覽月曆與主視窗
        self._update_nav_calendars(self.reference_date.year(), self.reference_date.month())
        self._refresh_main_calendar_views()

    def _refresh_main_calendar_views(self):
        """依目前 view mode 與日期，更新 Day/Week/Month 行事曆內容"""
        if not self.db_manager:
            return

        if self.current_view_mode == "list":
            self.label_current_range.setText("排程參數清單")
            self._refresh_schedule_list_view()
            return

        from datetime import datetime, timedelta

        # 確保一開始就以「今天」為基準，而不是某些地方把日期重設為該月 1 號
        qd = self.reference_date if self.reference_date.isValid() else QDate.currentDate()
        start = end = None

        if self.current_view_mode == "day":
            start = datetime(qd.year(), qd.month(), qd.day())
            end = start + timedelta(days=1)
            self.label_current_range.setText(start.strftime("%Y年%m月%d日"))
        elif self.current_view_mode == "week":
            # 以週日為一週起始（配合 UI 標頭「日 一 二 三 四 五 六」）
            weekday = qd.dayOfWeek()  # 1=Mon, 7=Sun
            # 將 Sunday 視為 0，其餘 1..6 對應 Mon..Sat
            offset_from_sunday = weekday % 7
            week_start = qd.addDays(-offset_from_sunday)
            week_end = week_start.addDays(6)
            start = datetime(week_start.year(), week_start.month(), week_start.day())
            end = datetime(week_end.year(), week_end.month(), week_end.day()) + timedelta(
                days=1
            )
            self.label_current_range.setText(
                f"{week_start.toString('yyyy/MM/dd')} - {week_end.toString('yyyy/MM/dd')}"
            )
        else:
            # 月視圖（Year 模式目前沿用月視圖）
            first = QDate(qd.year(), qd.month(), 1)
            # 月格固定顯示 6x7（42 天），包含前後月的灰色日期格
            # 因此查詢範圍需覆蓋整個可見月格，而非僅當月。
            days_to_sunday = first.dayOfWeek() % 7
            grid_start = first.addDays(-days_to_sunday)
            grid_end = grid_start.addDays(42)
            start = datetime(grid_start.year(), grid_start.month(), grid_start.day())
            end = datetime(grid_end.year(), grid_end.month(), grid_end.day())
            self.label_current_range.setText(qd.toString("yyyy年 M月"))

        if start is None or end is None:
            return

        occurrences = resolve_occurrences_for_range(
            self.schedules or [],
            start,
            end,
            self.schedule_exceptions or [],
            self.holiday_entries or [],
            self.db_manager,
        )

        # Day / Week / Month 視圖同步
        self.day_view.set_reference_date(self.reference_date)
        self.day_view.set_occurrences(occurrences)
        self.day_view.relayout_to_viewport()

        self.week_view.set_reference_date(self.reference_date)
        self.week_view.set_occurrences(occurrences)
        self.week_view.relayout_to_viewport()

        self.month_view.set_reference_date(
            QDate(self.reference_date.year(), self.reference_date.month(), 1)
        )
        self.month_view.set_selected_date(self.reference_date)
        self.month_view.set_holiday_checker(self._is_holiday_qdate)
        self.month_view.set_occurrences(occurrences)

    def _is_holiday_qdate(self, qdate: QDate) -> bool:
        """判斷日期是否為假日（週末或假日設定）。"""
        if qdate.dayOfWeek() in (6, 7):
            return True

        if not self.db_manager:
            return False

        try:
            check_date = dt_date(qdate.year(), qdate.month(), qdate.day())
            return bool(self.db_manager.is_holiday_on_date(check_date))
        except Exception:
            return False

    def _refresh_schedule_list_view(self):
        """更新右側排程參數清單視圖。"""
        self.schedule_list_view.clear()

        if not self.schedules:
            empty_item = QTreeWidgetItem(self.schedule_list_view)
            empty_item.setText(0, "提示")
            empty_item.setText(1, "目前沒有排程")
            return

        for schedule in self.schedules:
            schedule_id = int(schedule.get("id", 0) or 0)
            title = str(schedule.get("task_name", "")).strip() or f"任務{schedule_id}"

            root = QTreeWidgetItem(self.schedule_list_view)
            root.setText(0, f"{title} (ID:{schedule_id})")
            root.setText(1, "")

            rrule_str = str(schedule.get("rrule_str", "") or "")
            description = self._format_schedule_description(rrule_str, schedule_id=schedule_id)
            next_exec = self._calculate_next_execution_time(schedule)
            lock_text = "是" if bool(schedule.get("lock_enabled", 0)) else "否"

            fields = [
                ("啟用", "是" if bool(schedule.get("is_enabled", 1)) else "否"),
                ("Lock", lock_text),
                ("OPC URL", str(schedule.get("opc_url", "") or "")),
                ("Node ID", str(schedule.get("node_id", "") or "")),
                ("目標值", str(schedule.get("target_value", "") or "")),
                ("資料型別", str(schedule.get("data_type", "auto") or "auto")),
                ("週期規則", description),
                ("RRULE", rrule_str),
                ("下次執行", next_exec),
                ("最後狀態", str(schedule.get("last_execution_status", "") or "")),
                ("忽略假日", "是" if bool(schedule.get("ignore_holiday", 0)) else "否"),
            ]

            for key, value in fields:
                child = QTreeWidgetItem(root)
                child.setText(0, key)
                child.setText(1, value)

            root.setExpanded(True)

        self.schedule_list_view.resizeColumnToContents(0)

    def _on_calendar_context_action(self, action: str, payload: dict):
        """處理 Day/Week/Month 視圖發出的右鍵選單動作"""
        schedule_id = payload.get("schedule_id")
        schedule_ids_raw = payload.get("schedule_ids")
        schedule_ids: List[int] = []
        if isinstance(schedule_ids_raw, list):
            for value in schedule_ids_raw:
                if isinstance(value, int):
                    schedule_ids.append(value)
        if isinstance(schedule_id, int) and schedule_id not in schedule_ids:
            schedule_ids.insert(0, schedule_id)

        date_str = payload.get("date")
        hour = payload.get("hour")
        minute = payload.get("minute")
        month_mode = bool(payload.get("month_mode", False))

        # 記錄本次操作預設小時，供新增排程時帶入 RecurrenceDialog
        self._context_default_hour = hour if isinstance(hour, int) else None
        self._context_default_minute = minute if isinstance(minute, int) else 0

        if date_str:
            try:
                y, m, d = map(int, date_str.split("-"))
                self.reference_date = QDate(y, m, d)
                # 同步左側導覽月曆
                self._update_nav_calendars(y, m)
            except Exception:
                pass

        # 月視圖新增：以目前最近且「未來」的半小時作為預設開始時間。
        if action == "new" and month_mode:
            nearest = self._nearest_future_half_hour(datetime.now())
            self._context_default_hour = nearest.hour
            self._context_default_minute = nearest.minute

            # 若使用者在「今天」新增且取整已跨到隔日，預設日期同步到隔天。
            if self.reference_date == QDate.currentDate() and nearest.date() > datetime.now().date():
                self.reference_date = self.reference_date.addDays(1)
                self._update_nav_calendars(self.reference_date.year(), self.reference_date.month())

        if action == "new":
            self.add_schedule()
        elif action == "edit":
            if schedule_id:
                # 將當前點選的日期與小時帶入，用於 RecurrenceDialog 預設時間
                self.edit_schedule(
                    schedule_id,
                    default_date=self.reference_date,
                    default_hour=self._context_default_hour,
                    default_minute=self._context_default_minute,
                )
            else:
                QMessageBox.information(self, "提示", "此日期尚未有行程可編輯。")
        elif action == "delete":
            if schedule_ids:
                if len(schedule_ids) == 1:
                    self.delete_schedule(schedule_ids[0])
                else:
                    self._delete_schedules(schedule_ids)
            else:
                QMessageBox.information(self, "提示", "此日期尚未有行程可刪除。")
        elif action == "copy":
            if schedule_ids:
                self._copy_schedules(schedule_ids)
            else:
                QMessageBox.information(self, "提示", "此日期尚未有行程可複製。")
        elif action == "paste":
            self._paste_schedule(payload)
        elif action == "drag_update":
            self._apply_drag_time_update(payload)

    def _copy_schedule(self, schedule_id: int):
        self._copy_schedules([schedule_id])

    def _copy_schedules(self, schedule_ids: List[int]):
        unique_ids: List[int] = []
        seen: set[int] = set()
        for value in schedule_ids:
            if not isinstance(value, int):
                continue
            if value in seen:
                continue
            seen.add(value)
            unique_ids.append(value)

        if not unique_ids:
            QMessageBox.warning(self, "提示", "找不到可複製的行程。")
            return

        existing_ids = {int(s.get("id", 0) or 0) for s in self.schedules}
        valid_ids = [sid for sid in unique_ids if sid in existing_ids]
        if not valid_ids:
            QMessageBox.warning(self, "提示", "找不到可複製的行程。")
            return

        self._copied_schedule_ids = valid_ids
        if len(valid_ids) == 1:
            schedule_id = valid_ids[0]
            schedule = next((s for s in self.schedules if int(s.get("id", 0) or 0) == schedule_id), None)
            title = str(schedule.get("task_name", "")).strip() if schedule else ""
            title = title or f"任務{schedule_id}"
            self.status_bar.showMessage(f"已複製行程：{title} (ID:{schedule_id})", 2500)
        else:
            self.status_bar.showMessage(f"已複製 {len(valid_ids)} 筆行程", 2500)

    def _delete_schedules(self, schedule_ids: List[int]):
        if not self.db_manager:
            QMessageBox.warning(self, "警告", "資料庫未連線")
            return

        unique_ids: List[int] = []
        seen: set[int] = set()
        for value in schedule_ids:
            if not isinstance(value, int):
                continue
            if value in seen:
                continue
            seen.add(value)
            unique_ids.append(value)

        if not unique_ids:
            return

        if len(unique_ids) == 1:
            self.delete_schedule(unique_ids[0])
            return

        reply = QMessageBox.question(
            self,
            "確認刪除",
            f"確定要刪除這 {len(unique_ids)} 筆排程嗎？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        success_count = 0
        for sid in unique_ids:
            try:
                if self.db_manager.delete_schedule(sid):
                    success_count += 1
            except Exception:
                continue

        if success_count > 0:
            self.load_schedules()
            self.status_bar.showMessage(f"已刪除 {success_count} 筆排程", 3000)
        else:
            QMessageBox.critical(self, "錯誤", "批次刪除失敗")

    def _rrule_duration_minutes(self, rrule_str: str) -> int:
        match = re.search(r"(?:^|;)DURATION=PT(\d+)M(?:;|$)", str(rrule_str or "").upper())
        if match:
            try:
                return max(1, int(match.group(1)))
            except ValueError:
                pass
        return 5

    def _rrule_time_parts(self, rrule_str: str) -> tuple[int, int]:
        byhour = re.search(r"(?:^|;)BYHOUR=(\d{1,2})(?:;|$)", str(rrule_str or "").upper())
        byminute = re.search(r"(?:^|;)BYMINUTE=(\d{1,2})(?:;|$)", str(rrule_str or "").upper())

        hour = int(byhour.group(1)) if byhour else 9
        minute = int(byminute.group(1)) if byminute else 0
        return max(0, min(23, hour)), max(0, min(59, minute))

    def _paste_schedule(self, payload: dict):
        if not self.db_manager:
            QMessageBox.warning(self, "警告", "資料庫未連線")
            return

        copied_ids = list(self._copied_schedule_ids)
        if not copied_ids:
            QMessageBox.information(self, "提示", "尚未複製任何行程。")
            return

        date_text = str(payload.get("date", "")).strip()
        if not date_text:
            QMessageBox.warning(self, "提示", "貼上位置無效。")
            return

        try:
            y, m, d = map(int, date_text.split("-"))
            target_date = QDate(y, m, d)
        except Exception:
            QMessageBox.warning(self, "提示", "貼上日期格式錯誤。")
            return

        target_hour = payload.get("hour")
        target_minute = payload.get("minute")

        created_count = 0
        for copied_id in copied_ids:
            source = next((s for s in self.schedules if int(s.get("id", 0) or 0) == copied_id), None)
            if source is None:
                continue

            source_rrule = str(source.get("rrule_str", "") or "").strip()
            if not source_rrule:
                continue

            fallback_hour, fallback_minute = self._rrule_time_parts(source_rrule)
            hour = int(target_hour) if isinstance(target_hour, int) else fallback_hour
            minute = int(target_minute) if isinstance(target_minute, int) else fallback_minute
            hour = max(0, min(23, hour))
            minute = max(0, min(59, minute))

            start_dt = datetime(target_date.year(), target_date.month(), target_date.day(), hour, minute, 0)
            end_dt = start_dt + timedelta(minutes=self._rrule_duration_minutes(source_rrule))
            pasted_rrule = self._replace_rrule_time_fields(source_rrule, start_dt, end_dt)

            source_title = str(source.get("task_name", "")).strip() or f"任務{copied_id}"
            pasted_title = f"{source_title}_複製"

            new_id = self.db_manager.add_schedule(
                task_name=pasted_title,
                opc_url=str(source.get("opc_url", "") or ""),
                node_id=str(source.get("node_id", "") or ""),
                target_value=str(source.get("target_value", "") or ""),
                data_type=str(source.get("data_type", "auto") or "auto"),
                rrule_str=pasted_rrule,
                category_id=int(source.get("category_id", 1) or 1),
                opc_security_policy=str(source.get("opc_security_policy", "None") or "None"),
                opc_security_mode=str(source.get("opc_security_mode", "None") or "None"),
                opc_username=str(source.get("opc_username", "") or ""),
                opc_password=str(source.get("opc_password", "") or ""),
                opc_timeout=int(source.get("opc_timeout", 5) or 5),
                opc_write_timeout=int(source.get("opc_write_timeout", 3) or 3),
                lock_enabled=int(source.get("lock_enabled", 0) or 0),
                is_enabled=int(source.get("is_enabled", 1) or 1),
                ignore_holiday=int(source.get("ignore_holiday", 0) or 0),
            )
            if new_id:
                created_count += 1

        if created_count <= 0:
            QMessageBox.critical(self, "錯誤", "貼上行程失敗")
            return

        self.reference_date = target_date
        self._update_nav_calendars(target_date.year(), target_date.month())
        self.load_schedules()
        self.status_bar.showMessage(f"已貼上 {created_count} 筆行程", 3000)

    def _nearest_future_half_hour(self, base_dt: datetime) -> datetime:
        """回傳嚴格晚於 base_dt 的最近半小時時間。"""
        aligned = base_dt.replace(second=0, microsecond=0)
        minutes_mod = aligned.minute % 30
        if minutes_mod == 0:
            return aligned + timedelta(minutes=30)
        return aligned + timedelta(minutes=(30 - minutes_mod))

    def _weekday_code(self, dt: datetime) -> str:
        mapping = ["MO", "TU", "WE", "TH", "FR", "SA", "SU"]
        return mapping[dt.weekday()]

    def _snap_minute_to_scale(self, minute: int, scale: int) -> int:
        scale = max(1, int(scale))
        minute = max(0, min(24 * 60, int(minute)))
        snapped = int(round(minute / scale) * scale)
        return max(0, min(24 * 60, snapped))

    def _replace_rrule_time_fields(
        self,
        rrule_str: str,
        new_start: datetime,
        new_end: datetime,
    ) -> str:
        parts = [p.strip() for p in str(rrule_str or "").split(";") if p.strip()]
        params: Dict[str, str] = {}
        has_dtstart = False
        key_order: List[str] = []

        for part in parts:
            if part.startswith("DTSTART:"):
                has_dtstart = True
                continue
            if "=" not in part:
                continue
            key, value = part.split("=", 1)
            key = key.upper()
            params[key] = value
            key_order.append(key)

        freq = params.get("FREQ", "").upper()
        params["BYHOUR"] = str(new_start.hour)
        params["BYMINUTE"] = str(new_start.minute)
        params["X-RANGE-START"] = new_start.strftime("%Y%m%d")

        duration_minutes = max(1, int((new_end - new_start).total_seconds() // 60))
        params["DURATION"] = f"PT{duration_minutes}M"

        if freq == "MONTHLY" and "BYMONTHDAY" in params and "BYSETPOS" not in params:
            params["BYMONTHDAY"] = str(new_start.day)
        elif freq == "YEARLY" and "BYMONTHDAY" in params and "BYSETPOS" not in params:
            params["BYMONTH"] = str(new_start.month)
            params["BYMONTHDAY"] = str(new_start.day)
        elif freq == "WEEKLY":
            byday = params.get("BYDAY", "")
            if byday and "," not in byday:
                params["BYDAY"] = self._weekday_code(new_start)

        result_parts: List[str] = []
        emitted_keys: set[str] = set()

        for part in parts:
            if part.startswith("DTSTART:"):
                result_parts.append(f"DTSTART:{new_start.strftime('%Y%m%dT%H%M%S')}")
                continue
            if "=" not in part:
                result_parts.append(part)
                continue

            key, _value = part.split("=", 1)
            key_upper = key.upper()
            if key_upper in params:
                result_parts.append(f"{key_upper}={params[key_upper]}")
                emitted_keys.add(key_upper)
            else:
                result_parts.append(part)

        for key in key_order:
            if key in params and key not in emitted_keys:
                result_parts.append(f"{key}={params[key]}")
                emitted_keys.add(key)

        for key in ("BYHOUR", "BYMINUTE", "X-RANGE-START", "DURATION"):
            if key not in emitted_keys:
                result_parts.append(f"{key}={params[key]}")
                emitted_keys.add(key)

        if not has_dtstart:
            result_parts.append(f"DTSTART:{new_start.strftime('%Y%m%dT%H%M%S')}")

        return ";".join(result_parts)

    def _apply_drag_time_update(self, payload: dict):
        """套用視圖拖曳後的時間調整（移動/改開始/改結束）。"""
        schedule_id = payload.get("schedule_id")
        if not isinstance(schedule_id, int):
            return

        target_schedule = next((s for s in self.schedules if int(s.get("id", 0) or 0) == schedule_id), None)
        if target_schedule is None or not self.db_manager:
            return

        scale_minutes = payload.get("scale_minutes")
        if not isinstance(scale_minutes, int) or scale_minutes <= 0:
            scale_minutes = max(1, int(getattr(self.day_view, "time_scale_minutes", 60)))

        source = str(payload.get("source", "")).strip().lower()
        new_start: Optional[datetime] = None
        new_end: Optional[datetime] = None

        if source == "month_grid":
            start_text = str(payload.get("start_datetime", "")).strip()
            end_text = str(payload.get("end_datetime", "")).strip()
            try:
                new_start = datetime.strptime(start_text, "%Y-%m-%d %H:%M:%S")
                new_end = datetime.strptime(end_text, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                return
        else:
            date_text = str(payload.get("date", "")).strip()
            start_minute = payload.get("start_minute")
            end_minute = payload.get("end_minute")
            if not date_text or not isinstance(start_minute, int) or not isinstance(end_minute, int):
                return

            try:
                base_date = datetime.strptime(date_text, "%Y-%m-%d")
            except ValueError:
                return

            snapped_start = self._snap_minute_to_scale(start_minute, scale_minutes)
            snapped_end = self._snap_minute_to_scale(end_minute, scale_minutes)
            if snapped_end <= snapped_start:
                snapped_end = min(24 * 60, snapped_start + scale_minutes)

            new_start = base_date + timedelta(minutes=snapped_start)
            new_end = base_date + timedelta(minutes=snapped_end)

        if new_start is None or new_end is None or new_end <= new_start:
            return

        old_rrule = str(target_schedule.get("rrule_str", "") or "").strip()
        if not old_rrule:
            return

        new_rrule = self._replace_rrule_time_fields(old_rrule, new_start, new_end)
        if new_rrule == old_rrule:
            return

        success = self.db_manager.update_schedule(schedule_id=schedule_id, rrule_str=new_rrule)
        if not success:
            self.status_bar.showMessage("拖曳調整失敗：無法寫入排程", 4000)
            return

        self.reference_date = QDate(new_start.year, new_start.month, new_start.day)
        self.load_schedules()
        self._restart_scheduler_worker()
        self.status_bar.showMessage("已套用拖曳調整", 2500)

    def init_database(self):
        """初始化資料庫連線"""
        try:
            # 初始化 SQLite 管理器（使用預設資料庫路徑）
            self.db_manager = SQLiteManager()

            # 建立資料表
            if self.db_manager.init_db():
                self.db_status_label.setText("資料庫: 已連線")
                self.db_status_label.setStyleSheet("color: green;")

                self._load_time_scale_from_db()

                self.load_schedules()
                self.start_scheduler()
            else:
                self.db_status_label.setText("資料庫: 資料表建立失敗")
                self.db_status_label.setStyleSheet("color: red;")

        except Exception as e:
            self.db_status_label.setText("資料庫: 連線失敗")
            self.db_status_label.setStyleSheet("color: red;")
            self.db_manager = None
            QMessageBox.warning(
                self,
                "資料庫連線失敗",
                f"無法連線到資料庫:\n{str(e)}\n\n請檢查資料庫設定。",
            )

    def load_schedules(self):
        """載入排程列表"""
        if not self.db_manager:
            return

        self.schedules = self.db_manager.get_all_schedules()
        self.schedule_exceptions = self.db_manager.get_all_schedule_exceptions()
        self.holiday_entries = self.db_manager.get_all_holiday_entries()
        # 重置執行計數器（應用程式重啟時從 0 開始）
        self.execution_counts = {}

        # 更新主行事曆視圖
        self._refresh_main_calendar_views()
        self.status_bar.showMessage(f"已載入 {len(self.schedules)} 個排程")

    def _on_general_settings_changed(self):
        """全局設定變更時的處理"""
        self.status_bar.showMessage("全局設定已更新")
        # 若未來需要，可在此讀取 general_settings 調整排程行為

    def _load_time_scale_from_db(self):
        """從資料庫載入日/週視圖時間刻度。"""
        if not self.db_manager:
            return

        minutes = self.db_manager.get_time_scale_minutes()
        self.day_view.set_time_scale(minutes)
        self.week_view.set_time_scale(minutes)
        if hasattr(self, "month_view"):
            self.month_view.set_time_scale(minutes)

    def _on_time_scale_changed(self, minutes: int):
        """同步日/週視圖刻度並保存到資料庫。"""
        sender = self.sender()
        if sender is self.day_view and self.week_view.time_scale_minutes != minutes:
            self.week_view.set_time_scale(minutes)
        elif sender is self.week_view and self.day_view.time_scale_minutes != minutes:
            self.day_view.set_time_scale(minutes)

        if hasattr(self, "month_view"):
            self.month_view.set_time_scale(minutes)

        if self.db_manager:
            self.db_manager.save_time_scale_minutes(minutes)

    def _format_schedule_description(self, rrule_str: str, schedule_id: int = 0) -> str:
        """將 RRULE 轉換為中文簡易說明"""
        if not rrule_str:
            return "未設定"

        try:
            # 簡單的 RRULE 解析，轉換為中文說明
            parts = rrule_str.upper().split(';')
            freq_map = {
                'DAILY': '每天',
                'WEEKLY': '每週',
                'MONTHLY': '每月',
                'YEARLY': '每年',
                'HOURLY': '每小時',
                'MINUTELY': '每分鐘',
                'SECONDLY': '每秒'
            }
            
            freq = ""
            interval = 1
            byday = ""
            bymonthday = ""
            byhour = ""
            byminute = ""
            bysecond = ""
            count = ""
            until = ""
            bymonth = ""
            bysetpos = ""
            is_lunar = False
            
            for part in parts:
                if part.startswith('FREQ='):
                    freq_code = part.split('=')[1]
                    freq = freq_map.get(freq_code, freq_code)
                elif part.startswith('INTERVAL='):
                    interval = int(part.split('=')[1])
                elif part.startswith('BYDAY='):
                    byday = part.split('=')[1]
                elif part.startswith('BYMONTHDAY='):
                    bymonthday = part.split('=')[1]
                elif part.startswith('BYHOUR='):
                    byhour = part.split('=')[1]
                elif part.startswith('BYMINUTE='):
                    byminute = part.split('=')[1]
                elif part.startswith('BYSECOND='):
                    bysecond = part.split('=')[1]
                elif part.startswith('COUNT='):
                    count = part.split('=')[1]
                elif part.startswith('UNTIL='):
                    until = part.split('=')[1]
                elif part.startswith('BYMONTH='):
                    bymonth = part.split('=')[1]
                elif part.startswith('BYSETPOS='):
                    bysetpos = part.split('=')[1]
                elif part.startswith('X-LUNAR='):
                    is_lunar = part.split('=')[1] == '1'
            
            # 如果有 COUNT，計算剩餘次數
            if count and schedule_id:
                try:
                    original_count = int(count)
                    executed_count = self.execution_counts.get(schedule_id, 0)
                    remaining_count = max(0, original_count - executed_count)
                    count = str(remaining_count)
                except ValueError:
                    pass  # 如果解析失敗，保持原樣
            
            # 生成中文描述
            desc_parts = []
            
            # 頻率部分
            if interval > 1:
                desc_parts.append(f"每{interval}{freq[1:]}")  # 每3週
            else:
                desc_parts.append(freq)  # 每天
            
            # 範圍部分
            range_desc = ""
            if bymonth and bymonthday:
                # X月Y日
                month_map = {
                    '1': '一月', '2': '二月', '3': '三月', '4': '四月', '5': '五月', '6': '六月',
                    '7': '七月', '8': '八月', '9': '九月', '10': '十月', '11': '十一月', '12': '十二月'
                }
                month_name = month_map.get(bymonth, f"{bymonth}月")
                range_desc = f"{month_name}{bymonthday}日"
            elif bysetpos and byday:
                # X月的 第Y個 Z
                if bymonth:
                    month_map = {
                        '1': '一月', '2': '二月', '3': '三月', '4': '四月', '5': '五月', '6': '六月',
                        '7': '七月', '8': '八月', '9': '九月', '10': '十月', '11': '十一月', '12': '十二月'
                    }
                    month_name = month_map.get(bymonth, f"{bymonth}月")
                    pos_map = {'1': '第1個', '2': '第2個', '3': '第3個', '4': '第4個', '5': '第5個', '-1': '最後1個'}
                    pos = pos_map.get(bysetpos, f'第{bysetpos}個')
                    day_map = {
                        'MO': '週一', 'TU': '週二', 'WE': '週三', 'TH': '週四',
                        'FR': '週五', 'SA': '週六', 'SU': '週日'
                    }
                    days = [day_map.get(day, day) for day in byday.split(',')]
                    range_desc = f"{month_name}的 {pos} {','.join(days)}"
                else:
                    # 每月的 第Y個 Z
                    pos_map = {'1': '第1個', '2': '第2個', '3': '第3個', '4': '第4個', '5': '第5個', '-1': '最後1個'}
                    pos = pos_map.get(bysetpos, f'第{bysetpos}個')
                    day_map = {
                        'MO': '週一', 'TU': '週二', 'WE': '週三', 'TH': '週四',
                        'FR': '週五', 'SA': '週六', 'SU': '週日'
                    }
                    days = [day_map.get(day, day) for day in byday.split(',')]
                    if len(days) == 5 and set(days) == {'週一', '週二', '週三', '週四', '週五'}:
                        range_desc = f"{pos} 週一到週五"
                    else:
                        range_desc = f"{pos} {','.join(days)}"
            elif byday:
                # 星期幾
                day_map = {
                    'MO': '一', 'TU': '二', 'WE': '三', 'TH': '四',
                    'FR': '五', 'SA': '六', 'SU': '日'
                }
                days = [day_map.get(day, day) for day in byday.split(',')]
                if len(days) == 5 and set(days) == {'一', '二', '三', '四', '五'}:
                    range_desc = "工作天"
                else:
                    range_desc = "".join(days)
            elif bymonthday:
                range_desc = f"第{bymonthday}天"
            
            if range_desc:
                desc_parts.append(range_desc)
            
            # 時間部分
            time_parts = []
            if byhour:
                time_parts.append(byhour)
            if byminute:
                time_parts.append(byminute)
            if bysecond:
                time_parts.append(bysecond)
            
            if time_parts:
                # 根據頻率和可用參數決定時間顯示格式
                if byhour:
                    # 有小時信息，使用完整的時間格式
                    hour = int(byhour)
                    minute = int(byminute) if byminute else 0
                    second = int(bysecond) if bysecond else 0
                    
                    if hour < 12:
                        time_str = f"上午{hour}:{minute:02d}"
                    elif hour == 12:
                        time_str = f"中午{hour}:{minute:02d}"
                    else:
                        time_str = f"下午{hour-12}:{minute:02d}"
                    
                    # 如果有秒數參數，總是顯示秒數
                    if bysecond is not None:
                        time_str += f":{second:02d}"
                else:
                    # 沒有小時信息，只顯示分鐘和秒鐘
                    minute = int(byminute) if byminute else 0
                    second = int(bysecond) if bysecond else 0
                    
                    if freq_code in ['MINUTELY', 'SECONDLY']:
                        # 對於分鐘級或秒級頻率，顯示相對時間
                        if byminute and bysecond:
                            time_str = f"第{minute}分第{second}秒"
                        elif bysecond:
                            time_str = f"第{second}秒"
                        else:
                            time_str = f"第{minute}分"
                    else:
                        # 其他情況顯示絕對時間
                        time_str = f"{minute:02d}:{second:02d}"
                
                desc_parts.append(time_str)
            
            # 結束條件
            end_desc = ""
            if count:
                end_desc = f"剩餘{count}次之後結束"
            elif until:
                # 格式化日期，假設是 YYYYMMDD
                if len(until) >= 8:
                    year = until[:4]
                    month = until[4:6].lstrip('0')  # 移除前導零
                    day = until[6:8].lstrip('0')    # 移除前導零
                    end_desc = f"結束於{year}/{month}/{day}"
            
            if end_desc:
                desc_parts.append(end_desc)
            
            description = " ".join(desc_parts)
            if is_lunar:
                description = f"[農曆] {description}"
            return description
            
        except Exception:
            return rrule_str  # 如果解析失敗，返回原始字串

    def _format_node_name(self, node_id: str) -> str:
        """從 Node ID 提取節點名稱進行顯示"""
        if not node_id:
            return ""

        try:
            # 檢查是否包含 display_name|node_id 格式
            if "|" in node_id:
                parts = node_id.split("|")
                if len(parts) >= 2:
                    display_name = parts[0]
                    actual_node_id = parts[1]
                    # 優先從 actual_node_id 提取完整路徑
                    if actual_node_id.startswith("ns="):
                        ns_parts = actual_node_id.split(";")
                        if len(ns_parts) > 1:
                            last_part = ns_parts[-1]
                            if last_part.startswith("s="):
                                full_path = last_part[2:]
                                if full_path:
                                    return full_path
                    # 如果無法提取完整路徑，返回 display_name
                    return display_name
            
            # 處理 OPC UA NodeId 的字串表示
            if node_id.startswith("NodeId("):
                # 提取 Identifier 部分
                import re
                match = re.search(r"Identifier='([^']+)'", node_id)
                if match:
                    return match.group(1)
            
            # 特殊處理某些 OPC UA 實現的 Node ID 格式
            if node_id.startswith("String: "):
                identifier = node_id[7:]  # 移除 "String: " 前綴
                # 如果 identifier 是簡單的數字或短字串，嘗試提供更好的名稱
                if identifier == "3>":
                    return "Delta_42_1F.HPW1.DT1"
                return identifier
            elif node_id.startswith("Numeric: "):
                identifier = node_id[8:]  # 移除 "Numeric: " 前綴
                return f"Node_{identifier}"
            
            # 如果是標準 OPC UA Node ID 格式，提取最後一部分
            if node_id.startswith("ns="):
                # 格式如: ns=2;s=MyVariable 或 ns=2;i=12345
                parts = node_id.split(";")
                if len(parts) > 1:
                    last_part = parts[-1]
                    if last_part.startswith("s="):
                        return last_part[2:]  # 移除 "s=" 前綴
                    elif last_part.startswith("i="):
                        return f"Node_{last_part[2:]}"  # 數值 ID 轉換為可讀格式
                    else:
                        return last_part
            
            # 如果不是標準格式，返回最後一個點號之後的部分
            if "." in node_id:
                return node_id.split(".")[-1]
            
            # 如果都沒有特殊格式，返回原字串
            return node_id
            
        except Exception:
            return node_id  # 如果處理失敗，返回原始字串

    def _calculate_next_execution_time(self, schedule: Dict[str, Any]) -> str:
        """計算下次執行時間"""
        rrule_str = schedule.get("rrule_str", "")
        if not rrule_str:
            return "未設定"
        
        try:
            # 使用 RRuleParser 計算下次執行時間
            next_time = RRuleParser.get_next_trigger(rrule_str)
            
            # 檢查 UNTIL 過期
            parts = rrule_str.upper().split(';')
            until_expired = False
            for part in parts:
                if part.startswith('UNTIL='):
                    until_value = part.split('=', 1)[1]
                    if len(until_value) >= 8:
                        try:
                            year = int(until_value[:4])
                            month = int(until_value[4:6])
                            day = int(until_value[6:8])
                            until_date = datetime(year, month, day).date()
                            today = datetime.now().date()
                            if until_date < today:
                                until_expired = True
                        except ValueError:
                            pass
                    break
            
            if next_time:
                # 格式化時間顯示
                time_str = next_time.strftime("%Y/%m/%d %H:%M:%S")
                return time_str
            else:
                # 沒有下次執行時間，檢查是否因為過期
                if until_expired:
                    return "已過期"
                else:
                    return "已結束"
                
        except Exception:
            return "計算失敗"



    def add_schedule(self):
        """新增排程（可由主行事曆帶入預設日期/時間）"""
        # 嘗試從目前 reference_date 與右鍵 payload 帶入時間
        default_date = getattr(self, "reference_date", QDate.currentDate())
        default_hour = getattr(self, "_context_default_hour", None)
        default_minute = getattr(self, "_context_default_minute", 0)

        dialog = ScheduleEditDialog(
            self,
            default_date=default_date,
            default_hour=default_hour,
            default_minute=default_minute,
        )
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()

            if self.db_manager:
                schedule_id = self.db_manager.add_schedule(
                    task_name=data["task_name"],
                    opc_url=data["opc_url"],
                    node_id=data["node_id"],
                    target_value=data["target_value"],
                    data_type=data.get("data_type", "auto"),
                    rrule_str=data["rrule_str"],
                    category_id=data.get("category_id", 1),
                    opc_security_policy=data.get("opc_security_policy", "None"),
                    opc_security_mode=data.get("opc_security_mode", "None"),
                    opc_username=data.get("opc_username", ""),
                    opc_password=data.get("opc_password", ""),
                    opc_timeout=data.get("opc_timeout", 5),
                    opc_write_timeout=data.get("opc_write_timeout", 3),
                    lock_enabled=data.get("lock_enabled", 0),
                    is_enabled=data.get("is_enabled", 1),
                    ignore_holiday=data.get("ignore_holiday", 0),
                )

                if schedule_id:
                    QMessageBox.information(self, "成功", "排程已新增")
                    self.load_schedules()
                else:
                    QMessageBox.critical(self, "錯誤", "新增排程失敗")

    def edit_schedule(
        self,
        schedule_id: int = None,
        default_date: QDate | None = None,
        default_hour: int | None = None,
        default_minute: int = 0,
    ):
        """編輯排程（可帶入目前點選的日期/時間，供週期對話框預設使用）"""
        if schedule_id is None:
            QMessageBox.information(self, "提示", "請先選擇要編輯的排程")
            return
        
        schedule = next((s for s in self.schedules if s['id'] == schedule_id), None)
        if not schedule:
            return

        dialog = ScheduleEditDialog(
            self,
            schedule,
            default_date=default_date,
            default_hour=default_hour,
            default_minute=default_minute,
        )
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()

            if self.db_manager:
                success = self.db_manager.update_schedule(
                    schedule_id=schedule["id"],
                    task_name=data["task_name"],
                    opc_url=data["opc_url"],
                    node_id=data["node_id"],
                    target_value=data["target_value"],
                    data_type=data.get("data_type", "auto"),
                    rrule_str=data["rrule_str"],
                    category_id=data.get("category_id", 1),
                    opc_security_policy=data.get("opc_security_policy", "None"),
                    opc_security_mode=data.get("opc_security_mode", "None"),
                    opc_username=data.get("opc_username", ""),
                    opc_password=data.get("opc_password", ""),
                    opc_timeout=data.get("opc_timeout", 5),
                    opc_write_timeout=data.get("opc_write_timeout", 3),
                    lock_enabled=data.get("lock_enabled", 0),
                    is_enabled=data.get("is_enabled", 1),
                    ignore_holiday=data.get("ignore_holiday", 0),
                )

                if success:
                    QMessageBox.information(self, "成功", "排程已更新")
                    self.load_schedules()
                    self._restart_scheduler_worker()
                else:
                    QMessageBox.critical(self, "錯誤", "更新排程失敗")

    def delete_schedule(self, schedule_id: int = None):
        """刪除排程"""
        if schedule_id is None:
            QMessageBox.information(self, "提示", "請先選擇要刪除的排程")
            return
        
        schedule = next((s for s in self.schedules if s['id'] == schedule_id), None)
        if not schedule:
            return

        reply = QMessageBox.question(
            self,
            "確認刪除",
            f"確定要刪除排程 '{schedule.get('task_name')}' 嗎？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )

        if reply == QMessageBox.Yes:
            if self.db_manager:
                success = self.db_manager.delete_schedule(schedule["id"])

                if success:
                    QMessageBox.information(self, "成功", "排程已刪除")
                    self.load_schedules()
                else:
                    QMessageBox.critical(self, "錯誤", "刪除排程失敗")

    def edit_selected_schedule(self):
        """編輯目前選取的排程"""
        target_id = self.selected_schedule_id
        if target_id is None:
            target_id = self._pick_schedule_id("選擇要編輯的排程")
            if target_id is None:
                return
        self.edit_schedule(target_id)

    def delete_selected_schedule(self):
        """刪除目前選取的排程"""
        target_id = self.selected_schedule_id
        if target_id is None:
            target_id = self._pick_schedule_id("選擇要刪除的排程")
            if target_id is None:
                return
        self.delete_schedule(target_id)

    def _pick_schedule_id(self, title: str) -> Optional[int]:
        """從現有排程中挑選一筆 ID。"""
        if not self.schedules:
            QMessageBox.information(self, "提示", "目前沒有可操作的排程")
            return None

        items: List[str] = []
        mapping: Dict[str, int] = {}
        for schedule in self.schedules:
            schedule_id = int(schedule.get("id", 0))
            task_name = str(schedule.get("task_name", "未命名"))
            item_text = f"[{schedule_id}] {task_name}"
            items.append(item_text)
            mapping[item_text] = schedule_id

        selected_text, ok = QInputDialog.getItem(
            self,
            title,
            "排程:",
            items,
            0,
            False,
        )
        if not ok or not selected_text:
            return None

        return mapping.get(selected_text)

    def refresh_schedules(self):
        """重新載入排程資料 (F5)"""
        if not self.db_manager:
            QMessageBox.warning(self, "警告", "資料庫未連線")
            return
        
        self.status_bar.showMessage("正在重新載入排程...")
        self.load_schedules()
        self.status_bar.showMessage(f"已重新載入 {len(self.schedules)} 個排程", 3000)

    def apply_schedules(self):
        """套用排程變更"""
        if not self.db_manager:
            QMessageBox.warning(self, "警告", "資料庫未連線")
            return

        # 目前所有變更皆即時寫入，此處仍提供一致的操作入口
        self.load_schedules()
        self.status_bar.showMessage("已套用排程變更", 3000)

    def load_project_database(self):
        """開啟既有專案資料庫（.db）。"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "開啟 Project 資料庫",
            "",
            "SQLite 資料庫 (*.db);;所有檔案 (*)",
        )

        if not file_path:
            return

        if not os.path.exists(file_path):
            QMessageBox.warning(self, "檔案不存在", f"找不到資料庫檔案:\n{file_path}")
            return

        self.on_database_path_changed(file_path)
        self.status_bar.showMessage(f"已開啟 Project DB: {file_path}", 5000)
        QMessageBox.information(self, "開啟完成", f"已開啟 Project 資料庫:\n{file_path}")

    def show_database_settings(self):
        """顯示資料庫設定對話框"""
        if not self.db_manager:
            self.init_database()

        dialog = DatabaseSettingsDialog(self, self.db_manager)
        dialog.database_changed.connect(self.on_database_path_changed)
        dialog.exec()

    def show_holiday_settings(self):
        """顯示假日設定對話框。"""
        if not self.db_manager:
            QMessageBox.warning(self, "警告", "資料庫未連線")
            return

        dialog = HolidaySettingsDialog(self.db_manager, self)
        if dialog.exec() == QDialog.Accepted:
            self.holiday_entries = self.db_manager.get_all_holiday_entries()
            self._refresh_main_calendar_views()
            self._restart_scheduler_worker()
            self.status_bar.showMessage("假日設定已更新", 3000)

    def new_project(self):
        """建立新的專案資料庫（.db）。"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "建立新的 Project 資料庫",
            "calendarua_project.db",
            "SQLite 資料庫 (*.db);;所有檔案 (*)",
        )

        if not file_path:
            return

        if not file_path.lower().endswith(".db"):
            file_path = f"{file_path}.db"

        if os.path.exists(file_path):
            reply = QMessageBox.question(
                self,
                "檔案已存在",
                f"資料庫已存在，是否覆蓋重建？\n\n{file_path}",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return
            try:
                os.remove(file_path)
            except OSError as e:
                QMessageBox.critical(self, "錯誤", f"無法覆蓋既有檔案:\n{str(e)}")
                return

        self.on_database_path_changed(file_path)
        self.status_bar.showMessage(f"已建立新 Project DB: {file_path}", 5000)
        QMessageBox.information(self, "建立完成", f"已建立新的 Project 資料庫:\n{file_path}")

    def show_about(self):
        """顯示關於對話框"""
        about_text = """
        <h2>CalendarUA</h2>
        <p>版本: 1.0.0</p>
        <p>工業自動化排程管理系統</p>
        <p>採用 PySide6 開發，支援 OPC UA 通訊</p>
        <br>
        <p><b>主要功能:</b></p>
        <ul>
            <li>週期性排程管理 (RRULE)</li>
            <li>例外與假日支援</li>
            <li>General / Holidays / Exceptions 分頁</li>
            <li>Category 分類系統</li>
            <li>OPC UA 任務執行與監控</li>
        </ul>
        <br>
        <p>© 2026 CalendarUA Project</p>
        """
        QMessageBox.about(self, "關於 CalendarUA", about_text)


    def start_scheduler(self):
        """啟動排程背景工作"""
        if self.db_manager:
            self.scheduler_worker = SchedulerWorker(self.db_manager)
            self.scheduler_worker.trigger_task.connect(self.on_task_triggered)
            self.scheduler_worker.start()
            self.status_bar.showMessage("排程器已啟動")

    def _restart_scheduler_worker(self):
        """重啟排程背景工作，讓設定變更立即生效。"""
        if not self.db_manager:
            return

        if self.scheduler_worker:
            self.scheduler_worker.stop()
            self.scheduler_worker.wait()

        self.scheduler_worker = SchedulerWorker(self.db_manager)
        self.scheduler_worker.trigger_task.connect(self.on_task_triggered)
        self.scheduler_worker.start()
        self.status_bar.showMessage("排程器已重新載入設定", 3000)

    @Slot(dict)
    def on_task_triggered(self, schedule: Dict[str, Any]):
        """處理排程觸發"""
        schedule_id = schedule.get("id")
        trigger_time = schedule.get("_trigger_time")
        if not isinstance(trigger_time, datetime):
            trigger_time = datetime.now()
        
        # 檢查是否已經在執行，防止重複執行
        if schedule_id in self.running_tasks:
            self.status_bar.showMessage(f"任務 {schedule.get('task_name', '')} 已在執行中，跳過此次觸發", 3000)
            return
        
        # 標記為執行中
        self.running_tasks.add(schedule_id)
        
        self.status_bar.showMessage(f"執行排程: {schedule.get('task_name', '')}")

        # 執行 OPC UA 寫入
        asyncio.create_task(self.execute_task(schedule, trigger_time=trigger_time))

    def _extract_actual_node_id(self, node_id: Any) -> str:
        """解析 node_id，提取實際的 OPC UA Node ID。"""
        import re

        node_id_text = str(node_id).strip()

        std_nodeid_match = re.search(r"(ns=\d+;(?:s|i|g|b)=[^|\s]+)", node_id_text)
        if std_nodeid_match:
            return std_nodeid_match.group(1).strip()

        if "|" in node_id_text:
            return node_id_text.split("|")[-1].strip()

        match = re.search(r"Identifier='([^']+)'", node_id_text)
        if match:
            identifier = match.group(1)
            ns_match = re.search(r"NamespaceIndex=(\d+)", node_id_text)
            type_match = re.search(r"NodeIdType=<NodeIdType\.(\w+):", node_id_text)
            if ns_match and type_match:
                ns = ns_match.group(1)
                node_type = type_match.group(1)
                if node_type == "String":
                    return f"ns={ns};s={identifier}"
                if node_type == "Numeric":
                    return f"ns={ns};i={identifier}"
            return identifier

        return node_id_text

    def _get_runtime_target_value(self, runtime_schedule: Dict[str, Any], base_target_value: Any, now: datetime) -> str:
        """依據最新假日設定解析本次實際要寫入的 target value。"""
        resolved = str(base_target_value)

        if not self.db_manager:
            return resolved

        if bool(runtime_schedule.get("ignore_holiday", 0)):
            return resolved

        holiday_rule = self.db_manager.is_holiday_on_date(now.date())
        if holiday_rule and holiday_rule.get("override_target_value") not in (None, ""):
            return str(holiday_rule.get("override_target_value"))

        return resolved

    async def execute_task(self, schedule: Dict[str, Any], trigger_time: Optional[datetime] = None):
        """執行排程任務"""
        schedule_id = schedule.get("id")
        effective_trigger_time = (trigger_time or datetime.now()).replace(microsecond=0)

        handler: Optional[OPCHandler] = None
        connection_signature: Optional[tuple] = None

        try:
            # 更新狀態為執行中
            if self.db_manager:
                self.db_manager.update_execution_status(schedule_id, "執行中...")
            
            # 重新載入表格以顯示狀態更新
            self.load_schedules()

            attempt = 0
            success_once = False
            lock_poll_interval = 1
            lock_enabled = bool(schedule.get("lock_enabled", 0))
            duration_minutes = self._parse_duration_from_rrule(schedule.get("rrule_str", ""))
            status_msg = ""

            while True:
                runtime_schedule = schedule
                if self.db_manager:
                    latest_schedule = self.db_manager.get_schedule(int(schedule_id))
                    if latest_schedule:
                        runtime_schedule = latest_schedule

                if not bool(runtime_schedule.get("is_enabled", 1)):
                    status_msg = "排程已停用，停止執行"
                    break

                opc_url = runtime_schedule.get("opc_url", "")
                node_id = runtime_schedule.get("node_id", "")
                target_value = runtime_schedule.get("target_value", "")
                data_type = runtime_schedule.get("data_type", "auto")
                lock_enabled = bool(runtime_schedule.get("lock_enabled", 0))
                security_policy = runtime_schedule.get("opc_security_policy", "None")
                username = runtime_schedule.get("opc_username", "")
                password = runtime_schedule.get("opc_password", "")
                timeout = int(runtime_schedule.get("opc_timeout", 5) or 5)
                write_timeout = int(runtime_schedule.get("opc_write_timeout", 3) or 3)
                retry_delay = max(1, write_timeout)
                duration_minutes = self._parse_duration_from_rrule(runtime_schedule.get("rrule_str", ""))
                period_end_time = effective_trigger_time + timedelta(minutes=duration_minutes)
                actual_node_id = self._extract_actual_node_id(node_id)

                now = datetime.now()
                runtime_target_value = self._get_runtime_target_value(runtime_schedule, target_value, now)
                window_expired = duration_minutes > 0 and now >= period_end_time

                if lock_enabled and window_expired:
                    break

                if not lock_enabled:
                    if duration_minutes == 0 and attempt > 0:
                        break
                    if duration_minutes > 0 and window_expired and attempt > 0:
                        break

                new_signature = (
                    opc_url,
                    security_policy,
                    username,
                    password,
                    timeout,
                )

                if handler is None or connection_signature != new_signature or not handler.is_connected:
                    if handler and handler.is_connected:
                        await handler.disconnect()

                    handler = OPCHandler(opc_url, timeout=timeout)
                    if security_policy != "None":
                        handler.security_policy = security_policy
                    if username:
                        handler.username = username
                        handler.password = password

                    connected = await handler.connect()
                    if not connected:
                        status_msg = f"✗ 無法連線 OPC UA: {opc_url}"
                        if duration_minutes == 0 and not lock_enabled:
                            break
                        if duration_minutes > 0 and datetime.now() >= period_end_time:
                            break
                        await asyncio.sleep(retry_delay)
                        continue

                    connection_signature = new_signature

                attempt += 1

                try:
                    if lock_enabled and duration_minutes > 0 and success_once:
                        current_value = await handler.read_node(actual_node_id, suppress_errors=True)
                        if current_value is None:
                            status_msg = f"Lock 監控讀值暫時失敗，改以寫入維持鎖定: {node_id}"
                            logger.info("Lock 監控讀值暫時失敗，改以寫入維持鎖定")
                        elif self._is_target_value_matched(current_value, runtime_target_value, data_type):
                            status_msg = f"Lock 生效中，{node_id} 值正常，持續監控到 {period_end_time.strftime('%H:%M:%S')}"
                            await asyncio.sleep(lock_poll_interval)
                            continue

                    success = await handler.write_node(actual_node_id, runtime_target_value, data_type)
                    if success:
                        success_once = True

                        if not lock_enabled:
                            status_msg = f"✓ 成功寫入 {node_id} = {runtime_target_value}"
                            break

                        if duration_minutes <= 0:
                            status_msg = f"✓ 成功寫入 {node_id} = {runtime_target_value}"
                            break

                        status_msg = f"Lock 生效中，持續鎖定 {node_id} 到 {period_end_time.strftime('%H:%M:%S')}"
                        await asyncio.sleep(lock_poll_interval)
                        continue

                    status_msg = f"✗ 寫入失敗: {node_id}"

                    if duration_minutes == 0 and not lock_enabled:
                        break

                    if duration_minutes > 0 and datetime.now() >= period_end_time:
                        break

                    logger.warning(f"寫入失敗，等待 {retry_delay} 秒後重試")
                    await asyncio.sleep(retry_delay)
                    continue

                except Exception as e:
                    status_msg = f"✗ 執行錯誤: {str(e)[:50]}"

                    if duration_minutes == 0 and not lock_enabled:
                        break

                    if duration_minutes > 0 and datetime.now() >= period_end_time:
                        break

                    logger.warning(f"寫入錯誤: {e}，等待 {retry_delay} 秒後重試")
                    await asyncio.sleep(retry_delay)
                    continue

            if success_once:
                if self.db_manager:
                    if lock_enabled and duration_minutes > 0:
                        self.db_manager.update_execution_status(schedule_id, "鎖定期間完成")
                    else:
                        self.db_manager.update_execution_status(schedule_id, "執行成功")

                # 增加執行計數器
                self.execution_counts[schedule_id] = self.execution_counts.get(schedule_id, 0) + 1
                # 檢查是否達到 COUNT 上限
                self._check_and_disable_if_count_reached(schedule_id, schedule.get("rrule_str", ""))
            else:
                if self.db_manager:
                    self.db_manager.update_execution_status(schedule_id, "寫入失敗")
                if not status_msg:
                    status_msg = "✗ 寫入失敗"
                        
        except Exception as e:
            status_msg = f"✗ 執行錯誤: {str(e)}"
            if self.db_manager:
                self.db_manager.update_execution_status(schedule_id, f"執行錯誤: {str(e)[:50]}")
        finally:
            if handler and handler.is_connected:
                await handler.disconnect()

        # 更新狀態列和重新載入表格
        self.status_bar.showMessage(status_msg, 5000)
        self.load_schedules()
        
        # 從執行中任務集合中移除
        self.running_tasks.discard(schedule_id)

    def _is_target_value_matched(self, current_value: Any, target_value: Any, data_type: str) -> bool:
        """比較目前讀值是否已符合目標值。"""
        target_text = str(target_value).strip()

        def parse_bool(text: str) -> Optional[bool]:
            lowered = text.lower()
            if lowered in ("true", "1", "on", "yes"):
                return True
            if lowered in ("false", "0", "off", "no"):
                return False
            return None

        try:
            if data_type == "bool":
                parsed = parse_bool(target_text)
                if parsed is None:
                    return bool(current_value) == bool(target_text)
                return bool(current_value) == parsed

            if data_type == "int":
                return int(float(current_value)) == int(float(target_text))

            if data_type == "float":
                return abs(float(current_value) - float(target_text)) < 1e-6

            if data_type == "string":
                return str(current_value).strip() == target_text

            parsed_bool = parse_bool(target_text)
            if parsed_bool is not None and isinstance(current_value, (bool, int)):
                return bool(current_value) == parsed_bool

            try:
                return abs(float(current_value) - float(target_text)) < 1e-6
            except (TypeError, ValueError):
                return str(current_value).strip() == target_text
        except Exception:
            return str(current_value).strip() == target_text

    def _parse_duration_from_rrule(self, rrule_str: str) -> int:
        """從RRULE字串中解析期間參數（分鐘）"""
        if not rrule_str:
            return 0
        
        try:
            parts = rrule_str.upper().split(';')
            for part in parts:
                if part.startswith('DURATION=PT'):
                    duration_str = part.split('=')[1]
                    match = re.match(r"PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?", duration_str)
                    if not match:
                        continue

                    hours = int(match.group(1) or 0)
                    minutes = int(match.group(2) or 0)
                    seconds = int(match.group(3) or 0)
                    total_minutes = hours * 60 + minutes + (1 if seconds > 0 else 0)
                    return total_minutes
        except Exception:
            pass
        return 0

    def _check_and_disable_if_count_reached(self, schedule_id: int, rrule_str: str):
        """檢查是否達到 COUNT 上限，如果是則停用排程"""
        if not rrule_str:
            return

        try:
            # 解析 RRULE 中的 COUNT
            parts = rrule_str.upper().split(';')
            count_value = None
            
            for part in parts:
                if part.startswith('COUNT='):
                    try:
                        count_value = int(part.split('=')[1])
                        break
                    except ValueError:
                        return

            if count_value is not None:
                executed_count = self.execution_counts.get(schedule_id, 0)
                if executed_count >= count_value:
                    # 達到上限，停用排程
                    if self.db_manager:
                        self.db_manager.update_schedule(schedule_id, is_enabled=0)
                        print(f"排程 {schedule_id} 的執行次數已達上限 ({count_value})，已自動停用")

        except Exception as e:
            # 如果解析失敗，記錄錯誤但不中斷執行
            print(f"檢查 COUNT 上限失敗: {e}")

    def on_database_path_changed(self, new_path: str):
        """處理資料庫路徑變更"""
        # 重新初始化資料庫管理器
        self.db_manager = SQLiteManager(new_path)
        self.db_manager.init_db()

        self._load_time_scale_from_db()

        # 重新載入排程資料
        self.load_schedules()

        # 重新啟動排程工作執行緒
        if self.scheduler_worker:
            self.scheduler_worker.stop()
            self.scheduler_worker.wait()

        self.scheduler_worker = SchedulerWorker(self.db_manager)
        self.scheduler_worker.trigger_task.connect(self.on_task_triggered)
        self.scheduler_worker.start()

    def closeEvent(self, event):
        """處理視窗關閉事件"""
        import asyncio

        # 停止排程器
        if self.scheduler_worker:
            self.scheduler_worker.stop()

        # 取消所有正在執行的async任務
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                # 取消所有待處理的任務
                pending_tasks = [task for task in asyncio.all_tasks(loop) if not task.done()]
                for task in pending_tasks:
                    task.cancel()

                # 等待任務取消完成，最多等待2秒
                import time
                start_time = time.time()
                while pending_tasks and (time.time() - start_time) < 2.0:
                    # 簡單的輪詢等待
                    time.sleep(0.1)
                    pending_tasks = [task for task in pending_tasks if not task.done()]

        except RuntimeError:
            # 沒有事件迴圈或已經關閉
            pass

        event.accept()

    def changeEvent(self, event):
        """處理視窗狀態變化事件"""
        from PySide6.QtCore import QEvent
        if event.type() == QEvent.WindowStateChange:
            if self.windowState() & Qt.WindowMinimized:
                # 最小化時延遲隱藏視窗，確保托盤圖標正確顯示
                QTimer.singleShot(100, self._minimize_to_tray)
            elif self.windowState() == Qt.WindowNoState:
                # 從最小化恢復時顯示視窗
                self.show()
                self.raise_()
                self.activateWindow()
        super().changeEvent(event)

    def _minimize_to_tray(self):
        """將視窗最小化到系統托盤"""
        if not self._allow_minimize_to_tray:
            return
        if self.windowState() & Qt.WindowMinimized:
            self.hide()
            self.tray_icon.show()
            # 設定工具提示
            self.tray_icon.setToolTip("CalendarUA")

    def force_show_on_screen(self):
        """確保主視窗可見且落在目前螢幕可視範圍。"""
        self.setWindowState(self.windowState() & ~Qt.WindowMinimized)
        self.showNormal()

        screen = self.screen() or QGuiApplication.primaryScreen()
        if screen is not None:
            available = screen.availableGeometry()
            width = min(self.width(), available.width())
            height = min(self.height(), available.height())
            self.resize(width, height)

            frame = self.frameGeometry()
            frame.moveCenter(available.center())
            self.move(frame.topLeft())

        self.show()
        self.raise_()
        self.activateWindow()


class OPCNodeBrowserDialog(QDialog):
    """OPC UA 節點瀏覽對話框"""

    _last_selected_node_id: str = ""
    _ROLE_NODE_ID = Qt.ItemDataRole.UserRole
    _ROLE_DATA_TYPE = Qt.ItemDataRole.UserRole + 1
    _ROLE_CAN_WRITE = Qt.ItemDataRole.UserRole + 2
    _ROLE_NODE_CLASS = Qt.ItemDataRole.UserRole + 3
    _ROLE_CHILDREN_LOADED = Qt.ItemDataRole.UserRole + 4
    _ROLE_CHILDREN_LOADING = Qt.ItemDataRole.UserRole + 5
    _cached_children_by_url: Dict[str, Dict[str, List[Dict[str, Any]]]] = {}
    _last_selected_path_by_url: Dict[str, List[str]] = {}

    def __init__(self, parent=None, opc_url: str = "", preferred_node_id: str = ""):
        super().__init__(parent)
        self.opc_url = opc_url
        self.preferred_node_id = self._extract_plain_node_id(preferred_node_id)
        self.selected_node = ""
        self.opc_handler = None
        self.logger = logging.getLogger(__name__)
        self._disconnecting = False
        self.setup_ui()
        self.apply_style()
        # 自動連線並載入節點
        QTimer.singleShot(100, self.connect_and_load)

    @staticmethod
    def _extract_plain_node_id(raw_text: str) -> str:
        if not raw_text:
            return ""
        text = str(raw_text).strip()
        if not text:
            return ""
        if "|" in text:
            parts = [part.strip() for part in text.split("|") if part.strip()]
            for part in reversed(parts):
                if part.startswith("ns="):
                    return part
            return parts[-1] if parts else ""
        return text

    def _iter_tree_items(self, parent_item=None):
        if parent_item is None:
            for i in range(self.tree_widget.topLevelItemCount()):
                top = self.tree_widget.topLevelItem(i)
                yield top
                yield from self._iter_tree_items(top)
            return

        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            yield child
            yield from self._iter_tree_items(child)

    def _find_item_by_node_id(self, node_id: str):
        target = self._extract_plain_node_id(node_id)
        if not target:
            return None
        for item in self._iter_tree_items():
            if item.text(1) == target:
                return item
        return None

    def _expand_item_ancestors(self, item):
        parent = item.parent()
        while parent is not None:
            parent.setExpanded(True)
            parent = parent.parent()

    def _current_cache(self) -> Dict[str, List[Dict[str, Any]]]:
        return self._cached_children_by_url.setdefault(self.opc_url, {})

    def _find_direct_child_by_node_id(self, parent_item, node_id: str):
        target = self._extract_plain_node_id(node_id)
        if not parent_item or not target:
            return None
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            child_id = str(child.data(0, self._ROLE_NODE_ID) or child.text(1) or "").strip()
            if child_id == target:
                return child
        return None

    def _remember_selected_path(self, item):
        node_id = str(item.data(0, self._ROLE_NODE_ID) or item.text(1) or "").strip()
        if not node_id:
            return

        OPCNodeBrowserDialog._last_selected_node_id = node_id

        path: List[str] = []
        cursor = item
        while cursor is not None:
            cursor_node_id = str(cursor.data(0, self._ROLE_NODE_ID) or cursor.text(1) or "").strip()
            if cursor_node_id:
                path.insert(0, cursor_node_id)
            cursor = cursor.parent()

        if path:
            self._last_selected_path_by_url[self.opc_url] = path

    def _restore_last_position(self):
        target = self.preferred_node_id or self._last_selected_node_id
        if not target:
            return

        item = self._find_item_by_node_id(target)
        if item is None:
            return

        self._expand_item_ancestors(item)
        self.tree_widget.setCurrentItem(item)
        self.tree_widget.scrollToItem(item)

    async def _async_restore_last_position(self):
        """在懶載入模式下恢復上次選取節點。"""
        target = self.preferred_node_id or self._last_selected_node_id
        if not target:
            return

        # 先嘗試目前已載入的節點
        item = self._find_item_by_node_id(target)
        if item is not None:
            self._expand_item_ancestors(item)
            self.tree_widget.setCurrentItem(item)
            self.tree_widget.scrollToItem(item)
            return

        if not self.opc_handler or not self.opc_handler.is_connected or not self.opc_handler.client:
            return

        # 優先使用已記錄路徑直接還原，避免每次都重新搜尋。
        remembered_path = list(self._last_selected_path_by_url.get(self.opc_url, []))
        if remembered_path:
            current = self.tree_widget.topLevelItem(0) if self.tree_widget.topLevelItemCount() > 0 else None
            if current is not None:
                current_id = str(current.data(0, self._ROLE_NODE_ID) or current.text(1) or "").strip()
                start_index = 1 if remembered_path and remembered_path[0] == current_id else 0
                restore_ok = True

                for path_node_id in remembered_path[start_index:]:
                    await self._async_load_children_for_item(current)
                    next_item = self._find_direct_child_by_node_id(current, path_node_id)
                    if next_item is None:
                        restore_ok = False
                        break
                    current = next_item

                if restore_ok:
                    final_node_id = str(current.data(0, self._ROLE_NODE_ID) or current.text(1) or "").strip()
                    if final_node_id == target:
                        self._expand_item_ancestors(current)
                        self.tree_widget.setCurrentItem(current)
                        self.tree_widget.scrollToItem(current)
                        self.status_label.setText(f"已定位上次節點: {current.text(0)}")
                        self.status_label.setStyleSheet("color: #00aa00;")
                        return

        # 逐層展開 Object/View 節點直到找到目標，避免一次遞迴載完整棵樹
        queue = []
        for i in range(self.tree_widget.topLevelItemCount()):
            queue.append(self.tree_widget.topLevelItem(i))

        visited: set[str] = set()
        max_expand = 500
        expanded_count = 0

        while queue and expanded_count < max_expand:
            item = queue.pop(0)
            if item is None:
                continue

            node_id = str(item.data(0, self._ROLE_NODE_ID) or item.text(1) or "").strip()
            if not node_id or node_id in visited:
                continue
            visited.add(node_id)

            node_class_name = str(item.data(0, self._ROLE_NODE_CLASS) or item.text(2) or "")
            if node_class_name in ("Object", "View"):
                await self._async_load_children_for_item(item)
                expanded_count += 1

            found = self._find_item_by_node_id(target)
            if found is not None:
                self._expand_item_ancestors(found)
                self.tree_widget.setCurrentItem(found)
                self.tree_widget.scrollToItem(found)
                self.status_label.setText(f"已定位上次節點: {found.text(0)}")
                self.status_label.setStyleSheet("color: #00aa00;")
                return

            for j in range(item.childCount()):
                queue.append(item.child(j))

    def setup_ui(self):
        """設定介面"""
        self.setWindowTitle("瀏覽 OPC UA 節點")
        self.setWindowIcon(get_app_icon())
        self.setMinimumSize(500, 400)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # 顯示目前連線資訊
        info_layout = QHBoxLayout()
        info_layout.addWidget(QLabel("OPC URL:"))
        self.url_label = QLabel(self.opc_url)
        self.url_label.setStyleSheet("font-weight: bold;")
        info_layout.addWidget(self.url_label)
        info_layout.addStretch()
        layout.addLayout(info_layout)

        # 狀態標籤
        self.status_label = QLabel("正在連線...")
        self.status_label.setStyleSheet("color: #666;")
        layout.addWidget(self.status_label)

        # 節點樹狀列表
        self.tree_widget = QTreeWidget()
        self.tree_widget.setHeaderLabels(["節點名稱", "Node ID", "節點類型", "資料型別"])
        self.tree_widget.setColumnWidth(0, 200)
        self.tree_widget.setColumnWidth(1, 150)
        self.tree_widget.setColumnWidth(2, 100)
        self.tree_widget.setColumnWidth(3, 80)
        self.tree_widget.itemSelectionChanged.connect(self.on_selection_changed)
        self.tree_widget.itemClicked.connect(self.on_item_clicked)
        self.tree_widget.itemExpanded.connect(self.on_item_expanded)
        self.tree_widget.itemDoubleClicked.connect(self.on_item_double_clicked)
        layout.addWidget(self.tree_widget)

        # 按鈕
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        refresh_btn = QPushButton("重新整理")
        refresh_btn.clicked.connect(lambda: self.connect_and_load(force_refresh=True))
        button_layout.addWidget(refresh_btn)

        button_layout.addSpacing(20)

        self.select_btn = QPushButton("選擇")
        self.select_btn.setDefault(True)
        self.select_btn.setEnabled(False)
        self.select_btn.clicked.connect(self.accept)
        button_layout.addWidget(self.select_btn)

        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)

    def apply_style(self):
        """套用樣式"""
        parent = self.parent()
        is_dark = False
        if parent and hasattr(parent, "parent"):
            grandparent = parent.parent()
            if grandparent and hasattr(grandparent, "current_theme"):
                if grandparent.current_theme == "dark":
                    is_dark = True
                elif grandparent.current_theme == "system":
                    if hasattr(grandparent, "is_system_dark_mode"):
                        is_dark = grandparent.is_system_dark_mode()

        if is_dark:
            self.setStyleSheet("""
                QDialog {
                    background-color: #2b2b2b;
                }
                QLabel {
                    color: #cccccc;
                }
                QTreeWidget {
                    background-color: #1e1e1e;
                    border: 1px solid #3d3d3d;
                    color: #cccccc;
                }
                QTreeWidget::item:selected {
                    background-color: #094771;
                    color: white;
                }
                QTreeWidget::item:hover {
                    background-color: #2d2d2d;
                }
                QHeaderView::section {
                    background-color: #252526;
                    padding: 6px;
                    border: none;
                    border-bottom: 2px solid #0e639c;
                    font-weight: bold;
                    color: #cccccc;
                }
                QPushButton {
                    background-color: #0e639c;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 8px 16px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #1f89cd;
                }
                QPushButton:disabled {
                    background-color: #4a4a4a;
                    color: #808080;
                }
                QComboBox {
                    border: 1px solid #3d3d3d;
                    border-radius: 4px;
                    padding: 6px;
                    background-color: #1e1e1e;
                    color: #cccccc;
                }
                QComboBox::drop-down {
                    width: 0px;
                    border: none;
                }
                QComboBox::down-arrow {
                    image: none;
                    width: 0px;
                    height: 0px;
                }
            """)
        else:
            self.setStyleSheet("""
                QDialog {
                    background-color: #f5f5f5;
                }
                QLabel {
                    color: #333;
                }
                QTreeWidget {
                    background-color: white;
                    border: 1px solid #d0d0d0;
                    color: #333;
                }
                QTreeWidget::item:selected {
                    background-color: #9ec6f3;
                    color: #0f1f33;
                }
                QTreeWidget::item:hover {
                    background-color: #cfe3f8;
                }
                QHeaderView::section {
                    background-color: #f0f0f0;
                    padding: 6px;
                    border: none;
                    border-bottom: 2px solid #0078d4;
                    font-weight: bold;
                }
                QPushButton {
                    background-color: #e9ecef;
                    color: #111111;
                    border: none;
                    border-radius: 4px;
                    padding: 8px 16px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #c7d4e2;
                }
                QPushButton:disabled {
                    background-color: #cccccc;
                    color: #888888;
                }
                QComboBox {
                    border: 1px solid #d0d0d0;
                    border-radius: 4px;
                    padding: 6px;
                    background-color: white;
                    color: #333;
                }
                QComboBox::drop-down {
                    width: 0px;
                    border: none;
                }
                QComboBox::down-arrow {
                    image: none;
                    width: 0px;
                    height: 0px;
                }
            """)

    def connect_and_load(self, force_refresh: bool = False):
        """連線到 OPC UA 並載入節點 - 使用 qasync 整合"""
        self.tree_widget.clear()
        self.status_label.setText("正在連線...")
        self.status_label.setStyleSheet("color: #666;")

        if force_refresh:
            self._cached_children_by_url.pop(self.opc_url, None)

        # 使用 QTimer 稍後執行異步操作，避免阻塞 UI
        QTimer.singleShot(100, self._async_connect_and_load)

    def _async_connect_and_load(self):
        """異步連線和載入"""
        async def do_connect():
            try:
                self.opc_handler = OPCHandler(self.opc_url)

                # 連線到 OPC UA 伺服器
                success = await self.opc_handler.connect()

                if success:
                    self.status_label.setText("已連線，正在載入節點...")
                    self.status_label.setStyleSheet("color: green;")

                    # 載入節點
                    await self._async_load_nodes()
                else:
                    self.status_label.setText("連線失敗 - 請檢查 URL 和伺服器狀態")
                    self.status_label.setStyleSheet("color: red;")

            except Exception as e:
                self.status_label.setText(f"連線錯誤: {str(e)}")
                self.status_label.setStyleSheet("color: red;")

        # 使用現有的 qasync 事件迴圈執行
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                # 如果迴圈已在執行，建立 task
                asyncio.create_task(do_connect())
            else:
                loop.run_until_complete(do_connect())
        except RuntimeError:
            # 沒有事件迴圈的情況
            asyncio.run(do_connect())

    async def _async_load_nodes(self):
        """異步載入 OPC UA 節點樹（僅預載第一層）。"""
        try:
            # 取得 Objects 節點
            objects = await self.opc_handler.get_objects_node()

            if objects:
                root_item = QTreeWidgetItem(self.tree_widget)
                root_item.setText(0, "Objects")
                root_item.setText(1, "i=85")
                root_item.setText(2, "Object")

                root_item.setData(0, self._ROLE_NODE_ID, "i=85")
                root_item.setData(0, self._ROLE_NODE_CLASS, "Object")
                root_item.setData(0, self._ROLE_CAN_WRITE, None)
                root_item.setData(0, self._ROLE_CHILDREN_LOADED, False)
                root_item.setData(0, self._ROLE_CHILDREN_LOADING, True)

                # 僅載入第一層，深層在使用者點擊/展開時再載入
                await self._async_load_child_nodes(objects, root_item)
                root_item.setData(0, self._ROLE_CHILDREN_LOADED, True)
                root_item.setData(0, self._ROLE_CHILDREN_LOADING, False)
                root_item.setExpanded(True)

            self.status_label.setText("已載入節點")
            await self._async_restore_last_position()
            # 確保樹狀元件正確更新
            self.tree_widget.viewport().update()

        except Exception as e:
            self.status_label.setText(f"載入節點錯誤: {str(e)}")
            self.status_label.setStyleSheet("color: red;")

    async def _async_load_child_nodes(self, parent_node, parent_item):
        """異步載入指定節點的直屬子節點（單層）。"""
        if parent_item is None:
            return

        try:
            parent_node_id = str(parent_item.data(0, self._ROLE_NODE_ID) or parent_item.text(1) or "").strip()
            cache = self._current_cache()
            children_info = cache.get(parent_node_id)

            if children_info is None:
                # 取得子節點
                children = await parent_node.get_children()
                children_info = []
                seen_node_ids = set()

                for child in children:
                    try:
                        # 取得節點資訊
                        browse_name = await child.read_browse_name()
                        # 正確格式化 Node ID
                        node_id = child.nodeid.to_string()
                        if node_id in seen_node_ids:
                            continue
                        seen_node_ids.add(node_id)

                        node_class = await child.read_node_class()
                        node_class_name = node_class.name

                        # 讀取資料型別和存取權限（僅適用於變數節點）
                        data_type = "-"
                        can_write = None

                        if node_class_name == "Variable":
                            try:
                                # 讀取資料型別
                                detected_type = await self.opc_handler.read_node_data_type(node_id)
                                data_type = detected_type if detected_type else "未知"
                                self.logger.debug(f"Node {node_id} 資料型別: {data_type}")

                                # 讀取存取權限
                                try:
                                    from asyncua.ua import AttributeIds
                                    access_level_data = await child.read_attribute(AttributeIds.AccessLevel)
                                    access_level_value = access_level_data.Value.Value if hasattr(access_level_data, 'Value') and access_level_data.Value else None
                                    self.logger.debug(f"Node {node_id} AccessLevel: {access_level_value}")
                                    can_write = bool(access_level_value & 0x02) if access_level_value is not None and access_level_value > 0 else True
                                except Exception as e:
                                    self.logger.debug(f"無法讀取 Node {node_id} 的 AccessLevel: {e}")
                                    can_write = True

                                if not can_write:
                                    data_type = "唯讀"

                            except Exception as e:
                                self.logger.error(f"讀取 Node {node_id} 資料型別失敗: {e}")
                                data_type = "未知"
                                can_write = False

                        children_info.append(
                            {
                                "browse_name": browse_name.Name,
                                "node_id": node_id,
                                "node_class": node_class_name,
                                "data_type": data_type,
                                "can_write": can_write,
                            }
                        )

                    except Exception as e:
                        self.logger.warning(f"載入子節點失敗: {e}")
                        # 即使失敗也要繼續處理其他節點

                cache[parent_node_id] = children_info

            # 重新載入該層時先清空舊子節點
            parent_item.takeChildren()

            if not children_info:
                parent_item.setData(0, self._ROLE_CHILDREN_LOADED, True)
                parent_item.setData(0, self._ROLE_CHILDREN_LOADING, False)
                parent_item.setChildIndicatorPolicy(QTreeWidgetItem.ChildIndicatorPolicy.DontShowIndicatorWhenChildless)
                return

            loaded_children = []
            seen_node_ids = set()

            # 從快取資料建立同層節點，避免重覆網路讀取
            for info in children_info:
                try:
                    node_id = str(info.get("node_id", "") or "").strip()
                    if not node_id or node_id in seen_node_ids:
                        continue
                    seen_node_ids.add(node_id)

                    child_item = QTreeWidgetItem(parent_item)
                    browse_name = str(info.get("browse_name", "") or "")
                    node_class_name = str(info.get("node_class", "") or "")
                    data_type = str(info.get("data_type", "-") or "-")
                    can_write = info.get("can_write", None)

                    child_item.setText(0, browse_name)
                    child_item.setText(1, node_id)
                    child_item.setText(2, node_class_name)
                    child_item.setText(3, data_type)

                    # 儲存節點 ID 和資料型別
                    child_item.setData(0, self._ROLE_NODE_ID, node_id)
                    child_item.setData(0, self._ROLE_DATA_TYPE, data_type)
                    child_item.setData(0, self._ROLE_CAN_WRITE, can_write)
                    child_item.setData(0, self._ROLE_NODE_CLASS, node_class_name)
                    child_item.setData(0, self._ROLE_CHILDREN_LOADED, False)
                    child_item.setData(0, self._ROLE_CHILDREN_LOADING, False)

                    # Object/View 類型才顯示可展開符號，點擊時再動態載入其下一層
                    if node_class_name in ("Object", "View"):
                        child_item.setChildIndicatorPolicy(QTreeWidgetItem.ChildIndicatorPolicy.ShowIndicator)
                    else:
                        child_item.setChildIndicatorPolicy(QTreeWidgetItem.ChildIndicatorPolicy.DontShowIndicatorWhenChildless)

                    loaded_children.append(child_item)

                except Exception as e:
                    self.logger.warning(f"載入子節點失敗: {e}")
                    # 即使失敗也要繼續處理其他節點

            # 讓 UI 先刷新，保持快速回應
            if loaded_children:
                self.tree_widget.viewport().update()
                await asyncio.sleep(0)

            parent_item.setData(0, self._ROLE_CHILDREN_LOADED, True)
            parent_item.setData(0, self._ROLE_CHILDREN_LOADING, False)

        except Exception as e:
            parent_item.setData(0, self._ROLE_CHILDREN_LOADING, False)
            self.logger.error(f"載入子節點列表失敗: {e}")

    def _request_load_children(self, item):
        """需要時才載入該節點下一層子節點。"""
        if not item or not self.opc_handler or not self.opc_handler.is_connected:
            return

        node_class_name = str(item.data(0, self._ROLE_NODE_CLASS) or item.text(2) or "")
        if node_class_name not in ("Object", "View"):
            return

        if bool(item.data(0, self._ROLE_CHILDREN_LOADED)):
            return
        if bool(item.data(0, self._ROLE_CHILDREN_LOADING)):
            return

        node_id = str(item.data(0, self._ROLE_NODE_ID) or item.text(1) or "").strip()
        if not node_id:
            return

        async def do_load():
            try:
                await self._async_load_children_for_item(item)
                self.status_label.setText(f"已載入 {item.text(0)} 子節點")
                self.status_label.setStyleSheet("color: #00aa00;")
                item.setExpanded(True)
            except Exception as e:
                item.setData(0, self._ROLE_CHILDREN_LOADING, False)
                self.status_label.setText(f"載入子節點失敗: {e}")
                self.status_label.setStyleSheet("color: red;")

        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                asyncio.create_task(do_load())
            else:
                loop.run_until_complete(do_load())
        except RuntimeError:
            asyncio.run(do_load())

    async def _async_load_children_for_item(self, item):
        """以 await 方式載入指定 item 子節點，供 restore 與點擊載入共用。"""
        if not item or not self.opc_handler or not self.opc_handler.is_connected or not self.opc_handler.client:
            return

        node_class_name = str(item.data(0, self._ROLE_NODE_CLASS) or item.text(2) or "")
        if node_class_name not in ("Object", "View"):
            return
        if bool(item.data(0, self._ROLE_CHILDREN_LOADED)):
            return
        if bool(item.data(0, self._ROLE_CHILDREN_LOADING)):
            return

        node_id = str(item.data(0, self._ROLE_NODE_ID) or item.text(1) or "").strip()
        if not node_id:
            return

        item.setData(0, self._ROLE_CHILDREN_LOADING, True)
        node = self.opc_handler.client.get_node(node_id)
        await self._async_load_child_nodes(node, item)

    def on_item_clicked(self, item, column):
        self._request_load_children(item)

    def on_item_expanded(self, item):
        self._request_load_children(item)

    def on_selection_changed(self):
        """處理選擇變更"""
        selected_items = self.tree_widget.selectedItems()
        if selected_items:
            selected_item = selected_items[0]
            display_name = selected_item.text(0)
            node_id = selected_item.text(1)
            data_type = selected_item.text(3) if selected_item.text(3) != "-" else "未知"
            can_write = selected_item.data(0, self._ROLE_CAN_WRITE)
            
            if can_write is True:
                self.selected_node = f"{display_name}|{node_id}|{data_type}"
                self._remember_selected_path(selected_item)
                self.select_btn.setEnabled(True)
                self.status_label.setText("已選擇可寫入節點")
                self.status_label.setStyleSheet("color: green;")
            elif can_write is False:
                self.selected_node = ""
                self.select_btn.setEnabled(False)
                self.status_label.setText("選擇的節點為唯讀，無法寫入")
                self.status_label.setStyleSheet("color: red;")
            else:
                self.selected_node = ""
                self.select_btn.setEnabled(False)
                self.status_label.setText("此節點可展開瀏覽，請選擇可寫入變數節點")
                self.status_label.setStyleSheet("color: #666;")
        else:
            self.selected_node = ""
            self.select_btn.setEnabled(False)
            self.status_label.setText("")

    def on_item_double_clicked(self, item, column):
        """處理雙擊事件"""
        display_name = item.text(0)
        node_id = item.text(1)
        data_type = item.text(3) if item.text(3) != "-" else "未知"
        can_write = item.data(0, self._ROLE_CAN_WRITE)
        
        if can_write is True:
            self.selected_node = f"{display_name}|{node_id}|{data_type}"
            self._remember_selected_path(item)
            self.accept()
        else:
            self.status_label.setText("無法選擇唯讀節點")
            self.status_label.setStyleSheet("color: red;")

    def _close_and_disconnect(self):
        if self._disconnecting:
            return
        self._disconnecting = True

        async def do_disconnect():
            try:
                if self.opc_handler and self.opc_handler.is_connected:
                    await self.opc_handler.disconnect()
            finally:
                self._disconnecting = False

        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                asyncio.create_task(do_disconnect())
            else:
                loop.run_until_complete(do_disconnect())
        except RuntimeError:
            asyncio.run(do_disconnect())

    def accept(self):
        self._close_and_disconnect()
        super().accept()

    def reject(self):
        self._close_and_disconnect()
        super().reject()

    def get_selected_node(self) -> str:
        """取得選擇的節點 ID 和資料型別"""
        return self.selected_node


class OPCSettingsDialog(QDialog):
    """OPC UA 設定對話框"""

    class StepNumberComboBox(QComboBox):
        """數值下拉：支援滾輪，逐格切換下拉選單數值。"""

        def __init__(self, minimum: int, maximum: int, step: int = 1, parent=None):
            super().__init__(parent)
            self._minimum = minimum
            self._maximum = max(minimum, maximum)
            self._step = max(1, step)
            self._window_size = 10
            self._current_value = self._minimum

            self.setEditable(True)
            self.setInsertPolicy(QComboBox.NoInsert)
            self.setMaxVisibleItems(self._window_size)

            if self.lineEdit() is not None:
                self.lineEdit().setReadOnly(True)
                self.lineEdit().setAlignment(Qt.AlignCenter)
                self.lineEdit().setCursor(Qt.PointingHandCursor)
                self.lineEdit().installEventFilter(self)

            self.setCursor(Qt.PointingHandCursor)
            self.installEventFilter(self)
            self.view().installEventFilter(self)
            self.view().viewport().installEventFilter(self)

            self._rebuild_window(self._current_value)

        def value(self) -> int:
            data = self.currentData()
            if isinstance(data, int):
                return data
            try:
                return int(self.currentText().strip())
            except Exception:
                return self._current_value

        def setValue(self, value: int):
            if value < self._minimum:
                value = self._minimum
            if value > self._maximum:
                value = self._maximum
            self._current_value = value

            idx = self.findData(value)
            if idx < 0:
                self._rebuild_window(value)
                idx = self.findData(value)

            if idx >= 0:
                self.blockSignals(True)
                self.setCurrentIndex(idx)
                self.blockSignals(False)

        def _set_window(self, center_value: int, selected_value: int | None = None):
            half = self._window_size // 2
            start = center_value - half
            if start < self._minimum:
                start = self._minimum

            end = start + self._window_size - 1
            if end > self._maximum:
                end = self._maximum
                start = max(self._minimum, end - self._window_size + 1)

            target_value = selected_value if isinstance(selected_value, int) else center_value
            if target_value < start:
                target_value = start
            elif target_value > end:
                target_value = end

            values = list(range(start, end + 1))
            if not values:
                values = [center_value]

            self.blockSignals(True)
            self.clear()
            for n in values:
                self.addItem(str(n), n)

            idx = self.findData(target_value)
            if idx >= 0:
                self.setCurrentIndex(idx)
            self.blockSignals(False)
            self._current_value = self.value()

        def _rebuild_window(self, center_value: int):
            self._set_window(center_value, center_value)

        def _steps_from_wheel(self, event) -> int:
            delta = event.angleDelta().y()
            if delta == 0:
                return 0
            steps = int(delta / 120)
            if steps == 0:
                steps = 1 if delta > 0 else -1
            # 對齊主畫面 2026 行為：下滾增加、上滾減少
            return -steps

        def _shift_window_by_steps(self, steps: int):
            if steps == 0:
                return

            center_idx = min(self._window_size // 2, max(0, self.count() - 1))
            center_value = self.itemData(center_idx)
            if not isinstance(center_value, int):
                center_value = self.value()

            selected_value = self.value()
            self._set_window(center_value + steps, selected_value)

        def eventFilter(self, obj, event):
            if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
                line_edit = self.lineEdit()
                if obj in (self, line_edit):
                    QTimer.singleShot(0, self.showPopup)
                    event.accept()
                    return True

            if obj in (self, self.view(), self.view().viewport()) and event.type() == QEvent.Wheel:
                steps = self._steps_from_wheel(event)
                self._shift_window_by_steps(steps)
                event.accept()
                return True

            if obj is self.lineEdit() and event.type() == QEvent.Wheel:
                steps = self._steps_from_wheel(event)
                self._shift_window_by_steps(steps)
                event.accept()
                return True

            return super().eventFilter(obj, event)

    def __init__(self, parent=None, security_policy="None", username="", password="", timeout=5, write_timeout=3, security_mode="None", opc_url: str = ""):
        super().__init__(parent)
        self.security_policy = security_policy
        self.security_mode = security_mode
        self.username = username
        self.password = password
        self.timeout = timeout
        self.write_timeout = write_timeout
        self.opc_url = opc_url
        self._detected_supported = None
        
        self.setup_ui()
        attach_combo_wheel_behavior(self)
        
        # 連接信號
        self.chk_show_supported.toggled.connect(self.on_chk_show_supported_toggled)
        self.rb_anonymous.toggled.connect(self.on_auth_mode_changed)
        self.rb_username.toggled.connect(self.on_auth_mode_changed)
        self.rb_certificate.toggled.connect(self.on_auth_mode_changed)
        
        # 載入現有設定
        self.load_data()
        
        # 套用樣式
        self.apply_style()

    def setup_ui(self):
        self.setWindowTitle("OPC UA 連線設定")
        self.setWindowIcon(get_app_icon())
        self.setMinimumWidth(900)
        # 移除固定高度，讓視窗根據內容自動調整
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(15, 15, 15, 15)

        # 使用者認證（頂部）
        auth_top_layout = QHBoxLayout()
        auth_top_layout.setSpacing(12)
        auth_top_layout.setContentsMargins(0, 0, 0, 0)

        auth_group = QGroupBox("User Authentication")
        auth_group_layout = QVBoxLayout(auth_group)
        auth_group_layout.setContentsMargins(10, 12, 10, 12)
        auth_group_layout.setSpacing(8)
        
        # 認證方式單選按鈕容器（水平）
        rb_layout = QHBoxLayout()
        rb_layout.setContentsMargins(0, 0, 0, 0)
        rb_layout.setSpacing(15)
        self.rb_anonymous = QRadioButton("Anonymous")
        self.rb_username = QRadioButton("Username and Password")
        self.rb_certificate = QRadioButton("Certificate and Private Key")
        rb_layout.addWidget(self.rb_anonymous)
        rb_layout.addWidget(self.rb_username)
        rb_layout.addWidget(self.rb_certificate)

        auth_group_layout.addLayout(rb_layout)

        # 認證欄位（username/password）
        cred_layout = QGridLayout()
        cred_layout.setSpacing(8)
        cred_layout.setContentsMargins(0, 0, 0, 0)
        self.username_label = QLabel("Username:")
        cred_layout.addWidget(self.username_label, 0, 0)
        self.username_edit = QLineEdit()
        cred_layout.addWidget(self.username_edit, 0, 1)
        self.password_label = QLabel("Password:")
        cred_layout.addWidget(self.password_label, 0, 2)
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        cred_layout.addWidget(self.password_edit, 0, 3)

        # 憑證欄位（檔案選取）
        cert_layout = QGridLayout()
        cert_layout.setSpacing(8)
        cert_layout.setContentsMargins(0, 0, 0, 0)
        self.client_cert_label = QLabel("Client Cert:")
        cert_layout.addWidget(self.client_cert_label, 0, 0)
        self.client_cert_edit = QLineEdit()
        self.cert_browse_btn1 = QPushButton("Browse")
        self.cert_browse_btn1.clicked.connect(lambda: self._browse_file(self.client_cert_edit))
        cert_layout.addWidget(self.client_cert_edit, 0, 1)
        cert_layout.addWidget(self.cert_browse_btn1, 0, 2)

        self.client_key_label = QLabel("Private Key:")
        cert_layout.addWidget(self.client_key_label, 1, 0)
        self.client_key_edit = QLineEdit()
        self.cert_browse_btn2 = QPushButton("Browse")
        self.cert_browse_btn2.clicked.connect(lambda: self._browse_file(self.client_key_edit))
        cert_layout.addWidget(self.client_key_edit, 1, 1)
        cert_layout.addWidget(self.cert_browse_btn2, 1, 2)

        auth_group_layout.addLayout(cred_layout)
        auth_group_layout.addLayout(cert_layout)

        # 右側按鈕（Apply / Cancel）- 固定大小，頂部對齐
        btns_layout = QVBoxLayout()
        btns_layout.setContentsMargins(0, 0, 0, 0)
        btns_layout.setSpacing(8)
        self.apply_btn = QPushButton("Apply")
        self.apply_btn.setFixedSize(90, 40)
        self.apply_btn.clicked.connect(self._on_apply)
        btns_layout.addWidget(self.apply_btn, alignment=Qt.AlignmentFlag.AlignTop)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFixedSize(90, 40)
        cancel_btn.clicked.connect(self.reject)
        btns_layout.addWidget(cancel_btn, alignment=Qt.AlignmentFlag.AlignTop)
        btns_layout.addStretch()

        auth_top_layout.addWidget(auth_group, 1)
        # 按鈕布局容器，頂部對齐
        btns_container = QWidget()
        btns_container_layout = QVBoxLayout(btns_container)
        btns_container_layout.setContentsMargins(0, 0, 0, 0)
        btns_container_layout.addLayout(btns_layout)
        btns_container_layout.addStretch()
        auth_top_layout.addWidget(btns_container, 0)

        layout.addLayout(auth_top_layout)

        # 安全設定（中段）：左為 Security Mode，右為 Security Policy
        sec_mid_layout = QHBoxLayout()
        sec_mid_layout.setSpacing(12)
        sec_mid_layout.setContentsMargins(0, 0, 0, 0)

        # Security Mode
        mode_group = QGroupBox("Security Mode")
        mode_layout = QVBoxLayout(mode_group)
        mode_layout.setContentsMargins(10, 12, 10, 12)
        mode_layout.setSpacing(8)
        self.rb_mode_none = QRadioButton("None")
        self.rb_mode_sign = QRadioButton("Sign")
        self.rb_mode_sign_encrypt = QRadioButton("Sign & Encrypt")
        mode_layout.addWidget(self.rb_mode_none)
        mode_layout.addWidget(self.rb_mode_sign)
        mode_layout.addWidget(self.rb_mode_sign_encrypt)
        # 添加間距以與 Security Policy 對齊
        mode_layout.addStretch()

        # Security Policy list
        policy_group = QGroupBox("Security Policy")
        policy_layout = QVBoxLayout(policy_group)
        policy_layout.setContentsMargins(10, 12, 10, 12)
        policy_layout.setSpacing(8)
        self.policy_rb_none = QRadioButton("None")
        self.policy_rb_basic128 = QRadioButton("Basic128RSA15")
        self.policy_rb_basic256 = QRadioButton("Basic256")
        self.policy_rb_basic256sha = QRadioButton("Basic256Sha256")
        policy_layout.addWidget(self.policy_rb_none)
        policy_layout.addWidget(self.policy_rb_basic128)
        policy_layout.addWidget(self.policy_rb_basic256)
        policy_layout.addWidget(self.policy_rb_basic256sha)

        self.chk_show_supported = QCheckBox("Show only modes that are supported by the server")
        policy_layout.addWidget(self.chk_show_supported)

        sec_mid_layout.addWidget(mode_group, 1)
        sec_mid_layout.addWidget(policy_group, 1)

        layout.addLayout(sec_mid_layout)
        
        # 連線設定（下方）
        connection_group = QGroupBox("連線設定")
        connection_layout = QHBoxLayout(connection_group)
        connection_layout.setContentsMargins(10, 8, 10, 8)
        connection_layout.setSpacing(8)
        connection_layout.addWidget(QLabel("連線超時 (秒):"))
        self.timeout_spin = OPCSettingsDialog.StepNumberComboBox(1, 300, step=1)
        self.timeout_spin.setValue(5)
        self.timeout_spin.setFixedWidth(80)
        connection_layout.addWidget(self.timeout_spin)
        connection_layout.addWidget(QLabel("寫值重試延遲 (秒):"))
        self.write_timeout_spin = OPCSettingsDialog.StepNumberComboBox(1, 60, step=1)
        self.write_timeout_spin.setValue(3)
        self.write_timeout_spin.setFixedWidth(80)
        connection_layout.addWidget(self.write_timeout_spin)
        connection_layout.addStretch()
        layout.addWidget(connection_group)

    def apply_style(self):
        """套用樣式，根據父視窗主題選擇亮色或暗色"""
        is_dark = False
        parent = self.parent()
        
        # 嘗試找到主視窗以取得主題設定
        current_window = parent
        while current_window:
            if hasattr(current_window, "current_theme"):
                if current_window.current_theme == "dark":
                    is_dark = True
                elif current_window.current_theme == "system":
                    if hasattr(current_window, "is_system_dark_mode"):
                        is_dark = current_window.is_system_dark_mode()
                break
            current_window = current_window.parent() if hasattr(current_window, "parent") else None

        if is_dark:
            self._apply_dark_style()
        else:
            self._apply_light_style()

    def _apply_light_style(self):
        """套用亮色樣式"""
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
                color: #2c3e50;
            }
            QLabel {
                color: #333;
            }
            QPushButton {
                background-color: #e9ecef;
                color: #111111;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #c7d4e2;
            }
            QComboBox {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 6px;
                background-color: white;
                color: #333;
            }
            QComboBox::drop-down {
                width: 0px;
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
            }
            QLineEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 6px;
                background-color: white;
                color: #333;
            }
            QLineEdit:focus {
                border: 2px solid #0078d4;
            }
            QSpinBox {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 6px;
                background-color: white;
                color: #333;
            }
            QSpinBox:focus {
                border: 2px solid #0078d4;
            }
            QRadioButton {
                color: #333;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #999999;
                border-radius: 9px;
                background-color: white;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0078d4;
                background-color: #f0f0f0;
            }
            QRadioButton::indicator:checked {
                background-color: #0078d4;
                border: 2px solid #0078d4;
            }
            QCheckBox {
                color: #333;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #999999;
                border-radius: 2px;
                background-color: white;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #f0f0f0;
            }
            QCheckBox::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
                image: url(:/checkbox_check);
            }
        """)

    def _apply_dark_style(self):
        """套用暗色樣式"""
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
            }
            QGroupBox::title {
                color: #ffffff;
            }
            QLabel {
                color: #cccccc;
            }
            QPushButton {
                background-color: #0e639c;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1f89cd;
            }
            QComboBox {
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                background-color: #1e1e1e;
                color: #cccccc;
            }
            QComboBox::drop-down {
                width: 0px;
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
            }
            QLineEdit {
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                background-color: #1e1e1e;
                color: #cccccc;
            }
            QLineEdit:focus {
                border: 2px solid #0e639c;
            }
            QSpinBox {
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                background-color: #1e1e1e;
                color: #cccccc;
            }
            QSpinBox:focus {
                border: 2px solid #0e639c;
            }
            QRadioButton {
                color: #cccccc;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #666666;
                border-radius: 9px;
                background-color: #1e1e1e;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #252526;
            }
            QRadioButton::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
            }
            QCheckBox {
                color: #cccccc;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #666666;
                border-radius: 2px;
                background-color: #1e1e1e;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #252526;
            }
            QCheckBox::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
                image: url(:/checkbox_check);
            }
        """)

    def load_data(self):
        """載入現有的 OPC 連線設定"""
        # 設定安全策略單選按鈕（默認為 None）
        policy_mapping = {
            "None": self.policy_rb_none,
            "Basic128Rsa15": self.policy_rb_basic128,
            "Basic256": self.policy_rb_basic256,
            "Basic256Sha256": self.policy_rb_basic256sha,
        }
        policy_btn = policy_mapping.get(self.security_policy, self.policy_rb_none)
        policy_btn.setChecked(True)

        # 設定使用者認證相關欄位（默認為 Anonymous）
        if self.username or self.password:
            self.rb_username.setChecked(True)
        else:
            self.rb_anonymous.setChecked(True)

        self.username_edit.setText(self.username)
        self.password_edit.setText(self.password)
        self.timeout_spin.setValue(self.timeout)
        self.write_timeout_spin.setValue(self.write_timeout)

        # 初始化安全模式
        mode_mapping = {
            "None": self.rb_mode_none,
            "Sign": self.rb_mode_sign,
            "SignAndEncrypt": self.rb_mode_sign_encrypt,
        }
        mode_btn = mode_mapping.get(self.security_mode, self.rb_mode_none)
        mode_btn.setChecked(True)
        
        # 初始化認證欄位可見性
        self.on_auth_mode_changed()

        # 啟用「只顯示伺服器支援的模式」複選框並觸發檢測
        self.chk_show_supported.setChecked(True)

    def get_settings(self) -> Dict[str, Any]:
        """取得設定值"""
        # security policy
        if self.policy_rb_none.isChecked():
            policy = "None"
        elif self.policy_rb_basic128.isChecked():
            policy = "Basic128Rsa15"
        elif self.policy_rb_basic256.isChecked():
            policy = "Basic256"
        else:
            policy = "Basic256Sha256"

        # security mode
        if self.rb_mode_none.isChecked():
            mode = "None"
        elif self.rb_mode_sign.isChecked():
            mode = "Sign"
        else:
            mode = "SignAndEncrypt"

        # auth method
        if self.rb_anonymous.isChecked():
            auth = "Anonymous"
        elif self.rb_username.isChecked():
            auth = "Username"
        else:
            auth = "Certificate"

        return {
            "security_policy": policy,
            "security_mode": mode,
            "auth_method": auth,
            "username": self.username_edit.text(),
            "password": self.password_edit.text(),
            "client_cert": self.client_cert_edit.text(),
            "client_key": self.client_key_edit.text(),
            "timeout": self.timeout_spin.value(),
            "write_timeout": self.write_timeout_spin.value(),
            "show_only_supported": self.chk_show_supported.isChecked(),
        }

    def _browse_file(self, line_edit: QLineEdit):
        path, _ = QFileDialog.getOpenFileName(self, "Select file")
        if path:
            line_edit.setText(path)

    def on_auth_mode_changed(self):
        # 當選擇 Anonymous 時隱藏/停用其他認證欄位
        is_username = self.rb_username.isChecked()

        # Username/password visible only when username radio selected
        self.username_label.setVisible(is_username)
        self.username_edit.setVisible(is_username)
        self.password_label.setVisible(is_username)
        self.password_edit.setVisible(is_username)

        # Certificate fields visible only when certificate radio selected
        cert_visible = self.rb_certificate.isChecked()
        self.client_cert_label.setVisible(cert_visible)
        self.client_cert_edit.setVisible(cert_visible)
        self.cert_browse_btn1.setVisible(cert_visible)
        self.client_key_label.setVisible(cert_visible)
        self.client_key_edit.setVisible(cert_visible)
        self.cert_browse_btn2.setVisible(cert_visible)

        # 調整視窗大小以適應顯示的欄位
        self.adjustSize()

    def on_chk_show_supported_toggled(self, checked: bool):
        """檢測或隱藏伺服器不支援的安全模式"""
        if not checked:
            # 取消勾選時顯示所有模式
            self._show_all_policies_and_modes()
            return

        # 需要 OPC URL 才能進行檢測 - 無 OPC URL 時不進行檢測
        if not self.opc_url:
            return

        # 禁用複選框並開始檢測
        self.chk_show_supported.setEnabled(False)

        # 在非同步任務中執行檢測
        import asyncio
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                asyncio.create_task(self._detect_server_capabilities(self.opc_url))
            else:
                asyncio.run(self._detect_server_capabilities(self.opc_url))
        except RuntimeError:
            asyncio.run(self._detect_server_capabilities(self.opc_url))

    async def _detect_server_capabilities(self, opc_url: str):
        """偵測伺服器支援的安全策略和模式
        
        透過連接伺服器並解析日誌與異常信息來獲取支援的策略
        """
        supported_policies = set()
        supported_modes = set()
        error_message = None
        
        try:
            from asyncua import Client
            from asyncua.ua.uaerrors import UaError
            import logging
            import io
            
            # 攔截 asyncua 日誌以提取端點信息
            log_capture = io.StringIO()
            handler = logging.StreamHandler(log_capture)
            handler.setFormatter(logging.Formatter('%(message)s'))
            asyncua_logger = logging.getLogger('asyncua.client.client')
            asyncua_logger.addHandler(handler)
            
            try:
                client = Client(opc_url)
                await client.connect()
                endpoints = await client.get_endpoints()
                await client.disconnect()
                
                # 成功連接並獲取端點
                for idx, ep in enumerate(endpoints):
                    # 提取安全策略
                    uri = getattr(ep, "SecurityPolicyUri", None)
                    if uri:
                        frag = uri.split("#")[-1] if "#" in str(uri) else str(uri).rstrip('/').split('/')[-1]
                        norm = self._normalize_policy_name(frag)
                        if norm and norm != "":
                            supported_policies.add(norm)
                            print(f"[OPC UA 檢測] 策略: {norm}")
                    
                    # 提取安全模式
                    mode = getattr(ep, "SecurityMode", None)
                    if mode is not None:
                        name = getattr(mode, "name", None) or str(mode)
                        norm_mode = self._normalize_mode_name(str(name))
                        if norm_mode and norm_mode != "":
                            supported_modes.add(norm_mode)
                            print(f"[OPC UA 檢測] 模式: {norm_mode}")
                            
            except UaError as ua_exc:
                # 連接失敗，嘗試從日誌和異常消息中提取信息
                error_message = str(ua_exc)
                print(f"[OPC UA 檢測] 連接異常: {error_message}")
                
                # 從日誌中提取所有 SecurityPolicyUri
                log_content = log_capture.getvalue()
                print(f"[OPC UA 檢測] 日誌內容長度: {len(log_content)}")
                
                # 使用正則表達式查找所有 SecurityPolicyUri
                import re
                uri_matches = re.findall(r"SecurityPolicyUri='([^']+)'", log_content)
                print(f"[OPC UA 檢測] 從日誌找到 {len(uri_matches)} 個策略 URI: {set(uri_matches)}")
                
                for uri in uri_matches:
                    frag = uri.split("#")[-1] if "#" in uri else uri.rstrip('/').split('/')[-1]
                    norm = self._normalize_policy_name(frag)
                    if norm and norm != "" and norm != "None":  # 除外 None 可能只是認證策略
                        supported_policies.add(norm)
                        print(f"[OPC UA 檢測] 從日誌提取策略: {norm}")
                
                # 同時從日誌中提取 SecurityMode
                # 格式通常是: SecurityMode=<MessageSecurityMode.Sign: 2> 或 SecurityMode=<MessageSecurityMode.SignAndEncrypt: 3>
                mode_matches = re.findall(r"SecurityMode=<MessageSecurityMode\.(\w+(?:And\w+)?)", log_content)
                print(f"[OPC UA 檢測] 從日誌找到 {len(mode_matches)} 個安全模式: {set(mode_matches)}")
                
                for mode_str in set(mode_matches):  # 使用 set 避免重複
                    norm_mode = self._normalize_mode_name(mode_str)
                    if norm_mode and norm_mode != "":
                        supported_modes.add(norm_mode)
                        print(f"[OPC UA 檢測] 從日誌提取模式: {norm_mode}")

                
                # 如果還是沒有找到策略，至少報告所有找到的內容
                if not supported_policies:
                    print("[OPC UA 檢測] 警告: 未找到有效的安全策略")
                    
            finally:
                asyncua_logger.removeHandler(handler)
                log_capture.close()
                
        except Exception as exc:
            error_message = str(exc)
            print(f"[OPC UA 檢測] 未知錯誤: {error_message}")

        # 更新 UI
        def update_ui():
            self.chk_show_supported.setEnabled(True)
            
            if supported_policies or supported_modes:
                self._detected_supported = {"policies": supported_policies, "modes": supported_modes}
                self._apply_supported_filters()
            else:
                self._show_all_policies_and_modes()

        QTimer.singleShot(0, update_ui)
    def _apply_supported_filters(self):
        """套用伺服器支援的安全模式過濾"""
        data = self._detected_supported or {}
        policies = data.get("policies", set())
        modes = data.get("modes", set())

        # 只有當檢測到支援的模式時才進行過濾
        # 如果檢測到空集合，則顯示所有（伺服器可能不報告這些資訊）
        if not policies and not modes:
            self._show_all_policies_and_modes()
            return

        # 控制安全策略單選按鈕的可見性
        policy_buttons = [
            (self.policy_rb_none, "None"),
            (self.policy_rb_basic128, "Basic128Rsa15"),
            (self.policy_rb_basic256, "Basic256"),
            (self.policy_rb_basic256sha, "Basic256Sha256"),
        ]
        for btn, policy_name in policy_buttons:
            btn.setVisible(policy_name in policies)

        # 控制安全模式單選按鈕的可見性
        mode_buttons = [
            (self.rb_mode_none, "None"),
            (self.rb_mode_sign, "Sign"),
            (self.rb_mode_sign_encrypt, "SignAndEncrypt"),
        ]
        for btn, mode_name in mode_buttons:
            btn.setVisible(mode_name in modes)

    def _show_all_policies_and_modes(self):
        """顯示所有安全策略和模式（沒有過濾）"""
        # 顯示所有安全策略
        for btn in [self.policy_rb_none, self.policy_rb_basic128, 
                    self.policy_rb_basic256, self.policy_rb_basic256sha]:
            btn.setVisible(True)
        
        # 顯示所有安全模式
        for btn in [self.rb_mode_none, self.rb_mode_sign, self.rb_mode_sign_encrypt]:
            btn.setVisible(True)

    def _normalize_policy_name(self, fragment: str) -> str:
        """將各種 SecurityPolicy 片段標準化為 UI 使用的規範名稱
        
        例： 
            Basic128RSA15 -> Basic128Rsa15
            Basic256Sha256 -> Basic256Sha256
            None -> None
        """
        if not fragment:
            return ""
        
        # 轉為小寫並移除非字母數字字元進行比對
        normalized = re.sub(r'[^0-9a-z]', '', fragment.lower())
        
        # 根據關鍵字識別策略
        if normalized == "none":
            return "None"
        if "128" in normalized:
            return "Basic128Rsa15"
        if "sha256" in normalized or ("sha" in normalized and "256" in normalized):
            return "Basic256Sha256"
        if "256" in normalized and "128" not in normalized:
            return "Basic256"
        
        # 無法識別時，嘗試返回原始片段
        return fragment if fragment in ["None", "Basic128Rsa15", "Basic256", "Basic256Sha256"] else ""

    def _normalize_mode_name(self, name: str) -> str:
        """將安全模式名稱標準化為規範鍵值: None, Sign, SignAndEncrypt"""
        if not name:
            return ""
        
        # 移除前綴和特殊字元，轉為小寫進行比對
        cleaned = name.lower()
        cleaned = cleaned.replace("messagesecuritymode.", "")
        cleaned = cleaned.replace("_", "").replace(" ", "")
        
        # 識別模式（注意順序，SignAndEncrypt 要在 Sign 之前）
        if "signandencrypt" in cleaned or "signencrypt" in cleaned:
            return "SignAndEncrypt"
        if "sign" in cleaned and "encrypt" not in cleaned:
            return "Sign"
        if "none" in cleaned:
            return "None"
        
        # 無法識別
        return ""

    def _on_apply(self):
        """按下 Apply 時接受設定對話框"""
        self.accept()


class ScheduleEditDialog(QDialog):
    """排程編輯對話框"""

    def __init__(
        self,
        parent=None,
        schedule: Dict[str, Any] = None,
        default_date: Optional[QDate] = None,
        default_hour: Optional[int] = None,
        default_minute: int = 0,
    ):
        super().__init__(parent)
        self.schedule = schedule
        self.original_rrule = ""  # 儲存原始的 RRULE 字串

        # 從主行事曆帶入的預設日期/時間（例如右鍵點選的格子）
        self.default_date: Optional[QDate] = default_date
        self.default_hour: Optional[int] = default_hour
        self.default_minute: int = max(0, min(59, default_minute))

        # 取得資料庫管理器
        self.db_manager = parent.db_manager if parent and hasattr(parent, 'db_manager') else None

        # 初始化OPC設定
        self.opc_security_policy = schedule.get("opc_security_policy", "None") if schedule else "None"
        self.opc_security_mode = schedule.get("opc_security_mode", "None") if schedule else "None"
        self.opc_username = schedule.get("opc_username", "") if schedule else ""
        self.opc_password = schedule.get("opc_password", "") if schedule else ""
        self.opc_timeout = schedule.get("opc_timeout", 5) if schedule else 5
        self.opc_write_timeout = schedule.get("opc_write_timeout", 3) if schedule else 3
        self.is_enabled = schedule.get("is_enabled", 1) if schedule else 1
        self._last_opc_url = ""

        if not schedule and self.db_manager:
            self._load_last_opc_defaults()

        self.setup_ui()
        self.apply_style()

        # 如果是新增模式，設置預設任務名稱
        if not schedule and parent and hasattr(parent, 'db_manager'):
            default_name = parent.db_manager.get_next_task_name()
            self.task_name_edit.setText(default_name)

        if schedule:
            self.load_data()

    def setup_ui(self):
        self.setWindowTitle("編輯排程" if self.schedule else "新增排程")
        self.setWindowIcon(get_app_icon())
        self.setMinimumWidth(780)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # 基本資訊
        basic_group = QGroupBox("基本資訊")
        basic_layout = QGridLayout(basic_group)

        basic_layout.addWidget(QLabel("任務名稱:"), 0, 0)
        self.task_name_edit = QLineEdit()
        self.task_name_edit.setPlaceholderText("例如：每日早班開機")
        basic_layout.addWidget(self.task_name_edit, 0, 1)

        basic_layout.addWidget(QLabel("OPC URL:"), 1, 0)
        opc_url_layout = QHBoxLayout()
        opc_url_layout.setSpacing(5)
        self.opc_url_edit = QLineEdit()
        self.opc_url_edit.setPlaceholderText("localhost:4840")
        if self._last_opc_url:
            self.opc_url_edit.setText(self._last_opc_url)
        opc_url_layout.addWidget(self.opc_url_edit)
        # 添加協議標籤顯示
        self.opc_protocol_label = QLabel("opc.tcp://")
        self.opc_protocol_label.setStyleSheet("color: #666; padding-right: 5px;")
        opc_url_layout.insertWidget(0, self.opc_protocol_label)
        # 添加OPC設定按鈕
        self.btn_opc_settings = QPushButton("設定...")
        self.btn_opc_settings.setToolTip("OPC UA 連線設定")
        self.btn_opc_settings.clicked.connect(self.configure_opc_settings)
        self.btn_opc_settings.setMaximumWidth(80)
        opc_url_layout.addWidget(self.btn_opc_settings)
        basic_layout.addLayout(opc_url_layout, 1, 1)

        basic_layout.addWidget(QLabel("Node ID:"), 2, 0)
        node_id_layout = QHBoxLayout()
        node_id_layout.setSpacing(5)
        self.node_id_edit = QLineEdit()
        self.node_id_edit.setPlaceholderText("ns=2;i=1001")
        node_id_layout.addWidget(self.node_id_edit)
        # 添加瀏覽按鈕
        self.btn_browse_node = QPushButton("瀏覽...")
        self.btn_browse_node.setToolTip("瀏覽 OPC UA 節點")
        self.btn_browse_node.clicked.connect(self.browse_opcua_nodes)
        self.btn_browse_node.setMaximumWidth(80)
        node_id_layout.addWidget(self.btn_browse_node)
        basic_layout.addLayout(node_id_layout, 2, 1)

        basic_layout.addWidget(QLabel("目標值:"), 3, 0)
        target_layout = QHBoxLayout()
        self.target_value_edit = QLineEdit()
        self.target_value_edit.setPlaceholderText("1")
        target_layout.addWidget(self.target_value_edit)
        basic_layout.addLayout(target_layout, 3, 1)
        
        # 型別顯示 - 簡單文字標籤
        basic_layout.addWidget(QLabel("型別:"), 4, 0)
        type_layout = QHBoxLayout()
        self.data_type_label = QLabel("未偵測")
        self.data_type_label.setStyleSheet("font-weight: bold; color: #0066cc;")
        type_layout.addWidget(self.data_type_label)
        type_layout.addStretch()
        basic_layout.addLayout(type_layout, 4, 1)

        basic_layout.addWidget(QLabel("狀態:"), 5, 0)
        status_layout = QHBoxLayout()
        status_layout.setSpacing(14)
        self.enabled_checkbox = QCheckBox("啟用排程")
        self.enabled_checkbox.setChecked(True)  # 預設啟用
        self.enabled_checkbox.setToolTip("控制此排程是否會被執行")
        status_layout.addWidget(self.enabled_checkbox)

        self.ignore_holiday_checkbox = QCheckBox("忽略假日")
        self.ignore_holiday_checkbox.setChecked(False)
        self.ignore_holiday_checkbox.setToolTip("勾選後此排程不套用假日覆寫邏輯")
        status_layout.addWidget(self.ignore_holiday_checkbox)
        status_layout.addStretch()
        basic_layout.addLayout(status_layout, 5, 1)

        layout.addWidget(basic_group)

        # 週期設定
        recurrence_group = QGroupBox("週期設定")
        recurrence_layout = QVBoxLayout(recurrence_group)

        initial_time = (
            QTime(self.default_hour, self.default_minute, 0)
            if self.default_hour is not None
            else None
        )
        self.recurrence_editor = RecurrenceDialog(
            self,
            current_rrule=self.schedule.get("rrule_str", "") if self.schedule else "",
            initial_date=self.default_date,
            initial_time=initial_time,
            embedded=True,
        )
        recurrence_layout.addWidget(self.recurrence_editor)

        layout.addWidget(recurrence_group)

        # 按鈕
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        ok_btn = QPushButton("確定")
        ok_btn.setDefault(True)
        ok_btn.clicked.connect(self.on_ok_clicked)
        button_layout.addWidget(ok_btn)

        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)

    def apply_style(self):
        """套用樣式，根據父視窗主題選擇亮色或暗色"""
        # 判斷是否使用暗色模式
        is_dark = False
        parent = self.parent()
        if parent and hasattr(parent, "current_theme"):
            if parent.current_theme == "dark":
                is_dark = True
            elif parent.current_theme == "system":
                if hasattr(parent, "is_system_dark_mode"):
                    is_dark = parent.is_system_dark_mode()

        if is_dark:
            self._apply_dark_style()
        else:
            self._apply_light_style()

    def _apply_light_style(self):
        """套用亮色樣式"""
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
                color: #2c3e50;
            }
            QLabel {
                color: #333;
            }
            QPushButton {
                background-color: #e9ecef;
                color: #111111;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #c7d4e2;
            }
            QLineEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 6px;
                background-color: white;
                color: #333;
            }
            QLineEdit:focus {
                border: 2px solid #0078d4;
            }
            QRadioButton {
                color: #333;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #999999;
                border-radius: 9px;
                background-color: white;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0078d4;
                background-color: #f0f0f0;
            }
            QRadioButton::indicator:checked {
                background-color: #0078d4;
                border: 2px solid #0078d4;
            }
            QCheckBox {
                color: #333;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #999999;
                border-radius: 2px;
                background-color: white;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #f0f0f0;
            }
            QCheckBox::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
                image: url(:/checkbox_check);
            }
        """)

    def _apply_dark_style(self):
        """套用暗色樣式"""
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
            }
            QGroupBox::title {
                color: #ffffff;
            }
            QLabel {
                color: #cccccc;
            }
            QPushButton {
                background-color: #0e639c;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1f89cd;
            }
            QLineEdit {
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                background-color: #1e1e1e;
                color: #cccccc;
            }
            QLineEdit:focus {
                border: 2px solid #0e639c;
            }
            QRadioButton {
                color: #cccccc;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #666666;
                border-radius: 9px;
                background-color: #1e1e1e;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #252526;
            }
            QRadioButton::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
            }
            QCheckBox {
                color: #cccccc;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #666666;
                border-radius: 2px;
                background-color: #1e1e1e;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #252526;
            }
            QCheckBox::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
                image: url(:/checkbox_check);
            }
        """)

    def load_data(self):
        """載入現有資料"""
        self.task_name_edit.setText(self.schedule.get("task_name", ""))
        # 提取 ip:port 部分（去掉 opc.tcp:// 前綴）
        opc_url = self.schedule.get("opc_url", "")
        if opc_url.startswith("opc.tcp://"):
            opc_url = opc_url[10:]  # 去掉 "opc.tcp://"
        self.opc_url_edit.setText(opc_url)
        self.node_id_edit.setText(self.schedule.get("node_id", ""))
        self.target_value_edit.setText(self.schedule.get("target_value", ""))
        data_type = self.schedule.get("data_type", "auto")
        # 如果是"auto"，顯示為"未偵測"
        display_data_type = "未偵測" if data_type == "auto" else data_type
        self.data_type_label.setText(display_data_type)
        
        # 儲存原始 RRULE 字串
        self.original_rrule = self.schedule.get("rrule_str", "")
        
        self.enabled_checkbox.setChecked(bool(self.schedule.get("is_enabled", 1)))
        self.ignore_holiday_checkbox.setChecked(bool(self.schedule.get("ignore_holiday", 0)))
        self.recurrence_editor.set_lock_enabled(bool(self.schedule.get("lock_enabled", 0)))

    def get_data(self) -> Dict[str, Any]:
        """取得編輯的資料"""
        # 自動添加 opc.tcp:// 前綴
        opc_url = self.opc_url_edit.text().strip()
        if opc_url and not opc_url.startswith("opc.tcp://"):
            opc_url = f"opc.tcp://{opc_url}"
        
        return {
            "task_name": self.task_name_edit.text(),
            "opc_url": opc_url,
            "node_id": self.node_id_edit.text(),
            "target_value": self.target_value_edit.text(),
            # 處理資料型別：如果顯示"未偵測"，儲存為"auto"
            "data_type": "auto" if self.data_type_label.text() == "未偵測" else self.data_type_label.text(),
            "rrule_str": self.recurrence_editor.get_rrule(),
            "category_id": 1,
            "opc_security_policy": self.opc_security_policy,
            "opc_security_mode": self.opc_security_mode,
            "opc_username": self.opc_username,
            "opc_password": self.opc_password,
            "opc_timeout": self.opc_timeout,
            "opc_write_timeout": self.opc_write_timeout,
            "lock_enabled": 1 if self.recurrence_editor.get_lock_enabled() else 0,
            "is_enabled": 1 if self.enabled_checkbox.isChecked() else 0,
            "ignore_holiday": 1 if self.ignore_holiday_checkbox.isChecked() else 0,
        }

    def on_ok_clicked(self):
        """確定按鈕點擊處理"""
        # 檢查任務名稱
        task_name = self.task_name_edit.text().strip()
        if not task_name:
            QMessageBox.warning(
                self,
                "任務名稱未設定",
                "請輸入任務名稱。",
            )
            return

        # 檢查 OPC URL
        opc_url = self.opc_url_edit.text().strip()
        if not opc_url:
            QMessageBox.warning(
                self,
                "OPC URL 未設定",
                "請輸入 OPC URL。",
            )
            return

        # 檢查 Node ID
        node_id = self.node_id_edit.text().strip()
        if not node_id:
            QMessageBox.warning(
                self,
                "Node ID 未設定",
                "請輸入或瀏覽選擇 Node ID。",
            )
            return

        # 檢查目標值
        target_value = self.target_value_edit.text().strip()
        if not target_value:
            QMessageBox.warning(
                self,
                "目標值未設定",
                "請輸入目標值。",
            )
            return

        # 檢查 rrule 是否為空
        try:
            rrule_str = self.recurrence_editor.get_rrule().strip()
        except Exception as e:
            QMessageBox.warning(
                self,
                "週期規則錯誤",
                f"週期規則設定無效：{str(e)}",
            )
            return

        if not rrule_str:
            QMessageBox.warning(
                self,
                "週期規則未設定",
                "請設定排程的週期規則，無法儲存空的週期規則。",
            )
            return

        self.original_rrule = rrule_str

        # 記住本次使用的 OPC URL 與安全設定，供下次新增排程預設帶入
        if self.db_manager:
            self.db_manager.save_last_opc_defaults(
                {
                    "opc_url": self._normalize_opc_url(),
                    "opc_security_policy": self.opc_security_policy,
                    "opc_security_mode": self.opc_security_mode,
                    "opc_username": self.opc_username,
                    "opc_password": self.opc_password,
                    "opc_timeout": self.opc_timeout,
                    "opc_write_timeout": self.opc_write_timeout,
                }
            )

        # 如果檢查通過，接受對話框
        self.accept()

    def browse_opcua_nodes(self):
        """瀏覽 OPC UA 節點"""
        opc_url = self._normalize_opc_url()

        if not opc_url:
            QMessageBox.warning(
                self,
                "警告",
                "請先輸入 OPC URL (IP:Port)",
            )
            return

        # 開啟節點瀏覽對話框
        dialog = OPCNodeBrowserDialog(self, opc_url, self.node_id_edit.text().strip())
        if dialog.exec() == QDialog.Accepted:
            selected_node = dialog.get_selected_node()
            if selected_node:
                # 解析選擇的節點資訊: display_name|node_id|data_type
                parts = selected_node.split("|")
                if len(parts) >= 2:
                    display_name = parts[0]
                    node_id = parts[1]
                    data_type = parts[2] if len(parts) > 2 else "未知"
                    
                    # 設定節點 ID 和自動偵測的資料型別
                    self.node_id_edit.setText(f"{display_name}|{node_id}")
                    self.data_type_label.setText(data_type)

    def configure_opc_settings(self):
        """設定 OPC UA 連線參數"""
        opc_url = self._normalize_opc_url()

        dialog = OPCSettingsDialog(
            self,
            self.opc_security_policy,
            self.opc_username,
            self.opc_password,
            self.opc_timeout,
            self.opc_write_timeout,
            self.opc_security_mode,
            opc_url=opc_url,
        )

        if dialog.exec() == QDialog.Accepted:
            settings = dialog.get_settings()
            self.opc_security_policy = settings["security_policy"]
            self.opc_security_mode = settings["security_mode"]
            self.opc_username = settings["username"]
            self.opc_password = settings["password"]
            self.opc_timeout = settings["timeout"]
            self.opc_write_timeout = settings["write_timeout"]

    def _normalize_opc_url(self) -> str:
        """標準化 OPC URL，確保以 opc.tcp:// 開頭"""
        url = self.opc_url_edit.text().strip()
        if not url:
            return ""

        # 如果沒有協議前綴，添加預設的 opc.tcp://
        if not url.startswith(("opc.tcp://", "opc.https://", "opc.wss://")):
            url = f"opc.tcp://{url}"

        return url

    def _load_last_opc_defaults(self):
        """新增排程時讀取上一次使用的 OPC 設定。"""
        if not self.db_manager:
            return

        defaults = self.db_manager.get_last_opc_defaults() or {}
        last_opc_url = (defaults.get("opc_url") or "").strip()
        if last_opc_url.startswith("opc.tcp://"):
            last_opc_url = last_opc_url[10:]

        self._last_opc_url = last_opc_url
        self.opc_security_policy = defaults.get("opc_security_policy", self.opc_security_policy)
        self.opc_security_mode = defaults.get("opc_security_mode", self.opc_security_mode)
        self.opc_username = defaults.get("opc_username", self.opc_username)
        self.opc_password = defaults.get("opc_password", self.opc_password)
        self.opc_timeout = int(defaults.get("opc_timeout", self.opc_timeout) or self.opc_timeout)
        self.opc_write_timeout = int(defaults.get("opc_write_timeout", self.opc_write_timeout) or self.opc_write_timeout)


def main():
    """主程式進入點"""
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # 設定應用程式資訊
    app.setApplicationName("CalendarUA")
    app.setApplicationVersion("1.0.0")

    # 建立事件迴圈
    loop = qasync.QEventLoop(app)
    asyncio.set_event_loop(loop)

    # 建立主視窗
    window = CalendarUA()
    window.show()
    QTimer.singleShot(0, window.force_show_on_screen)
    QTimer.singleShot(300, window.force_show_on_screen)

    # 執行事件迴圈
    with loop:
        try:
            loop.run_forever()
        except KeyboardInterrupt:
            logger.info("收到中斷訊號，正在結束程式")


if __name__ == "__main__":
    main()

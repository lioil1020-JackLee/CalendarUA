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
    QCheckBox,
    QPushButton,
    QDateEdit,
    QCalendarWidget,
    QGroupBox,
    QRadioButton,
    QButtonGroup,
    QGridLayout,
    QFrame,
    QMessageBox,
    QToolButton,
    QAbstractSpinBox,
    QSizePolicy,
    QStyle,
)
from PySide6.QtCore import Qt, QDate, QTime, Signal, QEvent, QSize, QLocale, QPoint, QTimer
from PySide6.QtGui import QFont, QColor, QGuiApplication
import sys
from datetime import date as dt_date

from core.lunar_calendar import to_lunar, format_lunar_day_text
from ui.app_icon import get_app_icon


def _combo_steps_from_wheel(event) -> int:
    delta = event.angleDelta().y()
    if delta == 0:
        return 0
    steps = int(delta / 120)
    if steps == 0:
        steps = 1 if delta > 0 else -1
    return steps


class DropdownNavCalendar(QCalendarWidget):
    """自訂導覽列：上一月 / 年下拉 / 今日 / 月下拉 / 下一月。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._holiday_checker = None
        self._is_dark_theme = False
        self.setNavigationBarVisible(False)
        self.setGridVisible(True)
        self.setFirstDayOfWeek(Qt.Sunday)
        self.setLocale(QLocale(QLocale.Chinese, QLocale.Taiwan))
        self.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.setHorizontalHeaderFormat(QCalendarWidget.ShortDayNames)
        self.setMinimumSize(280, 270)

        header_widget = QWidget(self)
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(2, 2, 2, 2)
        header_layout.setSpacing(0)

        self.btn_prev = QToolButton(header_widget)
        self.btn_prev.setIcon(self.style().standardIcon(QStyle.SP_ArrowLeft))
        self.btn_prev.setAutoRaise(True)
        self.btn_prev.setIconSize(QSize(20, 20))
        self.btn_prev.setFixedSize(32, 28)

        self.btn_next = QToolButton(header_widget)
        self.btn_next.setIcon(self.style().standardIcon(QStyle.SP_ArrowRight))
        self.btn_next.setAutoRaise(True)
        self.btn_next.setIconSize(QSize(20, 20))
        self.btn_next.setFixedSize(32, 28)

        self.combo_year = QComboBox(header_widget)
        self.combo_month = QComboBox(header_widget)
        self.combo_year.setEditable(True)
        self.combo_month.setEditable(True)
        if self.combo_year.lineEdit() is not None:
            self.combo_year.lineEdit().setReadOnly(True)
            self.combo_year.lineEdit().setAlignment(Qt.AlignCenter)
            self.combo_year.lineEdit().setCursor(Qt.PointingHandCursor)
            self.combo_year.lineEdit().installEventFilter(self)
        if self.combo_month.lineEdit() is not None:
            self.combo_month.lineEdit().setReadOnly(True)
            self.combo_month.lineEdit().setAlignment(Qt.AlignCenter)
            self.combo_month.lineEdit().setCursor(Qt.PointingHandCursor)
            self.combo_month.lineEdit().installEventFilter(self)
        self.combo_year.setFixedWidth(72)
        self.combo_month.setFixedWidth(60)
        self.combo_year.setCursor(Qt.PointingHandCursor)
        self.combo_month.setCursor(Qt.PointingHandCursor)
        self.combo_year.setMaxVisibleItems(11)
        self.combo_month.setMaxVisibleItems(12)
        self.combo_year.view().setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.combo_year.view().setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.combo_month.view().setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.combo_month.view().setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.combo_year.setStyleSheet(
            """
            QComboBox {
                font-family: 'Segoe UI';
                font-size: 15px;
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
                width: 0px;
                border: none;
            }
            QComboBox::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
            }
            """
        )
        self.combo_month.setStyleSheet(self.combo_year.styleSheet())

        self.btn_today = QToolButton(header_widget)
        self.btn_today.setText("●")
        self.btn_today.setAutoRaise(True)
        self.btn_today.setFixedSize(22, 22)
        self.btn_today.setToolTip("跳到今天")

        header_layout.addStretch()
        header_layout.addWidget(self.btn_prev)
        header_layout.addWidget(self.combo_year)
        header_layout.addWidget(self.btn_today)
        header_layout.addWidget(self.combo_month)
        header_layout.addWidget(self.btn_next)
        header_layout.addStretch()

        calendar_layout = self.layout()
        if calendar_layout is not None and hasattr(calendar_layout, "setMenuBar"):
            calendar_layout.setMenuBar(header_widget)

        self._init_nav_values()
        self._sync_nav_from_page()

        self.btn_prev.clicked.connect(lambda: self._shift_month(-1))
        self.btn_next.clicked.connect(lambda: self._shift_month(1))
        self.btn_today.clicked.connect(self._go_today)
        self.combo_year.currentIndexChanged.connect(self._apply_page_from_nav)
        self.combo_month.currentIndexChanged.connect(self._apply_page_from_nav)
        self.currentPageChanged.connect(lambda _year, _month: self._sync_nav_from_page())
        self.combo_month.installEventFilter(self)
        self.combo_month.view().installEventFilter(self)
        self.combo_month.view().viewport().installEventFilter(self)
        self.combo_year.installEventFilter(self)
        self.combo_year.view().installEventFilter(self)
        self.combo_year.view().viewport().installEventFilter(self)
        self.apply_theme(False)

    def apply_theme(self, is_dark: bool):
        self._is_dark_theme = bool(is_dark)
        weekday_color = "#ffffff" if is_dark else "#111111"
        body_color = "#cccccc" if is_dark else "#111111"
        header_bg = "#363636" if is_dark else "#e0e0e0"
        calendar_bg = "#2b2b2b" if is_dark else "#ffffff"
        selected_bg = "#0e639c" if is_dark else "#9ec6f3"
        selected_fg = "#ffffff" if is_dark else "#0f1f33"
        self.setStyleSheet(
            f"""
            QCalendarWidget QWidget {{
                background-color: {calendar_bg};
                color: {body_color};
            }}
            QCalendarWidget QAbstractItemView:enabled {{
                background-color: {calendar_bg};
                color: {body_color};
                selection-background-color: {selected_bg};
                selection-color: {selected_fg};
            }}
            QCalendarWidget QTableView QHeaderView::section {{
                background-color: {header_bg};
                color: {weekday_color};
                font-weight: 600;
            }}
            QToolButton {{
                color: {body_color};
                background-color: transparent;
                border: none;
            }}
            """
        )
        weekday_fmt = self.weekdayTextFormat(Qt.Monday)
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
            self.setWeekdayTextFormat(day, weekday_fmt)

    def set_holiday_checker(self, checker):
        self._holiday_checker = checker
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

    def paintCell(self, painter, rect, date):
        shown_year = self.yearShown()
        shown_month = self.monthShown()
        is_this = (date.year() == shown_year and date.month() == shown_month)

        painter.save()
        cell_bg = QColor("#2b2b2b") if self._is_dark_theme else QColor("#ffffff")
        painter.fillRect(rect, cell_bg)

        is_dark_palette = self.palette().window().color().lightness() < 128
        is_holiday = self._is_holiday(date)

        lunar_text = ""
        try:
            lunar_info = to_lunar(dt_date(date.year(), date.month(), date.day()))
            if lunar_info:
                lunar_text = format_lunar_day_text(lunar_info)
        except Exception:
            lunar_text = ""

        if is_this:
            if is_holiday:
                day_fg = QColor("#ff6b6b") if is_dark_palette else QColor("#c62828")
            else:
                day_fg = QColor("#f0f0f0") if is_dark_palette else QColor("#202020")
        else:
            day_fg = QColor("#b36b6b") if is_holiday else QColor("#808080")

        # 今日框線
        if date == QDate.currentDate():
            today_pen = QColor("#ff8f00")
            painter.setPen(today_pen)
            painter.setBrush(Qt.NoBrush)
            today_rect = rect.adjusted(3, 3, -3, -3)
            painter.drawRect(today_rect)

        if date == self.selectedDate():
            sel = QColor("#0078d7") if is_dark_palette else QColor("#9ec6f3")
            painter.setPen(Qt.NoPen)
            painter.setBrush(sel)
            r = rect.adjusted(2, 2, -2, -2)
            painter.drawRect(r)

        painter.setPen(day_fg)
        solar_font = QFont(painter.font())
        solar_font.setFamily("Segoe UI")
        solar_font.setBold(True)
        solar_font.setPointSize(12)
        painter.setFont(solar_font)
        top_rect = rect.adjusted(0, 2, 0, -int(rect.height() * 0.48))
        painter.drawText(top_rect, Qt.AlignHCenter | Qt.AlignVCenter, str(date.day()))

        if lunar_text:
            lunar_font = QFont(painter.font())
            lunar_font.setFamily("Microsoft JhengHei")
            lunar_font.setBold(False)
            lunar_font.setPointSize(8)
            painter.setFont(lunar_font)
            bottom_rect = rect.adjusted(0, int(rect.height() * 0.50), 0, -1)
            painter.drawText(bottom_rect, Qt.AlignHCenter | Qt.AlignVCenter, lunar_text)

        painter.restore()

    def _init_nav_values(self):
        current_year = QDate.currentDate().year()
        self._set_year_window(current_year, current_year)

        self.combo_month.clear()
        for month in range(1, 13):
            self.combo_month.addItem(f"{month}月", month)

    def _shift_month(self, delta: int):
        page_date = QDate(self.yearShown(), self.monthShown(), 1).addMonths(delta)
        self.setCurrentPage(page_date.year(), page_date.month())

    def _go_today(self):
        today = QDate.currentDate()
        self.setSelectedDate(today)
        self.setCurrentPage(today.year(), today.month())

    def _sync_nav_from_page(self):
        year = self.yearShown()
        month = self.monthShown()

        self._ensure_year_available(year)

        year_index = self.combo_year.findData(year)
        month_index = self.combo_month.findData(month)

        self.combo_year.blockSignals(True)
        self.combo_month.blockSignals(True)
        if year_index >= 0:
            self.combo_year.setCurrentIndex(year_index)
        if month_index >= 0:
            self.combo_month.setCurrentIndex(month_index)
        self.combo_year.blockSignals(False)
        self.combo_month.blockSignals(False)

    def _set_year_window(self, center_year: int, selected_year: int | None = None):
        start_year = center_year - 5
        years = list(range(start_year, start_year + 11))

        target_year = selected_year if isinstance(selected_year, int) else center_year
        if target_year < years[0]:
            target_year = years[0]
        elif target_year > years[-1]:
            target_year = years[-1]

        self.combo_year.blockSignals(True)
        self.combo_year.clear()
        for y in years:
            self.combo_year.addItem(str(y), y)

        idx = self.combo_year.findData(target_year)
        if idx >= 0:
            self.combo_year.setCurrentIndex(idx)
        self.combo_year.blockSignals(False)

    def _ensure_year_available(self, year: int) -> int:
        idx = self.combo_year.findData(year)
        if idx >= 0:
            return idx

        self._set_year_window(year, year)
        return self.combo_year.findData(year)

    def _shift_year_window_by_steps(self, steps: int):
        if steps == 0:
            return

        center_idx = min(5, max(0, self.combo_year.count() - 1))
        center_year = self.combo_year.itemData(center_idx)
        if not isinstance(center_year, int):
            center_year = self.yearShown()

        selected_year = self.combo_year.currentData()
        if not isinstance(selected_year, int):
            selected_year = center_year

        self._set_year_window(center_year + steps, selected_year)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            year_line = self.combo_year.lineEdit()
            month_line = self.combo_month.lineEdit()
            if obj in (self.combo_year, year_line):
                QTimer.singleShot(0, self.combo_year.showPopup)
                event.accept()
                return True
            if obj in (self.combo_month, month_line):
                QTimer.singleShot(0, self.combo_month.showPopup)
                event.accept()
                return True

        if obj in (self.combo_year, self.combo_year.view(), self.combo_year.view().viewport()) and event.type() == QEvent.Wheel:
            delta = event.angleDelta().y()
            if delta != 0:
                steps = int(delta / 120)
                if steps == 0:
                    steps = 1 if delta > 0 else -1
                self._shift_year_window_by_steps(-steps)
            event.accept()
            return True

        if obj in (self.combo_month, self.combo_month.view(), self.combo_month.view().viewport()) and event.type() == QEvent.Wheel:
            steps = _combo_steps_from_wheel(event)
            if steps != 0 and self.combo_month.count() > 0:
                current_index = self.combo_month.currentIndex()
                if current_index < 0:
                    current_index = 0
                target_index = current_index - steps
                if target_index < 0:
                    target_index = 0
                elif target_index >= self.combo_month.count():
                    target_index = self.combo_month.count() - 1
                if target_index != self.combo_month.currentIndex():
                    self.combo_month.setCurrentIndex(target_index)
            event.accept()
            return True

        return super().eventFilter(obj, event)

    def _apply_page_from_nav(self):
        year = self.combo_year.currentData()
        month = self.combo_month.currentData()
        if isinstance(year, int) and isinstance(month, int):
            self.setCurrentPage(year, month)


class PopupDateEdit(QDateEdit):
    """移除右側箭頭，點擊日期欄位直接展開月曆。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCalendarPopup(False)
        self.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self.setReadOnly(True)
        self.setCursor(Qt.PointingHandCursor)
        self._calendar_popup = DropdownNavCalendar(self)
        self._calendar_popup.setWindowFlags(Qt.Popup)
        self._calendar_popup.clicked.connect(self._on_calendar_date_clicked)
        self.setStyleSheet(
            """
            QDateEdit::drop-down {
                width: 0px;
                border: none;
                padding: 0px;
                margin: 0px;
            }
            QDateEdit::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
            }
            """
        )

        if self.lineEdit() is not None:
            self.lineEdit().setCursor(Qt.PointingHandCursor)
            self.lineEdit().installEventFilter(self)

    def mousePressEvent(self, event):
        if self.isEnabled() and event.button() == Qt.LeftButton:
            self._show_calendar_popup()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event):
        # 忽略雙擊，避免觸發開啟編輯或其他行為
        event.accept()
        return

    def eventFilter(self, obj, event):
        if self.isEnabled() and obj is self.lineEdit() and event.type() == QEvent.MouseButtonPress:
            self._show_calendar_popup()
            event.accept()
            return True
        return super().eventFilter(obj, event)

    def _show_calendar_popup(self):
        if not self.isEnabled():
            return
        calendar = self._calendar_popup
        if calendar is None:
            return
        calendar.setSelectedDate(self.date())
        min_width = 280
        min_height = 270
        if calendar.width() < min_width or calendar.height() < min_height:
            calendar.resize(max(calendar.width(), min_width), max(calendar.height(), min_height))

        popup_pos = self.mapToGlobal(QPoint(0, self.height()))

        screen = QGuiApplication.screenAt(popup_pos)
        if screen is None:
            screen = QGuiApplication.primaryScreen()
        if screen is not None:
            available = screen.availableGeometry()

            # 優先顯示在輸入框下方，若空間不足則顯示於上方。
            if popup_pos.y() + calendar.height() > available.bottom():
                popup_pos.setY(self.mapToGlobal(QPoint(0, 0)).y() - calendar.height())

            if popup_pos.x() + calendar.width() > available.right():
                popup_pos.setX(max(available.left(), available.right() - calendar.width()))
            if popup_pos.x() < available.left():
                popup_pos.setX(available.left())
            if popup_pos.y() < available.top():
                popup_pos.setY(available.top())

        calendar.move(popup_pos)
        calendar.show()
        calendar.raise_()

    def _on_calendar_date_clicked(self, date: QDate):
        self.setDate(date)
        calendar = self._calendar_popup
        if calendar is not None:
            calendar.hide()


class RollingNumberComboBox(QComboBox):
    """數值下拉（預設顯示 10 個值），支援滑鼠滾輪連續增加/減少。"""

    def __init__(self, minimum: int = 1, maximum: int = 999, parent=None):
        super().__init__(parent)
        self._minimum = minimum
        self._maximum = max(minimum, maximum)
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

    def setMinimum(self, value: int):
        self._minimum = value
        if self._maximum < self._minimum:
            self._maximum = self._minimum
        self.setValue(self._current_value)

    def setMaximum(self, value: int):
        self._maximum = max(self._minimum, value)
        self.setValue(self._current_value)

    def setRange(self, minimum: int, maximum: int):
        self._minimum = minimum
        self._maximum = max(minimum, maximum)
        self.setValue(self._current_value)

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
        total_count = self._maximum - self._minimum + 1
        if total_count <= self._window_size:
            start = self._minimum
            end = self._maximum
        else:
            half = self._window_size // 2
            start = center_value - half
            if start < self._minimum:
                start = self._minimum
            end = start + self._window_size - 1
            if end > self._maximum:
                end = self._maximum
                start = end - self._window_size + 1

        target_value = selected_value if isinstance(selected_value, int) else center_value
        if target_value < start:
            target_value = start
        elif target_value > end:
            target_value = end

        self.blockSignals(True)
        self.clear()
        for number in range(start, end + 1):
            self.addItem(str(number), number)

        idx = self.findData(target_value)
        if idx >= 0:
            self.setCurrentIndex(idx)
        self.blockSignals(False)
        self._current_value = self.value()

    def _rebuild_window(self, center_value: int):
        self._set_window(center_value, center_value)

    def _shift_window_by_steps(self, steps: int):
        if steps == 0:
            return

        center_idx = min(self._window_size // 2, max(0, self.count() - 1))
        center_value = self.itemData(center_idx)
        if not isinstance(center_value, int):
            center_value = self.value()

        selected_value = self.value()
        self._set_window(center_value - steps, selected_value)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            line_edit = self.lineEdit()
            if obj in (self, line_edit):
                QTimer.singleShot(0, self.showPopup)
                event.accept()
                return True

        if obj in (self, self.view(), self.view().viewport()) and event.type() == QEvent.Wheel:
            steps = _combo_steps_from_wheel(event)
            self._shift_window_by_steps(steps)
            event.accept()
            return True

        if obj is self.lineEdit() and event.type() == QEvent.Wheel:
            steps = _combo_steps_from_wheel(event)
            self._shift_window_by_steps(steps)
            event.accept()
            return True

        return super().eventFilter(obj, event)


class RecurrenceDialog(QDialog):
    """週期性設定對話框 - Outlook 風格"""

    rrule_created = Signal(str)

    def __init__(
        self,
        parent=None,
        current_rrule: str = "",
        initial_date: QDate | None = None,
        initial_time: QTime | None = None,
        embedded: bool = False,
    ):
        super().__init__(parent)
        self.current_rrule = current_rrule
        self.initial_date: QDate | None = initial_date
        self.initial_time: QTime | None = initial_time
        self.embedded = embedded
        self._wheel_combo_targets: dict[object, QComboBox] = {}

        if not self.embedded:
            self.setWindowTitle("週期性約會")
            self.setWindowIcon(get_app_icon())
            self.setMinimumWidth(570)
            self.setMinimumHeight(480)
            self.setModal(True)
        else:
            self.setMinimumWidth(700)

        self.setup_ui()
        self.apply_modern_style()
        self._apply_popup_holiday_checkers()
        self.lock_recurrence_detail_height()
        self.connect_signals()

        # 設置預設時間
        self.set_default_times()

        # 初始化結束條件控制項狀態
        self.on_end_condition_changed(self.radio_end_never, True)

        # 初始化頻率選擇的顯示狀態
        self.on_frequency_changed()

        # 解析現有的 RRULE（如果有的話）
        if self.current_rrule:
            self.parse_existing_rrule()
        else:
            # 新增排程：套用預設值
            self.apply_new_schedule_defaults()

        # 初始化時間同步：確保結束時間根據期間正確計算
        self.on_start_time_changed(None)

    def _resolve_db_manager(self):
        parent = self.parent()
        if parent and hasattr(parent, "db_manager"):
            return parent.db_manager
        if parent and hasattr(parent, "parent"):
            gp = parent.parent()
            if gp and hasattr(gp, "db_manager"):
                return gp.db_manager
        return None

    def _is_holiday_qdate(self, qdate: QDate) -> bool:
        if qdate.dayOfWeek() in (6, 7):
            return True
        db_manager = self._resolve_db_manager()
        if db_manager is None:
            return False
        try:
            return bool(db_manager.is_holiday_on_date(dt_date(qdate.year(), qdate.month(), qdate.day())))
        except Exception:
            return False

    def _apply_popup_holiday_checkers(self):
        if hasattr(self, "start_date_edit") and hasattr(self.start_date_edit, "_calendar_popup"):
            self.start_date_edit._calendar_popup.set_holiday_checker(self._is_holiday_qdate)
        if hasattr(self, "end_date_edit") and hasattr(self.end_date_edit, "_calendar_popup"):
            self.end_date_edit._calendar_popup.set_holiday_checker(self._is_holiday_qdate)

    def apply_new_schedule_defaults(self):
        """新增排程時套用預設值。"""
        default_date = self.initial_date if self.initial_date is not None else QDate.currentDate()

        # 開始時間：優先使用外部帶入時間（例如日/週視圖格位），否則才用目前時間取整
        if self.initial_time is not None:
            self.set_start_time(self.initial_time)
        else:
            self.set_start_time(self._get_rounded_current_time())

        # 循環模式：每天 + 每個工作日
        self.radio_daily.setChecked(True)
        self.daily_weekday_radio.setChecked(True)

        # 約會時間：期間 5 分
        idx = self.duration_combo.findData(5)
        if idx >= 0:
            self.duration_combo.setCurrentIndex(idx)
        else:
            self.set_custom_duration(5)

        # 循環範圍：開始為預設日期；結束於開始 + 3 個月
        self.start_date_edit.setDate(default_date)
        self.end_date_edit.setDate(default_date.addMonths(3))
        self.radio_end_never.setChecked(True)

    def set_default_times(self):
        """設置預設時間"""
        if self.initial_time is not None:
            default_start_time = self.initial_time
        else:
            default_start_time = self._get_rounded_current_time()

        self.set_start_time(default_start_time)

    def _get_rounded_current_time(self) -> QTime:
        """取得目前時間向上取整到最近整點或 30 分。"""
        current_time = QTime.currentTime()
        minute = ((current_time.minute() + 29) // 30) * 30
        if minute >= 60:
            minute = 0
            hour = (current_time.hour() + 1) % 24
        else:
            hour = current_time.hour()
        return QTime(hour, minute, 0)

    def setup_ui(self):
        """設定主介面"""
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # 約會時間區塊
        main_layout.addWidget(self.create_time_group())

        # 循環模式區塊
        self.recurrence_pattern_group = self.create_recurrence_pattern_group()
        main_layout.addWidget(self.recurrence_pattern_group)

        # 循環範圍區塊
        main_layout.addWidget(self.create_range_group())

        # 按鈕（嵌入模式不顯示）
        if not self.embedded:
            main_layout.addWidget(self.create_button_group())

    def create_time_group(self) -> QGroupBox:
        """建立約會時間區塊"""
        group = QGroupBox("排程時間")
        layout = QGridLayout(group)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        # 開始時間
        start_label = QLabel("開始(T):")
        start_label.setObjectName("fieldLabel")
        layout.addWidget(start_label, 0, 0)
        self.start_time_combo = QComboBox()
        self.start_time_combo.setEditable(True)
        self.start_time_combo.setInsertPolicy(QComboBox.NoInsert)
        self.start_time_combo.setMaxVisibleItems(10)
        self.start_time_combo.setFixedWidth(130)
        self._populate_time_combo(self.start_time_combo)
        if self.start_time_combo.lineEdit() is not None:
            self.start_time_combo.lineEdit().setReadOnly(False)
            self.start_time_combo.lineEdit().setAlignment(Qt.AlignCenter)
            self.start_time_combo.lineEdit().setCursor(Qt.IBeamCursor)
        self.start_time_combo.setCursor(Qt.PointingHandCursor)
        layout.addWidget(self._build_combo_with_side_arrows(self.start_time_combo), 0, 1)

        # 結束時間
        end_label = QLabel("結束(N):")
        end_label.setObjectName("fieldLabel")
        layout.addWidget(end_label, 1, 0)
        self.end_time_combo = QComboBox()
        self.end_time_combo.setEditable(True)
        self.end_time_combo.setInsertPolicy(QComboBox.NoInsert)
        self.end_time_combo.setMaxVisibleItems(10)
        self.end_time_combo.setFixedWidth(130)
        self._populate_time_combo(self.end_time_combo)
        if self.end_time_combo.lineEdit() is not None:
            self.end_time_combo.lineEdit().setReadOnly(False)
            self.end_time_combo.lineEdit().setAlignment(Qt.AlignCenter)
            self.end_time_combo.lineEdit().setCursor(Qt.IBeamCursor)
        self.end_time_combo.setCursor(Qt.PointingHandCursor)
        layout.addWidget(self._build_combo_with_side_arrows(self.end_time_combo), 1, 1)

        # 期間
        duration_label = QLabel("期間(U):")
        duration_label.setObjectName("fieldLabel")
        layout.addWidget(duration_label, 2, 0)
        self.duration_combo = QComboBox()
        # 允許自訂輸入（可編輯），但不要自動插入新項目
        self.duration_combo.setEditable(True)
        self.duration_combo.setInsertPolicy(QComboBox.NoInsert)
        self.duration_combo.setMaxVisibleItems(10)
        self.duration_combo.setFixedWidth(130)
        self.update_duration_combo()
        if self.duration_combo.lineEdit() is not None:
            self.duration_combo.lineEdit().setReadOnly(False)
            self.duration_combo.lineEdit().setAlignment(Qt.AlignCenter)
            self.duration_combo.lineEdit().setCursor(Qt.IBeamCursor)
        self.duration_combo.setCursor(Qt.PointingHandCursor)
        layout.addWidget(self._build_combo_with_side_arrows(self.duration_combo), 2, 1)

        self.lock_checkbox = QCheckBox("Lock")
        self.lock_checkbox.setToolTip("勾選後在開始到結束期間持續鎖定 OPC UA Tag 值")
        layout.addWidget(self.lock_checkbox, 0, 2, 3, 1, alignment=Qt.AlignLeft | Qt.AlignTop)

        self.time_guide_label = QLabel(
            "操作方式:\n"
            "1. 點右側箭頭: 展開下拉選單\n"
            "2. 點中間文字區: 直接鍵入\n"
            "3. 雙擊中間文字區: 以滑鼠拖曳選取文字\n"
            "4. Lock: 在開始到結束期間持續鎖定 OPC UA Tag 值"
        )
        self.time_guide_label.setWordWrap(True)
        self.time_guide_label.setObjectName("timeGuideLabel")
        layout.addWidget(self.time_guide_label, 0, 3, 3, 1, alignment=Qt.AlignTop)
        self._apply_time_guide_label_style()

        layout.setColumnStretch(2, 1)
        layout.setColumnStretch(3, 2)
        return group

    def _build_combo_with_side_arrows(self, combo: QComboBox) -> QWidget:
        """建立右側箭頭 + 中央可輸入的複合控件。"""
        container = QWidget(self)
        row = QHBoxLayout(container)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(4)

        right_btn = QToolButton(container)
        right_btn.setText("▼")
        right_btn.setToolTip("展開選單")
        right_btn.setCursor(Qt.PointingHandCursor)
        right_btn.setAutoRaise(True)
        right_btn.setFixedWidth(20)
        right_btn.clicked.connect(combo.showPopup)

        row.addWidget(combo)
        row.addWidget(right_btn)
        return container

    def _apply_time_guide_label_style(self):
        """依主題更新右側操作說明文字顏色。"""
        if not hasattr(self, "time_guide_label"):
            return
        if self.is_dark_mode():
            self.time_guide_label.setStyleSheet("color: #ffffff;")
        else:
            self.time_guide_label.setStyleSheet("color: #666666;")

    def _populate_time_combo(self, combo: QComboBox):
        """填入 00:00 ~ 23:30（每 30 分）時間選項，下拉顯示 HH:mm。"""
        combo.clear()
        for hour in range(24):
            for minute in (0, 30):
                time_value = QTime(hour, minute, 0)
                combo.addItem(time_value.toString("HH:mm"), time_value)

    def _parse_combo_time(self, combo: QComboBox) -> QTime:
        """從時間下拉目前值解析為 QTime，支援 HH:mm:ss 與 HH:mm。"""
        text = combo.currentText().strip()
        parsed = QTime.fromString(text, "HH:mm:ss")
        if parsed.isValid():
            return parsed

        parsed = QTime.fromString(text, "HH:mm")
        if parsed.isValid():
            return QTime(parsed.hour(), parsed.minute(), 0)

        data = combo.currentData()
        if isinstance(data, QTime) and data.isValid():
            return data

        return QTime(0, 0, 0)

    def _set_combo_time(self, combo: QComboBox, value: QTime):
        """設定時間下拉目前值；若不在預設清單中，仍顯示為 HH:mm:ss。"""
        if not value.isValid():
            value = QTime(0, 0, 0)

        list_text = value.toString("HH:mm")
        display_text = value.toString("HH:mm:ss")
        index = combo.findText(list_text)
        if index < 0:
            index = self._nearest_half_hour_index(value)
        combo.blockSignals(True)
        if index >= 0 and index < combo.count():
            combo.setCurrentIndex(index)
            if combo.lineEdit() is not None:
                combo.lineEdit().setText(display_text)
        else:
            combo.setCurrentText(display_text)
        combo.blockSignals(False)

    def _nearest_half_hour_index(self, value: QTime) -> int:
        """取得最接近 30 分刻度的下拉索引（0..47）。"""
        if not value.isValid():
            return 0

        total_minutes = value.hour() * 60 + value.minute()
        quotient, remainder = divmod(total_minutes, 30)
        if remainder >= 15:
            quotient += 1

        if quotient > 47:
            quotient = 47
        if quotient < 0:
            quotient = 0
        return quotient

    def _normalize_time_combo_display(self, combo: QComboBox):
        """將欄位顯示統一成 HH:mm:ss；不影響下拉項目仍為 HH:mm。"""
        if combo.lineEdit() is None:
            return
        time_value = self._parse_combo_time(combo)
        nearest_index = self._nearest_half_hour_index(time_value)
        combo.blockSignals(True)
        if 0 <= nearest_index < combo.count():
            combo.setCurrentIndex(nearest_index)
        combo.blockSignals(False)
        combo.lineEdit().setText(time_value.toString("HH:mm:ss"))

    def get_start_time(self) -> QTime:
        return self._parse_combo_time(self.start_time_combo)

    def set_start_time(self, value: QTime):
        self._set_combo_time(self.start_time_combo, value)

    def get_end_time(self) -> QTime:
        return self._parse_combo_time(self.end_time_combo)

    def set_end_time(self, value: QTime):
        self._set_combo_time(self.end_time_combo, value)

    def get_lock_enabled(self) -> bool:
        return bool(self.lock_checkbox.isChecked())

    def set_lock_enabled(self, enabled: bool):
        self.lock_checkbox.setChecked(bool(enabled))

    def update_duration_combo(self):
        """更新期間下拉選單"""
        self.duration_combo.clear()
        durations = [
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
        self.duration_combo.setCurrentIndex(0)  # 預設為 5 分

    def connect_signals(self):
        """連接信號"""
        # 連接時間和期間的互動
        self.start_time_combo.currentIndexChanged.connect(self.on_start_time_changed)
        self.end_time_combo.currentIndexChanged.connect(self.on_end_time_changed)
        if self.start_time_combo.lineEdit() is not None:
            self.start_time_combo.lineEdit().editingFinished.connect(self.on_start_time_changed)
        if self.end_time_combo.lineEdit() is not None:
            self.end_time_combo.lineEdit().editingFinished.connect(self.on_end_time_changed)
        self.duration_combo.currentIndexChanged.connect(self.on_duration_changed)
        # 支援使用者直接在可編輯的 combo 中輸入自訂期間
        if self.duration_combo.isEditable() and self.duration_combo.lineEdit() is not None:
            self.duration_combo.lineEdit().editingFinished.connect(self.on_duration_text_edited)
            self.duration_combo.lineEdit().textChanged.connect(self.on_duration_text_changed)

        combo_targets = [self.start_time_combo, self.end_time_combo, self.duration_combo]
        for combo in combo_targets:
            combo.installEventFilter(self)
            if combo.lineEdit() is not None:
                combo.lineEdit().installEventFilter(self)

        self._register_combo_wheel_targets()

        # 頻率選擇變更
        self.radio_daily.toggled.connect(self.on_frequency_changed)
        self.radio_weekly.toggled.connect(self.on_frequency_changed)
        self.radio_monthly.toggled.connect(self.on_frequency_changed)
        self.radio_yearly.toggled.connect(self.on_frequency_changed)

        # 結束條件變更
        self.end_button_group.buttonToggled.connect(self.on_end_condition_changed)

    def _register_combo_wheel_targets(self):
        self._wheel_combo_targets.clear()
        for combo in self.findChildren(QComboBox):
            if type(combo) is not QComboBox:
                continue

            self._wheel_combo_targets[combo] = combo
            combo.installEventFilter(self)

            line_edit = combo.lineEdit()
            if line_edit is not None:
                self._wheel_combo_targets[line_edit] = combo
                line_edit.installEventFilter(self)

            view = combo.view()
            if view is not None:
                self._wheel_combo_targets[view] = combo
                self._wheel_combo_targets[view.viewport()] = combo
                view.installEventFilter(self)
                view.viewport().installEventFilter(self)

    def parse_existing_rrule(self):
        """解析現有的 RRULE 字串並設置控制項"""
        if not self.current_rrule:
            return

        try:
            # 解析 RRULE 參數
            params = {}
            dtstart_raw = ""
            parts = self.current_rrule.split(";")
            for part in parts:
                if "=" in part:
                    key, value = part.split("=", 1)
                    params[key] = value
                elif part.startswith("DTSTART:"):
                    dtstart_raw = part.split(":", 1)[1]

            # 設置頻率
            freq = params.get("FREQ", "DAILY")
            self.lunar_mode_checkbox.setChecked(params.get("X-LUNAR", "0") == "1")
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
                interval = max(1, min(12, interval))
                idx = self.monthly_interval.findData(interval)
                if idx >= 0:
                    self.monthly_interval.setCurrentIndex(idx)
                idx = self.monthly_week_interval.findData(interval)
                if idx >= 0:
                    self.monthly_week_interval.setCurrentIndex(idx)
            elif freq == "YEARLY":
                self.yearly_interval.setValue(interval)

            # 設置開始日期（優先使用 RRULE 的 DTSTART）
            range_start_raw = params.get("X-RANGE-START", "")
            if range_start_raw and len(range_start_raw) >= 8:
                try:
                    year = int(range_start_raw[:4])
                    month = int(range_start_raw[4:6])
                    day = int(range_start_raw[6:8])
                    self.start_date_edit.setDate(QDate(year, month, day))
                except (ValueError, IndexError):
                    pass
            elif dtstart_raw and len(dtstart_raw) >= 8:
                try:
                    year = int(dtstart_raw[:4])
                    month = int(dtstart_raw[4:6])
                    day = int(dtstart_raw[6:8])
                    self.start_date_edit.setDate(QDate(year, month, day))
                except (ValueError, IndexError):
                    pass

            # 設置開始時間
            # 編輯既有排程時，優先使用 RRULE 已儲存時間；僅在 RRULE 無時間時才回退到 initial_time
            byhour = params.get("BYHOUR")
            byminute = params.get("BYMINUTE", "0")
            if byhour:
                hour = int(byhour)
                minute = int(byminute)
                start_time = QTime(hour, minute, 0)
            elif dtstart_raw and "T" in dtstart_raw and len(dtstart_raw.split("T", 1)[1]) >= 4:
                try:
                    time_part = dtstart_raw.split("T", 1)[1]
                    hour = int(time_part[:2])
                    minute = int(time_part[2:4])
                    start_time = QTime(hour, minute, 0)
                except (ValueError, IndexError):
                    start_time = self.initial_time if self.initial_time is not None else QTime(9, 0, 0)
            elif self.initial_time is not None:
                start_time = self.initial_time
            else:
                # 如果沒有 BYHOUR，使用預設時間 (上午9:00)
                start_time = QTime(9, 0, 0)
            # 設置開始時間
            self.set_start_time(start_time)

            # 設置期間（優先使用 DURATION）
            duration_minutes = self._parse_duration_minutes(params.get("DURATION", ""))
            if duration_minutes is not None:
                idx = self.duration_combo.findData(duration_minutes)
                if idx >= 0:
                    self.duration_combo.setCurrentIndex(idx)
                else:
                    self.set_custom_duration(duration_minutes)

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

            # 依據已套用的開始時間與期間同步結束時間
            self.on_start_time_changed(None)

        except Exception as e:
            print(f"解析 RRULE 失敗: {e}")
            # 解析失敗時使用預設值

    def _parse_frequency_specific_params(self, params):
        """解析頻率特定的參數"""
        freq = params.get("FREQ", "DAILY")

        if freq == "DAILY":
            byday = params.get("BYDAY", "")
            if byday == "MO,TU,WE,TH,FR":
                self.daily_weekday_radio.setChecked(True)
            else:
                self.radio_daily_every.setChecked(True)

        elif freq == "WEEKLY":
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
                interval = max(1, min(12, int(params.get("INTERVAL", "1"))))
                idx = self.monthly_week_interval.findData(interval)
                if idx >= 0:
                    self.monthly_week_interval.setCurrentIndex(idx)
                setpos_to_index = {1: 0, 2: 1, 3: 2, 4: 3, -1: 4}
                bysetpos_num = int(bysetpos)
                self.monthly_week_num.setCurrentIndex(setpos_to_index.get(bysetpos_num, 0))
                # 設置星期幾
                if byday == "MO,TU,WE,TH,FR":
                    self.monthly_week_day.setCurrentIndex(0)
                else:
                    day_map = {
                        "SU": 1, "MO": 2, "TU": 3, "WE": 4, "TH": 5, "FR": 6, "SA": 7
                    }
                    if byday in day_map:
                        self.monthly_week_day.setCurrentIndex(day_map[byday])

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
                setpos_to_index = {1: 0, 2: 1, 3: 2, 4: 3, -1: 4}
                bysetpos_num = int(bysetpos)
                self.yearly_week_num.setCurrentIndex(setpos_to_index.get(bysetpos_num, 0))
                # 設置星期幾
                if byday == "MO,TU,WE,TH,FR":
                    self.yearly_week_day.setCurrentIndex(0)
                else:
                    day_map = {
                        "SU": 1, "MO": 2, "TU": 3, "WE": 4, "TH": 5, "FR": 6, "SA": 7
                    }
                    if byday in day_map:
                        self.yearly_week_day.setCurrentIndex(day_map[byday])

    def _parse_duration_minutes(self, duration_str: str):
        """解析 DURATION 參數（例如 PT5M）為分鐘數，失敗回傳 None。"""
        if not duration_str:
            return None

        s = duration_str.strip().upper()
        if not s.startswith("PT"):
            return None

        # 支援 PT#H#M（目前實際輸出主要為 PT#M）
        hours = 0
        minutes = 0
        body = s[2:]

        if "H" in body:
            h_part, body = body.split("H", 1)
            if h_part:
                hours = int(h_part)

        if "M" in body:
            m_part = body.split("M", 1)[0]
            if m_part:
                minutes = int(m_part)

        return hours * 60 + minutes

    def on_start_time_changed(self, value=None):
        """開始時間改變時更新結束時間"""
        if not hasattr(self, "_updating_times") or not self._updating_times:
            self._updating_times = True
            try:
                start_time = self.get_start_time()
                self._normalize_time_combo_display(self.start_time_combo)
                duration_minutes = self.get_duration_minutes()
                if duration_minutes is not None:
                    end_time = start_time.addSecs(duration_minutes * 60)
                    self.set_end_time(end_time)
            finally:
                self._updating_times = False

    def on_end_time_changed(self, value=None):
        """結束時間改變時更新期間"""
        if not hasattr(self, "_updating_times") or not self._updating_times:
            self._updating_times = True
            try:
                start_time = self.get_start_time()
                end_time = self.get_end_time()
                self._normalize_time_combo_display(self.end_time_combo)
                duration_seconds = start_time.secsTo(end_time)
                if duration_seconds < 0:
                    duration_seconds += 24 * 3600  # 跨日
                duration_minutes = duration_seconds // 60
                self.set_duration_to_minutes(duration_minutes)
            finally:
                self._updating_times = False

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
        root_layout = QVBoxLayout(group)
        root_layout.setSpacing(8)
        root_layout.setContentsMargins(12, 12, 12, 12)

        top_bar = QHBoxLayout()
        top_bar.setContentsMargins(0, 0, 0, 0)
        top_bar.setSpacing(6)
        top_bar.addStretch()
        self.lunar_mode_checkbox = QCheckBox("農曆")
        self.lunar_mode_checkbox.setToolTip("勾選後，排程將以農曆規則計算觸發日期")
        top_bar.addWidget(self.lunar_mode_checkbox)
        root_layout.addLayout(top_bar)

        layout = QHBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(0, 0, 0, 0)

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

        self.lock_recurrence_detail_height()

        layout.addWidget(self.detail_widget, 1)
        root_layout.addLayout(layout)
        return group

    def lock_recurrence_detail_height(self):
        """鎖定右側詳細設定高度，避免切換頻率時面板高度逐步增加。"""
        if not hasattr(self, "detail_widget"):
            return

        detail_panels = []
        for attr_name in ("daily_widget", "weekly_widget", "monthly_widget", "yearly_widget"):
            panel = getattr(self, attr_name, None)
            if panel is not None:
                detail_panels.append(panel)

        if not detail_panels:
            return

        max_height = max(panel.sizeHint().height() for panel in detail_panels)
        if max_height <= 0:
            return

        self.detail_widget.setFixedHeight(max_height)

    def create_daily_detail(self):
        """建立每天選項的詳細設定"""
        self.daily_widget = QWidget()
        layout = QHBoxLayout(self.daily_widget)
        layout.setSpacing(8)  # 增加間距
        layout.setContentsMargins(0, 0, 0, 0)

        self.radio_daily_every = QRadioButton("每(V)")
        self.radio_daily_every.setChecked(True)
        layout.addWidget(self.radio_daily_every)

        self.daily_interval = RollingNumberComboBox(1, 999)
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
        self.weekly_interval = RollingNumberComboBox(1, 52)
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

        self.monthly_interval = QComboBox()
        for value in range(1, 13):
            self.monthly_interval.addItem(str(value), value)
        self.monthly_interval.setCurrentIndex(0)
        self.monthly_interval.setFixedWidth(50)
        day_layout.addWidget(self.monthly_interval)

        month_label = QLabel("個月的第")
        month_label.setObjectName("fieldLabel")
        day_layout.addWidget(month_label)

        self.monthly_day = RollingNumberComboBox(1, 31)
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

        self.monthly_week_interval = QComboBox()
        for value in range(1, 13):
            self.monthly_week_interval.addItem(str(value), value)
        self.monthly_week_interval.setCurrentIndex(0)
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
        self.yearly_interval = RollingNumberComboBox(1, 999)
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

        self.yearly_day = RollingNumberComboBox(1, 31)
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
        week_layout.setSpacing(10)
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
        self.yearly_week_num.setFixedWidth(110)
        self.yearly_week_num.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        week_layout.addWidget(self.yearly_week_num)

        self.yearly_week_day = QComboBox()
        self.yearly_week_day.addItems(
            ["週一到週五", "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"]
        )
        self.yearly_week_day.setFixedWidth(130)
        self.yearly_week_day.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
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
        self.start_date_edit = PopupDateEdit()
        self.start_date_edit.setDisplayFormat("yyyy/M/d (ddd)")
        self.start_date_edit.setDate(self.initial_date or QDate.currentDate())
        self.start_date_edit.setFixedWidth(150)
        layout.addWidget(self.start_date_edit, 0, 1)

        # 結束選項
        self.end_button_group = QButtonGroup(self)

        # 結束於日期
        self.radio_end_by = QRadioButton("結束於(B):")
        self.end_button_group.addButton(self.radio_end_by)
        layout.addWidget(self.radio_end_by, 0, 2)

        self.end_date_edit = PopupDateEdit()
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
        self.end_count = RollingNumberComboBox(1, 999)
        self.end_count.setValue(1)
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
                start_time = self.get_start_time()
                duration_minutes = self.get_duration_minutes()
                # 選取內建項目時，取消自訂旗標
                if self.duration_combo.currentIndex() >= 0:
                    self._using_custom_duration = False

                if duration_minutes is not None:
                    end_time = start_time.addSecs(duration_minutes * 60)
                    self.set_end_time(end_time)
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
                    except (TypeError, ValueError):
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
        try:
            rrule_str = self.build_rrule()
            self.rrule_created.emit(rrule_str)
            self.accept()
        except Exception as e:
            QMessageBox.warning(self, "錯誤", f"建立週期規則時發生錯誤：{str(e)}")

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
        time = self.get_start_time()
        hour = time.hour()
        minute = time.minute()

        # 開始日期
        start_date = self.start_date_edit.date()
        dtstart_date = start_date
        range_start = f"{start_date.year()}{start_date.month():02d}{start_date.day():02d}"

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
                interval_data = self.monthly_interval.currentData()
                interval = int(interval_data) if isinstance(interval_data, int) else 1
                bymonthday = str(self.monthly_day.value())
                safe_day = min(self.monthly_day.value(), QDate(start_date.year(), start_date.month(), 1).daysInMonth())
                candidate = QDate(start_date.year(), start_date.month(), safe_day)
                if candidate.isValid():
                    dtstart_date = candidate
            else:
                interval_data = self.monthly_week_interval.currentData()
                interval = int(interval_data) if isinstance(interval_data, int) else 1
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
                target_month = self.yearly_month.currentIndex() + 1
                safe_day = min(self.yearly_day.value(), QDate(start_date.year(), target_month, 1).daysInMonth())
                candidate = QDate(start_date.year(), target_month, safe_day)
                if candidate.isValid():
                    dtstart_date = candidate
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

        dtstart = f"{dtstart_date.year()}{dtstart_date.month():02d}{dtstart_date.day():02d}T{hour:02d}{minute:02d}00"

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

        if self.lunar_mode_checkbox.isChecked():
            parts.append("X-LUNAR=1")

        parts.append(f"X-RANGE-START={range_start}")

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
                    border: 1px solid #2a8ccd;
                    border-radius: 4px;
                    padding: 6px 16px;
                    font-weight: bold;
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
                QComboBox::drop-down {
                    width: 0px;
                    border: none;
                }
                QComboBox::down-arrow {
                    image: none;
                    width: 0px;
                    height: 0px;
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
                    background-color: #e9ecef;
                    color: #111111;
                    border: 1px solid #9aa4ad;
                    border-radius: 4px;
                    padding: 6px 16px;
                    font-weight: bold;
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
                QComboBox::drop-down {
                    width: 0px;
                    border: none;
                }
                QComboBox::down-arrow {
                    image: none;
                    width: 0px;
                    height: 0px;
                }
                QSpinBox:focus, QComboBox:focus, QDateEdit:focus, QTimeEdit:focus {
                    border: 2px solid #0078d4;
                }
                QComboBox::item {
                    background-color: white;
                    color: #333;
                }
                QComboBox::item:selected {
                    background-color: #9ec6f3;
                    color: #0f1f33;
                }
                QCalendarWidget QWidget {
                    background-color: #f5f5f5;
                    color: #333;
                }
                QCalendarWidget QAbstractItemView:enabled {
                    background-color: white;
                    color: #333;
                    selection-background-color: #9ec6f3;
                    selection-color: #0f1f33;
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

        if hasattr(self, "start_date_edit") and hasattr(self.start_date_edit, "_calendar_popup"):
            self.start_date_edit._calendar_popup.apply_theme(is_dark)
        if hasattr(self, "end_date_edit") and hasattr(self.end_date_edit, "_calendar_popup"):
            self.end_date_edit._calendar_popup.apply_theme(is_dark)
        self._apply_time_guide_label_style()

    def get_rrule(self) -> str:
        """取得 RRULE 字串"""
        return self.build_rrule()

    def keyPressEvent(self, event):
        if self.embedded and event.key() == Qt.Key_Escape:
            parent = self.parentWidget()
            while parent is not None and not isinstance(parent, QDialog):
                parent = parent.parentWidget()
            if isinstance(parent, QDialog):
                parent.reject()
                event.accept()
                return
        super().keyPressEvent(event)

    def eventFilter(self, obj, event):
        """事件過濾器，用於處理時間 combo box 的鍵盤輸入"""
        combo_pairs = [
            (self.start_time_combo, self.start_time_combo.lineEdit()),
            (self.end_time_combo, self.end_time_combo.lineEdit()),
            (self.duration_combo, self.duration_combo.lineEdit()),
        ]

        wheel_targets: dict[object, QComboBox] = {}
        for combo, line_edit in combo_pairs:
            wheel_targets[combo] = combo
            if line_edit is not None:
                wheel_targets[line_edit] = combo
            view = combo.view()
            if view is not None:
                wheel_targets[view] = combo
                wheel_targets[view.viewport()] = combo

        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            for combo, line_edit in combo_pairs:
                # 點擊 combo 本體時直接展開清單；lineEdit 保留原生游標/選取/編輯行為。
                if obj is combo and combo.isEnabled():
                    QTimer.singleShot(0, combo.showPopup)
                    event.accept()
                    return True

                # 點擊 lineEdit 左右邊緣時展開；中間區域交給原生編輯行為。
                if obj is line_edit and combo.isEnabled():
                    x_pos = int(event.position().x()) if hasattr(event, "position") else int(event.x())
                    side_zone = 18
                    if x_pos >= max(side_zone, line_edit.width() - side_zone):
                        QTimer.singleShot(0, combo.showPopup)
                        event.accept()
                        return True

        if event.type() == QEvent.Wheel:
            combo = wheel_targets.get(obj)
            if combo is not None and combo.isEnabled() and combo.count() > 0:
                view = combo.view()
                if obj in (view, view.viewport()):
                    # 下拉展開時，讓清單使用原生滾輪捲動行為。
                    return False

                steps = _combo_steps_from_wheel(event)
                if steps != 0:
                    current_index = combo.currentIndex()
                    if current_index < 0:
                        current_index = 0
                    target_index = current_index - steps
                    if target_index < 0:
                        target_index = 0
                    elif target_index >= combo.count():
                        target_index = combo.count() - 1
                    if target_index != combo.currentIndex():
                        combo.setCurrentIndex(target_index)
                event.accept()
                return True

        return super().eventFilter(obj, event)


def show_recurrence_dialog(
    parent=None,
    current_rrule: str = "",
    initial_date: QDate | None = None,
    initial_time: QTime | None = None,
) -> str:
    """
    顯示週期性設定對話框並返回 RRULE 字串

    Args:
        parent: 父視窗
        current_rrule: 現有的 RRULE 字串（可選，用於編輯）
        initial_date: 建議的開始日期（例如從主行事曆點選的日期）
        initial_time: 建議的開始時間

    Returns:
        str: RRULE 字串，使用者取消則返回空字串
    """
    dialog = RecurrenceDialog(parent, current_rrule, initial_date, initial_time)
    if dialog.exec() == QDialog.Accepted:
        return dialog.get_rrule()
    return ""


if __name__ == "__main__":
    from PySide6.QtWidgets import QApplication

    app = QApplication(sys.argv)

    rrule = show_recurrence_dialog()
    print(f"生成的 RRULE: {rrule}")

    app.exit()

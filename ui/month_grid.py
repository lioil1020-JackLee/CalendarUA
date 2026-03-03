from __future__ import annotations

from collections import defaultdict
from datetime import datetime, date
from typing import Dict, List

from PySide6.QtCore import QDate, Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QHeaderView, QInputDialog, QLabel, QMenu, QTableWidget, QVBoxLayout, QWidget

from core.schedule_resolver import ResolvedOccurrence
from core.lunar_calendar import to_lunar, LunarDateInfo


def _month_grid_start(month_date: QDate) -> QDate:
    first_day = QDate(month_date.year(), month_date.month(), 1)
    days_to_sunday = first_day.dayOfWeek() % 7
    return first_day.addDays(-days_to_sunday)


def _format_lunar_day(info: LunarDateInfo) -> str:
    """將農曆日轉成簡短中文（初一、十五等），若無法判定則回傳空字串。"""
    n = info.lunar_day
    chinese_ten = ["初", "十", "廿", "卅"]
    numerals = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
    if n <= 0 or n > 30:
        return ""
    if n == 1:
        month_names = {
            1: "元",
            2: "二",
            3: "三",
            4: "四",
            5: "五",
            6: "六",
            7: "七",
            8: "八",
            9: "九",
            10: "十",
            11: "十一",
            12: "十二",
        }
        month_text = month_names.get(info.lunar_month, str(info.lunar_month))
        leap_prefix = "閏" if info.is_leap_month else ""
        return f"{leap_prefix}{month_text}月"
    if n == 10:
        return "初十"
    if n == 20:
        return "二十"
    if n == 30:
        return "三十"
    ten = chinese_ten[(n - 1) // 10]
    digit = numerals[(n - 1) % 10]
    return f"{ten}{digit}"


class EventChipLabel(QLabel):
    event_double_clicked = Signal(object)

    def __init__(self, occurrence: ResolvedOccurrence, text: str = "", parent=None):
        super().__init__(text, parent)
        self.occurrence = occurrence

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.event_double_clicked.emit(self.occurrence)
            event.accept()
            return
        super().mouseDoubleClickEvent(event)


class MonthViewWidget(QWidget):
    date_selected = Signal(QDate)
    context_action_requested = Signal(str, dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.reference_date = QDate.currentDate()
        self.selected_date = QDate.currentDate()
        self.occurrences: List[ResolvedOccurrence] = []
        self._cell_dates: Dict[tuple[int, int], QDate] = {}

        self.table = QTableWidget(6, 7)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionMode(QTableWidget.NoSelection)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        # 支援滑鼠左鍵雙擊：直接開啟編輯 / 新增視窗
        self.table.cellDoubleClicked.connect(self._on_cell_double_clicked)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setVisible(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 降低表頭高度，縮小主標題列與月格之間的視覺空白
        self.table.horizontalHeader().setFixedHeight(24)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        header_font = QFont(self.table.horizontalHeader().font())
        header_font.setFamily("Segoe UI")
        header_font.setPointSize(10)
        header_font.setBold(True)
        self.table.horizontalHeader().setFont(header_font)
        self.table.horizontalHeader().setStyleSheet(
            "QHeaderView::section { padding: 0px 4px; font-family: 'Segoe UI'; font-weight: 600; }"
        )
        self.table.cellClicked.connect(self._on_cell_clicked)

        self.table.setHorizontalHeaderLabels(["日", "一", "二", "三", "四", "五", "六"])
        self._grouped_events: Dict[QDate, List[ResolvedOccurrence]] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.table)

        self._render()

    def set_reference_date(self, qdate: QDate):
        self.reference_date = qdate
        self._render()

    def set_selected_date(self, qdate: QDate):
        self.selected_date = qdate
        self._render()

    def set_occurrences(self, occurrences: List[ResolvedOccurrence]):
        self.occurrences = occurrences
        self._render()

    def _group_by_date(self) -> Dict[QDate, List[ResolvedOccurrence]]:
        grouped: Dict[QDate, List[ResolvedOccurrence]] = defaultdict(list)
        for occurrence in self.occurrences:
            key = QDate(occurrence.start.year, occurrence.start.month, occurrence.start.day)
            grouped[key].append(occurrence)

        for date_key in grouped:
            grouped[date_key].sort(key=lambda item: item.start)

        return grouped

    def _build_cell_widget(self, qdate: QDate, events: List[ResolvedOccurrence]) -> QWidget:
        container = QWidget()
        is_selected = qdate == self.selected_date
        is_today = qdate == QDate.currentDate()

        if is_selected and is_today:
            container.setStyleSheet(
                "background-color: rgba(47, 115, 217, 0.22); border: 2px solid #f4c542; border-radius: 4px;"
            )
        elif is_selected:
            container.setStyleSheet(
                "background-color: rgba(47, 115, 217, 0.18); border: 1px solid #2f73d9; border-radius: 4px;"
            )
        elif is_today:
            container.setStyleSheet(
                "border: 1px solid #f4c542; border-radius: 4px;"
            )
        layout = QVBoxLayout(container)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(3)

        # 日期 + 農曆顯示（若有安裝農曆套件）
        text = str(qdate.day())
        lunar_text = ""
        try:
            info = to_lunar(date(qdate.year(), qdate.month(), qdate.day()))
            if info:
                lunar_text = _format_lunar_day(info)
        except Exception:
            lunar_text = ""

        if lunar_text:
            text = f"{qdate.day()} ({lunar_text})"

        date_label = QLabel(text)
        if qdate.month() != self.reference_date.month():
            date_label.setStyleSheet("color: #808080;")
        elif is_selected:
            date_label.setStyleSheet("font-weight: bold; color: #ffffff;")
        elif is_today:
            date_label.setStyleSheet("font-weight: bold; color: #f4c542;")
        else:
            date_label.setStyleSheet("font-weight: bold;")

        layout.addWidget(date_label)

        for occurrence in events[:3]:
            chip = EventChipLabel(occurrence, occurrence.title)
            chip.setStyleSheet(
                f"background-color: {occurrence.category_bg};"
                f"color: {occurrence.category_fg};"
                "border-radius: 8px;"
                "padding: 2px 6px;"
            )
            chip.setToolTip(
                f"{occurrence.title}\n"
                f"{occurrence.start.strftime('%H:%M')} - {occurrence.end.strftime('%H:%M')}\n"
                f"{occurrence.target_value}"
            )
            chip.event_double_clicked.connect(
                lambda occ, cell_date=qdate: self._on_chip_double_clicked(cell_date, occ)
            )
            layout.addWidget(chip)

        remain = len(events) - 3
        if remain > 0:
            more_label = QLabel(f"+{remain} more")
            more_label.setStyleSheet("color: #555555; font-size: 11px;")
            layout.addWidget(more_label)

        layout.addStretch()
        return container

    def _render(self):
        grouped = self._group_by_date()
        self._grouped_events = grouped
        self._cell_dates.clear()

        start = _month_grid_start(self.reference_date)

        for row in range(6):
            for col in range(7):
                qdate = start.addDays(row * 7 + col)
                self._cell_dates[(row, col)] = qdate
                events = grouped.get(qdate, [])
                self.table.setCellWidget(row, col, self._build_cell_widget(qdate, events))

    def _on_cell_clicked(self, row: int, col: int):
        qdate = self._cell_dates.get((row, col))
        if qdate is None:
            return

        self.selected_date = qdate
        self.date_selected.emit(qdate)
        self._render()

    def _on_cell_double_clicked(self, row: int, col: int):
        """滑鼠左鍵雙擊：有任務則選擇後編輯，否則新增。"""
        qdate = self._cell_dates.get((row, col))
        if qdate is None:
            return

        events = self._grouped_events.get(qdate, [])
        if events:
            target = self._pick_event(qdate, events, "編輯")
            if target is None:
                return
            payload = {
                "schedule_id": target.schedule_id,
                "date": qdate.toString("yyyy-MM-dd"),
                "hour": target.start.hour,
                "week_mode": False,
                "month_mode": True,
            }
            self.context_action_requested.emit("edit", payload)
            return

        payload = {
            "schedule_id": None,
            "date": qdate.toString("yyyy-MM-dd"),
            "hour": 8,
            "week_mode": False,
            "month_mode": True,
        }
        self.context_action_requested.emit("new", payload)

    def _on_chip_double_clicked(self, qdate: QDate, occurrence: ResolvedOccurrence):
        """雙擊任務 chip 時編輯該任務。"""
        payload = {
            "schedule_id": occurrence.schedule_id,
            "date": qdate.toString("yyyy-MM-dd"),
            "hour": occurrence.start.hour,
            "week_mode": False,
            "month_mode": True,
        }
        self.context_action_requested.emit("edit", payload)

    def _pick_event(self, qdate: QDate, events: List[ResolvedOccurrence], action_text: str) -> ResolvedOccurrence | None:
        """同一天多筆任務時，讓使用者選擇目標任務。"""
        if not events:
            return None
        if len(events) == 1:
            return events[0]

        option_map: Dict[str, ResolvedOccurrence] = {}
        options: List[str] = []
        for idx, occ in enumerate(events, start=1):
            label = (
                f"{occ.title} ({occ.start.strftime('%H:%M')} - {occ.end.strftime('%H:%M')}) "
                f"[ID:{occ.schedule_id}]"
            )
            if label in option_map:
                label = f"{label} #{idx}"
            option_map[label] = occ
            options.append(label)

        selected, ok = QInputDialog.getItem(
            self,
            f"選擇要{action_text}的行程",
            f"{qdate.toString('yyyy-MM-dd')} 有多筆任務，請選擇：",
            options,
            0,
            False,
        )
        if not ok or not selected:
            return None
        return option_map.get(selected)

    def _show_context_menu(self, position):
        index = self.table.indexAt(position)
        if not index.isValid():
            return

        row = index.row()
        col = index.column()
        qdate = self._cell_dates.get((row, col))
        if qdate is None:
            return

        events = self._grouped_events.get(qdate, [])
        first_event = events[0] if events else None

        payload = {
            "schedule_id": first_event.schedule_id if first_event else None,
            "date": qdate.toString("yyyy-MM-dd"),
            "hour": first_event.start.hour if first_event else 8,
            "week_mode": False,
            "month_mode": True,
        }

        menu = QMenu(self)
        new_action = menu.addAction("新增行程 (New Appointment)")
        edit_action = menu.addAction("編輯行程 (Edit)")
        delete_action = menu.addAction("刪除行程 (Delete)")
        menu.addSeparator()
        today_action = menu.addAction("移至今天 (Go to Today)")
        refresh_action = menu.addAction("重新整理 (Refresh)")

        has_event = first_event is not None
        edit_action.setEnabled(has_event)
        delete_action.setEnabled(has_event)

        selected_action = menu.exec(self.table.viewport().mapToGlobal(position))
        if selected_action is None:
            return

        if selected_action == new_action:
            self.context_action_requested.emit("new", payload)
        elif selected_action == edit_action:
            target = self._pick_event(qdate, events, "編輯")
            if target is None:
                return
            payload["schedule_id"] = target.schedule_id
            payload["hour"] = target.start.hour
            self.context_action_requested.emit("edit", payload)
        elif selected_action == delete_action:
            target = self._pick_event(qdate, events, "刪除")
            if target is None:
                return
            payload["schedule_id"] = target.schedule_id
            payload["hour"] = target.start.hour
            self.context_action_requested.emit("delete", payload)
        elif selected_action == today_action:
            self.context_action_requested.emit("today", payload)
        elif selected_action == refresh_action:
            self.context_action_requested.emit("refresh", payload)

from __future__ import annotations

from datetime import datetime
from typing import List

from PySide6.QtCore import QDate, Qt, Signal
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import QHeaderView, QMenu, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget

from core.schedule_resolver import ResolvedOccurrence


def _hour_label(hour: int) -> str:
    if hour == 0:
        return "上午 12:00"
    if 1 <= hour < 12:
        return f"上午 {hour:02d}:00"
    if hour == 12:
        return "下午 12:00"
    return f"下午 {hour - 12:02d}:00"


def _week_start_sunday(reference_date: QDate) -> QDate:
    days_to_sunday = reference_date.dayOfWeek() % 7
    return reference_date.addDays(-days_to_sunday)


class ScheduleTimeGridWidget(QWidget):
    context_action_requested = Signal(str, dict)

    def __init__(self, week_mode: bool, parent=None):
        super().__init__(parent)
        self.week_mode = week_mode
        self.reference_date = QDate.currentDate()
        self.occurrences: List[ResolvedOccurrence] = []
        self._cell_occurrence_map: dict[tuple[int, int], ResolvedOccurrence] = {}

        self.table = QTableWidget(24, 7 if week_mode else 1)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionMode(QTableWidget.NoSelection)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setFixedHeight(34)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.table.horizontalHeader().setStyleSheet("QHeaderView::section { padding: 0px 2px; font-family: 'Times New Roman'; font-weight: 700; }")
        header_font = QFont(self.table.horizontalHeader().font())
        header_font.setFamily("Times New Roman")
        header_font.setPointSize(14)
        header_font.setBold(True)
        self.table.horizontalHeader().setFont(header_font)

        for row in range(24):
            self.table.setVerticalHeaderItem(row, QTableWidgetItem(_hour_label(row)))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.table)

        self._refresh_headers()
        self._ensure_items()

    def set_reference_date(self, qdate: QDate):
        self.reference_date = qdate
        self._refresh_headers()
        self._render()

    def set_occurrences(self, occurrences: List[ResolvedOccurrence]):
        self.occurrences = occurrences
        self._render()

    def _refresh_headers(self):
        if not self.week_mode:
            header = f"{self.reference_date.toString('yyyy/MM/dd')}"
            self.table.setHorizontalHeaderLabels([header])
            return

        sunday = _week_start_sunday(self.reference_date)
        labels = []
        day_names = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"]
        for offset in range(7):
            day = sunday.addDays(offset)
            labels.append(f"{day_names[offset]} {day.day()}")
        self.table.setHorizontalHeaderLabels(labels)

    def _ensure_items(self):
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                if self.table.item(row, col) is None:
                    item = QTableWidgetItem("")
                    item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)
                    self.table.setItem(row, col, item)

    def _clear_grid(self):
        self._cell_occurrence_map.clear()
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item is None:
                    continue
                item.setText("")
                item.setBackground(QColor("#ffffff"))
                item.setForeground(QColor("#000000"))
                item.setToolTip("")

    def _column_for_date(self, dt: datetime) -> int:
        if not self.week_mode:
            return 0

        sunday = _week_start_sunday(self.reference_date)
        target = QDate(dt.year, dt.month, dt.day)
        return sunday.daysTo(target)

    def _render_occurrence(self, occurrence: ResolvedOccurrence):
        start = occurrence.start
        end = occurrence.end

        col = self._column_for_date(start)
        if col < 0 or col >= self.table.columnCount():
            return

        start_row = max(0, start.hour)
        if end.minute == 0 and end.second == 0:
            end_row = max(start_row, end.hour - 1)
        else:
            end_row = max(start_row, end.hour)
        end_row = min(23, end_row)

        for row in range(start_row, end_row + 1):
            item = self.table.item(row, col)
            if item is None:
                continue

            item.setBackground(QColor(occurrence.category_bg))
            item.setForeground(QColor(occurrence.category_fg))
            if row == start_row:
                item.setText(occurrence.title)
            tooltip = f"{occurrence.title}\n{start.strftime('%H:%M')} - {end.strftime('%H:%M')}\n{occurrence.target_value}"
            item.setToolTip(tooltip)
            self._cell_occurrence_map[(row, col)] = occurrence

    def _date_for_column(self, col: int) -> QDate:
        if not self.week_mode:
            return self.reference_date

        sunday = _week_start_sunday(self.reference_date)
        return sunday.addDays(col)

    def _show_context_menu(self, position):
        index = self.table.indexAt(position)
        if not index.isValid():
            return

        row = index.row()
        col = index.column()
        occurrence = self._cell_occurrence_map.get((row, col))
        schedule_id = occurrence.schedule_id if occurrence else None
        date_text = self._date_for_column(col).toString("yyyy-MM-dd")

        payload = {
            "schedule_id": schedule_id,
            "date": date_text,
            "hour": row,
            "week_mode": self.week_mode,
        }

        menu = QMenu(self)

        open_action = menu.addAction("Open")
        delete_action = menu.addAction("Delete")
        menu.addSeparator()
        new_action = menu.addAction("New Event")
        copy_action = menu.addAction("Copy")
        cut_action = menu.addAction("Cut")
        paste_action = menu.addAction("Paste")

        time_scale_menu = menu.addMenu("Time Scale")
        scale_30 = time_scale_menu.addAction("30 min")
        scale_60 = time_scale_menu.addAction("60 min")
        scale_120 = time_scale_menu.addAction("120 min")

        menu.addSeparator()
        refresh_action = menu.addAction("Refresh Schedule")
        apply_action = menu.addAction("Apply Schedule")

        has_event = schedule_id is not None
        open_action.setEnabled(has_event)
        delete_action.setEnabled(has_event)
        copy_action.setEnabled(has_event)
        cut_action.setEnabled(has_event)

        selected_action = menu.exec(self.table.viewport().mapToGlobal(position))
        if selected_action is None:
            return

        if selected_action == open_action:
            self.context_action_requested.emit("open", payload)
        elif selected_action == delete_action:
            self.context_action_requested.emit("delete", payload)
        elif selected_action == new_action:
            self.context_action_requested.emit("new", payload)
        elif selected_action == copy_action:
            self.context_action_requested.emit("copy", payload)
        elif selected_action == cut_action:
            self.context_action_requested.emit("cut", payload)
        elif selected_action == paste_action:
            self.context_action_requested.emit("paste", payload)
        elif selected_action == refresh_action:
            self.context_action_requested.emit("refresh", payload)
        elif selected_action == apply_action:
            self.context_action_requested.emit("apply", payload)
        elif selected_action == scale_30:
            self.context_action_requested.emit("time_scale", {**payload, "minutes": 30})
        elif selected_action == scale_60:
            self.context_action_requested.emit("time_scale", {**payload, "minutes": 60})
        elif selected_action == scale_120:
            self.context_action_requested.emit("time_scale", {**payload, "minutes": 120})

    def _render(self):
        self._ensure_items()
        self._clear_grid()

        for occurrence in self.occurrences:
            if self.week_mode:
                self._render_occurrence(occurrence)
            else:
                occurrence_date = QDate(occurrence.start.year, occurrence.start.month, occurrence.start.day)
                if occurrence_date == self.reference_date:
                    self._render_occurrence(occurrence)


class DayViewWidget(ScheduleTimeGridWidget):
    def __init__(self, parent=None):
        super().__init__(week_mode=False, parent=parent)


class WeekViewWidget(ScheduleTimeGridWidget):
    def __init__(self, parent=None):
        super().__init__(week_mode=True, parent=parent)

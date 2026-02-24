from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from typing import Dict, List

from PySide6.QtCore import QDate, Qt, Signal
from PySide6.QtWidgets import QLabel, QMenu, QTableWidget, QVBoxLayout, QWidget

from core.schedule_resolver import ResolvedOccurrence


def _month_grid_start(month_date: QDate) -> QDate:
    first_day = QDate(month_date.year(), month_date.month(), 1)
    days_to_sunday = first_day.dayOfWeek() % 7
    return first_day.addDays(-days_to_sunday)


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
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setVisible(True)
        self.table.verticalHeader().setDefaultSectionSize(112)
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
        layout = QVBoxLayout(container)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(3)

        date_label = QLabel(str(qdate.day()))
        if qdate == self.selected_date:
            date_label.setStyleSheet("font-weight: bold; color: #ffffff; background-color: #2f73d9; padding: 2px 4px; border-radius: 3px;")
        elif qdate.month() != self.reference_date.month():
            date_label.setStyleSheet("color: #808080;")
        else:
            date_label.setStyleSheet("font-weight: bold;")

        layout.addWidget(date_label)

        for occurrence in events[:3]:
            chip = QLabel(occurrence.title)
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
        open_action = menu.addAction("Open")
        delete_action = menu.addAction("Delete")
        menu.addSeparator()
        new_action = menu.addAction("New Event")
        copy_action = menu.addAction("Copy")
        cut_action = menu.addAction("Cut")
        paste_action = menu.addAction("Paste")
        menu.addSeparator()
        refresh_action = menu.addAction("Refresh Schedule")
        apply_action = menu.addAction("Apply Schedule")

        has_event = first_event is not None
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

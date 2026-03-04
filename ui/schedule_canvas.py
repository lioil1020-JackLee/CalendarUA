from __future__ import annotations

from datetime import datetime
from typing import List

from PySide6.QtCore import QDate, Qt, Signal
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import QHeaderView, QInputDialog, QMenu, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget

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
    time_scale_changed = Signal(int)
    TIME_SCALE_OPTIONS = (5, 6, 10, 15, 30, 60)

    def __init__(self, week_mode: bool, parent=None):
        super().__init__(parent)
        self.week_mode = week_mode
        self._is_dark_theme: bool | None = None
        self._fixed_row_height = 28
        self.time_scale_minutes = 60
        self.reference_date = QDate.currentDate()
        self.occurrences: List[ResolvedOccurrence] = []
        self._cell_occurrence_map: dict[tuple[int, int], List[ResolvedOccurrence]] = {}
        self._cell_start_titles: dict[tuple[int, int], List[str]] = {}

        self.table = QTableWidget(self._rows_per_day(), 7 if week_mode else 1)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        # 支援滑鼠左鍵雙擊：直接開啟編輯 / 新增視窗
        self.table.cellDoubleClicked.connect(self._on_cell_double_clicked)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table.verticalHeader().setDefaultSectionSize(self._fixed_row_height)
        self.table.verticalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.verticalHeader().customContextMenuRequested.connect(self._show_time_scale_menu)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setFixedHeight(32)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.table.horizontalHeader().setStyleSheet(
            "QHeaderView::section { padding: 0px 4px; font-family: 'Segoe UI'; font-weight: 600; }"
        )
        header_font = QFont(self.table.horizontalHeader().font())
        header_font.setFamily("Segoe UI")
        header_font.setPointSize(10)
        header_font.setBold(True)
        self.table.horizontalHeader().setFont(header_font)

        self._refresh_time_labels()
        self.apply_theme_style()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.table)

        self._refresh_headers()
        self._ensure_items()
        self._update_row_height_policy()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_row_height_policy()

    def showEvent(self, event):
        super().showEvent(event)
        self._update_row_height_policy()

    def relayout_to_viewport(self):
        self._update_row_height_policy()

    def _update_row_height_policy(self):
        # 需求：time span 改變時格高與時間字體保持固定，不隨視窗縮放。
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.table.verticalHeader().setDefaultSectionSize(self._fixed_row_height)

    def apply_theme_style(self, is_dark: bool | None = None):
        if isinstance(is_dark, bool):
            self._is_dark_theme = is_dark
        is_dark_palette = self._is_dark_theme if isinstance(self._is_dark_theme, bool) else (self.palette().window().color().lightness() < 128)
        if is_dark_palette:
            self.table.setStyleSheet(
                """
                QTableWidget {
                    background-color: #1e1e1e;
                    color: #cccccc;
                    gridline-color: #3d3d3d;
                }
                QTableWidget::item:selected {
                    background-color: #2f73d9;
                    color: #ffffff;
                }
                QHeaderView::section {
                    background-color: #252526;
                    color: #f0f0f0;
                    border: 1px solid #3d3d3d;
                    font-weight: 600;
                }
                """
            )
        else:
            self.table.setStyleSheet(
                """
                QTableWidget {
                    background-color: #ffffff;
                    color: #111111;
                    gridline-color: #e0e0e0;
                }
                QTableWidget::item:selected {
                    background-color: #2f73d9;
                    color: #ffffff;
                }
                QHeaderView::section {
                    background-color: #f0f0f0;
                    color: #111111;
                    border: 1px solid #d0d0d0;
                    font-weight: 600;
                }
                """
            )

    def _rows_per_day(self) -> int:
        return (24 * 60) // self.time_scale_minutes

    def _row_to_hour_minute(self, row: int) -> tuple[int, int]:
        total_minutes = row * self.time_scale_minutes
        return total_minutes // 60, total_minutes % 60

    def _minute_of_day_to_row(self, minute_of_day: int) -> int:
        minute_of_day = max(0, min((24 * 60) - 1, minute_of_day))
        return minute_of_day // self.time_scale_minutes

    def _refresh_time_labels(self):
        for row in range(self.table.rowCount()):
            hour, minute = self._row_to_hour_minute(row)
            if minute == 0:
                label = _hour_label(hour)
            else:
                label = f"{hour:02d}:{minute:02d}"
            self.table.setVerticalHeaderItem(row, QTableWidgetItem(label))

    def set_time_scale(self, minutes: int):
        if minutes not in self.TIME_SCALE_OPTIONS or minutes == self.time_scale_minutes:
            return

        self.time_scale_minutes = minutes
        self.table.setRowCount(self._rows_per_day())
        self._refresh_time_labels()
        self._ensure_items()
        self._render()
        self._update_row_height_policy()
        self.time_scale_changed.emit(minutes)

    def set_reference_date(self, qdate: QDate):
        self.reference_date = qdate
        self._refresh_headers()
        self._render()
        self._update_row_height_policy()

    def set_occurrences(self, occurrences: List[ResolvedOccurrence]):
        self.occurrences = occurrences
        self._render()
        self._update_row_height_policy()

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

    def _show_time_scale_menu(self, position):
        header = self.table.verticalHeader()
        if header.logicalIndexAt(position) < 0:
            return

        menu = QMenu(self)
        actions = {}
        for minutes in self.TIME_SCALE_OPTIONS:
            action = menu.addAction(f"{minutes}min")
            action.setCheckable(True)
            action.setChecked(minutes == self.time_scale_minutes)
            actions[action] = minutes

        selected = menu.exec(header.mapToGlobal(position))
        if selected in actions:
            self.set_time_scale(actions[selected])

    def _ensure_items(self):
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                if self.table.item(row, col) is None:
                    item = QTableWidgetItem("")
                    item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)
                    self.table.setItem(row, col, item)

    def _clear_grid(self):
        self._cell_occurrence_map.clear()
        self._cell_start_titles.clear()
        is_dark_palette = self._is_dark_theme if isinstance(self._is_dark_theme, bool) else (self.palette().window().color().lightness() < 128)
        empty_bg = QColor("#1e1e1e") if is_dark_palette else QColor("#ffffff")
        empty_fg = QColor("#cccccc") if is_dark_palette else QColor("#000000")
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item is None:
                    continue
                item.setText("")
                item.setBackground(empty_bg)
                item.setForeground(empty_fg)
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

        start_minute_of_day = (start.hour * 60) + start.minute
        end_minute_of_day = (end.hour * 60) + end.minute
        if end.second > 0 or end.microsecond > 0:
            end_minute_of_day += 1

        if end_minute_of_day <= start_minute_of_day:
            end_minute_of_day = start_minute_of_day + self.time_scale_minutes

        start_row = self._minute_of_day_to_row(start_minute_of_day)
        end_row = self._minute_of_day_to_row(end_minute_of_day - 1)
        end_row = min(self.table.rowCount() - 1, max(start_row, end_row))

        for row in range(start_row, end_row + 1):
            item = self.table.item(row, col)
            if item is None:
                continue

            cell_key = (row, col)
            self._cell_occurrence_map.setdefault(cell_key, []).append(occurrence)

            if row == start_row:
                self._cell_start_titles.setdefault(cell_key, []).append(occurrence.title)

            self._apply_cell_display(row, col)

    def _apply_cell_display(self, row: int, col: int):
        """根據同格 occurrence 數量更新文字、色彩與提示。"""
        item = self.table.item(row, col)
        if item is None:
            return

        cell_key = (row, col)
        occurrences = self._cell_occurrence_map.get(cell_key, [])
        if not occurrences:
            return

        titles = self._cell_start_titles.get(cell_key, [])
        item.setText("\n".join(titles))

        if len(occurrences) == 1:
            occ = occurrences[0]
            item.setBackground(QColor(occ.category_bg))
            item.setForeground(QColor(occ.category_fg))
        else:
            # 同一時間格有多筆任務時，使用統一藍底白字
            item.setBackground(QColor("#2f73d9"))
            item.setForeground(QColor("#ffffff"))

        tooltip_lines = []
        for occ in occurrences:
            tooltip_lines.append(
                f"{occ.title} ({occ.start.strftime('%H:%M')} - {occ.end.strftime('%H:%M')})"
            )
        item.setToolTip("\n".join(tooltip_lines))

    def _pick_occurrence(self, row: int, col: int, action_text: str) -> ResolvedOccurrence | None:
        """若同格有多筆任務，讓使用者選擇目標任務。"""
        occurrences = self._cell_occurrence_map.get((row, col), [])
        if not occurrences:
            return None
        if len(occurrences) == 1:
            return occurrences[0]

        options = []
        option_map: dict[str, ResolvedOccurrence] = {}
        for idx, occ in enumerate(occurrences, start=1):
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
            "同時段有多筆任務，請選擇：",
            options,
            0,
            False,
        )
        if not ok or not selected:
            return None
        return option_map.get(selected)

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
        occurrences = self._cell_occurrence_map.get((row, col), [])
        occurrence = occurrences[0] if occurrences else None
        schedule_id = occurrence.schedule_id if occurrence else None
        date_text = self._date_for_column(col).toString("yyyy-MM-dd")

        payload = {
            "schedule_id": schedule_id,
            "date": date_text,
            "hour": self._row_to_hour_minute(row)[0],
            "minute": self._row_to_hour_minute(row)[1],
            "week_mode": self.week_mode,
        }

        menu = QMenu(self)

        new_action = menu.addAction("新增行程 (New Appointment)")
        edit_action = menu.addAction("編輯行程 (Edit)")
        delete_action = menu.addAction("刪除行程 (Delete)")
        menu.addSeparator()
        today_action = menu.addAction("移至今天 (Go to Today)")
        refresh_action = menu.addAction("重新整理 (Refresh)")

        has_event = schedule_id is not None
        edit_action.setEnabled(has_event)
        delete_action.setEnabled(has_event)

        selected_action = menu.exec(self.table.viewport().mapToGlobal(position))
        if selected_action is None:
            return

        if selected_action == new_action:
            self.context_action_requested.emit("new", payload)
        elif selected_action == edit_action:
            target = self._pick_occurrence(row, col, "編輯")
            if target is None:
                return
            payload["schedule_id"] = target.schedule_id
            payload["hour"] = target.start.hour
            self.context_action_requested.emit("edit", payload)
        elif selected_action == delete_action:
            target = self._pick_occurrence(row, col, "刪除")
            if target is None:
                return
            payload["schedule_id"] = target.schedule_id
            payload["hour"] = target.start.hour
            self.context_action_requested.emit("delete", payload)
        elif selected_action == today_action:
            self.context_action_requested.emit("today", payload)
        elif selected_action == refresh_action:
            self.context_action_requested.emit("refresh", payload)

    def _on_cell_double_clicked(self, row: int, col: int):
        """滑鼠左鍵雙擊：僅在該格有行程時才進入編輯。"""
        occurrences = self._cell_occurrence_map.get((row, col), [])
        occurrence = occurrences[0] if occurrences else None
        schedule_id = occurrence.schedule_id if occurrence else None
        date_text = self._date_for_column(col).toString("yyyy-MM-dd")

        payload = {
            "schedule_id": schedule_id,
            "date": date_text,
            "hour": self._row_to_hour_minute(row)[0],
            "minute": self._row_to_hour_minute(row)[1],
            "week_mode": self.week_mode,
        }

        if schedule_id is None:
            return

        target = self._pick_occurrence(row, col, "編輯")
        if target is None:
            return
        payload["schedule_id"] = target.schedule_id
        payload["hour"] = target.start.hour
        self.context_action_requested.emit("edit", payload)

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

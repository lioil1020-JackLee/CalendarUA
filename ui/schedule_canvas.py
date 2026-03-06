from __future__ import annotations

from datetime import datetime
from typing import List

from PySide6.QtCore import QDate, QEvent, Qt, Signal
from PySide6.QtGui import QColor, QCursor, QFont, QPen, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHeaderView,
    QLabel,
    QListWidget,
    QMenu,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.schedule_resolver import ResolvedOccurrence
from ui.wheel_select_list import WheelSelectListWidget


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


class SelectedDayHeaderView(QHeaderView):
    """可依指定欄位繪製高亮背景的表頭。"""

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self._selected_column = -1
        self._is_dark_theme = False

    def set_selected_column(self, logical_index: int):
        if self._selected_column == logical_index:
            return
        self._selected_column = logical_index
        self.viewport().update()

    def set_dark_theme(self, is_dark: bool):
        if self._is_dark_theme == bool(is_dark):
            return
        self._is_dark_theme = bool(is_dark)
        self.viewport().update()

    def paintSection(self, painter, rect, logicalIndex):
        if not rect.isValid():
            return

        is_selected = logicalIndex == self._selected_column
        if self._is_dark_theme:
            bg = QColor("#2e7d32") if is_selected else QColor("#252526")
            fg = QColor("#ffffff") if is_selected else QColor("#f0f0f0")
            border = QColor("#3d3d3d")
        else:
            bg = QColor("#66bb6a") if is_selected else QColor("#f0f0f0")
            fg = QColor("#0f2d10") if is_selected else QColor("#111111")
            border = QColor("#d0d0d0")

        painter.save()
        painter.fillRect(rect, bg)
        painter.setPen(QPen(border))
        painter.drawRect(rect.adjusted(0, 0, -1, -1))

        text = ""
        model = self.model()
        if model is not None:
            header_value = model.headerData(logicalIndex, self.orientation(), Qt.DisplayRole)
            if header_value is not None:
                text = str(header_value)

        painter.setPen(fg)
        painter.setFont(self.font())
        painter.drawText(rect.adjusted(4, 0, -4, 0), Qt.AlignCenter, text)
        painter.restore()


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
        self._occurrence_bounds: dict[str, tuple[int, int, int, ResolvedOccurrence]] = {}
        self._drag_state: dict | None = None
        self._drag_preview_cells: list[tuple[int, int]] = []

        self.table = QTableWidget(self._rows_per_day(), 7 if week_mode else 1)
        self.table.setHorizontalHeader(SelectedDayHeaderView(Qt.Horizontal, self.table))
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.setFocusPolicy(Qt.StrongFocus)
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
        header_font = QFont(self.table.horizontalHeader().font())
        header_font.setFamily("Segoe UI")
        header_font.setPointSize(10)
        header_font.setBold(True)
        self.table.horizontalHeader().setFont(header_font)

        self._refresh_time_labels()
        self.apply_theme_style()
        self.table.viewport().setMouseTracking(True)
        self.table.viewport().installEventFilter(self)
        self.table.installEventFilter(self)

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
                    background-color: #2e7d32;
                    color: #ffffff;
                }
                QHeaderView::section {
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
                    background-color: #66bb6a;
                    color: #0f2d10;
                }
                QHeaderView::section {
                    border: 1px solid #d0d0d0;
                    font-weight: 600;
                }
                """
            )
        header = self.table.horizontalHeader()
        if isinstance(header, SelectedDayHeaderView):
            header.set_dark_theme(is_dark_palette)
        self._refresh_headers()

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
        header = self.table.horizontalHeader()

        if not self.week_mode:
            header = f"{self.reference_date.toString('yyyy/MM/dd')}"
            self.table.setHorizontalHeaderLabels([header])
            if isinstance(self.table.horizontalHeader(), SelectedDayHeaderView):
                self.table.horizontalHeader().set_selected_column(-1)
            return

        sunday = _week_start_sunday(self.reference_date)
        day_names = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"]
        selected_col = max(0, min(6, sunday.daysTo(self.reference_date)))
        labels = []
        for offset in range(7):
            day = sunday.addDays(offset)
            labels.append(f"{day_names[offset]} {day.day()}")
        self.table.setHorizontalHeaderLabels(labels)
        if isinstance(header, SelectedDayHeaderView):
            header.set_selected_column(selected_col)

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
        self._occurrence_bounds.clear()
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
        occurrence_key = self._occurrence_key(occurrence)
        self._occurrence_bounds[occurrence_key] = (start_row, end_row, col, occurrence)

        for row in range(start_row, end_row + 1):
            item = self.table.item(row, col)
            if item is None:
                continue

            cell_key = (row, col)
            self._cell_occurrence_map.setdefault(cell_key, []).append(occurrence)

            if row == start_row:
                self._cell_start_titles.setdefault(cell_key, []).append(occurrence.title)

            self._apply_cell_display(row, col)

    def _occurrence_key(self, occurrence: ResolvedOccurrence) -> str:
        key = str(getattr(occurrence, "occurrence_key", "") or "").strip()
        if key:
            return key
        return f"{occurrence.schedule_id}:{occurrence.start.isoformat()}"

    def _find_occurrence_at(self, row: int, col: int) -> ResolvedOccurrence | None:
        occurrences = self._cell_occurrence_map.get((row, col), [])
        if not occurrences:
            return None

        ranked: list[tuple[int, datetime, ResolvedOccurrence]] = []
        for occ in occurrences:
            bounds = self._occurrence_bounds.get(self._occurrence_key(occ))
            if not bounds:
                continue
            start_row, end_row, bound_col, _ = bounds
            if bound_col != col or row < start_row or row > end_row:
                continue
            ranked.append((end_row - start_row, occ.start, occ))

        if not ranked:
            return None

        ranked.sort(key=lambda item: (-item[0], item[1]))
        return ranked[0][2]

    def _drag_mode_for_position(self, occurrence: ResolvedOccurrence, col: int, y: int) -> str:
        bounds = self._occurrence_bounds.get(self._occurrence_key(occurrence))
        if not bounds:
            return "move"

        start_row, end_row, _bound_col, _ = bounds
        edge_px = 6

        start_item = self.table.item(start_row, col)
        end_item = self.table.item(end_row, col)
        if start_item is None or end_item is None:
            return "move"

        start_rect = self.table.visualItemRect(start_item)
        end_rect = self.table.visualItemRect(end_item)

        if abs(y - start_rect.top()) <= edge_px:
            return "resize_start"
        if abs(y - end_rect.bottom()) <= edge_px:
            return "resize_end"
        return "move"

    def _set_cursor_for_mode(self, mode: str | None):
        viewport = self.table.viewport()
        if mode == "move":
            viewport.setCursor(QCursor(Qt.OpenHandCursor))
        elif mode in ("resize_start", "resize_end"):
            viewport.setCursor(QCursor(Qt.SizeVerCursor))
        else:
            viewport.unsetCursor()

    def _update_hover_cursor(self, pos):
        index = self.table.indexAt(pos)
        if not index.isValid():
            self._set_cursor_for_mode(None)
            return

        row = index.row()
        col = index.column()
        occurrence = self._find_occurrence_at(row, col)
        if occurrence is None:
            self._set_cursor_for_mode(None)
            return

        mode = self._drag_mode_for_position(occurrence, col, pos.y())
        self._set_cursor_for_mode(mode)

    def _minute_range_for_occurrence(self, occurrence: ResolvedOccurrence) -> tuple[int, int]:
        start_minute = (occurrence.start.hour * 60) + occurrence.start.minute
        end_minute = (occurrence.end.hour * 60) + occurrence.end.minute
        if occurrence.end.second > 0 or occurrence.end.microsecond > 0:
            end_minute += 1
        if end_minute <= start_minute:
            end_minute = start_minute + self.time_scale_minutes
        end_minute = min(24 * 60, end_minute)
        return start_minute, end_minute

    def _compute_drag_target_minutes(
        self,
        mode: str,
        old_start: int,
        old_end: int,
        delta_minutes: int,
    ) -> tuple[int, int]:
        if mode == "move":
            duration = old_end - old_start
            new_start = old_start + delta_minutes
            new_start = max(0, min((24 * 60) - duration, new_start))
            new_end = new_start + duration
            return new_start, new_end

        if mode == "resize_start":
            min_end = old_end - self.time_scale_minutes
            new_start = max(0, min(min_end, old_start + delta_minutes))
            return new_start, old_end

        max_end = 24 * 60
        min_end = old_start + self.time_scale_minutes
        new_end = max(min_end, min(max_end, old_end + delta_minutes))
        return old_start, new_end

    def _clear_drag_preview(self):
        if not self._drag_preview_cells:
            return
        self._drag_preview_cells.clear()
        self._render()

    def _show_drag_preview(self, start_minute: int, end_minute: int, col: int):
        self._render()

        is_dark_palette = self._is_dark_theme if isinstance(self._is_dark_theme, bool) else (self.palette().window().color().lightness() < 128)
        preview_bg = QColor("#f59e0b")
        preview_bg.setAlpha(150 if is_dark_palette else 120)
        preview_fg = QColor("#ffffff") if is_dark_palette else QColor("#3d2b00")

        start_row = self._minute_of_day_to_row(start_minute)
        end_row = self._minute_of_day_to_row(max(start_minute, end_minute - 1))
        end_row = min(self.table.rowCount() - 1, max(start_row, end_row))

        preview_cells: list[tuple[int, int]] = []
        for row in range(start_row, end_row + 1):
            item = self.table.item(row, col)
            if item is None:
                continue
            item.setBackground(preview_bg)
            item.setForeground(preview_fg)
            preview_cells.append((row, col))

        self._drag_preview_cells = preview_cells

    def _finalize_drag(self, release_pos):
        if not self._drag_state:
            return

        state = self._drag_state
        if not state.get("active", False):
            self._clear_drag_preview()
            self._drag_state = None
            self._set_cursor_for_mode(None)
            return
        row_height = max(1, self.table.verticalHeader().defaultSectionSize())
        delta_y = release_pos.y() - state["press_y"]
        delta_rows = int(round(delta_y / row_height))
        delta_minutes = delta_rows * self.time_scale_minutes

        old_start = state["start_minute"]
        old_end = state["end_minute"]
        mode = state["mode"]

        new_start, new_end = self._compute_drag_target_minutes(mode, old_start, old_end, delta_minutes)

        if new_start != old_start or new_end != old_end:
            occurrence = state["occurrence"]
            payload = {
                "schedule_id": occurrence.schedule_id,
                "date": occurrence.start.strftime("%Y-%m-%d"),
                "start_minute": int(new_start),
                "end_minute": int(new_end),
                "scale_minutes": int(self.time_scale_minutes),
                "week_mode": self.week_mode,
                "source": "time_grid",
            }
            self.context_action_requested.emit("drag_update", payload)

        self._clear_drag_preview()
        self._drag_state = None
        self._set_cursor_for_mode(None)

    def eventFilter(self, watched, event):
        if watched is self.table and event.type() == QEvent.KeyPress:
            index = self.table.currentIndex()
            if not index.isValid():
                selected = self.table.selectedIndexes()
                if selected:
                    index = selected[0]

            if index.isValid():
                row = index.row()
                col = index.column()
                key = event.key()
                mods = event.modifiers()
                if (mods & Qt.ControlModifier) and key == Qt.Key_C:
                    self._trigger_action_for_cell("copy", row, col)
                    event.accept()
                    return True
                if (mods & Qt.ControlModifier) and key == Qt.Key_V:
                    self._trigger_action_for_cell("paste", row, col)
                    event.accept()
                    return True
                if key in (Qt.Key_Delete,):
                    self._trigger_action_for_cell("delete", row, col)
                    event.accept()
                    return True

        viewport = self.table.viewport()
        if watched is viewport:
            if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                index = self.table.indexAt(event.pos())
                if index.isValid():
                    row = index.row()
                    col = index.column()
                    occurrence = self._find_occurrence_at(row, col)
                    if occurrence is not None:
                        mode = self._drag_mode_for_position(occurrence, col, event.pos().y())
                        start_minute, end_minute = self._minute_range_for_occurrence(occurrence)
                        self._drag_state = {
                            "occurrence": occurrence,
                            "mode": mode,
                            "press_y": event.pos().y(),
                            "start_minute": start_minute,
                            "end_minute": end_minute,
                            "active": False,
                            "col": col,
                        }
                        return False

            elif event.type() == QEvent.MouseMove:
                if self._drag_state:
                    if not (event.buttons() & Qt.LeftButton):
                        self._clear_drag_preview()
                        self._drag_state = None
                        self._set_cursor_for_mode(None)
                        return False

                    if not self._drag_state.get("active", False):
                        if abs(event.pos().y() - self._drag_state["press_y"]) < 4:
                            return False
                        self._drag_state["active"] = True

                    mode = self._drag_state.get("mode")
                    row_height = max(1, self.table.verticalHeader().defaultSectionSize())
                    delta_y = event.pos().y() - self._drag_state["press_y"]
                    delta_rows = int(round(delta_y / row_height))
                    delta_minutes = delta_rows * self.time_scale_minutes
                    new_start, new_end = self._compute_drag_target_minutes(
                        str(mode),
                        int(self._drag_state["start_minute"]),
                        int(self._drag_state["end_minute"]),
                        int(delta_minutes),
                    )
                    self._show_drag_preview(new_start, new_end, int(self._drag_state["col"]))

                    if mode == "move":
                        self.table.viewport().setCursor(QCursor(Qt.ClosedHandCursor))
                    else:
                        self._set_cursor_for_mode(mode)
                    event.accept()
                    return True
                self._update_hover_cursor(event.pos())

            elif event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
                if self._drag_state:
                    if not self._drag_state.get("active", False):
                        self._clear_drag_preview()
                        self._drag_state = None
                        self._set_cursor_for_mode(None)
                        return False
                    self._finalize_drag(event.pos())
                    event.accept()
                    return True

            elif event.type() == QEvent.Leave:
                if not self._drag_state:
                    self._set_cursor_for_mode(None)

        return super().eventFilter(watched, event)

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
            # 同一時間格有多筆任務時，使用統一深青底白字
            item.setBackground(QColor("#0f766e"))
            item.setForeground(QColor("#ffffff"))

        tooltip_lines = []
        for occ in occurrences:
            tooltip_lines.append(
                f"{occ.title} ({occ.start.strftime('%H:%M')} - {occ.end.strftime('%H:%M')})"
            )
        item.setToolTip("\n".join(tooltip_lines))

    def _pick_occurrences(self, row: int, col: int, action_text: str, allow_multi: bool = False) -> list[ResolvedOccurrence] | None:
        """若同格有多筆任務，讓使用者選擇目標任務（可多選）。"""
        occurrences = sorted(
            self._cell_occurrence_map.get((row, col), []),
            key=lambda occ: occ.start,
        )
        if not occurrences:
            return None
        if len(occurrences) == 1:
            return [occurrences[0]]

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

        cell_date = self._date_for_column(col)

        dialog = QDialog(self)
        dialog.setWindowTitle(f"選擇要{action_text}的行程")
        dialog.setModal(True)
        dialog.setMinimumWidth(520)

        layout = QVBoxLayout(dialog)
        prompt = QLabel(f"{cell_date.toString('yyyy-MM-dd')} 有多筆任務，請選擇：")
        layout.addWidget(prompt)

        options_list = WheelSelectListWidget(dialog)
        options_list.addItems(options)
        options_list.setCurrentRow(0)
        options_list.setSelectionMode(QListWidget.ExtendedSelection if allow_multi else QListWidget.SingleSelection)
        options_list.set_hover_select_enabled(not allow_multi)
        options_list.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        row_height = max(24, options_list.fontMetrics().height() + 10)
        options_list.setStyleSheet(f"QListWidget::item {{ min-height: {row_height}px; max-height: {row_height}px; }}")
        frame = options_list.frameWidth() * 2
        options_list.setFixedHeight((row_height * 10) + frame)
        layout.addWidget(options_list)
        if allow_multi:
            tips = QLabel("提示：Ctrl 可不連續多選、Shift 可連續多選，選好後按 OK")
            layout.addWidget(tips)
            select_all_shortcut = QShortcut(QKeySequence("Ctrl+A"), dialog)
            select_all_shortcut.activated.connect(options_list.selectAll)
            buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=dialog)
            buttons.accepted.connect(dialog.accept)
            buttons.rejected.connect(dialog.reject)
            layout.addWidget(buttons)
        else:
            options_list.itemClicked.connect(lambda _item: dialog.accept())

        options_list.setFocus(Qt.OtherFocusReason)

        if dialog.exec() != QDialog.Accepted:
            return None
        selected_items = options_list.selectedItems() if allow_multi else [options_list.currentItem()]
        if allow_multi and not selected_items:
            current = options_list.currentItem()
            if current is not None:
                selected_items = [current]
        labels = [item.text() for item in selected_items if item is not None and item.text()]
        if not labels:
            return None
        targets = [option_map[label] for label in labels if label in option_map]
        return targets or None

    def _date_for_column(self, col: int) -> QDate:
        if not self.week_mode:
            return self.reference_date

        sunday = _week_start_sunday(self.reference_date)
        return sunday.addDays(col)

    def _selected_cells_or_current(self, row: int, col: int) -> list[tuple[int, int]]:
        selected_indexes = self.table.selectedIndexes()
        cells: list[tuple[int, int]] = []
        seen: set[tuple[int, int]] = set()

        for index in selected_indexes:
            cell = (index.row(), index.column())
            if cell in seen:
                continue
            seen.add(cell)
            cells.append(cell)

        if (row, col) not in seen:
            cells.append((row, col))
        return cells

    def _selected_single_occurrence_ids(self, row: int, col: int) -> list[int]:
        """取得目前選取格中「單一任務格」的 schedule_id，供批次操作。"""
        ids: list[int] = []
        seen: set[int] = set()
        for r, c in self._selected_cells_or_current(row, col):
            occurrences = self._cell_occurrence_map.get((r, c), [])
            if len(occurrences) != 1:
                continue
            sid = int(occurrences[0].schedule_id)
            if sid in seen:
                continue
            seen.add(sid)
            ids.append(sid)
        return ids

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

        new_action = menu.addAction("新增行程 (New)")
        copy_action = menu.addAction("複製行程 (Copy)")
        paste_action = menu.addAction("貼上行程 (Paste)")
        delete_action = menu.addAction("刪除行程 (Delete)")

        has_event = schedule_id is not None
        delete_action.setEnabled(has_event)
        copy_action.setEnabled(has_event)

        selected_action = menu.exec(self.table.viewport().mapToGlobal(position))
        if selected_action is None:
            return

        if selected_action == new_action:
            self.context_action_requested.emit("new", payload)
        elif selected_action == copy_action:
            batch_ids = self._selected_single_occurrence_ids(row, col)
            if len(batch_ids) > 1:
                payload["schedule_id"] = batch_ids[0]
                payload["schedule_ids"] = batch_ids
                self.context_action_requested.emit("copy", payload)
                return

            targets = self._pick_occurrences(row, col, "複製", allow_multi=True)
            if not targets:
                return
            first = targets[0]
            payload["schedule_id"] = first.schedule_id
            payload["schedule_ids"] = [t.schedule_id for t in targets]
            payload["hour"] = first.start.hour
            payload["minute"] = first.start.minute
            self.context_action_requested.emit("copy", payload)
        elif selected_action == paste_action:
            self.context_action_requested.emit("paste", payload)
        elif selected_action == delete_action:
            batch_ids = self._selected_single_occurrence_ids(row, col)
            if len(batch_ids) > 1:
                payload["schedule_id"] = batch_ids[0]
                payload["schedule_ids"] = batch_ids
                self.context_action_requested.emit("delete", payload)
                return

            targets = self._pick_occurrences(row, col, "刪除", allow_multi=True)
            if not targets:
                return
            first = targets[0]
            payload["schedule_id"] = first.schedule_id
            payload["schedule_ids"] = [t.schedule_id for t in targets]
            payload["hour"] = first.start.hour
            payload["minute"] = first.start.minute
            self.context_action_requested.emit("delete", payload)

    def _trigger_action_for_cell(self, action: str, row: int, col: int):
        occurrences = self._cell_occurrence_map.get((row, col), [])
        occurrence = occurrences[0] if occurrences else None
        schedule_id = occurrence.schedule_id if occurrence else None
        payload = {
            "schedule_id": schedule_id,
            "date": self._date_for_column(col).toString("yyyy-MM-dd"),
            "hour": self._row_to_hour_minute(row)[0],
            "minute": self._row_to_hour_minute(row)[1],
            "week_mode": self.week_mode,
        }

        if action == "copy":
            if schedule_id is None:
                return

            batch_ids = self._selected_single_occurrence_ids(row, col)
            if len(batch_ids) > 1:
                payload["schedule_id"] = batch_ids[0]
                payload["schedule_ids"] = batch_ids
                self.context_action_requested.emit("copy", payload)
                return

            targets = self._pick_occurrences(row, col, "複製", allow_multi=True)
            if not targets:
                return
            first = targets[0]
            payload["schedule_id"] = first.schedule_id
            payload["schedule_ids"] = [t.schedule_id for t in targets]
            payload["hour"] = first.start.hour
            payload["minute"] = first.start.minute
            self.context_action_requested.emit("copy", payload)
            return

        if action == "paste":
            self.context_action_requested.emit("paste", payload)
            return

        if action == "delete":
            if schedule_id is None:
                return

            batch_ids = self._selected_single_occurrence_ids(row, col)
            if len(batch_ids) > 1:
                payload["schedule_id"] = batch_ids[0]
                payload["schedule_ids"] = batch_ids
                self.context_action_requested.emit("delete", payload)
                return

            targets = self._pick_occurrences(row, col, "刪除", allow_multi=True)
            if not targets:
                return
            first = targets[0]
            payload["schedule_id"] = first.schedule_id
            payload["schedule_ids"] = [t.schedule_id for t in targets]
            payload["hour"] = first.start.hour
            payload["minute"] = first.start.minute
            self.context_action_requested.emit("delete", payload)

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

        targets = self._pick_occurrences(row, col, "編輯", allow_multi=False)
        if not targets:
            return
        target = targets[0]
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

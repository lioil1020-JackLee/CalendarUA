from __future__ import annotations

from collections import defaultdict
from datetime import datetime, date, timedelta
from typing import Callable, Dict, List, Optional

from PySide6.QtCore import QDate, Qt, Signal, QEvent
from PySide6.QtGui import QCursor, QFont, QColor, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHeaderView,
    QLabel,
    QListWidget,
    QMenu,
    QTableWidget,
    QVBoxLayout,
    QWidget,
)

from core.schedule_resolver import ResolvedOccurrence
from core.lunar_calendar import to_lunar, LunarDateInfo
from ui.wheel_select_list import WheelSelectListWidget


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


class MergedEventLabel(QLabel):
    merged_double_clicked = Signal()

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.merged_double_clicked.emit()
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
        self._holiday_checker: Optional[Callable[[QDate], bool]] = None
        self._cell_dates: Dict[tuple[int, int], QDate] = {}
        self.time_scale_minutes = 60
        self._drag_state: Optional[Dict[str, object]] = None
        self._drag_preview_date: Optional[QDate] = None

        self.table = QTableWidget(6, 7)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionMode(QTableWidget.NoSelection)
        self.table.setFocusPolicy(Qt.StrongFocus)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
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
        self.table.viewport().setMouseTracking(True)
        self.table.viewport().installEventFilter(self)
        self.table.installEventFilter(self)

        self.table.setHorizontalHeaderLabels(["週日", "週一", "週二", "週三", "週四", "週五", "週六"])
        self._grouped_events: Dict[QDate, List[ResolvedOccurrence]] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.table)

        self._render()

    def set_time_scale(self, minutes: int):
        if isinstance(minutes, int) and minutes > 0:
            self.time_scale_minutes = minutes

    def set_reference_date(self, qdate: QDate):
        self.reference_date = qdate
        self._render()

    def set_selected_date(self, qdate: QDate):
        self.selected_date = qdate
        self._render()

    def set_occurrences(self, occurrences: List[ResolvedOccurrence]):
        self.occurrences = occurrences
        self._render()

    def set_holiday_checker(self, checker: Optional[Callable[[QDate], bool]]):
        self._holiday_checker = checker
        self._render()

    def _is_holiday(self, qdate: QDate) -> bool:
        if qdate.dayOfWeek() in (6, 7):
            return True
        if self._holiday_checker is None:
            return False
        try:
            return bool(self._holiday_checker(qdate))
        except Exception:
            return False

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
        is_drag_preview = self._drag_preview_date is not None and qdate == self._drag_preview_date
        is_today = qdate == QDate.currentDate()
        is_holiday = self._is_holiday(qdate)
        is_dark_palette = self.palette().window().color().lightness() < 128

        if is_drag_preview:
            container.setStyleSheet(
                "background-color: rgba(245, 158, 11, 0.28); border: 2px solid #d97706; border-radius: 4px;"
            )
        elif is_selected and is_today:
            container.setStyleSheet(
                "background-color: rgba(56, 142, 60, 0.35); border: 2px solid #f4c542; border-radius: 4px;"
            )
        elif is_selected:
            container.setStyleSheet(
                "background-color: rgba(76, 175, 80, 0.25); border: 1px solid #2e7d32; border-radius: 4px;"
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
            cross_month_color = "#b36b6b" if is_holiday else "#808080"
            date_label.setStyleSheet(f"font-weight: bold; color: {cross_month_color};")
        elif is_selected and is_holiday:
            date_label.setStyleSheet("font-weight: bold; color: #c62828;")
        elif is_today:
            today_color = "#ff8f00"
            date_label.setStyleSheet(f"font-weight: bold; color: {today_color};")
        elif is_holiday:
            date_label.setStyleSheet("font-weight: bold; color: #c62828;")
        elif is_selected:
            selected_color = "#f0f0f0" if is_dark_palette else "#111111"
            date_label.setStyleSheet(f"font-weight: bold; color: {selected_color};")
        else:
            date_label.setStyleSheet("font-weight: bold;")

        layout.addWidget(date_label)

        if len(events) >= 3:
            merged_label = MergedEventLabel(f"{len(events)} 筆任務")
            merged_label.setStyleSheet(
                "background-color: #2f73d9;"
                "color: #ffffff;"
                "border-radius: 8px;"
                "padding: 2px 6px;"
                "font-weight: 600;"
            )
            merged_label.setToolTip(
                "\n".join(
                    f"{occ.title} ({occ.start.strftime('%H:%M')} - {occ.end.strftime('%H:%M')})"
                    for occ in events
                )
            )
            merged_label.merged_double_clicked.connect(
                lambda cell_date=qdate: self._on_merged_label_double_clicked(cell_date)
            )
            layout.addWidget(merged_label)
        else:
            for occurrence in events:
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

    def _set_month_cursor(self, mode: Optional[str]):
        viewport = self.table.viewport()
        if mode == "move":
            viewport.setCursor(QCursor(Qt.OpenHandCursor))
        elif mode in ("resize_start", "resize_end"):
            viewport.setCursor(QCursor(Qt.SizeHorCursor))
        else:
            viewport.unsetCursor()

    def _set_drag_preview_date(self, qdate: Optional[QDate]):
        if qdate is None and self._drag_preview_date is None:
            return
        if qdate is not None and self._drag_preview_date is not None and qdate == self._drag_preview_date:
            return
        self._drag_preview_date = qdate
        self._render()

    def _cell_mode_for_position(self, index, pos) -> str:
        rect = self.table.visualRect(index)
        edge = max(8, rect.width() // 6)
        if pos.x() - rect.left() <= edge:
            return "resize_start"
        if rect.right() - pos.x() <= edge:
            return "resize_end"
        return "move"

    def _update_month_hover_cursor(self, pos):
        index = self.table.indexAt(pos)
        if not index.isValid():
            self._set_month_cursor(None)
            return

        qdate = self._cell_dates.get((index.row(), index.column()))
        if qdate is None:
            self._set_month_cursor(None)
            return

        events = self._grouped_events.get(qdate, [])
        if not events:
            self._set_month_cursor(None)
            return

        mode = self._cell_mode_for_position(index, pos)
        self._set_month_cursor(mode)

    def _finalize_month_drag(self, release_pos):
        if not self._drag_state:
            return

        state = self._drag_state
        if not state.get("active", False):
            self._set_drag_preview_date(None)
            self._drag_state = None
            self._set_month_cursor(None)
            return
        index = self.table.indexAt(release_pos)
        if index.isValid():
            target_date = self._cell_dates.get((index.row(), index.column()), state["source_date"])
        else:
            target_date = state["source_date"]

        source_qdate: QDate = state["source_date"]
        target_qdate: QDate = target_date
        day_delta = source_qdate.daysTo(target_qdate)

        if day_delta != 0:
            occurrence: ResolvedOccurrence = state["occurrence"]
            old_start = occurrence.start
            old_end = occurrence.end
            step = timedelta(minutes=max(1, int(self.time_scale_minutes)))

            if state["mode"] == "move":
                new_start = old_start + timedelta(days=day_delta)
                new_end = old_end + timedelta(days=day_delta)
            elif state["mode"] == "resize_start":
                new_start = old_start + timedelta(days=day_delta)
                max_start = old_end - step
                if new_start > max_start:
                    new_start = max_start
                new_end = old_end
            else:
                new_end = old_end + timedelta(days=day_delta)
                min_end = old_start + step
                if new_end < min_end:
                    new_end = min_end
                new_start = old_start

            payload = {
                "schedule_id": occurrence.schedule_id,
                "start_datetime": new_start.strftime("%Y-%m-%d %H:%M:%S"),
                "end_datetime": new_end.strftime("%Y-%m-%d %H:%M:%S"),
                "scale_minutes": int(self.time_scale_minutes),
                "source": "month_grid",
            }
            self.context_action_requested.emit("drag_update", payload)

        self._set_drag_preview_date(None)
        self._drag_state = None
        self._set_month_cursor(None)

    def eventFilter(self, watched, event):
        if watched is self.table and event.type() == QEvent.KeyPress:
            index = self.table.currentIndex()
            if not index.isValid():
                for (row, col), qdate in self._cell_dates.items():
                    if qdate == self.selected_date:
                        index = self.table.model().index(row, col)
                        break

            if index.isValid():
                qdate = self._cell_dates.get((index.row(), index.column()))
                if qdate is not None:
                    key = event.key()
                    mods = event.modifiers()
                    if (mods & Qt.ControlModifier) and key == Qt.Key_C:
                        self._trigger_action_for_date("copy", qdate)
                        event.accept()
                        return True
                    if (mods & Qt.ControlModifier) and key == Qt.Key_V:
                        self._trigger_action_for_date("paste", qdate)
                        event.accept()
                        return True
                    if key in (Qt.Key_Delete,):
                        self._trigger_action_for_date("delete", qdate)
                        event.accept()
                        return True

        viewport = self.table.viewport()
        if watched is viewport:
            if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                index = self.table.indexAt(event.pos())
                if index.isValid():
                    qdate = self._cell_dates.get((index.row(), index.column()))
                    if qdate is not None:
                        events = self._grouped_events.get(qdate, [])
                        if events:
                            mode = self._cell_mode_for_position(index, event.pos())
                            self._drag_state = {
                                "source_date": qdate,
                                "mode": mode,
                                "events": events,
                                "active": False,
                                "press_pos": event.pos(),
                            }
                            return False

            elif event.type() == QEvent.MouseMove:
                if self._drag_state:
                    if not (event.buttons() & Qt.LeftButton):
                        self._set_drag_preview_date(None)
                        self._drag_state = None
                        self._set_month_cursor(None)
                        return False

                    if not self._drag_state.get("active", False):
                        press_pos = self._drag_state.get("press_pos")
                        if press_pos is None:
                            self._drag_state = None
                            self._set_month_cursor(None)
                            return False

                        delta = event.pos() - press_pos
                        if abs(delta.x()) < 6 and abs(delta.y()) < 6:
                            return False

                        source_qdate: QDate = self._drag_state["source_date"]
                        events: List[ResolvedOccurrence] = list(self._drag_state.get("events", []))
                        picked = self._pick_events(source_qdate, events, "拖曳調整", allow_multi=False)
                        if not picked:
                            self._set_drag_preview_date(None)
                            self._drag_state = None
                            self._set_month_cursor(None)
                            return False
                        occurrence = picked[0]

                        self._drag_state["occurrence"] = occurrence
                        self._drag_state["active"] = True

                    mode = str(self._drag_state.get("mode", "move"))
                    index = self.table.indexAt(event.pos())
                    if index.isValid():
                        target_date = self._cell_dates.get((index.row(), index.column()))
                    else:
                        target_date = self._drag_state.get("source_date")
                    if isinstance(target_date, QDate):
                        self._set_drag_preview_date(target_date)

                    if mode == "move":
                        self.table.viewport().setCursor(QCursor(Qt.ClosedHandCursor))
                    else:
                        self._set_month_cursor(mode)
                    event.accept()
                    return True
                self._update_month_hover_cursor(event.pos())

            elif event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
                if self._drag_state:
                    if not self._drag_state.get("active", False):
                        self._set_drag_preview_date(None)
                        self._drag_state = None
                        self._set_month_cursor(None)
                        return False
                    self._finalize_month_drag(event.pos())
                    event.accept()
                    return True

            elif event.type() == QEvent.Leave:
                if not self._drag_state:
                    self._set_drag_preview_date(None)
                    self._set_month_cursor(None)

        return super().eventFilter(watched, event)

    def _trigger_action_for_date(self, action: str, qdate: QDate):
        events = self._grouped_events.get(qdate, [])
        first_event = events[0] if events else None
        payload = {
            "schedule_id": first_event.schedule_id if first_event else None,
            "date": qdate.toString("yyyy-MM-dd"),
            "hour": first_event.start.hour if first_event else 8,
            "minute": first_event.start.minute if first_event else 0,
            "week_mode": False,
            "month_mode": True,
        }

        if action == "copy":
            if first_event is None:
                return
            targets = self._pick_events(qdate, events, "複製", allow_multi=True)
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
            if first_event is None:
                return
            targets = self._pick_events(qdate, events, "刪除", allow_multi=True)
            if not targets:
                return
            first = targets[0]
            payload["schedule_id"] = first.schedule_id
            payload["schedule_ids"] = [t.schedule_id for t in targets]
            payload["hour"] = first.start.hour
            payload["minute"] = first.start.minute
            self.context_action_requested.emit("delete", payload)

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

    def _on_merged_label_double_clicked(self, qdate: QDate):
        """雙擊合併任務區塊時，先讓使用者選擇目標任務再編輯。"""
        events = self._grouped_events.get(qdate, [])
        targets = self._pick_events(qdate, events, "編輯", allow_multi=False)
        if not targets:
            return
        target = targets[0]

        payload = {
            "schedule_id": target.schedule_id,
            "date": qdate.toString("yyyy-MM-dd"),
            "hour": target.start.hour,
            "week_mode": False,
            "month_mode": True,
        }
        self.context_action_requested.emit("edit", payload)

    def _pick_events(self, qdate: QDate, events: List[ResolvedOccurrence], action_text: str, allow_multi: bool = False) -> list[ResolvedOccurrence] | None:
        """同一天多筆任務時，讓使用者選擇目標任務（可多選）。"""
        events = sorted(events, key=lambda occ: occ.start)
        if not events:
            return None
        if len(events) == 1:
            return [events[0]]

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

        dialog = QDialog(self)
        dialog.setWindowTitle(f"選擇要{action_text}的行程")
        dialog.setModal(True)
        dialog.setMinimumWidth(520)

        layout = QVBoxLayout(dialog)
        prompt = QLabel(f"{qdate.toString('yyyy-MM-dd')} 有多筆任務，請選擇：")
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
            "minute": first_event.start.minute if first_event else 0,
            "week_mode": False,
            "month_mode": True,
        }

        menu = QMenu(self)
        new_action = menu.addAction("新增行程 (New)")
        copy_action = menu.addAction("複製行程 (Copy)")
        paste_action = menu.addAction("貼上行程 (Paste)")
        delete_action = menu.addAction("刪除行程 (Delete)")

        has_event = first_event is not None
        delete_action.setEnabled(has_event)
        copy_action.setEnabled(has_event)

        selected_action = menu.exec(self.table.viewport().mapToGlobal(position))
        if selected_action is None:
            return

        if selected_action == new_action:
            self.context_action_requested.emit("new", payload)
        elif selected_action == copy_action:
            targets = self._pick_events(qdate, events, "複製", allow_multi=True)
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
            targets = self._pick_events(qdate, events, "刪除", allow_multi=True)
            if not targets:
                return
            first = targets[0]
            payload["schedule_id"] = first.schedule_id
            payload["schedule_ids"] = [t.schedule_id for t in targets]
            payload["hour"] = first.start.hour
            payload["minute"] = first.start.minute
            self.context_action_requested.emit("delete", payload)

"""
Weekly Tab - 週間班表編輯面板
用於建立與編輯每週重複的事件（recurring weekly schedules）
"""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QMenu, QDialog
)
from PySide6.QtCore import Qt, Signal, QTime
from PySide6.QtGui import QColor, QBrush, QPainter
from datetime import datetime, time, timedelta
from typing import List, Dict, Any, Optional
from core.schedule_resolver import resolve_occurrences_for_range


class WeeklyEventCell(QTableWidgetItem):
    """週間事件格子，儲存事件資訊"""
    def __init__(self, event_data: Optional[Dict[str, Any]] = None):
        super().__init__()
        self.event_data = event_data
        if event_data:
            title = event_data.get("title", "")
            value = event_data.get("target_value", "")
            self.setText(f"{title}\n{value}" if value else title)

            bg_color = event_data.get("bg_color")
            fg_color = event_data.get("fg_color")
            if bg_color:
                self.setBackground(QBrush(QColor(bg_color)))
            if fg_color:
                self.setForeground(QBrush(QColor(fg_color)))


class WeeklyGridWidget(QTableWidget):
    """週間時間軸格子（7天 × 24小時）"""
    
    event_double_clicked = Signal(int, int, dict)  # (day_index, hour, event_data)
    cell_double_clicked = Signal(int, int)  # (day_index, hour) - 空格子建立新事件
    event_context_requested = Signal(str, dict)  # (action, payload)
    selection_changed = Signal(object)  # event_data or None
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.events_map = {}  # {(day_idx, hour): event_data}
        self._init_ui()
        
    def _init_ui(self):
        # 24 行（小時） × 7 列（星期）
        self.setRowCount(24)
        self.setColumnCount(7)
        
        # 標題
        days = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"]
        self.setHorizontalHeaderLabels(days)
        
        hours = [f"{h:02d}:00" for h in range(24)]
        self.setVerticalHeaderLabels(hours)
        
        # 初始化空格子
        for row in range(24):
            for col in range(7):
                item = WeeklyEventCell(None)
                item.setTextAlignment(Qt.AlignCenter)
                self.setItem(row, col, item)
        
        # 樣式
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.setEditTriggers(QTableWidget.NoEditTriggers)
        self.setSelectionMode(QTableWidget.SingleSelection)
        
        # 訊號
        self.cellDoubleClicked.connect(self._on_cell_double_clicked)
        self.itemSelectionChanged.connect(self._on_selection_changed)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_context_menu)
    
    def load_weekly_events(self, schedules: List[Dict[str, Any]]):
        """載入週間事件（篩選出 FREQ=WEEKLY 的 recurring schedules）"""
        self.events_map.clear()
        
        # 篩選週間 recurring 排程
        weekly_schedules = []
        for sch in schedules:
            rrule_str = sch.get("rrule_str", "")
            if "FREQ=WEEKLY" in rrule_str:
                weekly_schedules.append(sch)
        
        # 展開到本週
        from datetime import datetime, timedelta
        today = datetime.now().date()
        # 找到本週日
        weekday = today.weekday()  # 0=Monday, 6=Sunday
        week_start = today - timedelta(days=(weekday + 1) % 7)
        range_start = datetime.combine(week_start, time.min)
        range_end = range_start + timedelta(days=7)
        
        occurrences = resolve_occurrences_for_range(
            weekly_schedules,
            range_start,
            range_end,
            [],
            None,
            self.parent().db_manager if hasattr(self.parent(), "db_manager") else None,
        )
        
        # 映射到格子
        for occ in occurrences:
            start_dt = occ.start
            day_idx = (start_dt.weekday() + 1) % 7  # 轉換成週日=0
            hour = start_dt.hour
            
            event_data = {
                "schedule_id": occ.schedule_id,
                "title": occ.title,
                "target_value": occ.target_value,
                "bg_color": occ.category_bg,
                "fg_color": occ.category_fg,
                "start": occ.start,
                "end": occ.end,
            }
            self.events_map[(day_idx, hour)] = event_data
        
        self._refresh_grid()
    
    def _refresh_grid(self):
        """重繪格子"""
        for row in range(24):
            for col in range(7):
                event_data = self.events_map.get((col, row))
                item = WeeklyEventCell(event_data)
                item.setTextAlignment(Qt.AlignCenter)
                self.setItem(row, col, item)
    
    def _on_cell_double_clicked(self, row: int, col: int):
        """雙擊格子"""
        event_data = self.events_map.get((col, row))
        if event_data:
            # 編輯既有事件
            self.event_double_clicked.emit(col, row, event_data)
        else:
            # 建立新事件
            self.cell_double_clicked.emit(col, row)

    def _on_selection_changed(self):
        selected = self.selectedItems()
        if not selected:
            self.selection_changed.emit(None)
            return

        item = selected[0]
        row = self.row(item)
        col = self.column(item)
        event_data = self.events_map.get((col, row))
        self.selection_changed.emit(event_data)
    
    def _show_context_menu(self, pos):
        """右鍵選單"""
        item = self.itemAt(pos)
        if not item:
            return
        
        row = self.row(item)
        col = self.column(item)
        event_data = self.events_map.get((col, row))
        
        menu = QMenu(self)
        
        if event_data:
            edit_action = menu.addAction("編輯事件 (Edit Event)")
            delete_action = menu.addAction("刪除事件 (Delete Event)")
            menu.addSeparator()
        
        new_action = menu.addAction("新增事件 (New Event)")
        menu.addSeparator()
        refresh_action = menu.addAction("重新整理 (Refresh)")
        
        action = menu.exec(self.mapToGlobal(pos))
        
        if not action:
            return
        
        payload = {
            "day_index": col,
            "hour": row,
            "event_data": event_data,
        }
        
        if event_data and action.text().startswith("編輯"):
            self.event_context_requested.emit("edit", payload)
        elif event_data and action.text().startswith("刪除"):
            self.event_context_requested.emit("delete", payload)
        elif action.text().startswith("新增"):
            self.event_context_requested.emit("new", payload)
        elif action.text().startswith("重新整理"):
            self.event_context_requested.emit("refresh", payload)


class WeeklyPanel(QWidget):
    """Weekly Tab 主面板"""
    
    schedule_changed = Signal()  # 通知主視窗重新載入排程
    schedule_selected = Signal(object)  # schedule_id or None
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.schedules = []
        self.db_manager = None
        self._init_ui()
    
    def _init_ui(self):
        layout = QVBoxLayout(self)
        
        # 頂部工具列
        toolbar = QHBoxLayout()
        self.btn_new = QPushButton("新增事件 (New)")
        self.btn_refresh = QPushButton("重新整理 (Refresh)")
        self.btn_help = QPushButton("說明")
        
        toolbar.addWidget(QLabel("週間班表編輯器"))
        toolbar.addStretch()
        toolbar.addWidget(self.btn_new)
        toolbar.addWidget(self.btn_refresh)
        toolbar.addWidget(self.btn_help)
        
        layout.addLayout(toolbar)
        
        # 週間格子
        self.weekly_grid = WeeklyGridWidget(self)
        layout.addWidget(self.weekly_grid)
        
        # 底部說明
        info_label = QLabel(
            "提示：雙擊空白格子建立新的週間重複事件，雙擊事件可編輯。"
            "此 Tab 僅顯示 FREQ=WEEKLY 的排程。"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: gray; font-size: 10pt; padding: 5px;")
        layout.addWidget(info_label)
        
        # 訊號連接
        self.btn_new.clicked.connect(self._create_new_event)
        self.btn_refresh.clicked.connect(self.refresh)
        self.btn_help.clicked.connect(self._show_help)
        
        self.weekly_grid.cell_double_clicked.connect(self._on_cell_double_clicked)
        self.weekly_grid.event_double_clicked.connect(self._on_event_double_clicked)
        self.weekly_grid.event_context_requested.connect(self._on_context_action)
        self.weekly_grid.selection_changed.connect(self._on_selection_changed)
    
    def set_db_manager(self, db_manager):
        """設定資料庫管理器"""
        self.db_manager = db_manager
    
    def load_schedules(self, schedules: List[Dict[str, Any]]):
        """載入排程列表"""
        self.schedules = schedules
        self.weekly_grid.load_weekly_events(schedules)
    
    def refresh(self):
        """重新整理"""
        if self.db_manager:
            self.schedules = self.db_manager.get_all_schedules()
            self.weekly_grid.load_weekly_events(self.schedules)
    
    def _create_new_event(self):
        """建立新事件（預設週一 08:00）"""
        self._open_event_editor(day_index=1, hour=8, event_data=None)
    
    def _on_cell_double_clicked(self, day_index: int, hour: int):
        """雙擊空白格子 - 建立新事件"""
        self._open_event_editor(day_index, hour, None)
    
    def _on_event_double_clicked(self, day_index: int, hour: int, event_data: Dict[str, Any]):
        """雙擊事件 - 編輯事件"""
        self._open_event_editor(day_index, hour, event_data)
    
    def _on_context_action(self, action: str, payload: Dict[str, Any]):
        """右鍵選單動作"""
        day_index = payload.get("day_index", 0)
        hour = payload.get("hour", 8)
        event_data = payload.get("event_data")
        
        if action == "edit" and event_data:
            self._open_event_editor(day_index, hour, event_data)
        elif action == "delete" and event_data:
            self._delete_event(event_data)
        elif action == "new":
            self._open_event_editor(day_index, hour, None)
        elif action == "refresh":
            self.refresh()

    def _on_selection_changed(self, event_data: Optional[Dict[str, Any]]):
        schedule_id = event_data.get("schedule_id") if event_data else None
        self.schedule_selected.emit(schedule_id)
    
    def _open_event_editor(self, day_index: int, hour: int, event_data: Optional[Dict[str, Any]]):
        """開啟事件編輯器"""
        from ui.weekly_event_dialog import WeeklyEventDialog
        
        # 轉換星期索引到 BYDAY 格式
        day_names = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"]
        byday = day_names[day_index]
        
        initial_data = {
            "byday": byday,
            "hour": hour,
            "minute": 0,
            "duration_minutes": 60,
            "title": "",
            "target_value": "",
            "opc_url": "opc.tcp://localhost:4840",
            "node_id": "",
            "data_type": "auto",
        }
        
        if event_data:
            # 編輯模式
            schedule_id = event_data.get("schedule_id")
            if not schedule_id:
                return
            
            # 從排程列表找到完整資料
            schedule = next((s for s in self.schedules if s.get("id") == schedule_id), None)
            if not schedule:
                return
            
            initial_data.update({
                "schedule_id": schedule_id,
                "title": schedule.get("task_name", ""),
                "target_value": schedule.get("target_value", ""),
                "opc_url": schedule.get("opc_url", ""),
                "node_id": schedule.get("node_id", ""),
                "data_type": schedule.get("data_type", "auto"),
                "rrule_str": schedule.get("rrule_str", ""),
            })
            
            # 解析 RRULE 取得時間
            rrule_str = schedule.get("rrule_str", "")
            if "BYHOUR=" in rrule_str:
                hour_match = rrule_str.split("BYHOUR=")[1].split(";")[0]
                initial_data["hour"] = int(hour_match)
            if "BYMINUTE=" in rrule_str:
                minute_match = rrule_str.split("BYMINUTE=")[1].split(";")[0]
                initial_data["minute"] = int(minute_match)
            if "BYDAY=" in rrule_str:
                byday_match = rrule_str.split("BYDAY=")[1].split(";")[0]
                initial_data["byday"] = byday_match
        
        dialog = WeeklyEventDialog(self, initial_data)
        if dialog.exec() != QDialog.Accepted:
            return
        
        data = dialog.get_data()
        self._save_weekly_event(data)
    
    def _save_weekly_event(self, data: Dict[str, Any]):
        """儲存週間事件"""
        if not self.db_manager:
            return
        
        # 建立 RRULE (FREQ=WEEKLY;BYDAY=MO;BYHOUR=8;BYMINUTE=0)
        byday = data.get("byday", "MO")
        hour = data.get("hour", 8)
        minute = data.get("minute", 0)
        
        rrule_str = f"FREQ=WEEKLY;BYDAY={byday};BYHOUR={hour};BYMINUTE={minute}"
        
        schedule_id = data.get("schedule_id")
        
        if schedule_id:
            # 更新既有排程
            self.db_manager.update_schedule(
                schedule_id=schedule_id,
                task_name=data["title"],
                opc_url=data["opc_url"],
                node_id=data["node_id"],
                target_value=data["target_value"],
                data_type=data.get("data_type", "auto"),
                rrule_str=rrule_str,
            )
        else:
            # 新增排程
            self.db_manager.add_schedule(
                task_name=data["title"],
                opc_url=data["opc_url"],
                node_id=data["node_id"],
                target_value=data["target_value"],
                data_type=data.get("data_type", "auto"),
                rrule_str=rrule_str,
                is_enabled=1,
            )
        
        self.schedule_changed.emit()
        self.refresh()
    
    def _delete_event(self, event_data: Dict[str, Any]):
        """刪除事件"""
        from PySide6.QtWidgets import QMessageBox
        
        schedule_id = event_data.get("schedule_id")
        if not schedule_id or not self.db_manager:
            return
        
        title = event_data.get("title", "此事件")
        reply = QMessageBox.question(
            self,
            "確認刪除",
            f"確定要刪除「{title}」嗎？\n此操作將刪除整個週間重複系列。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.db_manager.delete_schedule(schedule_id)
            self.schedule_changed.emit()
            self.refresh()
    
    def _show_help(self):
        """顯示說明"""
        from PySide6.QtWidgets import QMessageBox
        
        QMessageBox.information(
            self,
            "Weekly Tab 說明",
            "此面板用於管理週間重複事件（FREQ=WEEKLY）。\n\n"
            "操作方式：\n"
            "• 雙擊空白格子建立新事件\n"
            "• 雙擊事件可編輯\n"
            "• 右鍵選單可刪除/編輯\n"
            "• 事件會在 Preview 視圖顯示\n\n"
            "注意：此處僅顯示 FREQ=WEEKLY 的排程，"
            "其他頻率（DAILY/MONTHLY）請在排程列表編輯。"
        )

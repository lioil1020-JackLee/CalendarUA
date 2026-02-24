"""
Exceptions Panel - 例外記錄管理面板
用於查看和管理 schedule exceptions (occurrence overrides and cancellations)
"""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QDialog, QFormLayout, QComboBox, QDateEdit, QLineEdit,
    QDateTimeEdit, QGroupBox, QRadioButton, QButtonGroup
)
from PySide6.QtCore import Qt, Signal, QDate, QDateTime
from PySide6.QtGui import QColor, QBrush
from datetime import datetime, date
from typing import List, Dict, Any, Optional


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
        self._init_ui()
    
    def _init_ui(self):
        layout = QVBoxLayout(self)
        
        # 頂部工具列
        toolbar = QHBoxLayout()
        self.btn_new = QPushButton("新增例外 (New Exception)")
        self.btn_edit = QPushButton("編輯 (Edit)")
        self.btn_delete = QPushButton("刪除 (Delete)")
        self.btn_refresh = QPushButton("重新整理 (Refresh)")
        
        toolbar.addWidget(QLabel("例外記錄管理"))
        toolbar.addStretch()
        toolbar.addWidget(self.btn_new)
        toolbar.addWidget(self.btn_edit)
        toolbar.addWidget(self.btn_delete)
        toolbar.addWidget(self.btn_refresh)
        
        layout.addLayout(toolbar)
        
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
        
        layout.addWidget(self.table)
        
        # 底部說明
        info_label = QLabel(
            "提示：此面板管理所有排程的例外記錄。"
            "「取消」會隱藏該次 occurrence；「覆寫」會替換時間/標題/數值。"
            "雙擊列可編輯。"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: gray; font-size: 10pt; padding: 5px;")
        layout.addWidget(info_label)
        
        # 訊號
        self.btn_new.clicked.connect(self._create_exception)
        self.btn_edit.clicked.connect(self._edit_selected_exception)
        self.btn_delete.clicked.connect(self._delete_selected_exception)
        self.btn_refresh.clicked.connect(self.refresh)
        self.table.doubleClicked.connect(self._edit_selected_exception)
    
    def set_db_manager(self, db_manager):
        """設定資料庫管理器"""
        self.db_manager = db_manager
    
    def load_data(self, schedules: List[Dict[str, Any]], exceptions: List[Dict[str, Any]]):
        """載入排程與例外資料"""
        self.schedules = schedules
        self.exceptions = exceptions
        self._populate_table()
    
    def refresh(self):
        """重新整理"""
        if self.db_manager:
            self.schedules = self.db_manager.get_all_schedules()
            self.exceptions = self.db_manager.get_all_schedule_exceptions()
            self._populate_table()
    
    def _populate_table(self):
        """填充表格"""
        self.table.setRowCount(0)
        
        # 建立 schedule_id -> task_name 映射
        schedule_map = {s.get("id"): s.get("task_name", "未知") for s in self.schedules}
        
        for exc in self.exceptions:
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

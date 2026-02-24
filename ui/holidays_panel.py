"""
Holidays Panel - 假日日曆管理面板
用於管理 holiday calendars 和 holiday entries
"""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QDialog, QFormLayout, QLineEdit, QDateEdit, QCheckBox,
    QTimeEdit, QGroupBox, QListWidget, QSplitter, QListWidgetItem
)
from PySide6.QtCore import Qt, Signal, QDate, QTime
from PySide6.QtGui import QColor, QBrush
from datetime import datetime, date
from typing import List, Dict, Any, Optional


class HolidayEntryDialog(QDialog):
    """假日條目編輯對話框"""
    
    def __init__(self, parent=None, initial_data: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.initial_data = initial_data or {}
        self.setWindowTitle("編輯假日 - Edit Holiday")
        self.resize(450, 350)
        self._init_ui()
        self._load_data()
    
    def _init_ui(self):
        layout = QVBoxLayout(self)
        
        # 基本設定
        group_basic = QGroupBox("假日資訊 (Holiday Info)")
        form_basic = QFormLayout()
        
        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText("例如：春節、國慶日...")
        form_basic.addRow("假日名稱 (Name):", self.edit_name)
        
        self.date_holiday = QDateEdit()
        self.date_holiday.setCalendarPopup(True)
        self.date_holiday.setDate(QDate.currentDate())
        self.date_holiday.setDisplayFormat("yyyy-MM-dd")
        form_basic.addRow("日期 (Date):", self.date_holiday)
        
        self.check_full_day = QCheckBox("全天 (All Day)")
        self.check_full_day.setChecked(True)
        form_basic.addRow("", self.check_full_day)
        
        group_basic.setLayout(form_basic)
        layout.addWidget(group_basic)
        
        # 時間設定（部分時段假日）
        self.group_time = QGroupBox("時段設定 (Time Range)")
        form_time = QFormLayout()
        
        self.time_start = QTimeEdit()
        self.time_start.setDisplayFormat("HH:mm")
        self.time_start.setTime(QTime(9, 0))
        form_time.addRow("開始時間 (Start):", self.time_start)
        
        self.time_end = QTimeEdit()
        self.time_end.setDisplayFormat("HH:mm")
        self.time_end.setTime(QTime(17, 0))
        form_time.addRow("結束時間 (End):", self.time_end)
        
        self.group_time.setLayout(form_time)
        self.group_time.setEnabled(False)
        layout.addWidget(self.group_time)
        
        # 按鈕
        button_layout = QHBoxLayout()
        self.btn_ok = QPushButton("確定 (OK)")
        self.btn_cancel = QPushButton("取消 (Cancel)")
        
        button_layout.addStretch()
        button_layout.addWidget(self.btn_ok)
        button_layout.addWidget(self.btn_cancel)
        
        layout.addLayout(button_layout)
        
        # 訊號
        self.check_full_day.toggled.connect(self._on_full_day_toggled)
        self.btn_ok.clicked.connect(self._on_ok)
        self.btn_cancel.clicked.connect(self.reject)
    
    def _on_full_day_toggled(self):
        """全天選項變更時啟用/禁用時段設定"""
        is_full_day = self.check_full_day.isChecked()
        self.group_time.setEnabled(not is_full_day)
    
    def _load_data(self):
        """載入初始資料"""
        if not self.initial_data:
            return
        
        self.edit_name.setText(self.initial_data.get("name", ""))
        
        holiday_date_str = self.initial_data.get("holiday_date")
        if holiday_date_str:
            try:
                dt = datetime.strptime(holiday_date_str, "%Y-%m-%d")
                self.date_holiday.setDate(QDate(dt.year, dt.month, dt.day))
            except:
                pass
        
        is_full_day = self.initial_data.get("is_full_day", 1)
        self.check_full_day.setChecked(bool(is_full_day))
        
        start_time_str = self.initial_data.get("start_time")
        if start_time_str:
            try:
                h, m = map(int, start_time_str.split(":"))
                self.time_start.setTime(QTime(h, m))
            except:
                pass
        
        end_time_str = self.initial_data.get("end_time")
        if end_time_str:
            try:
                h, m = map(int, end_time_str.split(":"))
                self.time_end.setTime(QTime(h, m))
            except:
                pass
    
    def _on_ok(self):
        """驗證並確定"""
        if not self.edit_name.text().strip():
            QMessageBox.warning(self, "驗證錯誤", "請輸入假日名稱")
            self.edit_name.setFocus()
            return
        
        if not self.check_full_day.isChecked():
            start = self.time_start.time()
            end = self.time_end.time()
            if end <= start:
                QMessageBox.warning(self, "驗證錯誤", "結束時間必須大於開始時間")
                return
        
        self.accept()
    
    def get_data(self) -> Dict[str, Any]:
        """取得編輯結果"""
        holiday_date = self.date_holiday.date().toPython()
        is_full_day = self.check_full_day.isChecked()
        
        result = {
            "entry_id": self.initial_data.get("id"),
            "name": self.edit_name.text().strip(),
            "holiday_date": holiday_date.strftime("%Y-%m-%d"),
            "is_full_day": 1 if is_full_day else 0,
        }
        
        if not is_full_day:
            start_time = self.time_start.time()
            end_time = self.time_end.time()
            result["start_time"] = start_time.toString("HH:mm:ss")
            result["end_time"] = end_time.toString("HH:mm:ss")
        else:
            result["start_time"] = None
            result["end_time"] = None
        
        return result


class HolidaysPanel(QWidget):
    """Holidays 面板 - 管理假日日曆"""
    
    holiday_changed = Signal()  # 通知主視窗重新載入
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db_manager = None
        self.calendars = []
        self.current_calendar_id = None
        self._init_ui()
    
    def _init_ui(self):
        layout = QVBoxLayout(self)
        
        # 頂部工具列
        toolbar = QHBoxLayout()
        self.btn_new_calendar = QPushButton("新增日曆 (New Calendar)")
        self.btn_delete_calendar = QPushButton("刪除日曆 (Delete Calendar)")
        self.btn_new_holiday = QPushButton("新增假日 (New Holiday)")
        self.btn_refresh = QPushButton("重新整理 (Refresh)")
        
        toolbar.addWidget(QLabel("假日日曆管理"))
        toolbar.addStretch()
        toolbar.addWidget(self.btn_new_calendar)
        toolbar.addWidget(self.btn_delete_calendar)
        toolbar.addWidget(self.btn_new_holiday)
        toolbar.addWidget(self.btn_refresh)
        
        layout.addLayout(toolbar)
        
        # 分割版面：左側日曆列表，右側假日條目
        splitter = QSplitter(Qt.Horizontal)
        
        # 左側：日曆列表
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        left_layout.addWidget(QLabel("假日日曆列表 (Holiday Calendars)"))
        self.calendar_list = QListWidget()
        self.calendar_list.currentItemChanged.connect(self._on_calendar_selected)
        left_layout.addWidget(self.calendar_list)
        
        splitter.addWidget(left_widget)
        
        # 右側：假日條目
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        entry_toolbar = QHBoxLayout()
        self.label_current_calendar = QLabel("請選擇日曆")
        self.btn_edit_holiday = QPushButton("編輯假日 (Edit)")
        self.btn_delete_holiday = QPushButton("刪除假日 (Delete)")
        
        entry_toolbar.addWidget(self.label_current_calendar)
        entry_toolbar.addStretch()
        entry_toolbar.addWidget(self.btn_edit_holiday)
        entry_toolbar.addWidget(self.btn_delete_holiday)
        
        right_layout.addLayout(entry_toolbar)
        
        self.holiday_table = QTableWidget()
        self.holiday_table.setColumnCount(5)
        self.holiday_table.setHorizontalHeaderLabels([
            "ID", "日期", "名稱", "全天/時段", "時間範圍"
        ])
        self.holiday_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.holiday_table.setSelectionMode(QTableWidget.SingleSelection)
        self.holiday_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.holiday_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.holiday_table.horizontalHeader().setStretchLastSection(True)
        self.holiday_table.doubleClicked.connect(self._edit_selected_holiday)
        
        right_layout.addWidget(self.holiday_table)
        
        splitter.addWidget(right_widget)
        
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        
        layout.addWidget(splitter)
        
        # 底部說明
        info_label = QLabel(
            "提示：左側管理假日日曆，右側管理各日曆的假日條目。"
            "假日可設為全天或指定時段，用於覆寫排程顯示。"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: gray; font-size: 10pt; padding: 5px;")
        layout.addWidget(info_label)
        
        # 訊號
        self.btn_new_calendar.clicked.connect(self._create_calendar)
        self.btn_delete_calendar.clicked.connect(self._delete_calendar)
        self.btn_new_holiday.clicked.connect(self._create_holiday)
        self.btn_edit_holiday.clicked.connect(self._edit_selected_holiday)
        self.btn_delete_holiday.clicked.connect(self._delete_selected_holiday)
        self.btn_refresh.clicked.connect(self.refresh)
    
    def set_db_manager(self, db_manager):
        """設定資料庫管理器"""
        self.db_manager = db_manager
    
    def refresh(self):
        """重新整理"""
        if not self.db_manager:
            return
        
        self.calendars = self.db_manager.get_all_holiday_calendars()
        self._populate_calendar_list()
        
        # 重新載入當前日曆的假日
        if self.current_calendar_id:
            self._load_holiday_entries(self.current_calendar_id)
    
    def _populate_calendar_list(self):
        """填充日曆列表"""
        self.calendar_list.clear()
        
        for cal in self.calendars:
            cal_id = cal.get("id")
            name = cal.get("name", "未命名")
            is_default = cal.get("is_default", 0)
            
            display = f"★ {name}" if is_default else name
            
            item = QListWidgetItem(display)
            item.setData(Qt.UserRole, cal_id)
            
            if is_default:
                item.setBackground(QBrush(QColor(255, 255, 200)))  # 淡黃色
            
            self.calendar_list.addItem(item)
        
        # 自動選擇第一個
        if self.calendar_list.count() > 0:
            self.calendar_list.setCurrentRow(0)
    
    def _on_calendar_selected(self, current, previous):
        """日曆選擇變更"""
        if not current:
            self.current_calendar_id = None
            self.label_current_calendar.setText("請選擇日曆")
            self.holiday_table.setRowCount(0)
            return
        
        calendar_id = current.data(Qt.UserRole)
        self.current_calendar_id = calendar_id
        
        calendar_name = current.text().replace("★ ", "")
        self.label_current_calendar.setText(f"當前日曆：{calendar_name}")
        
        self._load_holiday_entries(calendar_id)
    
    def _load_holiday_entries(self, calendar_id: int):
        """載入假日條目"""
        if not self.db_manager:
            return
        
        entries = self.db_manager.get_holiday_entries_by_calendar(calendar_id)
        
        self.holiday_table.setRowCount(0)
        
        for entry in entries:
            row = self.holiday_table.rowCount()
            self.holiday_table.insertRow(row)
            
            # ID
            item_id = QTableWidgetItem(str(entry.get("id", "")))
            item_id.setData(Qt.UserRole, entry)
            self.holiday_table.setItem(row, 0, item_id)
            
            # 日期
            self.holiday_table.setItem(row, 1, QTableWidgetItem(entry.get("holiday_date", "")))
            
            # 名稱
            self.holiday_table.setItem(row, 2, QTableWidgetItem(entry.get("name", "")))
            
            # 全天/時段
            is_full_day = entry.get("is_full_day", 1)
            type_text = "全天" if is_full_day else "時段"
            self.holiday_table.setItem(row, 3, QTableWidgetItem(type_text))
            
            # 時間範圍
            if is_full_day:
                time_text = "-"
            else:
                start_time = entry.get("start_time", "")
                end_time = entry.get("end_time", "")
                time_text = f"{start_time} ~ {end_time}" if start_time and end_time else "-"
            self.holiday_table.setItem(row, 4, QTableWidgetItem(time_text))
        
        self.holiday_table.resizeColumnsToContents()
    
    def _create_calendar(self):
        """建立新日曆"""
        from PySide6.QtWidgets import QInputDialog
        
        name, ok = QInputDialog.getText(self, "新增日曆", "日曆名稱:")
        if not ok or not name.strip():
            return
        
        if not self.db_manager:
            return
        
        calendar_id = self.db_manager.add_holiday_calendar(name.strip())
        if calendar_id:
            self.holiday_changed.emit()
            self.refresh()
    
    def _delete_calendar(self):
        """刪除選中的日曆"""
        current_item = self.calendar_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "提示", "請先選擇一個日曆")
            return
        
        calendar_id = current_item.data(Qt.UserRole)
        calendar_name = current_item.text().replace("★ ", "")
        
        reply = QMessageBox.question(
            self,
            "確認刪除",
            f"確定要刪除日曆「{calendar_name}」嗎？\n此操作將同時刪除該日曆的所有假日條目。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes and self.db_manager:
            self.db_manager.delete_holiday_calendar(calendar_id)
            self.holiday_changed.emit()
            self.refresh()
    
    def _create_holiday(self):
        """建立新假日"""
        if not self.current_calendar_id:
            QMessageBox.warning(self, "提示", "請先選擇一個日曆")
            return
        
        dialog = HolidayEntryDialog(self, None)
        if dialog.exec() != QDialog.Accepted:
            return
        
        data = dialog.get_data()
        self._save_holiday_entry(data)
    
    def _edit_selected_holiday(self):
        """編輯選中的假日"""
        selected_items = self.holiday_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "提示", "請先選擇一筆假日記錄")
            return
        
        row = self.holiday_table.currentRow()
        entry_data = self.holiday_table.item(row, 0).data(Qt.UserRole)
        
        dialog = HolidayEntryDialog(self, entry_data)
        if dialog.exec() != QDialog.Accepted:
            return
        
        data = dialog.get_data()
        self._save_holiday_entry(data)
    
    def _delete_selected_holiday(self):
        """刪除選中的假日"""
        selected_items = self.holiday_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "提示", "請先選擇一筆假日記錄")
            return
        
        row = self.holiday_table.currentRow()
        entry_data = self.holiday_table.item(row, 0).data(Qt.UserRole)
        entry_id = entry_data.get("id")
        entry_name = entry_data.get("name")
        
        reply = QMessageBox.question(
            self,
            "確認刪除",
            f"確定要刪除假日「{entry_name}」嗎？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes and self.db_manager:
            self.db_manager.delete_holiday_entry(entry_id)
            self.holiday_changed.emit()
            self._load_holiday_entries(self.current_calendar_id)
    
    def _save_holiday_entry(self, data: Dict[str, Any]):
        """儲存假日條目"""
        if not self.db_manager or not self.current_calendar_id:
            return
        
        entry_id = data.get("entry_id")
        
        if entry_id:
            # 更新
            self.db_manager.update_holiday_entry(
                entry_id=entry_id,
                holiday_date=data["holiday_date"],
                name=data["name"],
                is_full_day=data["is_full_day"],
                start_time=data.get("start_time"),
                end_time=data.get("end_time"),
            )
        else:
            # 新增
            self.db_manager.add_holiday_entry(
                calendar_id=self.current_calendar_id,
                holiday_date=data["holiday_date"],
                name=data["name"],
                is_full_day=data["is_full_day"],
                start_time=data.get("start_time"),
                end_time=data.get("end_time"),
            )
        
        self.holiday_changed.emit()
        self._load_holiday_entries(self.current_calendar_id)

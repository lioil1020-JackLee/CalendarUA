"""
Weekly Event Dialog - 週間事件編輯對話框
用於建立/編輯 FREQ=WEEKLY 的重複事件
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QComboBox, QSpinBox, QTimeEdit, QPushButton,
    QLabel, QGroupBox, QMessageBox
)
from PySide6.QtCore import Qt, QTime
from typing import Dict, Any, Optional


class WeeklyEventDialog(QDialog):
    """週間事件編輯對話框"""
    
    def __init__(self, parent=None, initial_data: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.initial_data = initial_data or {}
        self.setWindowTitle("週間事件編輯 - Weekly Event Editor")
        self.resize(500, 450)
        self._init_ui()
        self._load_data()
    
    def _init_ui(self):
        layout = QVBoxLayout(self)
        
        # 基本資訊
        group_basic = QGroupBox("基本資訊 (Basic Info)")
        form_basic = QFormLayout()
        
        self.edit_title = QLineEdit()
        self.edit_title.setPlaceholderText("例如：開啟設備、備份資料...")
        form_basic.addRow("事件名稱 (Subject):", self.edit_title)
        
        self.edit_target_value = QLineEdit()
        self.edit_target_value.setPlaceholderText("例如：1, true, 開啟...")
        form_basic.addRow("目標數值 (ValueSet Value):", self.edit_target_value)
        
        group_basic.setLayout(form_basic)
        layout.addWidget(group_basic)
        
        # 時間設定
        group_time = QGroupBox("時間設定 (Time Settings)")
        form_time = QFormLayout()
        
        self.combo_weekday = QComboBox()
        weekdays = [
            ("星期日 (Sunday)", "SU"),
            ("星期一 (Monday)", "MO"),
            ("星期二 (Tuesday)", "TU"),
            ("星期三 (Wednesday)", "WE"),
            ("星期四 (Thursday)", "TH"),
            ("星期五 (Friday)", "FR"),
            ("星期六 (Saturday)", "SA"),
        ]
        for display, value in weekdays:
            self.combo_weekday.addItem(display, value)
        form_time.addRow("重複星期 (Day of Week):", self.combo_weekday)
        
        self.time_edit = QTimeEdit()
        self.time_edit.setDisplayFormat("HH:mm")
        self.time_edit.setTime(QTime(8, 0))
        form_time.addRow("觸發時間 (Start Time):", self.time_edit)
        
        self.spin_duration = QSpinBox()
        self.spin_duration.setRange(1, 1440)  # 1分鐘到24小時
        self.spin_duration.setValue(60)
        self.spin_duration.setSuffix(" 分鐘")
        form_time.addRow("持續時間 (Duration):", self.spin_duration)
        
        group_time.setLayout(form_time)
        layout.addWidget(group_time)
        
        # OPC 設定
        group_opc = QGroupBox("OPC UA 設定 (OPC Settings)")
        form_opc = QFormLayout()
        
        self.edit_opc_url = QLineEdit()
        self.edit_opc_url.setPlaceholderText("opc.tcp://localhost:4840")
        form_opc.addRow("OPC UA 伺服器:", self.edit_opc_url)
        
        self.edit_node_id = QLineEdit()
        self.edit_node_id.setPlaceholderText('ns=2;s=Channel1.Device1.Tag1')
        form_opc.addRow("節點 ID (Node ID):", self.edit_node_id)
        
        self.combo_data_type = QComboBox()
        data_types = ["auto", "Boolean", "Int16", "Int32", "Float", "Double", "String"]
        self.combo_data_type.addItems(data_types)
        form_opc.addRow("資料型態 (Data Type):", self.combo_data_type)
        
        group_opc.setLayout(form_opc)
        layout.addWidget(group_opc)
        
        # 按鈕
        button_layout = QHBoxLayout()
        self.btn_ok = QPushButton("確定 (OK)")
        self.btn_cancel = QPushButton("取消 (Cancel)")
        
        button_layout.addStretch()
        button_layout.addWidget(self.btn_ok)
        button_layout.addWidget(self.btn_cancel)
        
        layout.addLayout(button_layout)
        
        # 訊號
        self.btn_ok.clicked.connect(self._on_ok)
        self.btn_cancel.clicked.connect(self.reject)
    
    def _load_data(self):
        """載入初始資料"""
        if not self.initial_data:
            return
        
        # 基本資訊
        self.edit_title.setText(self.initial_data.get("title", ""))
        self.edit_target_value.setText(self.initial_data.get("target_value", ""))
        
        # 時間設定
        byday = self.initial_data.get("byday", "MO")
        for i in range(self.combo_weekday.count()):
            if self.combo_weekday.itemData(i) == byday:
                self.combo_weekday.setCurrentIndex(i)
                break
        
        hour = self.initial_data.get("hour", 8)
        minute = self.initial_data.get("minute", 0)
        self.time_edit.setTime(QTime(hour, minute))
        
        duration = self.initial_data.get("duration_minutes", 60)
        self.spin_duration.setValue(duration)
        
        # OPC 設定
        self.edit_opc_url.setText(self.initial_data.get("opc_url", "opc.tcp://localhost:4840"))
        self.edit_node_id.setText(self.initial_data.get("node_id", ""))
        
        data_type = self.initial_data.get("data_type", "auto")
        idx = self.combo_data_type.findText(data_type)
        if idx >= 0:
            self.combo_data_type.setCurrentIndex(idx)
    
    def _on_ok(self):
        """驗證並確定"""
        # 驗證必填欄位
        if not self.edit_title.text().strip():
            QMessageBox.warning(self, "驗證錯誤", "請輸入事件名稱")
            self.edit_title.setFocus()
            return
        
        if not self.edit_node_id.text().strip():
            QMessageBox.warning(self, "驗證錯誤", "請輸入 OPC UA 節點 ID")
            self.edit_node_id.setFocus()
            return
        
        self.accept()
    
    def get_data(self) -> Dict[str, Any]:
        """取得編輯結果"""
        time_value = self.time_edit.time()
        
        return {
            "schedule_id": self.initial_data.get("schedule_id"),
            "title": self.edit_title.text().strip(),
            "target_value": self.edit_target_value.text().strip(),
            "byday": self.combo_weekday.currentData(),
            "hour": time_value.hour(),
            "minute": time_value.minute(),
            "duration_minutes": self.spin_duration.value(),
            "opc_url": self.edit_opc_url.text().strip(),
            "node_id": self.edit_node_id.text().strip(),
            "data_type": self.combo_data_type.currentText(),
        }

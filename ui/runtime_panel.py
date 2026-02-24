"""
Runtime Panel - 運行時狀態與覆寫面板
用於手動覆寫輸出值、查看當前狀態和下一事件
"""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QPushButton,
    QLabel, QLineEdit, QComboBox, QGroupBox, QSpinBox, QRadioButton,
    QButtonGroup, QMessageBox, QTextEdit
)
from PySide6.QtCore import Qt, Signal, QTimer
from datetime import datetime, timedelta
from typing import Dict, Any, Optional


class RuntimePanel(QWidget):
    """Runtime Panel - 運行時狀態與覆寫"""
    
    override_changed = Signal()  # 覆寫變更訊號
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db_manager = None
        self.schedules = []
        self.current_override = None
        self._init_ui()
        
        # 定時更新狀態
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self._update_status_display)
        self.update_timer.start(1000)  # 每秒更新
    
    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # 頂部標題
        title_layout = QHBoxLayout()
        title_layout.addWidget(QLabel("<b>Runtime Override & Status</b>"))
        title_layout.addStretch()
        
        layout.addLayout(title_layout)
        
        # Override 區域
        group_override = QGroupBox("手動覆寫 (Manual Override)")
        override_layout = QVBoxLayout()
        
        # Override 選項
        form_override = QFormLayout()
        
        self.edit_override_value = QLineEdit()
        self.edit_override_value.setPlaceholderText("輸入覆寫值，例如：100, true, 開啟...")
        form_override.addRow("覆寫值 (Override Value):", self.edit_override_value)
        
        override_layout.addLayout(form_override)
        
        # Temporary Override 有效期
        temp_layout = QHBoxLayout()
        temp_layout.addWidget(QLabel("有效期 (Duration):"))
        
        self.duration_group = QButtonGroup(self)
        durations = [
            ("永久 (Permanent)", 0),
            ("30 秒", 30),
            ("1 分鐘", 60),
            ("5 分鐘", 300),
            ("1 小時", 3600),
            ("1 天", 86400),
            ("1 週", 604800),
        ]
        
        for i, (text, seconds) in enumerate(durations):
            radio = QRadioButton(text)
            radio.setProperty("duration_seconds", seconds)
            self.duration_group.addButton(radio, i)
            temp_layout.addWidget(radio)
            if i == 0:
                radio.setChecked(True)
        
        temp_layout.addStretch()
        override_layout.addLayout(temp_layout)
        
        # Override 按鈕
        button_layout = QHBoxLayout()
        self.btn_apply_override = QPushButton("套用覆寫 (Apply Override)")
        self.btn_apply_override.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 8px;")
        self.btn_clear_override = QPushButton("清除覆寫 (Clear Override)")
        self.btn_clear_override.setStyleSheet("background-color: #f44336; color: white; font-weight: bold; padding: 8px;")
        
        button_layout.addWidget(self.btn_apply_override)
        button_layout.addWidget(self.btn_clear_override)
        button_layout.addStretch()
        
        override_layout.addLayout(button_layout)
        
        group_override.setLayout(override_layout)
        layout.addWidget(group_override)
        
        # Current Status 區域
        group_status = QGroupBox("目前狀態 (Current Status)")
        form_status = QFormLayout()
        
        self.label_current_value = QLabel("-")
        self.label_current_value.setStyleSheet("font-weight: bold; color: blue; font-size: 14pt;")
        form_status.addRow("當前值 (Current Value):", self.label_current_value)
        
        self.label_current_subject = QLabel("-")
        form_status.addRow("事件主題 (Subject):", self.label_current_subject)
        
        self.label_current_type = QLabel("-")
        form_status.addRow("類型 (Type):", self.label_current_type)
        
        self.label_busy_period = QLabel("-")
        form_status.addRow("忙碌期間 (Busy Period):", self.label_busy_period)
        
        self.label_override_value = QLabel("-")
        self.label_override_value.setStyleSheet("color: red; font-weight: bold;")
        form_status.addRow("覆寫值 (Override Value):", self.label_override_value)
        
        self.label_override_until = QLabel("-")
        self.label_override_until.setStyleSheet("color: orange;")
        form_status.addRow("覆寫有效至 (Override Until):", self.label_override_until)
        
        group_status.setLayout(form_status)
        layout.addWidget(group_status)
        
        # Next Event 區域
        group_next = QGroupBox("下一事件 (Next Event)")
        form_next = QFormLayout()
        
        self.label_next_value = QLabel("-")
        self.label_next_value.setStyleSheet("font-weight: bold; color: green;")
        form_next.addRow("下一事件值 (Next Event Value):", self.label_next_value)
        
        self.label_next_subject = QLabel("-")
        form_next.addRow("事件主題 (Subject):", self.label_next_subject)
        
        self.label_next_date = QLabel("-")
        self.label_next_date.setStyleSheet("font-weight: bold; color: purple;")
        form_next.addRow("發生時間 (Date/Time):", self.label_next_date)
        
        self.label_countdown = QLabel("-")
        self.label_countdown.setStyleSheet("color: gray; font-style: italic;")
        form_next.addRow("倒數計時 (Countdown):", self.label_countdown)
        
        group_next.setLayout(form_next)
        layout.addWidget(group_next)
        
        # 底部說明
        layout.addStretch()
        
        info_label = QLabel(
            "提示：Runtime Override 優先於所有排程規則。\n"
            "「永久」覆寫會持續到手動清除；「臨時」覆寫到期後自動恢復排程值。"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: gray; font-size: 9pt; padding: 5px;")
        layout.addWidget(info_label)
        
        # 訊號連接
        self.btn_apply_override.clicked.connect(self._apply_override)
        self.btn_clear_override.clicked.connect(self._clear_override)
    
    def set_db_manager(self, db_manager):
        """設定資料庫管理器"""
        self.db_manager = db_manager
        self._load_current_override()
        self._update_status_display()
    
    def load_schedules(self, schedules):
        """載入排程列表"""
        self.schedules = schedules
        self._update_status_display()
    
    def _load_current_override(self):
        """載入當前覆寫狀態"""
        if not self.db_manager:
            return
        
        override_data = self.db_manager.get_runtime_override()
        if override_data:
            self.current_override = override_data
            self.edit_override_value.setText(override_data.get("override_value", ""))
    
    def _apply_override(self):
        """套用覆寫"""
        override_value = self.edit_override_value.text().strip()
        if not override_value:
            QMessageBox.warning(self, "輸入錯誤", "請輸入覆寫值")
            return
        
        # 取得選中的有效期
        selected_button = self.duration_group.checkedButton()
        duration_seconds = selected_button.property("duration_seconds") if selected_button else 0
        
        override_until = None
        if duration_seconds > 0:
            override_until = (datetime.now() + timedelta(seconds=duration_seconds)).isoformat()
        
        if not self.db_manager:
            QMessageBox.warning(self, "錯誤", "資料庫未連線")
            return
        
        success = self.db_manager.set_runtime_override(override_value, override_until)
        
        if success:
            self.current_override = {
                "override_value": override_value,
                "override_until": override_until,
            }
            self._update_status_display()
            self.override_changed.emit()
            
            duration_text = selected_button.text() if selected_button else "永久"
            QMessageBox.information(self, "套用成功", f"已套用覆寫值：{override_value}\n有效期：{duration_text}")
        else:
            QMessageBox.warning(self, "套用失敗", "無法套用覆寫")
    
    def _clear_override(self):
        """清除覆寫"""
        if not self.current_override:
            QMessageBox.information(self, "提示", "目前沒有覆寫")
            return
        
        reply = QMessageBox.question(
            self,
            "確認清除",
            "確定要清除目前的覆寫嗎？\n系統將恢復使用排程值。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes and self.db_manager:
            success = self.db_manager.clear_runtime_override()
            if success:
                self.current_override = None
                self.edit_override_value.clear()
                self._update_status_display()
                self.override_changed.emit()
                QMessageBox.information(self, "清除成功", "已清除覆寫，恢復排程值")
    
    def _update_status_display(self):
        """更新狀態顯示"""
        # 更新 Current Status
        if self.current_override:
            override_value = self.current_override.get("override_value", "-")
            override_until = self.current_override.get("override_until")
            
            # 檢查是否過期
            if override_until:
                try:
                    until_dt = datetime.fromisoformat(override_until)
                    if datetime.now() > until_dt:
                        # 已過期，自動清除
                        if self.db_manager:
                            self.db_manager.clear_runtime_override()
                        self.current_override = None
                        self.label_current_value.setText("-")
                        self.label_override_value.setText("-")
                        self.label_override_until.setText("-")
                        return
                    else:
                        self.label_override_until.setText(override_until)
                except:
                    pass
            else:
                self.label_override_until.setText("永久 (Permanent)")
            
            self.label_current_value.setText(override_value)
            self.label_override_value.setText(override_value)
            self.label_current_subject.setText("手動覆寫 (Manual Override)")
            self.label_current_type.setText("Runtime Override")
            self.label_busy_period.setText("-")
        else:
            self.label_current_value.setText("-")
            self.label_override_value.setText("-")
            self.label_override_until.setText("-")
            self.label_current_subject.setText("-")
            self.label_current_type.setText("-")
            self.label_busy_period.setText("-")
        
        # 更新 Next Event
        next_event = self._find_next_event()
        if next_event:
            self.label_next_value.setText(next_event.get("target_value", "-"))
            self.label_next_subject.setText(next_event.get("task_name", "-"))
            next_time = next_event.get("next_execution_time")
            if next_time:
                self.label_next_date.setText(str(next_time))
                
                # 計算倒數
                try:
                    next_dt = datetime.fromisoformat(str(next_time))
                    delta = next_dt - datetime.now()
                    if delta.total_seconds() > 0:
                        days = delta.days
                        hours, remainder = divmod(delta.seconds, 3600)
                        minutes, seconds = divmod(remainder, 60)
                        
                        if days > 0:
                            countdown = f"{days}天 {hours}小時 {minutes}分 {seconds}秒"
                        elif hours > 0:
                            countdown = f"{hours}小時 {minutes}分 {seconds}秒"
                        elif minutes > 0:
                            countdown = f"{minutes}分 {seconds}秒"
                        else:
                            countdown = f"{seconds}秒"
                        
                        self.label_countdown.setText(f"還剩 {countdown}")
                    else:
                        self.label_countdown.setText("即將執行")
                except:
                    self.label_countdown.setText("-")
            else:
                self.label_next_date.setText("-")
                self.label_countdown.setText("-")
        else:
            self.label_next_value.setText("-")
            self.label_next_subject.setText("-")
            self.label_next_date.setText("-")
            self.label_countdown.setText("-")
    
    def _find_next_event(self) -> Optional[Dict[str, Any]]:
        """找出下一個要執行的事件"""
        if not self.schedules:
            return None
        
        # 找出最近的 next_execution_time
        from core.rrule_parser import RRuleParser
        parser = RRuleParser()
        
        next_event = None
        nearest_time = None
        
        for schedule in self.schedules:
            if not schedule.get("is_enabled"):
                continue
            
            rrule_str = schedule.get("rrule_str", "")
            if not rrule_str:
                continue
            
            try:
                next_time = parser.get_next_occurrence(rrule_str)
                if next_time:
                    if nearest_time is None or next_time < nearest_time:
                        nearest_time = next_time
                        next_event = {
                            "task_name": schedule.get("task_name", ""),
                            "target_value": schedule.get("target_value", ""),
                            "next_execution_time": next_time.isoformat(),
                        }
            except:
                continue
        
        return next_event

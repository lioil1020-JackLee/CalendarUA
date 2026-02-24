"""
General Panel - 全局設定面板
用於配置 Profile 全局參數（Name, Description, Enable, Scan Rate 等）
"""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QPushButton,
    QLabel, QLineEdit, QSpinBox, QCheckBox, QTextEdit, QGroupBox,
    QDateTimeEdit, QComboBox, QMessageBox
)
from PySide6.QtCore import Qt, Signal, QDateTime
from datetime import datetime
from typing import Dict, Any, Optional


class GeneralPanel(QWidget):
    """General Panel - Profile 全局設定"""
    
    settings_changed = Signal()  # 設定變更訊號
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db_manager = None
        self.current_settings = {}
        self._init_ui()
    
    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # 頂部工具列
        toolbar = QHBoxLayout()
        toolbar.addWidget(QLabel("全局設定 (General Settings)"))
        toolbar.addStretch()
        
        self.btn_save = QPushButton("儲存設定 (Save)")
        self.btn_reset = QPushButton("重設 (Reset)")
        
        toolbar.addWidget(self.btn_save)
        toolbar.addWidget(self.btn_reset)
        
        layout.addLayout(toolbar)
        
        # Profile 基本資訊
        group_profile = QGroupBox("Profile 資訊")
        form_profile = QFormLayout()
        
        self.edit_profile_name = QLineEdit()
        self.edit_profile_name.setPlaceholderText("預設 Profile")
        form_profile.addRow("Profile 名稱 (Name):", self.edit_profile_name)
        
        self.edit_description = QTextEdit()
        self.edit_description.setMaximumHeight(80)
        self.edit_description.setPlaceholderText("描述此 Profile 的用途...")
        form_profile.addRow("描述 (Description):", self.edit_description)
        
        group_profile.setLayout(form_profile)
        layout.addWidget(group_profile)
        
        # 排程控制
        group_schedule = QGroupBox("排程控制 (Schedule Control)")
        form_schedule = QFormLayout()
        
        self.check_enable_schedule = QCheckBox("啟用排程系統")
        self.check_enable_schedule.setChecked(True)
        form_schedule.addRow("Enable Schedule:", self.check_enable_schedule)
        
        self.spin_scan_rate = QSpinBox()
        self.spin_scan_rate.setRange(1, 3600)
        self.spin_scan_rate.setValue(1)
        self.spin_scan_rate.setSuffix(" 秒")
        form_schedule.addRow("掃描間隔 (Scan Rate):", self.spin_scan_rate)
        
        self.spin_refresh_rate = QSpinBox()
        self.spin_refresh_rate.setRange(1, 300)
        self.spin_refresh_rate.setValue(5)
        self.spin_refresh_rate.setSuffix(" 秒")
        form_schedule.addRow("更新頻率 (Refresh Rate):", self.spin_refresh_rate)
        
        group_schedule.setLayout(form_schedule)
        layout.addWidget(group_schedule)
        
        # 有效期間
        group_active = QGroupBox("有效期間 (Active Period)")
        form_active = QFormLayout()
        
        self.check_use_active_period = QCheckBox("啟用有效期間限制")
        self.check_use_active_period.setChecked(False)
        form_active.addRow("", self.check_use_active_period)
        
        self.datetime_active_from = QDateTimeEdit()
        self.datetime_active_from.setCalendarPopup(True)
        self.datetime_active_from.setDisplayFormat("yyyy-MM-dd HH:mm")
        self.datetime_active_from.setDateTime(QDateTime.currentDateTime())
        self.datetime_active_from.setEnabled(False)
        form_active.addRow("開始時間 (Active From):", self.datetime_active_from)
        
        self.datetime_active_to = QDateTimeEdit()
        self.datetime_active_to.setCalendarPopup(True)
        self.datetime_active_to.setDisplayFormat("yyyy-MM-dd HH:mm")
        self.datetime_active_to.setDateTime(QDateTime.currentDateTime().addMonths(12))
        self.datetime_active_to.setEnabled(False)
        form_active.addRow("結束時間 (Active To):", self.datetime_active_to)
        
        group_active.setLayout(form_active)
        layout.addWidget(group_active)
        
        # 輸出設定
        group_output = QGroupBox("輸出設定 (Output Settings)")
        form_output = QFormLayout()
        
        self.combo_output_type = QComboBox()
        self.combo_output_type.addItems(["OPC UA Write", "Database Log", "File Export", "HTTP POST"])
        form_output.addRow("輸出類型 (Output Type):", self.combo_output_type)
        
        self.check_refresh_output = QCheckBox("重新整理時刷新輸出")
        self.check_refresh_output.setChecked(True)
        form_output.addRow("Refresh Output:", self.check_refresh_output)
        
        self.check_generate_events = QCheckBox("產生事件記錄")
        self.check_generate_events.setChecked(True)
        form_output.addRow("Generate Events:", self.check_generate_events)
        
        group_output.setLayout(form_output)
        layout.addWidget(group_output)
        
        # 底部空白
        layout.addStretch()
        
        # 狀態標籤
        self.label_status = QLabel("狀態：未儲存變更")
        self.label_status.setStyleSheet("color: gray; font-size: 9pt;")
        layout.addWidget(self.label_status)
        
        # 訊號連接
        self.btn_save.clicked.connect(self._save_settings)
        self.btn_reset.clicked.connect(self._reset_settings)
        self.check_use_active_period.toggled.connect(self._on_active_period_toggled)
        
        # 監聽變更
        self.edit_profile_name.textChanged.connect(self._on_settings_modified)
        self.edit_description.textChanged.connect(self._on_settings_modified)
        self.check_enable_schedule.toggled.connect(self._on_settings_modified)
        self.spin_scan_rate.valueChanged.connect(self._on_settings_modified)
        self.spin_refresh_rate.valueChanged.connect(self._on_settings_modified)
    
    def _on_active_period_toggled(self):
        """有效期間選項變更"""
        enabled = self.check_use_active_period.isChecked()
        self.datetime_active_from.setEnabled(enabled)
        self.datetime_active_to.setEnabled(enabled)
        self._on_settings_modified()
    
    def _on_settings_modified(self):
        """設定已變更（未儲存）"""
        self.label_status.setText("狀態：有未儲存的變更")
        self.label_status.setStyleSheet("color: orange; font-size: 9pt;")
    
    def set_db_manager(self, db_manager):
        """設定資料庫管理器"""
        self.db_manager = db_manager
        self._load_settings()
    
    def _load_settings(self):
        """從資料庫載入設定"""
        if not self.db_manager:
            return
        
        settings = self.db_manager.get_general_settings()
        if settings:
            self.current_settings = settings
            
            self.edit_profile_name.setText(settings.get("profile_name", "預設 Profile"))
            self.edit_description.setPlainText(settings.get("description", ""))
            
            self.check_enable_schedule.setChecked(bool(settings.get("enable_schedule", 1)))
            self.spin_scan_rate.setValue(settings.get("scan_rate", 1))
            self.spin_refresh_rate.setValue(settings.get("refresh_rate", 5))
            
            use_active_period = bool(settings.get("use_active_period", 0))
            self.check_use_active_period.setChecked(use_active_period)
            
            active_from_str = settings.get("active_from")
            if active_from_str:
                try:
                    dt = datetime.fromisoformat(active_from_str)
                    self.datetime_active_from.setDateTime(QDateTime(dt.year, dt.month, dt.day, dt.hour, dt.minute))
                except:
                    pass
            
            active_to_str = settings.get("active_to")
            if active_to_str:
                try:
                    dt = datetime.fromisoformat(active_to_str)
                    self.datetime_active_to.setDateTime(QDateTime(dt.year, dt.month, dt.day, dt.hour, dt.minute))
                except:
                    pass
            
            output_type = settings.get("output_type", "OPC UA Write")
            idx = self.combo_output_type.findText(output_type)
            if idx >= 0:
                self.combo_output_type.setCurrentIndex(idx)
            
            self.check_refresh_output.setChecked(bool(settings.get("refresh_output", 1)))
            self.check_generate_events.setChecked(bool(settings.get("generate_events", 1)))
            
            self.label_status.setText("狀態：已載入設定")
            self.label_status.setStyleSheet("color: green; font-size: 9pt;")
    
    def _save_settings(self):
        """儲存設定到資料庫"""
        if not self.db_manager:
            QMessageBox.warning(self, "錯誤", "資料庫未連線")
            return
        
        settings = {
            "profile_name": self.edit_profile_name.text().strip() or "預設 Profile",
            "description": self.edit_description.toPlainText().strip(),
            "enable_schedule": 1 if self.check_enable_schedule.isChecked() else 0,
            "scan_rate": self.spin_scan_rate.value(),
            "refresh_rate": self.spin_refresh_rate.value(),
            "use_active_period": 1 if self.check_use_active_period.isChecked() else 0,
            "active_from": self.datetime_active_from.dateTime().toPython().isoformat() if self.check_use_active_period.isChecked() else None,
            "active_to": self.datetime_active_to.dateTime().toPython().isoformat() if self.check_use_active_period.isChecked() else None,
            "output_type": self.combo_output_type.currentText(),
            "refresh_output": 1 if self.check_refresh_output.isChecked() else 0,
            "generate_events": 1 if self.check_generate_events.isChecked() else 0,
        }
        
        success = self.db_manager.save_general_settings(settings)
        
        if success:
            self.current_settings = settings
            self.label_status.setText("狀態：設定已儲存")
            self.label_status.setStyleSheet("color: green; font-size: 9pt;")
            QMessageBox.information(self, "儲存成功", "全局設定已儲存")
            self.settings_changed.emit()
        else:
            QMessageBox.warning(self, "儲存失敗", "無法儲存設定到資料庫")
    
    def _reset_settings(self):
        """重設為預設值"""
        reply = QMessageBox.question(
            self,
            "確認重設",
            "確定要重設為預設值嗎？\n未儲存的變更將遺失。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self._load_settings()
    
    def get_current_settings(self) -> Dict[str, Any]:
        """取得當前設定（不儲存）"""
        return {
            "profile_name": self.edit_profile_name.text().strip() or "預設 Profile",
            "description": self.edit_description.toPlainText().strip(),
            "enable_schedule": 1 if self.check_enable_schedule.isChecked() else 0,
            "scan_rate": self.spin_scan_rate.value(),
            "refresh_rate": self.spin_refresh_rate.value(),
            "use_active_period": 1 if self.check_use_active_period.isChecked() else 0,
            "output_type": self.combo_output_type.currentText(),
            "refresh_output": 1 if self.check_refresh_output.isChecked() else 0,
            "generate_events": 1 if self.check_generate_events.isChecked() else 0,
        }

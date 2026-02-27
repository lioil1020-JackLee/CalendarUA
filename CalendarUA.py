#!/usr/bin/env python3
"""
CalendarUA - 工業自動化排程管理系統主程式
採用 PySide6 開發，結合 Office/Outlook 風格行事曆介面
"""

import sys
import asyncio
from datetime import datetime, timedelta, time
from typing import Optional, Dict, Any, List
import logging

# 設定統一的日誌記錄級別
logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QSplitter,
    QCalendarWidget,
    QTableWidget,
    QTableWidgetItem,
    QPushButton,
    QLabel,
    QLineEdit,
    QInputDialog,
    QGroupBox,
    QMessageBox,
    QHeaderView,
    QMenu,
    QSystemTrayIcon,
    QStyle,
    QDialog,
    QComboBox,
    QSpinBox,
    QTextEdit,
    QStatusBar,
    QToolBar,
    QTreeWidget,
    QRadioButton,
    QCheckBox,
    QFileDialog,
    QTreeWidgetItem,
    QTabWidget,
    QStackedWidget,
    QDateEdit,
)
from PySide6.QtCore import Qt, QTimer, Signal, Slot, QThread, QDate, QSize
from PySide6.QtGui import QAction, QColor, QIcon
import qasync
import re

from database.sqlite_manager import SQLiteManager
from core.opc_handler import OPCHandler
from core.rrule_parser import RRuleParser
from ui.recurrence_dialog import RecurrenceDialog, show_recurrence_dialog
from ui.database_settings_dialog import DatabaseSettingsDialog
from ui.exceptions_panel import ExceptionsPanel
from ui.holidays_panel import HolidaysPanel
from ui.general_panel import GeneralPanel


def get_app_icon():
    """獲取應用程式圖示，支援打包環境"""
    import sys
    import os

    # 優先檢查打包環境中的圖示
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包環境
        base_path = sys._MEIPASS
        icon_name = 'lioil.ico' if os.name == 'nt' else 'lioil.icns'
        icon_path = os.path.join(base_path, icon_name)
        if os.path.exists(icon_path):
            return QIcon(icon_path)

    # 開發環境：檢查當前目錄
    icon_name = 'lioil.ico' if os.name == 'nt' else 'lioil.icns'
    if os.path.exists(icon_name):
        return QIcon(icon_name)

    # 預設圖示
    return QIcon()


class SchedulerWorker(QThread):
    """背景排程工作執行緒"""

    trigger_task = Signal(dict)

    def __init__(self, db_manager: SQLiteManager, check_interval: int = 1):
        super().__init__()
        self.db_manager = db_manager
        self.check_interval = check_interval
        self.running = True
        self.last_check = datetime.now()
        # 記錄每個排程的上次觸發時間，防止重複觸發
        self.last_trigger_times: Dict[int, datetime] = {}

    def run(self):
        """持續檢查排程"""
        while self.running:
            try:
                current_time = datetime.now()

                # 取得所有啟用的排程
                schedules = self.db_manager.get_all_schedules(enabled_only=True)

                for schedule in schedules:
                    if self.should_trigger(schedule, current_time):
                        self.trigger_task.emit(schedule)

                self.last_check = current_time

                # 休眠指定秒數
                for _ in range(self.check_interval):
                    if not self.running:
                        break
                    self.msleep(1000)

            except Exception as e:
                print(f"排程檢查錯誤: {e}")
                self.msleep(5000)

    def should_trigger(self, schedule: Dict[str, Any], current_time: datetime) -> bool:
        """檢查是否應該觸發排程"""
        schedule_id = schedule.get("id")
        rrule_str = schedule.get("rrule_str", "")
        if not rrule_str:
            return False

        # 檢查是否已經在最近60秒內觸發過，防止重複觸發
        last_trigger = self.last_trigger_times.get(schedule_id)
        if last_trigger and (current_time - last_trigger).total_seconds() < 60:
            return False

        # 使用 RRuleParser 檢查是否為觸發時間
        if RRuleParser.is_trigger_time(rrule_str, current_time, tolerance_seconds=30):
            # 記錄觸發時間
            self.last_trigger_times[schedule_id] = current_time
            return True

        return False

    def stop(self):
        """停止工作執行緒"""
        self.running = False
        self.wait(2000)


class CalendarUA(QMainWindow):
    """CalendarUA 主視窗"""

    def __init__(self):
        super().__init__()

        self.db_manager: Optional[SQLiteManager] = None
        self.scheduler_worker: Optional[SchedulerWorker] = None
        self.schedules: List[Dict[str, Any]] = []
        
        # 執行計數器：schedule_id -> 已執行次數
        self.execution_counts: Dict[int, int] = {}
        
        # 正在執行的任務ID集合，防止重複執行
        self.running_tasks: set[int] = set()

        # 目前選取的排程 ID (Ribbon Edit/Delete 使用)
        self.selected_schedule_id: Optional[int] = None

        # 主題模式: "light", "dark", "system"
        self.current_theme = "system"

        self.setup_ui()
        self.apply_modern_style()
        self.setup_connections()
        self.setup_system_tray()

        # 初始化資料庫連線
        self.init_database()

        # 設定系統主題監聽
        self.setup_theme_listener()

    def setup_ui(self):
        """設定使用者介面"""
        self.setWindowTitle("CalendarUA")
        self.setWindowIcon(get_app_icon())
        self.setMinimumSize(1200, 800)

        # 建立中央widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # 主要內容面板 (全寬顯示)
        main_panel = self.create_main_panel()
        main_layout.addWidget(main_panel)

        # 建立選單列
        self.create_menu_bar()

        # 建立狀態列
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就緒")

    def create_main_panel(self) -> QWidget:
        """建立主要內容面板（全寬顯示三個 Tab）"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        # Tab 系統
        self.schedule_tabs = QTabWidget()

        # Exceptions Tab - 例外記錄管理
        self.exceptions_panel = ExceptionsPanel()
        self.exceptions_panel.exception_changed.connect(self.load_schedules)
        self.schedule_tabs.addTab(self.exceptions_panel, "Exceptions")
        
        # General Tab - 全局設定
 
        self.general_panel = GeneralPanel()
        self.general_panel.settings_changed.connect(self._on_general_settings_changed)
        self.schedule_tabs.addTab(self.general_panel, "General")
        
        # Holidays Tab - 假日管理
        self.holidays_panel = HolidaysPanel()
        self.holidays_panel.holiday_changed.connect(self.load_schedules)
        self.schedule_tabs.addTab(self.holidays_panel, "Holidays")

        self.schedule_tabs.setCurrentIndex(0)
        layout.addWidget(self.schedule_tabs)

        # 資料庫狀態列
        status_layout = QHBoxLayout()
        self.db_status_label = QLabel("資料庫: 未連線")
        status_layout.addWidget(self.db_status_label)
        status_layout.addStretch()
        layout.addLayout(status_layout)

        return panel

    def create_menu_bar(self):
        """建立選單列"""
        menubar = self.menuBar()
        
        # File 選單
        file_menu = menubar.addMenu("&File")
        
        self.action_new = QAction("&New Schedule", self)
        self.action_new.setShortcut("Ctrl+N")
        self.action_new.setStatusTip("新增排程")
        self.action_new.triggered.connect(self.add_schedule)
        file_menu.addAction(self.action_new)
        
        self.action_refresh = QAction("&Refresh", self)
        self.action_refresh.setShortcut("F5")
        self.action_refresh.setStatusTip("重新載入排程資料")
        self.action_refresh.triggered.connect(self.refresh_schedules)
        file_menu.addAction(self.action_refresh)

        self.action_apply = QAction("&Apply Schedule", self)
        self.action_apply.setShortcut("Ctrl+Shift+A")
        self.action_apply.setStatusTip("套用排程變更")
        self.action_apply.triggered.connect(self.apply_schedules)
        file_menu.addAction(self.action_apply)
        
        file_menu.addSeparator()
        
        self.action_load_profile = QAction("&Load Profile...", self)
        self.action_load_profile.setStatusTip("載入設定檔")
        self.action_load_profile.triggered.connect(self.load_profile)
        file_menu.addAction(self.action_load_profile)
        
        self.action_save_profile = QAction("&Save Profile...", self)
        self.action_save_profile.setStatusTip("儲存設定檔")
        self.action_save_profile.triggered.connect(self.save_profile)
        file_menu.addAction(self.action_save_profile)
        
        file_menu.addSeparator()
        
        self.action_exit = QAction("E&xit", self)
        self.action_exit.setShortcut("Ctrl+Q")
        self.action_exit.setStatusTip("離開程式")
        self.action_exit.triggered.connect(self.close)
        file_menu.addAction(self.action_exit)
        
        # Edit 選單
        edit_menu = menubar.addMenu("&Edit")
        
        self.action_edit = QAction("&Edit Schedule", self)
        self.action_edit.setShortcut("Ctrl+E")
        self.action_edit.setStatusTip("編輯排程")
        self.action_edit.setEnabled(True)
        self.action_edit.triggered.connect(self.edit_selected_schedule)
        edit_menu.addAction(self.action_edit)
        
        self.action_delete = QAction("&Delete Schedule", self)
        self.action_delete.setShortcut("Delete")
        self.action_delete.setStatusTip("刪除排程")
        self.action_delete.setEnabled(True)
        self.action_delete.triggered.connect(self.delete_selected_schedule)
        edit_menu.addAction(self.action_delete)
        
        edit_menu.addSeparator()
        
        self.action_manage_categories = QAction("Manage &Categories...", self)
        self.action_manage_categories.setStatusTip("管理 Category 分類")
        self.action_manage_categories.triggered.connect(self.manage_categories)
        edit_menu.addAction(self.action_manage_categories)
        
        # Tools 選單
        tools_menu = menubar.addMenu("&Tools")
        
        self.action_db_settings = QAction("&Database Settings...", self)
        self.action_db_settings.setStatusTip("資料庫連線設定")
        self.action_db_settings.triggered.connect(self.show_database_settings)
        tools_menu.addAction(self.action_db_settings)
        
        # Help 選單
        help_menu = menubar.addMenu("&Help")
        
        self.action_about = QAction("&About CalendarUA", self)
        self.action_about.triggered.connect(self.show_about)
        help_menu.addAction(self.action_about)

    def _on_theme_menu_triggered(self, action):
        """處理主題選單點擊，確保只有一個選項被選中"""
        for theme, act in self.theme_action_group.items():
            if act != action:
                act.setChecked(False)

    def create_tool_bar(self):
        """建立工具列"""
        toolbar = QToolBar("主工具列")
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(24, 24))
        self.addToolBar(toolbar)
        
        # Create 按鈕
        self.btn_toolbar_new = QPushButton("New")
        self.btn_toolbar_new.setToolTip("新增排程 (Ctrl+N)")
        self.btn_toolbar_new.setFixedWidth(80)
        self.btn_toolbar_new.clicked.connect(self.add_schedule)
        toolbar.addWidget(self.btn_toolbar_new)
        
        toolbar.addSeparator()
        
        # Edit 按鈕
        self.btn_toolbar_edit = QPushButton("Edit")
        self.btn_toolbar_edit.setToolTip("編輯排程 (Ctrl+E)")
        self.btn_toolbar_edit.setFixedWidth(80)
        self.btn_toolbar_edit.setEnabled(True)
        self.btn_toolbar_edit.clicked.connect(self.edit_selected_schedule)
        toolbar.addWidget(self.btn_toolbar_edit)
        
        # Delete 按鈕
        self.btn_toolbar_delete = QPushButton("Delete")
        self.btn_toolbar_delete.setToolTip("刪除排程 (Del)")
        self.btn_toolbar_delete.setFixedWidth(80)
        self.btn_toolbar_delete.setEnabled(True)
        self.btn_toolbar_delete.clicked.connect(self.delete_selected_schedule)
        toolbar.addWidget(self.btn_toolbar_delete)
        
        toolbar.addSeparator()
        
        # Refresh 按鈕
        self.btn_toolbar_refresh = QPushButton("Refresh")
        self.btn_toolbar_refresh.setToolTip("重新載入排程資料 (F5)")
        self.btn_toolbar_refresh.setFixedWidth(80)
        self.btn_toolbar_refresh.clicked.connect(self.refresh_schedules)
        toolbar.addWidget(self.btn_toolbar_refresh)

        # Apply 按鈕
        self.btn_toolbar_apply = QPushButton("Apply")
        self.btn_toolbar_apply.setToolTip("套用排程變更 (Ctrl+Shift+A)")
        self.btn_toolbar_apply.setFixedWidth(80)
        self.btn_toolbar_apply.clicked.connect(self.apply_schedules)
        toolbar.addWidget(self.btn_toolbar_apply)
        
        toolbar.addSeparator()
        
        # Categories 按鈕
        self.btn_toolbar_categories = QPushButton("Categories")
        self.btn_toolbar_categories.setToolTip("管理 Category 分類")
        self.btn_toolbar_categories.setFixedWidth(100)
        self.btn_toolbar_categories.clicked.connect(self.manage_categories)
        toolbar.addWidget(self.btn_toolbar_categories)
        
        toolbar.addWidget(QLabel(""))  # Spacer
        toolbar.addSeparator()
        
        # Scheduler 控制
        scheduler_label = QLabel("Scheduler:")
        toolbar.addWidget(scheduler_label)
        
        self.scheduler_status_label = QLabel("Running")
        self.scheduler_status_label.setStyleSheet("color: green; font-weight: bold;")
        toolbar.addWidget(self.scheduler_status_label)

    def setup_connections(self):
        """設定信號連接"""
        return

    def setup_system_tray(self):
        """設定系統托盤"""
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(get_app_icon())  # 使用應用圖標
        self.tray_icon.setToolTip("CalendarUA - 工業自動化排程管理系統")

        tray_menu = QMenu()
        show_action = QAction("顯示", self)
        show_action.triggered.connect(self.show_window)
        tray_menu.addAction(show_action)

        tray_menu.addSeparator()

        quit_action = QAction("結束", self)
        quit_action.triggered.connect(self.close)
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()  # 總是顯示托盤圖標

    def on_tray_activated(self, reason):
        """處理托盤圖示點擊"""
        if reason == QSystemTrayIcon.DoubleClick:
            self.setWindowState(self.windowState() & ~Qt.WindowMinimized)
            self.show()
            self.raise_()
            QTimer.singleShot(100, self.activateWindow)

    def show_window(self):
        """顯示視窗並置於最上層"""
        self.setWindowState(self.windowState() & ~Qt.WindowMinimized)
        self.show()
        self.raise_()
        QTimer.singleShot(100, self.activateWindow)

    def apply_modern_style(self):
        """套用現代化樣式，根據主題模式選擇亮色或暗色主題"""
        # 判斷是否使用暗色模式
        is_dark = False
        if self.current_theme == "dark":
            is_dark = True
        elif self.current_theme == "system":
            is_dark = self.is_system_dark_mode()

        if is_dark:
            self._apply_dark_theme()
        else:
            self._apply_light_theme()

    def _apply_light_theme(self):
        """套用亮色主題"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #d0d0d0;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 12px;
                background-color: white;
                color: #333;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 8px;
                color: #2c3e50;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
            QTableWidget {
                background-color: white;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                gridline-color: #e0e0e0;
            }
            QTableWidget::item:selected {
                background-color: #0078d4;
                color: white;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 8px;
                border: none;
                border-bottom: 2px solid #0078d4;
                font-weight: bold;
                color: #333;
            }
            QTextEdit {
                background-color: white;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
                color: #333;
            }
            QLabel {
                color: #333;
            }
            QMenuBar {
                background-color: #f0f0f0;
                border-bottom: 1px solid #d0d0d0;
            }
            QMenuBar::item:selected {
                background-color: #0078d4;
                color: white;
            }
            QStatusBar {
                background-color: #f0f0f0;
                border-top: 1px solid #d0d0d0;
            }
            QRadioButton {
                color: #333;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #999999;
                border-radius: 9px;
                background-color: white;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0078d4;
                background-color: #f0f0f0;
            }
            QRadioButton::indicator:checked {
                background-color: #0078d4;
                border: 2px solid #0078d4;
            }
            QCheckBox {
                color: #333;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #999999;
                border-radius: 2px;
                background-color: white;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #f0f0f0;
            }
            QCheckBox::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
                image: url(:/checkbox_check);
            }
        """)

    def _apply_dark_theme(self):
        """套用暗色主題"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2b2b2b;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #3d3d3d;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 12px;
                background-color: #363636;
                color: #cccccc;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 8px;
                color: #ffffff;
            }
            QPushButton {
                background-color: #0e639c;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #1177bb;
            }
            QPushButton:pressed {
                background-color: #094771;
            }
            QPushButton:disabled {
                background-color: #4a4a4a;
                color: #808080;
            }
            QTableWidget {
                background-color: #1e1e1e;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                gridline-color: #3d3d3d;
                color: #cccccc;
            }
            QTableWidget::item:selected {
                background-color: #094771;
                color: white;
            }
            QHeaderView::section {
                background-color: #252526;
                padding: 8px;
                border: none;
                border-bottom: 2px solid #0e639c;
                font-weight: bold;
                color: #cccccc;
            }
            QTextEdit {
                background-color: #1e1e1e;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 8px;
                color: #cccccc;
            }
            QLabel {
                color: #cccccc;
            }
            QMenuBar {
                background-color: #2b2b2b;
                border-bottom: 1px solid #3d3d3d;
                color: #cccccc;
            }
            QMenuBar::item {
                color: #cccccc;
            }
            QMenuBar::item:selected {
                background-color: #094771;
                color: white;
            }
            QMenu {
                background-color: #2b2b2b;
                border: 1px solid #3d3d3d;
                color: #cccccc;
            }
            QMenu::item:selected {
                background-color: #094771;
                color: white;
            }
            QStatusBar {
                background-color: #2b2b2b;
                border-top: 1px solid #3d3d3d;
                color: #cccccc;
            }
            QLineEdit {
                background-color: #1e1e1e;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                color: #cccccc;
            }
            QLineEdit:focus {
                border: 2px solid #0e639c;
            }
            QComboBox {
                background-color: #1e1e1e;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                color: #cccccc;
            }
            QSpinBox {
                background-color: #1e1e1e;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                color: #cccccc;
            }
            QCalendarWidget {
                background-color: #2b2b2b;
            }
            QCalendarWidget QTableView {
                selection-background-color: #094771;
                selection-color: white;
                background-color: #1e1e1e;
                color: #cccccc;
            }
            QCalendarWidget QWidget#qt_calendar_navigationbar {
                background-color: #0e639c;
            }
            QCalendarWidget QToolButton {
                color: white;
                background-color: transparent;
                border: none;
                font-weight: bold;
            }
            QRadioButton {
                color: #cccccc;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #666666;
                border-radius: 9px;
                background-color: #1e1e1e;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #252526;
            }
            QRadioButton::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
            }
            QCheckBox {
                color: #cccccc;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #666666;
                border-radius: 2px;
                background-color: #1e1e1e;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #252526;
            }
            QCheckBox::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
                image: url(:/checkbox_check);
            }
        """)

    def setup_theme_listener(self):
        """設定系統主題監聽"""
        # 在 Windows 上監聽系統主題變化
        try:
            import winreg

            self._theme_timer = QTimer(self)
            self._theme_timer.timeout.connect(self.check_system_theme)
            self._theme_timer.start(2000)  # 每2秒檢查一次
            self._last_theme = self.is_system_dark_mode()
        except ImportError:
            pass

    def is_system_dark_mode(self) -> bool:
        """檢查系統是否使用暗色模式"""
        try:
            import winreg

            registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            key = winreg.OpenKey(registry, key_path)
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            winreg.CloseKey(key)
            return value == 0
        except Exception:
            return False

    def check_system_theme(self):
        """檢查系統主題是否變化"""
        if self.current_theme == "system":
            current_system_theme = self.is_system_dark_mode()
            if current_system_theme != self._last_theme:
                self._last_theme = current_system_theme
                self.apply_modern_style()

    def set_theme(self, theme: str):
        """設定主題模式

        Args:
            theme: "light", "dark", 或 "system"
        """
        if theme in ["light", "dark", "system"]:
            self.current_theme = theme
            self.apply_modern_style()

    def on_theme_changed(self):
        """處理主題選擇改變"""
        theme_data = self.theme_combo.currentData()
        self.set_theme(theme_data)

    def init_database(self):
        """初始化資料庫連線"""
        try:
            # 初始化 SQLite 管理器（使用預設資料庫路徑）
            self.db_manager = SQLiteManager()

            # 建立資料表
            if self.db_manager.init_db():
                self.db_status_label.setText("資料庫: 已連線")
                self.db_status_label.setStyleSheet("color: green;")
                
                # 設定 General Panel 的 db_manager
                if hasattr(self, 'general_panel'):
                    self.general_panel.set_db_manager(self.db_manager)
                
                # 設定 Holidays Panel 的 db_manager
                if hasattr(self, 'holidays_panel'):
                    self.holidays_panel.set_db_manager(self.db_manager)
                
                # 設定 Exceptions Panel 的 db_manager
                if hasattr(self, 'exceptions_panel'):
                    self.exceptions_panel.set_db_manager(self.db_manager)
                
                self.load_schedules()
                self.start_scheduler()
            else:
                self.db_status_label.setText("資料庫: 資料表建立失敗")
                self.db_status_label.setStyleSheet("color: red;")

        except Exception as e:
            self.db_status_label.setText(f"資料庫: 連線失敗")
            self.db_status_label.setStyleSheet("color: red;")
            self.db_manager = None
            QMessageBox.warning(
                self,
                "資料庫連線失敗",
                f"無法連線到資料庫:\n{str(e)}\n\n請檢查資料庫設定。",
            )

    def load_schedules(self):
        """載入排程列表"""
        if not self.db_manager:
            return

        self.schedules = self.db_manager.get_all_schedules()
        self.schedule_exceptions = self.db_manager.get_all_schedule_exceptions()
        self.holiday_entries = self.db_manager.get_all_holiday_entries()
        # 重置執行計數器（應用程式重啟時從 0 開始）
        self.execution_counts = {}
        
        # 載入到 Holidays Panel
        if hasattr(self, 'holidays_panel'):
            self.holidays_panel.refresh()
        
        # 載入到 Exceptions Panel
        if hasattr(self, 'exceptions_panel'):
            self.exceptions_panel.load_data(self.schedules, self.schedule_exceptions)
        
        self.status_bar.showMessage(f"已載入 {len(self.schedules)} 個排程")

    def _on_general_settings_changed(self):
        """全局設定變更時的處理"""
        self.status_bar.showMessage("全局設定已更新")
        # 可以在這裡根據新設定調整行為，例如更新掃描間隔等
        if hasattr(self, 'general_panel'):
            settings = self.general_panel.get_current_settings()
            # 如果需要，可以根據 settings['enable_schedule'] 啟用/禁用排程
            logger.info(f"全局設定已更新: {settings.get('profile_name', 'Unknown')}")

    def _format_schedule_description(self, rrule_str: str, schedule_id: int = 0) -> str:
        """將 RRULE 轉換為中文簡易說明"""
        if not rrule_str:
            return "未設定"

        try:
            # 簡單的 RRULE 解析，轉換為中文說明
            parts = rrule_str.upper().split(';')
            freq_map = {
                'DAILY': '每天',
                'WEEKLY': '每週',
                'MONTHLY': '每月',
                'YEARLY': '每年',
                'HOURLY': '每小時',
                'MINUTELY': '每分鐘',
                'SECONDLY': '每秒'
            }
            
            freq = ""
            interval = 1
            byday = ""
            bymonthday = ""
            byhour = ""
            byminute = ""
            bysecond = ""
            count = ""
            until = ""
            bymonth = ""
            bysetpos = ""
            
            for part in parts:
                if part.startswith('FREQ='):
                    freq_code = part.split('=')[1]
                    freq = freq_map.get(freq_code, freq_code)
                elif part.startswith('INTERVAL='):
                    interval = int(part.split('=')[1])
                elif part.startswith('BYDAY='):
                    byday = part.split('=')[1]
                elif part.startswith('BYMONTHDAY='):
                    bymonthday = part.split('=')[1]
                elif part.startswith('BYHOUR='):
                    byhour = part.split('=')[1]
                elif part.startswith('BYMINUTE='):
                    byminute = part.split('=')[1]
                elif part.startswith('BYSECOND='):
                    bysecond = part.split('=')[1]
                elif part.startswith('COUNT='):
                    count = part.split('=')[1]
                elif part.startswith('UNTIL='):
                    until = part.split('=')[1]
                elif part.startswith('BYMONTH='):
                    bymonth = part.split('=')[1]
                elif part.startswith('BYSETPOS='):
                    bysetpos = part.split('=')[1]
            
            # 如果有 COUNT，計算剩餘次數
            if count and schedule_id:
                try:
                    original_count = int(count)
                    executed_count = self.execution_counts.get(schedule_id, 0)
                    remaining_count = max(0, original_count - executed_count)
                    count = str(remaining_count)
                except ValueError:
                    pass  # 如果解析失敗，保持原樣
            
            # 生成中文描述
            desc_parts = []
            
            # 頻率部分
            if interval > 1:
                desc_parts.append(f"每{interval}{freq[1:]}")  # 每3週
            else:
                desc_parts.append(freq)  # 每天
            
            # 範圍部分
            range_desc = ""
            if bymonth and bymonthday:
                # X月Y日
                month_map = {
                    '1': '一月', '2': '二月', '3': '三月', '4': '四月', '5': '五月', '6': '六月',
                    '7': '七月', '8': '八月', '9': '九月', '10': '十月', '11': '十一月', '12': '十二月'
                }
                month_name = month_map.get(bymonth, f"{bymonth}月")
                range_desc = f"{month_name}{bymonthday}日"
            elif bysetpos and byday:
                # X月的 第Y個 Z
                if bymonth:
                    month_map = {
                        '1': '一月', '2': '二月', '3': '三月', '4': '四月', '5': '五月', '6': '六月',
                        '7': '七月', '8': '八月', '9': '九月', '10': '十月', '11': '十一月', '12': '十二月'
                    }
                    month_name = month_map.get(bymonth, f"{bymonth}月")
                    pos_map = {'1': '第1個', '2': '第2個', '3': '第3個', '4': '第4個', '5': '第5個', '-1': '最後1個'}
                    pos = pos_map.get(bysetpos, f'第{bysetpos}個')
                    day_map = {
                        'MO': '週一', 'TU': '週二', 'WE': '週三', 'TH': '週四',
                        'FR': '週五', 'SA': '週六', 'SU': '週日'
                    }
                    days = [day_map.get(day, day) for day in byday.split(',')]
                    range_desc = f"{month_name}的 {pos} {','.join(days)}"
                else:
                    # 每月的 第Y個 Z
                    pos_map = {'1': '第1個', '2': '第2個', '3': '第3個', '4': '第4個', '5': '第5個', '-1': '最後1個'}
                    pos = pos_map.get(bysetpos, f'第{bysetpos}個')
                    day_map = {
                        'MO': '週一', 'TU': '週二', 'WE': '週三', 'TH': '週四',
                        'FR': '週五', 'SA': '週六', 'SU': '週日'
                    }
                    days = [day_map.get(day, day) for day in byday.split(',')]
                    if len(days) == 5 and set(days) == {'週一', '週二', '週三', '週四', '週五'}:
                        range_desc = f"{pos} 週一到週五"
                    else:
                        range_desc = f"{pos} {','.join(days)}"
            elif byday:
                # 星期幾
                day_map = {
                    'MO': '一', 'TU': '二', 'WE': '三', 'TH': '四',
                    'FR': '五', 'SA': '六', 'SU': '日'
                }
                days = [day_map.get(day, day) for day in byday.split(',')]
                if len(days) == 5 and set(days) == {'一', '二', '三', '四', '五'}:
                    range_desc = "工作天"
                else:
                    range_desc = "".join(days)
            elif bymonthday:
                range_desc = f"第{bymonthday}天"
            
            if range_desc:
                desc_parts.append(range_desc)
            
            # 時間部分
            time_parts = []
            if byhour:
                time_parts.append(byhour)
            if byminute:
                time_parts.append(byminute)
            if bysecond:
                time_parts.append(bysecond)
            
            if time_parts:
                # 根據頻率和可用參數決定時間顯示格式
                if byhour:
                    # 有小時信息，使用完整的時間格式
                    hour = int(byhour)
                    minute = int(byminute) if byminute else 0
                    second = int(bysecond) if bysecond else 0
                    
                    if hour < 12:
                        time_str = f"上午{hour}:{minute:02d}"
                    elif hour == 12:
                        time_str = f"中午{hour}:{minute:02d}"
                    else:
                        time_str = f"下午{hour-12}:{minute:02d}"
                    
                    # 如果有秒數參數，總是顯示秒數
                    if bysecond is not None:
                        time_str += f":{second:02d}"
                else:
                    # 沒有小時信息，只顯示分鐘和秒鐘
                    minute = int(byminute) if byminute else 0
                    second = int(bysecond) if bysecond else 0
                    
                    if freq_code in ['MINUTELY', 'SECONDLY']:
                        # 對於分鐘級或秒級頻率，顯示相對時間
                        if byminute and bysecond:
                            time_str = f"第{minute}分第{second}秒"
                        elif bysecond:
                            time_str = f"第{second}秒"
                        else:
                            time_str = f"第{minute}分"
                    else:
                        # 其他情況顯示絕對時間
                        time_str = f"{minute:02d}:{second:02d}"
                
                desc_parts.append(time_str)
            
            # 結束條件
            end_desc = ""
            if count:
                end_desc = f"剩餘{count}次之後結束"
            elif until:
                # 格式化日期，假設是 YYYYMMDD
                if len(until) >= 8:
                    year = until[:4]
                    month = until[4:6].lstrip('0')  # 移除前導零
                    day = until[6:8].lstrip('0')    # 移除前導零
                    end_desc = f"結束於{year}/{month}/{day}"
            
            if end_desc:
                desc_parts.append(end_desc)
            
            return " ".join(desc_parts)
            
        except Exception:
            return rrule_str  # 如果解析失敗，返回原始字串

    def _format_node_name(self, node_id: str) -> str:
        """從 Node ID 提取節點名稱進行顯示"""
        if not node_id:
            return ""

        try:
            # 檢查是否包含 display_name|node_id 格式
            if "|" in node_id:
                parts = node_id.split("|")
                if len(parts) >= 2:
                    display_name = parts[0]
                    actual_node_id = parts[1]
                    # 優先從 actual_node_id 提取完整路徑
                    if actual_node_id.startswith("ns="):
                        ns_parts = actual_node_id.split(";")
                        if len(ns_parts) > 1:
                            last_part = ns_parts[-1]
                            if last_part.startswith("s="):
                                full_path = last_part[2:]
                                if full_path:
                                    return full_path
                    # 如果無法提取完整路徑，返回 display_name
                    return display_name
            
            # 處理 OPC UA NodeId 的字串表示
            if node_id.startswith("NodeId("):
                # 提取 Identifier 部分
                import re
                match = re.search(r"Identifier='([^']+)'", node_id)
                if match:
                    return match.group(1)
            
            # 特殊處理某些 OPC UA 實現的 Node ID 格式
            if node_id.startswith("String: "):
                identifier = node_id[7:]  # 移除 "String: " 前綴
                # 如果 identifier 是簡單的數字或短字串，嘗試提供更好的名稱
                if identifier == "3>":
                    return "Delta_42_1F.HPW1.DT1"
                return identifier
            elif node_id.startswith("Numeric: "):
                identifier = node_id[8:]  # 移除 "Numeric: " 前綴
                return f"Node_{identifier}"
            
            # 如果是標準 OPC UA Node ID 格式，提取最後一部分
            if node_id.startswith("ns="):
                # 格式如: ns=2;s=MyVariable 或 ns=2;i=12345
                parts = node_id.split(";")
                if len(parts) > 1:
                    last_part = parts[-1]
                    if last_part.startswith("s="):
                        return last_part[2:]  # 移除 "s=" 前綴
                    elif last_part.startswith("i="):
                        return f"Node_{last_part[2:]}"  # 數值 ID 轉換為可讀格式
                    else:
                        return last_part
            
            # 如果不是標準格式，返回最後一個點號之後的部分
            if "." in node_id:
                return node_id.split(".")[-1]
            
            # 如果都沒有特殊格式，返回原字串
            return node_id
            
        except Exception:
            return node_id  # 如果處理失敗，返回原始字串

    def _calculate_next_execution_time(self, schedule: Dict[str, Any]) -> str:
        """計算下次執行時間"""
        rrule_str = schedule.get("rrule_str", "")
        if not rrule_str:
            return "未設定"
        
        try:
            # 使用 RRuleParser 計算下次執行時間
            next_time = RRuleParser.get_next_trigger(rrule_str)
            
            # 檢查 UNTIL 過期
            parts = rrule_str.upper().split(';')
            until_expired = False
            for part in parts:
                if part.startswith('UNTIL='):
                    until_value = part.split('=', 1)[1]
                    if len(until_value) >= 8:
                        try:
                            year = int(until_value[:4])
                            month = int(until_value[4:6])
                            day = int(until_value[6:8])
                            until_date = datetime(year, month, day).date()
                            today = datetime.now().date()
                            if until_date < today:
                                until_expired = True
                        except ValueError:
                            pass
                    break
            
            if next_time:
                # 格式化時間顯示
                time_str = next_time.strftime("%Y/%m/%d %H:%M:%S")
                return time_str
            else:
                # 沒有下次執行時間，檢查是否因為過期
                if until_expired:
                    return "已過期"
                else:
                    return "已結束"
                
        except Exception:
            return "計算失敗"



    def add_schedule(self):
        """新增排程"""
        dialog = ScheduleEditDialog(self)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()

            if self.db_manager:
                schedule_id = self.db_manager.add_schedule(
                    task_name=data["task_name"],
                    opc_url=data["opc_url"],
                    node_id=data["node_id"],
                    target_value=data["target_value"],
                    data_type=data.get("data_type", "auto"),
                    rrule_str=data["rrule_str"],
                    category_id=data.get("category_id", 1),
                    opc_security_policy=data.get("opc_security_policy", "None"),
                    opc_security_mode=data.get("opc_security_mode", "None"),
                    opc_username=data.get("opc_username", ""),
                    opc_password=data.get("opc_password", ""),
                    opc_timeout=data.get("opc_timeout", 5),
                    opc_write_timeout=data.get("opc_write_timeout", 3),
                    is_enabled=data.get("is_enabled", 1),
                )

                if schedule_id:
                    QMessageBox.information(self, "成功", "排程已新增")
                    self.load_schedules()
                else:
                    QMessageBox.critical(self, "錯誤", "新增排程失敗")

    def edit_schedule(self, schedule_id: int = None):
        """編輯排程"""
        if schedule_id is None:
            QMessageBox.information(self, "提示", "請先選擇要編輯的排程")
            return
        
        schedule = next((s for s in self.schedules if s['id'] == schedule_id), None)
        if not schedule:
            return

        dialog = ScheduleEditDialog(self, schedule)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()

            if self.db_manager:
                success = self.db_manager.update_schedule(
                    schedule_id=schedule["id"],
                    task_name=data["task_name"],
                    opc_url=data["opc_url"],
                    node_id=data["node_id"],
                    target_value=data["target_value"],
                    data_type=data.get("data_type", "auto"),
                    rrule_str=data["rrule_str"],
                    category_id=data.get("category_id", 1),
                    opc_security_policy=data.get("opc_security_policy", "None"),
                    opc_security_mode=data.get("opc_security_mode", "None"),
                    opc_username=data.get("opc_username", ""),
                    opc_password=data.get("opc_password", ""),
                    opc_timeout=data.get("opc_timeout", 5),
                    opc_write_timeout=data.get("opc_write_timeout", 3),
                    is_enabled=data.get("is_enabled", 1),
                )

                if success:
                    QMessageBox.information(self, "成功", "排程已更新")
                    self.load_schedules()
                else:
                    QMessageBox.critical(self, "錯誤", "更新排程失敗")

    def delete_schedule(self, schedule_id: int = None):
        """刪除排程"""
        if schedule_id is None:
            QMessageBox.information(self, "提示", "請先選擇要刪除的排程")
            return
        
        schedule = next((s for s in self.schedules if s['id'] == schedule_id), None)
        if not schedule:
            return

        reply = QMessageBox.question(
            self,
            "確認刪除",
            f"確定要刪除排程 '{schedule.get('task_name')}' 嗎？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )

        if reply == QMessageBox.Yes:
            if self.db_manager:
                success = self.db_manager.delete_schedule(schedule["id"])

                if success:
                    QMessageBox.information(self, "成功", "排程已刪除")
                    self.load_schedules()
                else:
                    QMessageBox.critical(self, "錯誤", "刪除排程失敗")

    def manage_categories(self):
        """開啟 Category 管理對話框"""
        from ui.category_manager_dialog import CategoryManagerDialog
        
        dialog = CategoryManagerDialog(self, self.db_manager)
        dialog.category_changed.connect(self.on_category_changed)
        dialog.exec()

    def on_category_changed(self):
        """當 Category 變更時重新載入視圖"""
        self.load_schedules()  # 重新載入排程以更新 category 顏色

    def edit_selected_schedule(self):
        """編輯目前選取的排程"""
        target_id = self.selected_schedule_id
        if target_id is None:
            target_id = self._pick_schedule_id("選擇要編輯的排程")
            if target_id is None:
                return
        self.edit_schedule(target_id)

    def delete_selected_schedule(self):
        """刪除目前選取的排程"""
        target_id = self.selected_schedule_id
        if target_id is None:
            target_id = self._pick_schedule_id("選擇要刪除的排程")
            if target_id is None:
                return
        self.delete_schedule(target_id)

    def _pick_schedule_id(self, title: str) -> Optional[int]:
        """從現有排程中挑選一筆 ID。"""
        if not self.schedules:
            QMessageBox.information(self, "提示", "目前沒有可操作的排程")
            return None

        items: List[str] = []
        mapping: Dict[str, int] = {}
        for schedule in self.schedules:
            schedule_id = int(schedule.get("id", 0))
            task_name = str(schedule.get("task_name", "未命名"))
            item_text = f"[{schedule_id}] {task_name}"
            items.append(item_text)
            mapping[item_text] = schedule_id

        selected_text, ok = QInputDialog.getItem(
            self,
            title,
            "排程:",
            items,
            0,
            False,
        )
        if not ok or not selected_text:
            return None

        return mapping.get(selected_text)

    def refresh_schedules(self):
        """重新載入排程資料 (F5)"""
        if not self.db_manager:
            QMessageBox.warning(self, "警告", "資料庫未連線")
            return
        
        self.status_bar.showMessage("正在重新載入排程...")
        self.load_schedules()
        self.status_bar.showMessage(f"已重新載入 {len(self.schedules)} 個排程", 3000)

    def apply_schedules(self):
        """套用排程變更"""
        if not self.db_manager:
            QMessageBox.warning(self, "警告", "資料庫未連線")
            return

        # 目前所有變更皆即時寫入，此處仍提供一致的操作入口
        self.load_schedules()
        self.status_bar.showMessage("已套用排程變更", 3000)

    def load_profile(self):
        """載入設定檔"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "載入 Profile 設定檔",
            "",
            "JSON Files (*.json);;All Files (*.*)"
        )
        
        if not file_path:
            return
        
        try:
            import json
            with open(file_path, 'r', encoding='utf-8') as f:
                profile_data = json.load(f)
            
            # 載入 General Settings
            if 'general_settings' in profile_data and hasattr(self, 'general_panel'):
                self.general_panel.load_from_dict(profile_data['general_settings'])

            if self.db_manager:
                # 載入 Categories (僅使用者自訂)
                categories = profile_data.get('categories', [])
                if categories:
                    existing_categories = {c['name']: c for c in self.db_manager.get_all_categories()}
                    for cat in categories:
                        name = cat.get('name', '').strip()
                        if not name:
                            continue
                        bg_color = cat.get('bg_color', '#FFFFFF')
                        fg_color = cat.get('fg_color', '#000000')
                        sort_order = int(cat.get('sort_order', 0) or 0)
                        existing = existing_categories.get(name)
                        if existing:
                            if existing.get('is_system'):
                                continue
                            self.db_manager.update_category(
                                existing['id'],
                                name=name,
                                bg_color=bg_color,
                                fg_color=fg_color,
                                sort_order=sort_order,
                            )
                        else:
                            self.db_manager.add_category(name, bg_color, fg_color, sort_order)

                # 載入 Schedules
                schedules = profile_data.get('schedules', [])
                if schedules:
                    existing_schedules = self.db_manager.get_all_schedules()
                    for sch in schedules:
                        task_name = str(sch.get('task_name', '')).strip()
                        rrule_str = str(sch.get('rrule_str', '')).strip()
                        opc_url = str(sch.get('opc_url', '')).strip()
                        node_id = str(sch.get('node_id', '')).strip()
                        if not task_name or not rrule_str:
                            continue

                        matched = next(
                            (
                                s for s in existing_schedules
                                if s.get('task_name') == task_name
                                and str(s.get('rrule_str', '')).strip() == rrule_str
                                and str(s.get('opc_url', '')).strip() == opc_url
                                and str(s.get('node_id', '')).strip() == node_id
                            ),
                            None,
                        )

                        data = {
                            "task_name": task_name,
                            "opc_url": opc_url,
                            "node_id": node_id,
                            "target_value": str(sch.get('target_value', '')),
                            "data_type": str(sch.get('data_type', 'auto')),
                            "rrule_str": rrule_str,
                            "category_id": int(sch.get('category_id', 1) or 1),
                            "opc_security_policy": str(sch.get('opc_security_policy', 'None')),
                            "opc_security_mode": str(sch.get('opc_security_mode', 'None')),
                            "opc_username": str(sch.get('opc_username', '')),
                            "opc_password": str(sch.get('opc_password', '')),
                            "opc_timeout": int(sch.get('opc_timeout', 5) or 5),
                            "opc_write_timeout": int(sch.get('opc_write_timeout', 3) or 3),
                            "is_enabled": int(sch.get('is_enabled', 1) or 1),
                        }

                        if matched:
                            self.db_manager.update_schedule(matched['id'], **data)
                        else:
                            self.db_manager.add_schedule(**data)

                self.load_schedules()
            
            QMessageBox.information(self, "成功", f"已載入設定檔:\n{file_path}")
            self.status_bar.showMessage(f"已載入 Profile: {file_path}", 5000)
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"載入設定檔失敗:\n{str(e)}")

    def save_profile(self):
        """儲存設定檔"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "儲存 Profile 設定檔",
            "",
            "JSON Files (*.json);;All Files (*.*)"
        )
        
        if not file_path:
            return

        if not file_path.lower().endswith('.json'):
            file_path = f"{file_path}.json"
        
        try:
            import json
            profile_data = {}
            
            # 儲存 General Settings
            if hasattr(self, 'general_panel'):
                profile_data['general_settings'] = self.general_panel.get_current_settings()

            # 儲存 Categories (僅使用者自訂)
            if self.db_manager:
                categories = [
                    {
                        "name": c.get("name"),
                        "bg_color": c.get("bg_color"),
                        "fg_color": c.get("fg_color"),
                        "sort_order": c.get("sort_order", 0),
                    }
                    for c in self.db_manager.get_all_categories()
                    if not c.get("is_system")
                ]
                profile_data['categories'] = categories

                schedules = []
                for s in self.db_manager.get_all_schedules():
                    schedules.append(
                        {
                            "task_name": s.get("task_name"),
                            "opc_url": s.get("opc_url"),
                            "node_id": s.get("node_id"),
                            "target_value": s.get("target_value"),
                            "data_type": s.get("data_type", "auto"),
                            "rrule_str": s.get("rrule_str"),
                            "category_id": s.get("category_id", 1),
                            "opc_security_policy": s.get("opc_security_policy", "None"),
                            "opc_security_mode": s.get("opc_security_mode", "None"),
                            "opc_username": s.get("opc_username", ""),
                            "opc_password": s.get("opc_password", ""),
                            "opc_timeout": s.get("opc_timeout", 5),
                            "opc_write_timeout": s.get("opc_write_timeout", 3),
                            "is_enabled": s.get("is_enabled", 1),
                        }
                    )
                profile_data['schedules'] = schedules
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(profile_data, f, indent=2, ensure_ascii=False)
            
            QMessageBox.information(self, "成功", f"已儲存設定檔:\n{file_path}")
            self.status_bar.showMessage(f"已儲存 Profile: {file_path}", 5000)
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"儲存設定檔失敗:\n{str(e)}")

    def show_database_settings(self):
        """顯示資料庫設定對話框"""
        dialog = DatabaseSettingsDialog(self)
        if dialog.exec() == QDialog.Accepted:
            # 重新連線資料庫
            self.init_database()

    def show_about(self):
        """顯示關於對話框"""
        about_text = """
        <h2>CalendarUA</h2>
        <p>版本: 1.0.0</p>
        <p>工業自動化排程管理系統</p>
        <p>採用 PySide6 開發，支援 OPC UA 通訊</p>
        <br>
        <p><b>主要功能:</b></p>
        <ul>
            <li>週期性排程管理 (RRULE)</li>
            <li>例外與假日支援</li>
            <li>General / Holidays / Exceptions 分頁</li>
            <li>Category 分類系統</li>
            <li>OPC UA 任務執行與監控</li>
        </ul>
        <br>
        <p>© 2026 CalendarUA Project</p>
        """
        QMessageBox.about(self, "關於 CalendarUA", about_text)


    def start_scheduler(self):
        """啟動排程背景工作"""
        if self.db_manager:
            self.scheduler_worker = SchedulerWorker(self.db_manager)
            self.scheduler_worker.trigger_task.connect(self.on_task_triggered)
            self.scheduler_worker.start()
            self.status_bar.showMessage("排程器已啟動")

    @Slot(dict)
    def on_task_triggered(self, schedule: Dict[str, Any]):
        """處理排程觸發"""
        schedule_id = schedule.get("id")
        
        # 檢查是否已經在執行，防止重複執行
        if schedule_id in self.running_tasks:
            self.status_bar.showMessage(f"任務 {schedule.get('task_name', '')} 已在執行中，跳過此次觸發", 3000)
            return
        
        # 標記為執行中
        self.running_tasks.add(schedule_id)
        
        self.status_bar.showMessage(f"執行排程: {schedule.get('task_name', '')}")

        # 執行 OPC UA 寫入
        asyncio.create_task(self.execute_task(schedule))

    async def execute_task(self, schedule: Dict[str, Any]):
        """執行排程任務"""
        schedule_id = schedule.get("id")
        opc_url = schedule.get("opc_url", "")
        node_id = schedule.get("node_id", "")
        target_value = schedule.get("target_value", "")
        data_type = schedule.get("data_type", "auto")

        # 解析 node_id，提取實際的 OPC UA Node ID
        import re
        if "|" in node_id:
            _, temp = node_id.split("|", 1)
            actual_node_id = temp
        else:
            # 嘗試從 NodeId 字串表示提取
            match = re.search(r"Identifier='([^']+)'", node_id)
            if match:
                identifier = match.group(1)
                # 從 NodeId 字串提取 namespace 和類型資訊
                ns_match = re.search(r"NamespaceIndex=(\d+)", node_id)
                type_match = re.search(r"NodeIdType=<NodeIdType\.(\w+):", node_id)
                if ns_match and type_match:
                    ns = ns_match.group(1)
                    node_type = type_match.group(1)
                    if node_type == "String":
                        actual_node_id = f"ns={ns};s={identifier}"
                    elif node_type == "Numeric":
                        actual_node_id = f"ns={ns};i={identifier}"
                    else:
                        actual_node_id = identifier
                else:
                    actual_node_id = identifier
            else:
                actual_node_id = node_id

        # 取得OPC設定
        security_policy = schedule.get("opc_security_policy", "None")
        username = schedule.get("opc_username", "")
        password = schedule.get("opc_password", "")
        timeout = schedule.get("opc_timeout", 5)
        write_timeout = schedule.get("opc_write_timeout", 3)

        try:
            # 更新狀態為執行中
            if self.db_manager:
                self.db_manager.update_execution_status(schedule_id, "執行中...")
            
            # 重新載入表格以顯示狀態更新
            self.load_schedules()

            # 建立OPCHandler並設定安全參數
            handler = OPCHandler(opc_url, timeout=timeout)

            # 設定安全策略
            if security_policy != "None":
                handler.security_policy = security_policy

            # 設定認證
            if username:
                handler.username = username
                handler.password = password

            async with handler:
                if handler.is_connected:
                    # 重試機制：根據期間決定重試策略
                    duration_minutes = self._parse_duration_from_rrule(schedule.get("rrule_str", ""))
                    
                    if duration_minutes == 0:
                        # 期間=0分：只嘗試寫入一次，不重試
                        max_retries = 1
                        retry_delay = write_timeout
                    else:
                        # 期間>0分：持續寫入直到成功或結束時間到
                        max_retries = float('inf')  # 無限重試，直到成功或時間到
                        retry_delay = write_timeout
                    
                    attempt = 0
                    success_once = False
                    
                    while attempt < max_retries and not success_once:
                        # 檢查是否超過結束時間（對於期間>0的情況）
                        if duration_minutes > 0:
                            current_time = datetime.now()
                            # 這裡需要解析結束時間，簡化處理：假設結束時間是開始時間 + 期間
                            # 實際上應該從RRULE解析結束時間
                            if current_time >= self._calculate_end_time(schedule):
                                break  # 超過結束時間，停止重試
                        
                        try:
                            success = await handler.write_node(actual_node_id, target_value, data_type)
                            if success:
                                status_msg = f"✓ 成功寫入 {node_id} = {target_value}"
                                success_once = True
                                if self.db_manager:
                                    self.db_manager.update_execution_status(schedule_id, "執行成功")
                                    # 增加執行計數器
                                    self.execution_counts[schedule_id] = self.execution_counts.get(schedule_id, 0) + 1
                                    # 檢查是否達到 COUNT 上限
                                    self._check_and_disable_if_count_reached(schedule_id, schedule.get("rrule_str", ""))
                                break  # 成功一次就結束
                            else:
                                if attempt < max_retries - 1 or max_retries == float('inf'):
                                    if duration_minutes == 0:
                                        # 期間=0分，只嘗試一次
                                        status_msg = f"✗ 寫入失敗: {node_id}"
                                        if self.db_manager:
                                            self.db_manager.update_execution_status(schedule_id, "寫入失敗")
                                    else:
                                        # 期間>0分，正在重試
                                        status_msg = f"寫入失敗，正在等待 {retry_delay} 秒後重試..."
                                        logger.warning(f"寫入失敗，正在等待 {retry_delay} 秒後重試")
                                        await asyncio.sleep(retry_delay)
                                else:
                                    # 最後一次嘗試失敗
                                    status_msg = f"✗ 寫入失敗: {node_id}"
                                    if self.db_manager:
                                        self.db_manager.update_execution_status(schedule_id, "寫入失敗")
                        except Exception as e:
                            if attempt < max_retries - 1 or max_retries == float('inf'):
                                if duration_minutes == 0:
                                    # 期間=0分，只嘗試一次
                                    status_msg = f"✗ 執行錯誤: {str(e)[:50]}"
                                    if self.db_manager:
                                        self.db_manager.update_execution_status(schedule_id, f"執行錯誤: {str(e)[:50]}")
                                else:
                                    # 期間>0分，正在重試
                                    status_msg = f"執行錯誤，正在等待 {retry_delay} 秒後重試..."
                                    logger.warning(f"寫入錯誤: {e}，正在等待 {retry_delay} 秒後重試")
                                    await asyncio.sleep(retry_delay)
                            else:
                                # 最後一次嘗試失敗
                                status_msg = f"✗ 執行錯誤: {str(e)[:50]}"
                                if self.db_manager:
                                    self.db_manager.update_execution_status(schedule_id, f"執行錯誤: {str(e)[:50]}")
                                break
                        attempt += 1
                    else:
                        # 如果所有重試都失敗，這裡不會執行，因為break會跳出
                        pass
                else:
                    status_msg = f"✗ 無法連線 OPC UA: {opc_url}"
                    if self.db_manager:
                        self.db_manager.update_execution_status(schedule_id, "連線失敗")
                        
        except Exception as e:
            status_msg = f"✗ 執行錯誤: {str(e)}"
            if self.db_manager:
                self.db_manager.update_execution_status(schedule_id, f"執行錯誤: {str(e)[:50]}")

        # 更新狀態列和重新載入表格
        self.status_bar.showMessage(status_msg, 5000)
        self.load_schedules()
        
        # 從執行中任務集合中移除
        self.running_tasks.discard(schedule_id)

    def _parse_duration_from_rrule(self, rrule_str: str) -> int:
        """從RRULE字串中解析期間參數（分鐘）"""
        if not rrule_str:
            return 0
        
        try:
            parts = rrule_str.upper().split(';')
            for part in parts:
                if part.startswith('DURATION=PT'):
                    # 格式如: DURATION=PT5M
                    duration_str = part.split('=')[1]  # PT5M
                    if duration_str.endswith('M'):
                        minutes = int(duration_str[2:-1])  # 移除PT和M
                        return minutes
        except Exception:
            pass
        return 0

    def _calculate_end_time(self, schedule: Dict[str, Any]) -> datetime:
        """計算任務的結束時間"""
        # 簡化實作：從開始時間 + 期間計算
        # 實際應該從RRULE解析完整的結束時間
        start_time_str = schedule.get("start_time", "")
        duration_minutes = self._parse_duration_from_rrule(schedule.get("rrule_str", ""))
        
        try:
            # 假設start_time是HH:MM格式，轉換為今天的datetime
            today = datetime.now().date()
            time_part = datetime.strptime(start_time_str, "%H:%M").time()
            start_datetime = datetime.combine(today, time_part)
            
            # 如果開始時間已經過去，可能是明天
            if start_datetime < datetime.now():
                start_datetime += timedelta(days=1)
            
            return start_datetime + timedelta(minutes=duration_minutes)
        except Exception:
            # 預設1小時後結束
            return datetime.now() + timedelta(hours=1)

    def _check_and_disable_if_count_reached(self, schedule_id: int, rrule_str: str):
        """檢查是否達到 COUNT 上限，如果是則停用排程"""
        if not rrule_str:
            return

        try:
            # 解析 RRULE 中的 COUNT
            parts = rrule_str.upper().split(';')
            count_value = None
            
            for part in parts:
                if part.startswith('COUNT='):
                    try:
                        count_value = int(part.split('=')[1])
                        break
                    except ValueError:
                        return

            if count_value is not None:
                executed_count = self.execution_counts.get(schedule_id, 0)
                if executed_count >= count_value:
                    # 達到上限，停用排程
                    if self.db_manager:
                        self.db_manager.update_schedule(schedule_id, is_enabled=0)
                        print(f"排程 {schedule_id} 的執行次數已達上限 ({count_value})，已自動停用")

        except Exception as e:
            # 如果解析失敗，記錄錯誤但不中斷執行
            print(f"檢查 COUNT 上限失敗: {e}")

    def show_db_settings(self):
        """顯示資料庫設定對話框"""
        dialog = DatabaseSettingsDialog(self, self.db_manager)
        dialog.database_changed.connect(self.on_database_path_changed)
        dialog.exec()

    def on_database_path_changed(self, new_path: str):
        """處理資料庫路徑變更"""
        # 重新初始化資料庫管理器
        self.db_manager = SQLiteManager(new_path)

        # 重新載入排程資料
        self.load_schedules()

        # 重新啟動排程工作執行緒
        if self.scheduler_worker:
            self.scheduler_worker.stop()
            self.scheduler_worker.wait()

        self.scheduler_worker = SchedulerWorker(self.db_manager)
        self.scheduler_worker.trigger_task.connect(self.on_task_triggered)
        self.scheduler_worker.start()

    def show_about(self):
        """顯示關於對話框"""
        QMessageBox.about(
            self,
            "關於 CalendarUA",
            """<h2>CalendarUA v1.0</h2>
            <p>工業自動化排程管理系統</p>
            <p>採用 Python 3.12 + PySide6 開發</p>
            <p>結合 OPC UA 與 MySQL 技術</p>
            """,
        )

    def closeEvent(self, event):
        """處理視窗關閉事件"""
        import asyncio

        # 停止排程器
        if self.scheduler_worker:
            self.scheduler_worker.stop()

        # 取消所有正在執行的async任務
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                # 取消所有待處理的任務
                pending_tasks = [task for task in asyncio.all_tasks(loop) if not task.done()]
                for task in pending_tasks:
                    task.cancel()

                # 等待任務取消完成，最多等待2秒
                import time
                start_time = time.time()
                while pending_tasks and (time.time() - start_time) < 2.0:
                    # 簡單的輪詢等待
                    time.sleep(0.1)
                    pending_tasks = [task for task in pending_tasks if not task.done()]

        except RuntimeError:
            # 沒有事件迴圈或已經關閉
            pass

        event.accept()

    def changeEvent(self, event):
        """處理視窗狀態變化事件"""
        from PySide6.QtCore import QEvent
        if event.type() == QEvent.WindowStateChange:
            if self.windowState() & Qt.WindowMinimized:
                # 最小化時延遲隱藏視窗，確保托盤圖標正確顯示
                QTimer.singleShot(100, self._minimize_to_tray)
            elif self.windowState() == Qt.WindowNoState:
                # 從最小化恢復時顯示視窗
                self.show()
                self.raise_()
                self.activateWindow()
        super().changeEvent(event)

    def _minimize_to_tray(self):
        """將視窗最小化到系統托盤"""
        if self.windowState() & Qt.WindowMinimized:
            self.hide()
            self.tray_icon.show()
            # 設定工具提示
            self.tray_icon.setToolTip("CalendarUA")


class OPCNodeBrowserDialog(QDialog):
    """OPC UA 節點瀏覽對話框"""

    def __init__(self, parent=None, opc_url: str = ""):
        super().__init__(parent)
        self.opc_url = opc_url
        self.selected_node = ""
        self.opc_handler = None
        self.logger = logging.getLogger(__name__)
        self.setup_ui()
        self.apply_style()
        # 自動連線並載入節點
        QTimer.singleShot(100, self.connect_and_load)

    def setup_ui(self):
        """設定介面"""
        self.setWindowTitle("瀏覽 OPC UA 節點")
        self.setWindowIcon(get_app_icon())
        self.setMinimumSize(500, 400)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # 顯示目前連線資訊
        info_layout = QHBoxLayout()
        info_layout.addWidget(QLabel("OPC URL:"))
        self.url_label = QLabel(self.opc_url)
        self.url_label.setStyleSheet("font-weight: bold;")
        info_layout.addWidget(self.url_label)
        info_layout.addStretch()
        layout.addLayout(info_layout)

        # 狀態標籤
        self.status_label = QLabel("正在連線...")
        self.status_label.setStyleSheet("color: #666;")
        layout.addWidget(self.status_label)

        # 節點樹狀列表
        self.tree_widget = QTreeWidget()
        self.tree_widget.setHeaderLabels(["節點名稱", "Node ID", "節點類型", "資料型別"])
        self.tree_widget.setColumnWidth(0, 200)
        self.tree_widget.setColumnWidth(1, 150)
        self.tree_widget.setColumnWidth(2, 100)
        self.tree_widget.setColumnWidth(3, 80)
        self.tree_widget.itemSelectionChanged.connect(self.on_selection_changed)
        self.tree_widget.itemDoubleClicked.connect(self.on_item_double_clicked)
        layout.addWidget(self.tree_widget)

        # 按鈕
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        refresh_btn = QPushButton("重新整理")
        refresh_btn.clicked.connect(self.connect_and_load)
        button_layout.addWidget(refresh_btn)

        button_layout.addSpacing(20)

        self.select_btn = QPushButton("選擇")
        self.select_btn.setDefault(True)
        self.select_btn.setEnabled(False)
        self.select_btn.clicked.connect(self.accept)
        button_layout.addWidget(self.select_btn)

        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)

    def apply_style(self):
        """套用樣式"""
        parent = self.parent()
        is_dark = False
        if parent and hasattr(parent, "parent"):
            grandparent = parent.parent()
            if grandparent and hasattr(grandparent, "current_theme"):
                if grandparent.current_theme == "dark":
                    is_dark = True
                elif grandparent.current_theme == "system":
                    if hasattr(grandparent, "is_system_dark_mode"):
                        is_dark = grandparent.is_system_dark_mode()

        if is_dark:
            self.setStyleSheet("""
                QDialog {
                    background-color: #2b2b2b;
                }
                QLabel {
                    color: #cccccc;
                }
                QTreeWidget {
                    background-color: #1e1e1e;
                    border: 1px solid #3d3d3d;
                    color: #cccccc;
                }
                QTreeWidget::item:selected {
                    background-color: #094771;
                    color: white;
                }
                QTreeWidget::item:hover {
                    background-color: #2d2d2d;
                }
                QHeaderView::section {
                    background-color: #252526;
                    padding: 6px;
                    border: none;
                    border-bottom: 2px solid #0e639c;
                    font-weight: bold;
                    color: #cccccc;
                }
                QPushButton {
                    background-color: #0e639c;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 8px 16px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #1177bb;
                }
                QPushButton:disabled {
                    background-color: #4a4a4a;
                    color: #808080;
                }
            """)
        else:
            self.setStyleSheet("""
                QDialog {
                    background-color: #f5f5f5;
                }
                QLabel {
                    color: #333;
                }
                QTreeWidget {
                    background-color: white;
                    border: 1px solid #d0d0d0;
                    color: #333;
                }
                QTreeWidget::item:selected {
                    background-color: #0078d4;
                    color: white;
                }
                QTreeWidget::item:hover {
                    background-color: #e5f3ff;
                }
                QHeaderView::section {
                    background-color: #f0f0f0;
                    padding: 6px;
                    border: none;
                    border-bottom: 2px solid #0078d4;
                    font-weight: bold;
                }
                QPushButton {
                    background-color: #0078d4;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 8px 16px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #106ebe;
                }
                QPushButton:disabled {
                    background-color: #cccccc;
                    color: #888888;
                }
            """)

    def connect_and_load(self):
        """連線到 OPC UA 並載入節點 - 使用 qasync 整合"""
        self.tree_widget.clear()
        self.status_label.setText("正在連線...")
        self.status_label.setStyleSheet("color: #666;")

        # 使用 QTimer 稍後執行異步操作，避免阻塞 UI
        QTimer.singleShot(100, self._async_connect_and_load)

    def _async_connect_and_load(self):
        """異步連線和載入"""
        import asyncio

        async def do_connect():
            try:
                from core.opc_handler import OPCHandler

                self.opc_handler = OPCHandler(self.opc_url)

                # 連線到 OPC UA 伺服器
                success = await self.opc_handler.connect()

                if success:
                    self.status_label.setText("已連線，正在載入節點...")
                    self.status_label.setStyleSheet("color: green;")

                    # 載入節點
                    await self._async_load_nodes()

                    # 斷開連線
                    await self.opc_handler.disconnect()
                else:
                    self.status_label.setText("連線失敗 - 請檢查 URL 和伺服器狀態")
                    self.status_label.setStyleSheet("color: red;")

            except Exception as e:
                self.status_label.setText(f"連線錯誤: {str(e)}")
                self.status_label.setStyleSheet("color: red;")

        # 使用現有的 qasync 事件迴圈執行
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                # 如果迴圈已在執行，建立 task
                asyncio.create_task(do_connect())
            else:
                loop.run_until_complete(do_connect())
        except RuntimeError:
            # 沒有事件迴圈的情況
            asyncio.run(do_connect())

    async def _async_load_nodes(self):
        """異步載入 OPC UA 節點樹"""
        try:
            # 取得 Objects 節點
            objects = await self.opc_handler.get_objects_node()

            if objects:
                root_item = QTreeWidgetItem(self.tree_widget)
                root_item.setText(0, "Objects")
                root_item.setText(1, "i=85")
                root_item.setText(2, "Object")
                root_item.setExpanded(True)

                # 載入子節點
                await self._async_load_child_nodes(objects, root_item, depth=0)

            self.status_label.setText("已載入節點")
            # 確保樹狀元件正確更新
            self.tree_widget.update()
            self.tree_widget.repaint()

        except Exception as e:
            self.status_label.setText(f"載入節點錯誤: {str(e)}")
            self.status_label.setStyleSheet("color: red;")

    async def _async_load_child_nodes(self, parent_node, parent_item, depth=0):
        """異步遞迴載入子節點"""
        if depth > 5:  # 增加深度限制到 5，以載入更深層的節點
            return

        try:
            # 取得子節點
            children = await parent_node.get_children()

            for child in children:
                try:
                    child_item = QTreeWidgetItem(parent_item)

                    # 取得節點資訊
                    browse_name = await child.read_browse_name()
                    # 正確格式化 Node ID
                    node_id = child.nodeid.to_string()
                    node_class = await child.read_node_class()

                    # 讀取資料型別和存取權限（僅適用於變數節點）
                    data_type = "-"
                    access_level = ""
                    can_write = False
                    
                    if node_class.name == "Variable":
                        try:
                            # 讀取資料型別
                            detected_type = await self.opc_handler.read_node_data_type(node_id)
                            data_type = detected_type if detected_type else "未知"
                            self.logger.debug(f"Node {node_id} 資料型別: {data_type}")
                            
                            # 讀取存取權限
                            try:
                                from asyncua.ua import AttributeIds
                                access_level_data = await child.read_attribute(AttributeIds.AccessLevel)
                                # 從 DataValue 中提取實際值
                                access_level_value = access_level_data.Value.Value if hasattr(access_level_data, 'Value') and access_level_data.Value else None
                                self.logger.debug(f"Node {node_id} AccessLevel: {access_level_value}")
                                # 檢查是否有 Write 權限 (0x02)，或者如果無法確定，預設為可寫入
                                can_write = bool(access_level_value & 0x02) if access_level_value is not None and access_level_value > 0 else True
                            except Exception as e:
                                self.logger.debug(f"無法讀取 Node {node_id} 的 AccessLevel: {e}")
                                # 如果無法讀取AccessLevel，預設為可寫入
                                can_write = True
                            
                            if not can_write:
                                data_type = "唯讀"
                                access_level = "唯讀"
                            
                        except Exception as e:
                            self.logger.error(f"讀取 Node {node_id} 資料型別失敗: {e}")
                            data_type = "未知"
                            can_write = False

                    child_item.setText(0, browse_name.Name)
                    child_item.setText(1, node_id)
                    child_item.setText(2, str(node_class))
                    child_item.setText(3, data_type)

                    # 儲存節點 ID 和資料型別
                    child_item.setData(0, Qt.ItemDataRole.UserRole, node_id)
                    child_item.setData(0, Qt.ItemDataRole.UserRole + 1, data_type)
                    child_item.setData(0, Qt.ItemDataRole.UserRole + 2, can_write)

                    # 繼續載入子節點
                    await self._async_load_child_nodes(child, child_item, depth + 1)

                except Exception as e:
                    self.logger.warning(f"載入子節點失敗 (深度 {depth + 1}): {e}")
                    # 即使失敗也要繼續處理其他節點

        except Exception as e:
            self.logger.error(f"載入子節點列表失敗 (深度 {depth}): {e}")

    def on_selection_changed(self):
        """處理選擇變更"""
        selected_items = self.tree_widget.selectedItems()
        if selected_items:
            selected_item = selected_items[0]
            display_name = selected_item.text(0)
            node_id = selected_item.text(1)
            data_type = selected_item.text(3) if selected_item.text(3) != "-" else "未知"
            can_write = selected_item.data(0, Qt.ItemDataRole.UserRole + 2)
            
            if can_write:
                self.selected_node = f"{display_name}|{node_id}|{data_type}"
                self.select_btn.setEnabled(True)
                self.status_label.setText("已選擇可寫入節點")
                self.status_label.setStyleSheet("color: green;")
            else:
                self.selected_node = ""
                self.select_btn.setEnabled(False)
                self.status_label.setText("選擇的節點為唯讀，無法寫入")
                self.status_label.setStyleSheet("color: red;")
        else:
            self.selected_node = ""
            self.select_btn.setEnabled(False)
            self.status_label.setText("")

    def on_item_double_clicked(self, item, column):
        """處理雙擊事件"""
        display_name = item.text(0)
        node_id = item.text(1)
        data_type = item.text(3) if item.text(3) != "-" else "未知"
        can_write = item.data(0, Qt.ItemDataRole.UserRole + 2)
        
        if can_write:
            self.selected_node = f"{display_name}|{node_id}|{data_type}"
            self.accept()
        else:
            self.status_label.setText("無法選擇唯讀節點")
            self.status_label.setStyleSheet("color: red;")

    def get_selected_node(self) -> str:
        """取得選擇的節點 ID 和資料型別"""
        return self.selected_node


class OPCSettingsDialog(QDialog):
    """OPC UA 設定對話框"""

    def __init__(self, parent=None, security_policy="None", username="", password="", timeout=5, write_timeout=3, security_mode="None", opc_url: str = ""):
        super().__init__(parent)
        self.security_policy = security_policy
        self.security_mode = security_mode
        self.username = username
        self.password = password
        self.timeout = timeout
        self.write_timeout = write_timeout
        self.opc_url = opc_url
        self._detected_supported = None
        
        self.setup_ui()
        
        # 連接信號
        self.chk_show_supported.toggled.connect(self.on_chk_show_supported_toggled)
        self.rb_anonymous.toggled.connect(self.on_auth_mode_changed)
        self.rb_username.toggled.connect(self.on_auth_mode_changed)
        self.rb_certificate.toggled.connect(self.on_auth_mode_changed)
        
        # 載入現有設定
        self.load_data()
        
        # 套用樣式
        self.apply_style()

    def setup_ui(self):
        self.setWindowTitle("OPC UA 連線設定")
        self.setWindowIcon(get_app_icon())
        self.setMinimumWidth(900)
        # 移除固定高度，讓視窗根據內容自動調整
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(15, 15, 15, 15)

        # 使用者認證（頂部）
        auth_top_layout = QHBoxLayout()
        auth_top_layout.setSpacing(12)
        auth_top_layout.setContentsMargins(0, 0, 0, 0)

        auth_group = QGroupBox("User Authentication")
        auth_group_layout = QVBoxLayout(auth_group)
        auth_group_layout.setContentsMargins(10, 12, 10, 12)
        auth_group_layout.setSpacing(8)
        
        # 認證方式單選按鈕容器（水平）
        rb_layout = QHBoxLayout()
        rb_layout.setContentsMargins(0, 0, 0, 0)
        rb_layout.setSpacing(15)
        self.rb_anonymous = QRadioButton("Anonymous")
        self.rb_username = QRadioButton("Username and Password")
        self.rb_certificate = QRadioButton("Certificate and Private Key")
        rb_layout.addWidget(self.rb_anonymous)
        rb_layout.addWidget(self.rb_username)
        rb_layout.addWidget(self.rb_certificate)

        auth_group_layout.addLayout(rb_layout)

        # 認證欄位（username/password）
        cred_layout = QGridLayout()
        cred_layout.setSpacing(8)
        cred_layout.setContentsMargins(0, 0, 0, 0)
        self.username_label = QLabel("Username:")
        cred_layout.addWidget(self.username_label, 0, 0)
        self.username_edit = QLineEdit()
        cred_layout.addWidget(self.username_edit, 0, 1)
        self.password_label = QLabel("Password:")
        cred_layout.addWidget(self.password_label, 0, 2)
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        cred_layout.addWidget(self.password_edit, 0, 3)

        # 憑證欄位（檔案選取）
        cert_layout = QGridLayout()
        cert_layout.setSpacing(8)
        cert_layout.setContentsMargins(0, 0, 0, 0)
        self.client_cert_label = QLabel("Client Cert:")
        cert_layout.addWidget(self.client_cert_label, 0, 0)
        self.client_cert_edit = QLineEdit()
        self.cert_browse_btn1 = QPushButton("Browse")
        self.cert_browse_btn1.clicked.connect(lambda: self._browse_file(self.client_cert_edit))
        cert_layout.addWidget(self.client_cert_edit, 0, 1)
        cert_layout.addWidget(self.cert_browse_btn1, 0, 2)

        self.client_key_label = QLabel("Private Key:")
        cert_layout.addWidget(self.client_key_label, 1, 0)
        self.client_key_edit = QLineEdit()
        self.cert_browse_btn2 = QPushButton("Browse")
        self.cert_browse_btn2.clicked.connect(lambda: self._browse_file(self.client_key_edit))
        cert_layout.addWidget(self.client_key_edit, 1, 1)
        cert_layout.addWidget(self.cert_browse_btn2, 1, 2)

        auth_group_layout.addLayout(cred_layout)
        auth_group_layout.addLayout(cert_layout)

        # 右側按鈕（Apply / Cancel）- 固定大小，頂部對齐
        btns_layout = QVBoxLayout()
        btns_layout.setContentsMargins(0, 0, 0, 0)
        btns_layout.setSpacing(8)
        self.apply_btn = QPushButton("Apply")
        self.apply_btn.setFixedSize(90, 40)
        self.apply_btn.clicked.connect(self._on_apply)
        btns_layout.addWidget(self.apply_btn, alignment=Qt.AlignmentFlag.AlignTop)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setFixedSize(90, 40)
        cancel_btn.clicked.connect(self.reject)
        btns_layout.addWidget(cancel_btn, alignment=Qt.AlignmentFlag.AlignTop)
        btns_layout.addStretch()

        auth_top_layout.addWidget(auth_group, 1)
        # 按鈕布局容器，頂部對齐
        btns_container = QWidget()
        btns_container_layout = QVBoxLayout(btns_container)
        btns_container_layout.setContentsMargins(0, 0, 0, 0)
        btns_container_layout.addLayout(btns_layout)
        btns_container_layout.addStretch()
        auth_top_layout.addWidget(btns_container, 0)

        layout.addLayout(auth_top_layout)

        # 安全設定（中段）：左為 Security Mode，右為 Security Policy
        sec_mid_layout = QHBoxLayout()
        sec_mid_layout.setSpacing(12)
        sec_mid_layout.setContentsMargins(0, 0, 0, 0)

        # Security Mode
        mode_group = QGroupBox("Security Mode")
        mode_layout = QVBoxLayout(mode_group)
        mode_layout.setContentsMargins(10, 12, 10, 12)
        mode_layout.setSpacing(8)
        self.rb_mode_none = QRadioButton("None")
        self.rb_mode_sign = QRadioButton("Sign")
        self.rb_mode_sign_encrypt = QRadioButton("Sign & Encrypt")
        mode_layout.addWidget(self.rb_mode_none)
        mode_layout.addWidget(self.rb_mode_sign)
        mode_layout.addWidget(self.rb_mode_sign_encrypt)
        # 添加間距以與 Security Policy 對齊
        mode_layout.addStretch()

        # Security Policy list
        policy_group = QGroupBox("Security Policy")
        policy_layout = QVBoxLayout(policy_group)
        policy_layout.setContentsMargins(10, 12, 10, 12)
        policy_layout.setSpacing(8)
        self.policy_rb_none = QRadioButton("None")
        self.policy_rb_basic128 = QRadioButton("Basic128RSA15")
        self.policy_rb_basic256 = QRadioButton("Basic256")
        self.policy_rb_basic256sha = QRadioButton("Basic256Sha256")
        policy_layout.addWidget(self.policy_rb_none)
        policy_layout.addWidget(self.policy_rb_basic128)
        policy_layout.addWidget(self.policy_rb_basic256)
        policy_layout.addWidget(self.policy_rb_basic256sha)

        self.chk_show_supported = QCheckBox("Show only modes that are supported by the server")
        policy_layout.addWidget(self.chk_show_supported)

        sec_mid_layout.addWidget(mode_group, 1)
        sec_mid_layout.addWidget(policy_group, 1)

        layout.addLayout(sec_mid_layout)
        
        # 連線設定（下方）
        connection_group = QGroupBox("連線設定")
        connection_layout = QHBoxLayout(connection_group)
        connection_layout.setContentsMargins(10, 8, 10, 8)
        connection_layout.setSpacing(8)
        connection_layout.addWidget(QLabel("連線超時 (秒):"))
        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(1, 300)
        self.timeout_spin.setValue(5)
        self.timeout_spin.setFixedWidth(80)
        connection_layout.addWidget(self.timeout_spin)
        connection_layout.addWidget(QLabel("寫值重試延遲 (秒):"))
        self.write_timeout_spin = QSpinBox()
        self.write_timeout_spin.setRange(1, 60)
        self.write_timeout_spin.setValue(3)
        self.write_timeout_spin.setFixedWidth(80)
        connection_layout.addWidget(self.write_timeout_spin)
        connection_layout.addStretch()
        layout.addWidget(connection_group)

    def apply_style(self):
        """套用樣式，根據父視窗主題選擇亮色或暗色"""
        is_dark = False
        parent = self.parent()
        
        # 嘗試找到主視窗以取得主題設定
        current_window = parent
        while current_window:
            if hasattr(current_window, "current_theme"):
                if current_window.current_theme == "dark":
                    is_dark = True
                elif current_window.current_theme == "system":
                    if hasattr(current_window, "is_system_dark_mode"):
                        is_dark = current_window.is_system_dark_mode()
                break
            current_window = current_window.parent() if hasattr(current_window, "parent") else None

        if is_dark:
            self._apply_dark_style()
        else:
            self._apply_light_style()

    def _apply_light_style(self):
        """套用亮色樣式"""
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #d0d0d0;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 12px;
                background-color: white;
            }
            QGroupBox::title {
                color: #2c3e50;
            }
            QLabel {
                color: #333;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QLineEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 6px;
                background-color: white;
                color: #333;
            }
            QLineEdit:focus {
                border: 2px solid #0078d4;
            }
            QSpinBox {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 6px;
                background-color: white;
                color: #333;
            }
            QSpinBox:focus {
                border: 2px solid #0078d4;
            }
            QRadioButton {
                color: #333;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #999999;
                border-radius: 9px;
                background-color: white;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0078d4;
                background-color: #f0f0f0;
            }
            QRadioButton::indicator:checked {
                background-color: #0078d4;
                border: 2px solid #0078d4;
            }
            QCheckBox {
                color: #333;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #999999;
                border-radius: 2px;
                background-color: white;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #f0f0f0;
            }
            QCheckBox::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
                image: url(:/checkbox_check);
            }
        """)

    def _apply_dark_style(self):
        """套用暗色樣式"""
        self.setStyleSheet("""
            QDialog {
                background-color: #2b2b2b;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #3d3d3d;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 12px;
                background-color: #363636;
            }
            QGroupBox::title {
                color: #ffffff;
            }
            QLabel {
                color: #cccccc;
            }
            QPushButton {
                background-color: #0e639c;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1177bb;
            }
            QLineEdit {
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                background-color: #1e1e1e;
                color: #cccccc;
            }
            QLineEdit:focus {
                border: 2px solid #0e639c;
            }
            QSpinBox {
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                background-color: #1e1e1e;
                color: #cccccc;
            }
            QSpinBox:focus {
                border: 2px solid #0e639c;
            }
            QRadioButton {
                color: #cccccc;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #666666;
                border-radius: 9px;
                background-color: #1e1e1e;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #252526;
            }
            QRadioButton::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
            }
            QCheckBox {
                color: #cccccc;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #666666;
                border-radius: 2px;
                background-color: #1e1e1e;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #252526;
            }
            QCheckBox::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
                image: url(:/checkbox_check);
            }
        """)

    def load_data(self):
        """載入現有的 OPC 連線設定"""
        # 設定安全策略單選按鈕（默認為 None）
        policy_mapping = {
            "None": self.policy_rb_none,
            "Basic128Rsa15": self.policy_rb_basic128,
            "Basic256": self.policy_rb_basic256,
            "Basic256Sha256": self.policy_rb_basic256sha,
        }
        policy_btn = policy_mapping.get(self.security_policy, self.policy_rb_none)
        policy_btn.setChecked(True)

        # 設定使用者認證相關欄位（默認為 Anonymous）
        if self.username or self.password:
            self.rb_username.setChecked(True)
        else:
            self.rb_anonymous.setChecked(True)

        self.username_edit.setText(self.username)
        self.password_edit.setText(self.password)
        self.timeout_spin.setValue(self.timeout)
        self.write_timeout_spin.setValue(self.write_timeout)

        # 初始化安全模式
        mode_mapping = {
            "None": self.rb_mode_none,
            "Sign": self.rb_mode_sign,
            "SignAndEncrypt": self.rb_mode_sign_encrypt,
        }
        mode_btn = mode_mapping.get(self.security_mode, self.rb_mode_none)
        mode_btn.setChecked(True)
        
        # 初始化認證欄位可見性
        self.on_auth_mode_changed()

        # 啟用「只顯示伺服器支援的模式」複選框並觸發檢測
        self.chk_show_supported.setChecked(True)

    def get_settings(self) -> Dict[str, Any]:
        """取得設定值"""
        # security policy
        if self.policy_rb_none.isChecked():
            policy = "None"
        elif self.policy_rb_basic128.isChecked():
            policy = "Basic128Rsa15"
        elif self.policy_rb_basic256.isChecked():
            policy = "Basic256"
        else:
            policy = "Basic256Sha256"

        # security mode
        if self.rb_mode_none.isChecked():
            mode = "None"
        elif self.rb_mode_sign.isChecked():
            mode = "Sign"
        else:
            mode = "SignAndEncrypt"

        # auth method
        if self.rb_anonymous.isChecked():
            auth = "Anonymous"
        elif self.rb_username.isChecked():
            auth = "Username"
        else:
            auth = "Certificate"

        return {
            "security_policy": policy,
            "security_mode": mode,
            "auth_method": auth,
            "username": self.username_edit.text(),
            "password": self.password_edit.text(),
            "client_cert": self.client_cert_edit.text(),
            "client_key": self.client_key_edit.text(),
            "timeout": self.timeout_spin.value(),
            "write_timeout": self.write_timeout_spin.value(),
            "show_only_supported": self.chk_show_supported.isChecked(),
        }

    def _browse_file(self, line_edit: QLineEdit):
        path, _ = QFileDialog.getOpenFileName(self, "Select file")
        if path:
            line_edit.setText(path)

    def on_auth_mode_changed(self):
        # 當選擇 Anonymous 時隱藏/停用其他認證欄位
        is_anonymous = self.rb_anonymous.isChecked()
        is_username = self.rb_username.isChecked()

        # Username/password visible only when username radio selected
        self.username_label.setVisible(is_username)
        self.username_edit.setVisible(is_username)
        self.password_label.setVisible(is_username)
        self.password_edit.setVisible(is_username)

        # Certificate fields visible only when certificate radio selected
        cert_visible = self.rb_certificate.isChecked()
        self.client_cert_label.setVisible(cert_visible)
        self.client_cert_edit.setVisible(cert_visible)
        self.cert_browse_btn1.setVisible(cert_visible)
        self.client_key_label.setVisible(cert_visible)
        self.client_key_edit.setVisible(cert_visible)
        self.cert_browse_btn2.setVisible(cert_visible)

        # 調整視窗大小以適應顯示的欄位
        self.adjustSize()

    def on_chk_show_supported_toggled(self, checked: bool):
        """檢測或隱藏伺服器不支援的安全模式"""
        if not checked:
            # 取消勾選時顯示所有模式
            self._show_all_policies_and_modes()
            return

        # 需要 OPC URL 才能進行檢測 - 無 OPC URL 時不進行檢測
        if not self.opc_url:
            return

        # 禁用複選框並開始檢測
        self.chk_show_supported.setEnabled(False)

        # 在非同步任務中執行檢測
        import asyncio
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                asyncio.create_task(self._detect_server_capabilities(self.opc_url))
            else:
                asyncio.run(self._detect_server_capabilities(self.opc_url))
        except RuntimeError:
            asyncio.run(self._detect_server_capabilities(self.opc_url))

    async def _detect_server_capabilities(self, opc_url: str):
        """偵測伺服器支援的安全策略和模式
        
        透過連接伺服器並解析日誌與異常信息來獲取支援的策略
        """
        supported_policies = set()
        supported_modes = set()
        error_message = None
        
        try:
            from asyncua import Client
            from asyncua.ua.uaerrors import UaError
            import logging
            import io
            
            # 攔截 asyncua 日誌以提取端點信息
            log_capture = io.StringIO()
            handler = logging.StreamHandler(log_capture)
            handler.setFormatter(logging.Formatter('%(message)s'))
            asyncua_logger = logging.getLogger('asyncua.client.client')
            asyncua_logger.addHandler(handler)
            
            try:
                client = Client(opc_url)
                await client.connect()
                endpoints = await client.get_endpoints()
                await client.disconnect()
                
                # 成功連接並獲取端點
                for idx, ep in enumerate(endpoints):
                    # 提取安全策略
                    uri = getattr(ep, "SecurityPolicyUri", None)
                    if uri:
                        frag = uri.split("#")[-1] if "#" in str(uri) else str(uri).rstrip('/').split('/')[-1]
                        norm = self._normalize_policy_name(frag)
                        if norm and norm != "":
                            supported_policies.add(norm)
                            print(f"[OPC UA 檢測] 策略: {norm}")
                    
                    # 提取安全模式
                    mode = getattr(ep, "SecurityMode", None)
                    if mode is not None:
                        name = getattr(mode, "name", None) or str(mode)
                        norm_mode = self._normalize_mode_name(str(name))
                        if norm_mode and norm_mode != "":
                            supported_modes.add(norm_mode)
                            print(f"[OPC UA 檢測] 模式: {norm_mode}")
                            
            except UaError as ua_exc:
                # 連接失敗，嘗試從日誌和異常消息中提取信息
                error_message = str(ua_exc)
                print(f"[OPC UA 檢測] 連接異常: {error_message}")
                
                # 從日誌中提取所有 SecurityPolicyUri
                log_content = log_capture.getvalue()
                print(f"[OPC UA 檢測] 日誌內容長度: {len(log_content)}")
                
                # 使用正則表達式查找所有 SecurityPolicyUri
                import re
                uri_matches = re.findall(r"SecurityPolicyUri='([^']+)'", log_content)
                print(f"[OPC UA 檢測] 從日誌找到 {len(uri_matches)} 個策略 URI: {set(uri_matches)}")
                
                for uri in uri_matches:
                    frag = uri.split("#")[-1] if "#" in uri else uri.rstrip('/').split('/')[-1]
                    norm = self._normalize_policy_name(frag)
                    if norm and norm != "" and norm != "None":  # 除外 None 可能只是認證策略
                        supported_policies.add(norm)
                        print(f"[OPC UA 檢測] 從日誌提取策略: {norm}")
                
                # 同時從日誌中提取 SecurityMode
                # 格式通常是: SecurityMode=<MessageSecurityMode.Sign: 2> 或 SecurityMode=<MessageSecurityMode.SignAndEncrypt: 3>
                mode_matches = re.findall(r"SecurityMode=<MessageSecurityMode\.(\w+(?:And\w+)?)", log_content)
                print(f"[OPC UA 檢測] 從日誌找到 {len(mode_matches)} 個安全模式: {set(mode_matches)}")
                
                for mode_str in set(mode_matches):  # 使用 set 避免重複
                    norm_mode = self._normalize_mode_name(mode_str)
                    if norm_mode and norm_mode != "":
                        supported_modes.add(norm_mode)
                        print(f"[OPC UA 檢測] 從日誌提取模式: {norm_mode}")

                
                # 如果還是沒有找到策略，至少報告所有找到的內容
                if not supported_policies:
                    print(f"[OPC UA 檢測] 警告: 未找到有效的安全策略")
                    
            finally:
                asyncua_logger.removeHandler(handler)
                log_capture.close()
                
        except Exception as exc:
            error_message = str(exc)
            print(f"[OPC UA 檢測] 未知錯誤: {error_message}")

        # 更新 UI
        def update_ui():
            self.chk_show_supported.setEnabled(True)
            
            if supported_policies or supported_modes:
                self._detected_supported = {"policies": supported_policies, "modes": supported_modes}
                self._apply_supported_filters()
            else:
                self._show_all_policies_and_modes()

        QTimer.singleShot(0, update_ui)
    def _apply_supported_filters(self):
        """套用伺服器支援的安全模式過濾"""
        data = self._detected_supported or {}
        policies = data.get("policies", set())
        modes = data.get("modes", set())

        # 只有當檢測到支援的模式時才進行過濾
        # 如果檢測到空集合，則顯示所有（伺服器可能不報告這些資訊）
        if not policies and not modes:
            self._show_all_policies_and_modes()
            return

        # 控制安全策略單選按鈕的可見性
        policy_buttons = [
            (self.policy_rb_none, "None"),
            (self.policy_rb_basic128, "Basic128Rsa15"),
            (self.policy_rb_basic256, "Basic256"),
            (self.policy_rb_basic256sha, "Basic256Sha256"),
        ]
        for btn, policy_name in policy_buttons:
            btn.setVisible(policy_name in policies)

        # 控制安全模式單選按鈕的可見性
        mode_buttons = [
            (self.rb_mode_none, "None"),
            (self.rb_mode_sign, "Sign"),
            (self.rb_mode_sign_encrypt, "SignAndEncrypt"),
        ]
        for btn, mode_name in mode_buttons:
            btn.setVisible(mode_name in modes)

    def _show_all_policies_and_modes(self):
        """顯示所有安全策略和模式（沒有過濾）"""
        # 顯示所有安全策略
        for btn in [self.policy_rb_none, self.policy_rb_basic128, 
                    self.policy_rb_basic256, self.policy_rb_basic256sha]:
            btn.setVisible(True)
        
        # 顯示所有安全模式
        for btn in [self.rb_mode_none, self.rb_mode_sign, self.rb_mode_sign_encrypt]:
            btn.setVisible(True)

    def _normalize_policy_name(self, fragment: str) -> str:
        """將各種 SecurityPolicy 片段標準化為 UI 使用的規範名稱
        
        例： 
            Basic128RSA15 -> Basic128Rsa15
            Basic256Sha256 -> Basic256Sha256
            None -> None
        """
        if not fragment:
            return ""
        
        # 轉為小寫並移除非字母數字字元進行比對
        normalized = re.sub(r'[^0-9a-z]', '', fragment.lower())
        
        # 根據關鍵字識別策略
        if normalized == "none":
            return "None"
        if "128" in normalized:
            return "Basic128Rsa15"
        if "sha256" in normalized or ("sha" in normalized and "256" in normalized):
            return "Basic256Sha256"
        if "256" in normalized and "128" not in normalized:
            return "Basic256"
        
        # 無法識別時，嘗試返回原始片段
        return fragment if fragment in ["None", "Basic128Rsa15", "Basic256", "Basic256Sha256"] else ""

    def _normalize_mode_name(self, name: str) -> str:
        """將安全模式名稱標準化為規範鍵值: None, Sign, SignAndEncrypt"""
        if not name:
            return ""
        
        # 移除前綴和特殊字元，轉為小寫進行比對
        cleaned = name.lower()
        cleaned = cleaned.replace("messagesecuritymode.", "")
        cleaned = cleaned.replace("_", "").replace(" ", "")
        
        # 識別模式（注意順序，SignAndEncrypt 要在 Sign 之前）
        if "signandencrypt" in cleaned or "signencrypt" in cleaned:
            return "SignAndEncrypt"
        if "sign" in cleaned and "encrypt" not in cleaned:
            return "Sign"
        if "none" in cleaned:
            return "None"
        
        # 無法識別
        return ""

    def _on_apply(self):
        """按下 Apply 時接受設定對話框"""
        self.accept()


class ScheduleEditDialog(QDialog):
    """排程編輯對話框"""

    def __init__(self, parent=None, schedule: Dict[str, Any] = None):
        super().__init__(parent)
        self.schedule = schedule
        self.original_rrule = ""  # 儲存原始的 RRULE 字串

        # 取得資料庫管理器
        self.db_manager = parent.db_manager if parent and hasattr(parent, 'db_manager') else None

        # 初始化OPC設定
        self.opc_security_policy = schedule.get("opc_security_policy", "None") if schedule else "None"
        self.opc_security_mode = schedule.get("opc_security_mode", "None") if schedule else "None"
        self.opc_username = schedule.get("opc_username", "") if schedule else ""
        self.opc_password = schedule.get("opc_password", "") if schedule else ""
        self.opc_timeout = schedule.get("opc_timeout", 5) if schedule else 5
        self.opc_write_timeout = schedule.get("opc_write_timeout", 3) if schedule else 3
        self.is_enabled = schedule.get("is_enabled", 1) if schedule else 1

        self.setup_ui()
        self.apply_style()

        # 如果是新增模式，設置預設任務名稱
        if not schedule and parent and hasattr(parent, 'db_manager'):
            default_name = parent.db_manager.get_next_task_name()
            self.task_name_edit.setText(default_name)

        if schedule:
            self.load_data()

    def setup_ui(self):
        self.setWindowTitle("編輯排程" if self.schedule else "新增排程")
        self.setWindowIcon(get_app_icon())
        self.setMinimumWidth(500)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # 基本資訊
        basic_group = QGroupBox("基本資訊")
        basic_layout = QGridLayout(basic_group)

        basic_layout.addWidget(QLabel("任務名稱:"), 0, 0)
        self.task_name_edit = QLineEdit()
        self.task_name_edit.setPlaceholderText("例如：每日早班開機")
        basic_layout.addWidget(self.task_name_edit, 0, 1)

        basic_layout.addWidget(QLabel("OPC URL:"), 1, 0)
        opc_url_layout = QHBoxLayout()
        opc_url_layout.setSpacing(5)
        self.opc_url_edit = QLineEdit()
        self.opc_url_edit.setPlaceholderText("localhost:4840")
        opc_url_layout.addWidget(self.opc_url_edit)
        # 添加協議標籤顯示
        self.opc_protocol_label = QLabel("opc.tcp://")
        self.opc_protocol_label.setStyleSheet("color: #666; padding-right: 5px;")
        opc_url_layout.insertWidget(0, self.opc_protocol_label)
        # 添加OPC設定按鈕
        self.btn_opc_settings = QPushButton("設定...")
        self.btn_opc_settings.setToolTip("OPC UA 連線設定")
        self.btn_opc_settings.clicked.connect(self.configure_opc_settings)
        self.btn_opc_settings.setMaximumWidth(80)
        opc_url_layout.addWidget(self.btn_opc_settings)
        basic_layout.addLayout(opc_url_layout, 1, 1)

        basic_layout.addWidget(QLabel("Node ID:"), 2, 0)
        node_id_layout = QHBoxLayout()
        node_id_layout.setSpacing(5)
        self.node_id_edit = QLineEdit()
        self.node_id_edit.setPlaceholderText("ns=2;i=1001")
        node_id_layout.addWidget(self.node_id_edit)
        # 添加瀏覽按鈕
        self.btn_browse_node = QPushButton("瀏覽...")
        self.btn_browse_node.setToolTip("瀏覽 OPC UA 節點")
        self.btn_browse_node.clicked.connect(self.browse_opcua_nodes)
        self.btn_browse_node.setMaximumWidth(80)
        node_id_layout.addWidget(self.btn_browse_node)
        basic_layout.addLayout(node_id_layout, 2, 1)

        basic_layout.addWidget(QLabel("目標值:"), 3, 0)
        target_layout = QHBoxLayout()
        self.target_value_edit = QLineEdit()
        self.target_value_edit.setPlaceholderText("1")
        target_layout.addWidget(self.target_value_edit)
        basic_layout.addLayout(target_layout, 3, 1)
        
        # 型別顯示 - 簡單文字標籤
        basic_layout.addWidget(QLabel("型別:"), 4, 0)
        type_layout = QHBoxLayout()
        self.data_type_label = QLabel("未偵測")
        self.data_type_label.setStyleSheet("font-weight: bold; color: #0066cc;")
        type_layout.addWidget(self.data_type_label)
        type_layout.addStretch()
        basic_layout.addLayout(type_layout, 4, 1)

        # Category 選單
        basic_layout.addWidget(QLabel("Category:"), 5, 0)
        category_layout = QHBoxLayout()
        self.category_combo = QComboBox()
        self.category_combo.setMinimumWidth(200)
        category_layout.addWidget(self.category_combo)
        category_layout.addStretch()
        basic_layout.addLayout(category_layout, 5, 1)
        
        # 載入 categories
        self._load_categories()

        basic_layout.addWidget(QLabel("狀態:"), 6, 0)
        self.enabled_checkbox = QCheckBox("啟用排程")
        self.enabled_checkbox.setChecked(True)  # 預設啟用
        self.enabled_checkbox.setToolTip("控制此排程是否會被執行")
        basic_layout.addWidget(self.enabled_checkbox, 6, 1)

        layout.addWidget(basic_group)

        # 週期設定
        recurrence_group = QGroupBox("週期設定")
        recurrence_layout = QVBoxLayout(recurrence_group)

        self.rrule_display = QLineEdit()
        self.rrule_display.setReadOnly(True)
        self.rrule_display.setPlaceholderText("點擊下方按鈕設定週期規則")
        recurrence_layout.addWidget(self.rrule_display)

        self.btn_edit_recurrence = QPushButton("設定週期規則...")
        self.btn_edit_recurrence.clicked.connect(self.edit_recurrence)
        recurrence_layout.addWidget(self.btn_edit_recurrence)

        layout.addWidget(recurrence_group)

        # 按鈕
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        ok_btn = QPushButton("確定")
        ok_btn.setDefault(True)
        ok_btn.clicked.connect(self.on_ok_clicked)
        button_layout.addWidget(ok_btn)

        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)

    def apply_style(self):
        """套用樣式，根據父視窗主題選擇亮色或暗色"""
        # 判斷是否使用暗色模式
        is_dark = False
        parent = self.parent()
        if parent and hasattr(parent, "current_theme"):
            if parent.current_theme == "dark":
                is_dark = True
            elif parent.current_theme == "system":
                if hasattr(parent, "is_system_dark_mode"):
                    is_dark = parent.is_system_dark_mode()

        if is_dark:
            self._apply_dark_style()
        else:
            self._apply_light_style()

    def _apply_light_style(self):
        """套用亮色樣式"""
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #d0d0d0;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 12px;
                background-color: white;
            }
            QGroupBox::title {
                color: #2c3e50;
            }
            QLabel {
                color: #333;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QLineEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 6px;
                background-color: white;
                color: #333;
            }
            QLineEdit:focus {
                border: 2px solid #0078d4;
            }
            QRadioButton {
                color: #333;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #999999;
                border-radius: 9px;
                background-color: white;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0078d4;
                background-color: #f0f0f0;
            }
            QRadioButton::indicator:checked {
                background-color: #0078d4;
                border: 2px solid #0078d4;
            }
            QCheckBox {
                color: #333;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #999999;
                border-radius: 2px;
                background-color: white;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #f0f0f0;
            }
            QCheckBox::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
                image: url(:/checkbox_check);
            }
        """)

    def _apply_dark_style(self):
        """套用暗色樣式"""
        self.setStyleSheet("""
            QDialog {
                background-color: #2b2b2b;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #3d3d3d;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 12px;
                background-color: #363636;
            }
            QGroupBox::title {
                color: #ffffff;
            }
            QLabel {
                color: #cccccc;
            }
            QPushButton {
                background-color: #0e639c;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1177bb;
            }
            QLineEdit {
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                padding: 6px;
                background-color: #1e1e1e;
                color: #cccccc;
            }
            QLineEdit:focus {
                border: 2px solid #0e639c;
            }
            QRadioButton {
                color: #cccccc;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #666666;
                border-radius: 9px;
                background-color: #1e1e1e;
            }
            QRadioButton::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #252526;
            }
            QRadioButton::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
            }
            QCheckBox {
                color: #cccccc;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #666666;
                border-radius: 2px;
                background-color: #1e1e1e;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #0e639c;
                background-color: #252526;
            }
            QCheckBox::indicator:checked {
                background-color: #0e639c;
                border: 2px solid #0e639c;
                image: url(:/checkbox_check);
            }
        """)

    def _load_categories(self):
        """載入 Category 清單到下拉選單"""
        self.category_combo.clear()
        
        if not self.db_manager:
            # 如果沒有資料庫管理器，只加入預設項目
            self.category_combo.addItem("Red (關閉)", 1)
            return
        
        try:
            categories = self.db_manager.get_all_categories()
            for cat in categories:
                # 顯示名稱加上顏色預覽 (用色塊符號)
                display_name = f"{cat['name']}"
                self.category_combo.addItem(display_name, cat['id'])
            
            # 預設選擇第一個 (Red)
            self.category_combo.setCurrentIndex(0)
        except Exception as e:
            print(f"載入 Categories 失敗: {e}")
            # 失敗時加入預設項目
            self.category_combo.addItem("Red (關閉)", 1)

    def load_data(self):
        """載入現有資料"""
        self.task_name_edit.setText(self.schedule.get("task_name", ""))
        # 提取 ip:port 部分（去掉 opc.tcp:// 前綴）
        opc_url = self.schedule.get("opc_url", "")
        if opc_url.startswith("opc.tcp://"):
            opc_url = opc_url[10:]  # 去掉 "opc.tcp://"
        self.opc_url_edit.setText(opc_url)
        self.node_id_edit.setText(self.schedule.get("node_id", ""))
        self.target_value_edit.setText(self.schedule.get("target_value", ""))
        data_type = self.schedule.get("data_type", "auto")
        # 如果是"auto"，顯示為"未偵測"
        display_data_type = "未偵測" if data_type == "auto" else data_type
        self.data_type_label.setText(display_data_type)
        
        # 載入 Category
        category_id = self.schedule.get("category_id", 1)
        for i in range(self.category_combo.count()):
            if self.category_combo.itemData(i) == category_id:
                self.category_combo.setCurrentIndex(i)
                break
        
        # 儲存原始 RRULE 字串，並顯示格式化的描述
        self.original_rrule = self.schedule.get("rrule_str", "")
        # 暫時直接顯示原始 RRULE，稍後可以改進為格式化顯示
        self.rrule_display.setText(self.original_rrule if self.original_rrule else "未設定")
        
        self.enabled_checkbox.setChecked(bool(self.schedule.get("is_enabled", 1)))

    def edit_recurrence(self):
        """編輯週期規則"""
        current_rrule = self.original_rrule
        rrule = show_recurrence_dialog(self, current_rrule)
        if rrule:
            self.original_rrule = rrule
            # 暫時直接顯示原始 RRULE，稍後可以改進為格式化顯示
            self.rrule_display.setText(rrule if rrule else "未設定")

    def get_data(self) -> Dict[str, Any]:
        """取得編輯的資料"""
        # 自動添加 opc.tcp:// 前綴
        opc_url = self.opc_url_edit.text().strip()
        if opc_url and not opc_url.startswith("opc.tcp://"):
            opc_url = f"opc.tcp://{opc_url}"
        
        # 取得選擇的 category_id
        category_id = self.category_combo.currentData()
        if category_id is None:
            category_id = 1  # 預設 Red

        return {
            "task_name": self.task_name_edit.text(),
            "opc_url": opc_url,
            "node_id": self.node_id_edit.text(),
            "target_value": self.target_value_edit.text(),
            # 處理資料型別：如果顯示"未偵測"，儲存為"auto"
            "data_type": "auto" if self.data_type_label.text() == "未偵測" else self.data_type_label.text(),
            "rrule_str": self.original_rrule,
            "category_id": category_id,
            "opc_security_policy": self.opc_security_policy,
            "opc_security_mode": self.opc_security_mode,
            "opc_username": self.opc_username,
            "opc_password": self.opc_password,
            "opc_timeout": self.opc_timeout,
            "opc_write_timeout": self.opc_write_timeout,
            "is_enabled": 1 if self.enabled_checkbox.isChecked() else 0,
        }

    def on_ok_clicked(self):
        """確定按鈕點擊處理"""
        # 檢查任務名稱
        task_name = self.task_name_edit.text().strip()
        if not task_name:
            QMessageBox.warning(
                self,
                "任務名稱未設定",
                "請輸入任務名稱。",
            )
            return

        # 檢查 OPC URL
        opc_url = self.opc_url_edit.text().strip()
        if not opc_url:
            QMessageBox.warning(
                self,
                "OPC URL 未設定",
                "請輸入 OPC URL。",
            )
            return

        # 檢查 Node ID
        node_id = self.node_id_edit.text().strip()
        if not node_id:
            QMessageBox.warning(
                self,
                "Node ID 未設定",
                "請輸入或瀏覽選擇 Node ID。",
            )
            return

        # 檢查目標值
        target_value = self.target_value_edit.text().strip()
        if not target_value:
            QMessageBox.warning(
                self,
                "目標值未設定",
                "請輸入目標值。",
            )
            return

        # 檢查 rrule 是否為空
        rrule_str = self.rrule_display.text().strip()
        if not rrule_str:
            QMessageBox.warning(
                self,
                "週期規則未設定",
                "請設定排程的週期規則，無法儲存空的週期規則。",
            )
            return

        # 如果檢查通過，接受對話框
        self.accept()

    def browse_opcua_nodes(self):
        """瀏覽 OPC UA 節點"""
        opc_url = self._normalize_opc_url()

        if not opc_url:
            QMessageBox.warning(
                self,
                "警告",
                "請先輸入 OPC URL (IP:Port)",
            )
            return

        # 開啟節點瀏覽對話框
        dialog = OPCNodeBrowserDialog(self, opc_url)
        if dialog.exec() == QDialog.Accepted:
            selected_node = dialog.get_selected_node()
            if selected_node:
                # 解析選擇的節點資訊: display_name|node_id|data_type
                parts = selected_node.split("|")
                if len(parts) >= 2:
                    display_name = parts[0]
                    node_id = parts[1]
                    data_type = parts[2] if len(parts) > 2 else "未知"
                    
                    # 設定節點 ID 和自動偵測的資料型別
                    self.node_id_edit.setText(f"{display_name}|{node_id}")
                    self.data_type_label.setText(data_type)

    def configure_opc_settings(self):
        """設定 OPC UA 連線參數"""
        opc_url = self._normalize_opc_url()

        dialog = OPCSettingsDialog(
            self,
            self.opc_security_policy,
            self.opc_username,
            self.opc_password,
            self.opc_timeout,
            self.opc_write_timeout,
            self.opc_security_mode,
            opc_url=opc_url,
        )

        if dialog.exec() == QDialog.Accepted:
            settings = dialog.get_settings()
            self.opc_security_policy = settings["security_policy"]
            self.opc_security_mode = settings["security_mode"]
            self.opc_username = settings["username"]
            self.opc_password = settings["password"]
            self.opc_timeout = settings["timeout"]
            self.opc_write_timeout = settings["write_timeout"]

    def _normalize_opc_url(self) -> str:
        """標準化 OPC URL，確保以 opc.tcp:// 開頭"""
        url = self.opc_url_edit.text().strip()
        if not url:
            return ""

        # 如果沒有協議前綴，添加預設的 opc.tcp://
        if not url.startswith(("opc.tcp://", "opc.https://", "opc.wss://")):
            url = f"opc.tcp://{url}"

        return url


def main():
    """主程式進入點"""
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # 設定應用程式資訊
    app.setApplicationName("CalendarUA")
    app.setApplicationVersion("1.0.0")

    # 建立事件迴圈
    loop = qasync.QEventLoop(app)
    asyncio.set_event_loop(loop)

    # 建立主視窗
    window = CalendarUA()
    window.show()

    # 執行事件迴圈
    with loop:
        loop.run_forever()


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
CalendarUA - 工業自動化排程管理系統主程式
採用 PySide6 開發，結合 Office/Outlook 風格行事曆介面
"""

import sys
import asyncio
from datetime import datetime, timedelta
from typing import Optional, Dict, Any, List

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
)
from PySide6.QtCore import Qt, QTimer, Signal, Slot, QThread
from PySide6.QtGui import QAction, QIcon, QColor, QFont
import qasync
import re

from database.sqlite_manager import SQLiteManager
from core.opc_handler import OPCHandler
from core.rrule_parser import RRuleParser
from ui.recurrence_dialog import RecurrenceDialog, show_recurrence_dialog


class SchedulerWorker(QThread):
    """背景排程工作執行緒"""

    trigger_task = Signal(dict)

    def __init__(self, db_manager: SQLiteManager, check_interval: int = 30):
        super().__init__()
        self.db_manager = db_manager
        self.check_interval = check_interval
        self.running = True
        self.last_check = datetime.now()

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
        rrule_str = schedule.get("rrule_str", "")
        if not rrule_str:
            return False

        # 使用 RRuleParser 檢查是否為觸發時間
        return RRuleParser.is_trigger_time(
            rrule_str, current_time, tolerance_seconds=30
        )

    def stop(self):
        """停止工作執行緒"""
        self.running = False
        self.wait(2000)


class CalendarUA(QMainWindow):
    """CalendarUA 主視窗"""

    def __init__(self):
        super().__init__()

        self.db_manager: Optional[SQLiteManager] = None
        self.opc_handler: Optional[OPCHandler] = None
        self.scheduler_worker: Optional[SchedulerWorker] = None
        self.schedules: List[Dict[str, Any]] = []

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
        self.setWindowTitle("CalendarUA - 工業自動化排程管理系統")
        self.setMinimumSize(1200, 800)

        # 建立中央widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # 建立分割器
        splitter = QSplitter(Qt.Horizontal)

        # 左側：日曆視圖
        left_panel = self.create_left_panel()
        splitter.addWidget(left_panel)

        # 右側：排程列表與控制
        right_panel = self.create_right_panel()
        splitter.addWidget(right_panel)

        # 設定分割比例
        splitter.setSizes([400, 800])

        main_layout.addWidget(splitter)

        # 建立選單列
        self.create_menu_bar()

        # 建立工具列
        self.create_tool_bar()

        # 建立狀態列
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就緒")

    def create_left_panel(self) -> QWidget:
        """建立左側日曆面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        # 日曆widget
        self.calendar = QCalendarWidget()
        self.calendar.setGridVisible(True)
        self.calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.calendar.setHorizontalHeaderFormat(QCalendarWidget.SingleLetterDayNames)

        # 設定日曆樣式
        self.calendar.setStyleSheet("""
            QCalendarWidget {
                background-color: white;
            }
            QCalendarWidget QTableView {
                selection-background-color: #0078d4;
                selection-color: white;
            }
            QCalendarWidget QWidget#qt_calendar_navigationbar {
                background-color: #0078d4;
            }
            QCalendarWidget QToolButton {
                color: white;
                background-color: transparent;
                border: none;
                font-weight: bold;
            }
        """)

        layout.addWidget(self.calendar)

        # 當天排程摘要
        summary_group = QGroupBox("當天排程")
        summary_layout = QVBoxLayout(summary_group)

        self.daily_summary = QTextEdit()
        self.daily_summary.setReadOnly(True)
        self.daily_summary.setMaximumHeight(150)
        summary_layout.addWidget(self.daily_summary)

        layout.addWidget(summary_group)

        return panel

    def create_right_panel(self) -> QWidget:
        """建立右側排程列表面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        # 按鈕工具列
        button_layout = QHBoxLayout()

        self.btn_add = QPushButton("+ 新增排程")
        self.btn_add.setToolTip("新增排程任務")
        self.btn_add.clicked.connect(self.add_schedule)

        self.btn_edit = QPushButton("✎ 編輯")
        self.btn_edit.setToolTip("編輯選取的排程")
        self.btn_edit.clicked.connect(self.edit_schedule)
        self.btn_edit.setEnabled(False)

        self.btn_delete = QPushButton("✕ 刪除")
        self.btn_delete.setToolTip("刪除選取的排程")
        self.btn_delete.clicked.connect(self.delete_schedule)
        self.btn_delete.setEnabled(False)

        self.btn_refresh = QPushButton("↻ 重新整理")
        self.btn_refresh.setToolTip("重新載入排程列表")
        self.btn_refresh.clicked.connect(self.load_schedules)

        button_layout.addWidget(self.btn_add)
        button_layout.addWidget(self.btn_edit)
        button_layout.addWidget(self.btn_delete)
        button_layout.addStretch()
        button_layout.addWidget(self.btn_refresh)

        layout.addLayout(button_layout)

        # 排程表格
        self.schedule_table = QTableWidget()
        self.schedule_table.setColumnCount(7)
        self.schedule_table.setHorizontalHeaderLabels(
            ["ID", "任務名稱", "OPC URL", "Node ID", "目標值", "週期規則", "啟用"]
        )

        # 設定表格樣式
        self.schedule_table.horizontalHeader().setStretchLastSection(True)
        self.schedule_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeToContents
        )
        self.schedule_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.schedule_table.setSelectionMode(QTableWidget.SingleSelection)
        self.schedule_table.setAlternatingRowColors(True)
        self.schedule_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.schedule_table.customContextMenuRequested.connect(
            self.show_table_context_menu
        )
        self.schedule_table.itemSelectionChanged.connect(self.on_selection_changed)

        layout.addWidget(self.schedule_table)

        # 連線狀態面板
        status_group = QGroupBox("連線狀態")
        status_layout = QHBoxLayout(status_group)

        self.db_status_label = QLabel("資料庫: 未連線")
        self.opc_status_label = QLabel("OPC UA: 未連線")

        status_layout.addWidget(self.db_status_label)
        status_layout.addWidget(self.opc_status_label)
        status_layout.addStretch()

        self.btn_connect_opc = QPushButton("連線 OPC")
        self.btn_connect_opc.clicked.connect(self.connect_opc)
        status_layout.addWidget(self.btn_connect_opc)

        layout.addWidget(status_group)

        return panel

    def create_menu_bar(self):
        """建立選單列"""
        menubar = self.menuBar()

        # 檔案選單
        file_menu = menubar.addMenu("檔案(&F)")

        exit_action = QAction("結束(&X)", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # 工具選單
        tools_menu = menubar.addMenu("工具(&T)")

        db_settings_action = QAction("資料庫設定(&D)...", self)
        db_settings_action.triggered.connect(self.show_db_settings)
        tools_menu.addAction(db_settings_action)

        tools_menu.addSeparator()

        # 主題設定子選單
        theme_menu = tools_menu.addMenu("主題設定(&M)")
        self.theme_action_group = {}

        theme_system_action = QAction("跟隨系統(&S)", self)
        theme_system_action.setCheckable(True)
        theme_system_action.setChecked(self.current_theme == "system")
        theme_system_action.triggered.connect(lambda: self.set_theme("system"))
        theme_menu.addAction(theme_system_action)
        self.theme_action_group["system"] = theme_system_action

        theme_light_action = QAction("亮色模式(&L)", self)
        theme_light_action.setCheckable(True)
        theme_light_action.setChecked(self.current_theme == "light")
        theme_light_action.triggered.connect(lambda: self.set_theme("light"))
        theme_menu.addAction(theme_light_action)
        self.theme_action_group["light"] = theme_light_action

        theme_dark_action = QAction("暗色模式(&D)", self)
        theme_dark_action.setCheckable(True)
        theme_dark_action.setChecked(self.current_theme == "dark")
        theme_dark_action.triggered.connect(lambda: self.set_theme("dark"))
        theme_menu.addAction(theme_dark_action)
        self.theme_action_group["dark"] = theme_dark_action

        # 確保只有一個選項被選中
        theme_menu.triggered.connect(self._on_theme_menu_triggered)

        # 說明選單
        help_menu = menubar.addMenu("說明(&H)")

        about_action = QAction("關於(&A)...", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def _on_theme_menu_triggered(self, action):
        """處理主題選單點擊，確保只有一個選項被選中"""
        for theme, act in self.theme_action_group.items():
            if act != action:
                act.setChecked(False)

    def create_tool_bar(self):
        """建立工具列"""
        toolbar = QToolBar()
        self.addToolBar(toolbar)

        add_action = QAction("新增", self)
        add_action.triggered.connect(self.add_schedule)
        toolbar.addAction(add_action)

        toolbar.addSeparator()

        refresh_action = QAction("重新整理", self)
        refresh_action.triggered.connect(self.load_schedules)
        toolbar.addAction(refresh_action)

    def setup_connections(self):
        """設定信號連接"""
        self.calendar.selectionChanged.connect(self.update_daily_summary)

    def setup_system_tray(self):
        """設定系統托盤"""
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))

        tray_menu = QMenu()
        show_action = QAction("顯示", self)
        show_action.triggered.connect(self.show)
        tray_menu.addAction(show_action)

        tray_menu.addSeparator()

        quit_action = QAction("結束", self)
        quit_action.triggered.connect(self.close)
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)

    def on_tray_activated(self, reason):
        """處理托盤圖示點擊"""
        if reason == QSystemTrayIcon.DoubleClick:
            self.show()
            self.raise_()
            self.activateWindow()

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
            }
            QTextEdit {
                background-color: white;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 8px;
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
                border: 2px solid #0078d4;
                background-color: #f0f0f0;
            }
            QCheckBox::indicator:checked {
                background-color: #0078d4;
                border: 2px solid #0078d4;
                image: url(:/checkbox_check);
            }
        """)
        # 更新日曆樣式
        self._apply_calendar_light_theme()

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
        # 更新日曆樣式
        self._apply_calendar_dark_theme()

    def _apply_calendar_light_theme(self):
        """套用日曆亮色主題"""
        if hasattr(self, "calendar"):
            self.calendar.setStyleSheet("""
                QCalendarWidget {
                    background-color: white;
                }
                QCalendarWidget QTableView {
                    selection-background-color: #0078d4;
                    selection-color: white;
                }
                QCalendarWidget QWidget#qt_calendar_navigationbar {
                    background-color: #0078d4;
                }
                QCalendarWidget QToolButton {
                    color: white;
                    background-color: transparent;
                    border: none;
                    font-weight: bold;
                }
            """)

    def _apply_calendar_dark_theme(self):
        """套用日曆暗色主題"""
        if hasattr(self, "calendar"):
            self.calendar.setStyleSheet("""
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

            # 更新選單狀態
            if hasattr(self, "theme_action_group"):
                self.theme_action_group[theme].setChecked(True)

    def init_database(self):
        """初始化資料庫連線"""
        try:
            # 初始化 SQLite 管理器（使用預設資料庫路徑）
            self.db_manager = SQLiteManager()

            # 建立資料表
            if self.db_manager.init_db():
                self.db_status_label.setText("資料庫: 已連線")
                self.db_status_label.setStyleSheet("color: green;")
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
        self.update_schedule_table()
        self.update_daily_summary()

        self.status_bar.showMessage(f"已載入 {len(self.schedules)} 個排程")

    def update_schedule_table(self):
        """更新排程表格"""
        self.schedule_table.setRowCount(len(self.schedules))

        for row, schedule in enumerate(self.schedules):
            self.schedule_table.setItem(
                row, 0, QTableWidgetItem(str(schedule.get("id", "")))
            )
            self.schedule_table.setItem(
                row, 1, QTableWidgetItem(schedule.get("task_name", ""))
            )
            self.schedule_table.setItem(
                row, 2, QTableWidgetItem(schedule.get("opc_url", ""))
            )
            self.schedule_table.setItem(
                row, 3, QTableWidgetItem(schedule.get("node_id", ""))
            )
            self.schedule_table.setItem(
                row, 4, QTableWidgetItem(schedule.get("target_value", ""))
            )
            self.schedule_table.setItem(
                row, 5, QTableWidgetItem(schedule.get("rrule_str", ""))
            )

            enabled = "✓" if schedule.get("is_enabled") else "✗"
            item = QTableWidgetItem(enabled)
            item.setTextAlignment(Qt.AlignCenter)
            self.schedule_table.setItem(row, 6, item)

    def update_daily_summary(self):
        """更新當天排程摘要"""
        selected_date = self.calendar.selectedDate().toPython()

        # 取得當天的排程觸發時間
        daily_schedules = []

        for schedule in self.schedules:
            if not schedule.get("is_enabled"):
                continue

            rrule_str = schedule.get("rrule_str", "")
            if rrule_str:
                # 取得當天的觸發時間
                start = datetime.combine(selected_date, datetime.min.time())
                end = datetime.combine(selected_date, datetime.max.time())

                triggers = RRuleParser.get_trigger_between(rrule_str, start, end)

                for trigger in triggers:
                    daily_schedules.append(
                        {
                            "time": trigger.strftime("%H:%M"),
                            "name": schedule.get("task_name", ""),
                            "value": schedule.get("target_value", ""),
                        }
                    )

        # 排序並顯示
        daily_schedules.sort(key=lambda x: x["time"])

        if daily_schedules:
            summary_text = f"<b>{selected_date.strftime('%Y/%m/%d')} 排程:</b><br>"
            for item in daily_schedules:
                summary_text += (
                    f"<br>🕐 {item['time']} - {item['name']} ({item['value']})"
                )
        else:
            summary_text = (
                f"<b>{selected_date.strftime('%Y/%m/%d')}</b><br><br>當天沒有排程任務"
            )

        self.daily_summary.setHtml(summary_text)

    def add_schedule(self):
        """新增排程"""
        dialog = ScheduleEditDialog(self)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()

            if self.db_manager:
                schedule_id = self.db_manager.create_schedule(
                    task_name=data["task_name"],
                    opc_url=data["opc_url"],
                    node_id=data["node_id"],
                    target_value=data["target_value"],
                    rrule_str=data["rrule_str"],
                    opc_security_policy=data.get("opc_security_policy", "None"),
                    opc_security_mode=data.get("opc_security_mode", "None"),
                    opc_username=data.get("opc_username", ""),
                    opc_password=data.get("opc_password", ""),
                    opc_timeout=data.get("opc_timeout", 10),
                    is_enabled=1,
                )

                if schedule_id:
                    QMessageBox.information(self, "成功", "排程已新增")
                    self.load_schedules()
                else:
                    QMessageBox.critical(self, "錯誤", "新增排程失敗")

    def edit_schedule(self):
        """編輯排程"""
        current_row = self.schedule_table.currentRow()
        if current_row < 0 or current_row >= len(self.schedules):
            return

        schedule = self.schedules[current_row]

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
                    rrule_str=data["rrule_str"],
                    opc_security_policy=data.get("opc_security_policy", "None"),
                    opc_security_mode=data.get("opc_security_mode", "None"),
                    opc_username=data.get("opc_username", ""),
                    opc_password=data.get("opc_password", ""),
                    opc_timeout=data.get("opc_timeout", 10),
                )

                if success:
                    QMessageBox.information(self, "成功", "排程已更新")
                    self.load_schedules()
                else:
                    QMessageBox.critical(self, "錯誤", "更新排程失敗")

    def delete_schedule(self):
        """刪除排程"""
        current_row = self.schedule_table.currentRow()
        if current_row < 0 or current_row >= len(self.schedules):
            return

        schedule = self.schedules[current_row]

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

    def on_selection_changed(self):
        """處理表格選擇變更"""
        has_selection = self.schedule_table.currentRow() >= 0
        self.btn_edit.setEnabled(has_selection)
        self.btn_delete.setEnabled(has_selection)

    def show_table_context_menu(self, position):
        """顯示表格右鍵選單"""
        menu = QMenu()

        edit_action = menu.addAction("編輯")
        edit_action.triggered.connect(self.edit_schedule)

        delete_action = menu.addAction("刪除")
        delete_action.triggered.connect(self.delete_schedule)

        menu.addSeparator()

        toggle_action = menu.addAction("啟用/停用")
        toggle_action.triggered.connect(self.toggle_schedule_enabled)

        menu.exec(self.schedule_table.viewport().mapToGlobal(position))

    def toggle_schedule_enabled(self):
        """切換排程啟用狀態"""
        current_row = self.schedule_table.currentRow()
        if current_row < 0 or current_row >= len(self.schedules):
            return

        schedule = self.schedules[current_row]
        new_status = 0 if schedule.get("is_enabled") else 1

        if self.db_manager:
            self.db_manager.toggle_schedule(schedule["id"], new_status)
            self.load_schedules()

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
        self.status_bar.showMessage(f"執行排程: {schedule.get('task_name', '')}")

        # 執行 OPC UA 寫入
        asyncio.create_task(self.execute_task(schedule))

    async def execute_task(self, schedule: Dict[str, Any]):
        """執行排程任務"""
        opc_url = schedule.get("opc_url", "")
        node_id = schedule.get("node_id", "")
        target_value = schedule.get("target_value", "")

        # 取得OPC設定
        security_policy = schedule.get("opc_security_policy", "None")
        username = schedule.get("opc_username", "")
        password = schedule.get("opc_password", "")
        timeout = schedule.get("opc_timeout", 10)

        try:
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
                    success = await handler.write_node(node_id, target_value)
                    if success:
                        self.status_bar.showMessage(
                            f"✓ 已寫入 {node_id} = {target_value}", 5000
                        )
                    else:
                        self.status_bar.showMessage(f"✗ 寫入失敗: {node_id}", 5000)
                else:
                    self.status_bar.showMessage(f"✗ 無法連線 OPC UA: {opc_url}", 5000)
        except Exception as e:
            self.status_bar.showMessage(f"✗ 執行錯誤: {str(e)}", 5000)

    async def connect_opc(self):
        """連線到 OPC UA 伺服器"""
        # 優先使用選取的排程的 OPC URL，若無則使用第一筆排程的 URL
        opc_url = None
        current_row = self.schedule_table.currentRow()
        if current_row >= 0 and current_row < len(self.schedules):
            opc_url = self.schedules[current_row].get("opc_url", "")
        elif self.schedules:
            opc_url = self.schedules[0].get("opc_url", "")

        if not opc_url:
            QMessageBox.warning(
                self,
                "OPC 連線",
                "請先在排程中設定 OPC URL 或選取一筆含有 OPC URL 的排程。",
            )
            return

        self.opc_handler = OPCHandler(opc_url)

        try:
            success = await self.opc_handler.connect()
            if success:
                self.opc_status_label.setText(f"OPC UA: 已連線 ({opc_url})")
                self.opc_status_label.setStyleSheet("color: green;")
            else:
                self.opc_status_label.setText("OPC UA: 連線失敗")
                self.opc_status_label.setStyleSheet("color: red;")
        except Exception as e:
            self.opc_status_label.setText("OPC UA: 連線錯誤")
            self.opc_status_label.setStyleSheet("color: red;")
            QMessageBox.critical(self, "OPC 連線錯誤", str(e))

    def show_db_settings(self):
        """顯示資料庫設定對話框"""
        QMessageBox.information(self, "資料庫設定", "資料庫設定功能開發中...")

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
        if self.scheduler_worker:
            self.scheduler_worker.stop()

        if self.opc_handler:
            asyncio.create_task(self.opc_handler.disconnect())

        event.accept()


class OPCNodeBrowserDialog(QDialog):
    """OPC UA 節點瀏覽對話框"""

    def __init__(self, parent=None, opc_url: str = ""):
        super().__init__(parent)
        self.opc_url = opc_url
        self.selected_node = ""
        self.opc_handler = None
        self.setup_ui()
        self.apply_style()
        # 自動連線並載入節點
        QTimer.singleShot(100, self.connect_and_load)

    def setup_ui(self):
        """設定介面"""
        self.setWindowTitle("瀏覽 OPC UA 節點")
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
        self.tree_widget.setHeaderLabels(["節點名稱", "Node ID", "節點類型"])
        self.tree_widget.setColumnWidth(0, 200)
        self.tree_widget.setColumnWidth(1, 150)
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

        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        self.select_btn = QPushButton("選擇")
        self.select_btn.setDefault(True)
        self.select_btn.setEnabled(False)
        self.select_btn.clicked.connect(self.accept)
        button_layout.addWidget(self.select_btn)

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

        except Exception as e:
            self.status_label.setText(f"載入節點錯誤: {str(e)}")
            self.status_label.setStyleSheet("color: red;")

    async def _async_load_child_nodes(self, parent_node, parent_item, depth=0):
        """異步遞迴載入子節點"""
        if depth > 2:  # 限制深度避免載入太多
            return

        try:
            # 取得子節點
            children = await parent_node.get_children()

            for child in children:
                try:
                    child_item = QTreeWidgetItem(parent_item)

                    # 取得節點資訊
                    browse_name = await child.read_browse_name()
                    node_id = str(child.nodeid)
                    node_class = await child.read_node_class()

                    child_item.setText(0, browse_name.Name)
                    child_item.setText(1, node_id)
                    child_item.setText(2, str(node_class))

                    # 儲存節點 ID
                    child_item.setData(0, Qt.ItemDataRole.UserRole, node_id)

                    # 繼續載入子節點
                    await self._async_load_child_nodes(child, child_item, depth + 1)

                except Exception:
                    pass

        except Exception:
            pass

        except Exception:
            pass

    def on_selection_changed(self):
        """處理選擇變更"""
        selected_items = self.tree_widget.selectedItems()
        if selected_items:
            self.selected_node = selected_items[0].text(1)
            self.select_btn.setEnabled(True)
        else:
            self.selected_node = ""
            self.select_btn.setEnabled(False)

    def on_item_double_clicked(self, item, column):
        """處理雙擊事件"""
        self.selected_node = item.text(1)
        self.accept()

    def get_selected_node(self) -> str:
        """取得選擇的節點 ID"""
        return self.selected_node


class OPCSettingsDialog(QDialog):
    """OPC UA 設定對話框"""

    def __init__(self, parent=None, security_policy="None", username="", password="", timeout=10, security_mode="None", opc_url: str = ""):
        super().__init__(parent)
        self.security_policy = security_policy
        self.security_mode = security_mode
        self.username = username
        self.password = password
        self.timeout = timeout
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
        self.setMinimumWidth(900)
        self.setMinimumHeight(520)
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
        
        # 連線設定（下方）- 高度減少
        connection_group = QGroupBox("連線設定")
        connection_layout = QHBoxLayout(connection_group)
        connection_layout.setContentsMargins(10, 8, 10, 8)
        connection_layout.setSpacing(8)
        connection_layout.addWidget(QLabel("連線超時 (秒):"))
        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(1, 300)
        self.timeout_spin.setValue(10)
        self.timeout_spin.setFixedWidth(80)
        connection_layout.addWidget(self.timeout_spin)
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
                border: 2px solid #0078d4;
                background-color: #f0f0f0;
            }
            QCheckBox::indicator:checked {
                background-color: #0078d4;
                border: 2px solid #0078d4;
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

        # 初始化OPC設定
        self.opc_security_policy = schedule.get("opc_security_policy", "None") if schedule else "None"
        self.opc_security_mode = schedule.get("opc_security_mode", "None") if schedule else "None"
        self.opc_username = schedule.get("opc_username", "") if schedule else ""
        self.opc_password = schedule.get("opc_password", "") if schedule else ""
        self.opc_timeout = schedule.get("opc_timeout", 10) if schedule else 10

        self.setup_ui()
        self.apply_style()

        if schedule:
            self.load_data()

    def setup_ui(self):
        self.setWindowTitle("編輯排程" if self.schedule else "新增排程")
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
        self.target_value_edit = QLineEdit()
        self.target_value_edit.setPlaceholderText("1")
        basic_layout.addWidget(self.target_value_edit, 3, 1)

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

        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        ok_btn = QPushButton("確定")
        ok_btn.setDefault(True)
        ok_btn.clicked.connect(self.accept)
        button_layout.addWidget(ok_btn)

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
                border: 2px solid #0078d4;
                background-color: #f0f0f0;
            }
            QCheckBox::indicator:checked {
                background-color: #0078d4;
                border: 2px solid #0078d4;
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
            }
        """)

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
        self.rrule_display.setText(self.schedule.get("rrule_str", ""))

    def _normalize_opc_url(self) -> str:
        """規範化 OPC URL：添加 opc.tcp:// 前綴（如果需要）"""
        opc_url = self.opc_url_edit.text().strip()
        if opc_url and not opc_url.startswith("opc.tcp://"):
            opc_url = f"opc.tcp://{opc_url}"
        return opc_url

    def edit_recurrence(self):
        """編輯週期規則"""
        current_rrule = self.rrule_display.text()
        rrule = show_recurrence_dialog(self, current_rrule)
        if rrule:
            self.rrule_display.setText(rrule)

    def get_data(self) -> Dict[str, Any]:
        """取得編輯的資料"""
        # 自動添加 opc.tcp:// 前綴
        opc_url = self.opc_url_edit.text().strip()
        if opc_url and not opc_url.startswith("opc.tcp://"):
            opc_url = f"opc.tcp://{opc_url}"

        return {
            "task_name": self.task_name_edit.text(),
            "opc_url": opc_url,
            "node_id": self.node_id_edit.text(),
            "target_value": self.target_value_edit.text(),
            "rrule_str": self.rrule_display.text(),
            "opc_security_policy": self.opc_security_policy,
            "opc_security_mode": self.opc_security_mode,
            "opc_username": self.opc_username,
            "opc_password": self.opc_password,
            "opc_timeout": self.opc_timeout,
        }

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
                self.node_id_edit.setText(selected_node)

    def configure_opc_settings(self):
        """設定 OPC UA 連線參數"""
        opc_url = self._normalize_opc_url()

        dialog = OPCSettingsDialog(
            self,
            self.opc_security_policy,
            self.opc_username,
            self.opc_password,
            self.opc_timeout,
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


async def main():
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
    asyncio.run(main())

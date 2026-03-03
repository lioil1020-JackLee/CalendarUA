"""
資料庫設定對話框
提供資料庫路徑設定與資訊檢視功能
"""

from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QWidget,
    QLabel,
    QPushButton,
    QLineEdit,
    QGroupBox,
    QMessageBox,
    QFileDialog,
    QTextEdit,
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QIcon
import os
from pathlib import Path
from datetime import datetime
from database.sqlite_manager import SQLiteManager


def get_app_icon():
    """獲取應用程式圖示，支援打包環境"""
    import sys
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


class DatabaseSettingsDialog(QDialog):
    """資料庫設定對話框"""

    database_changed = Signal(str)  # 當資料庫路徑改變時發出訊號

    def __init__(self, parent=None, db_manager: SQLiteManager = None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.setWindowTitle("資料庫設定")
        self.setWindowIcon(get_app_icon())
        self.setMinimumWidth(500)
        self.setMinimumHeight(560)
        self.setModal(True)

        self.setup_ui()
        self.apply_modern_style()
        self.connect_signals()
        self.load_current_settings()

    def setup_ui(self):
        """設定主介面"""
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # 資料庫路徑設定區塊
        main_layout.addWidget(self.create_path_group())

        # 資料庫資訊區塊
        main_layout.addWidget(self.create_info_group())

        # 按鈕區塊
        main_layout.addWidget(self.create_button_group())

    def create_path_group(self) -> QGroupBox:
        """建立資料庫路徑設定區塊"""
        group = QGroupBox("資料庫路徑設定")
        layout = QVBoxLayout(group)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        # 當前路徑顯示
        path_layout = QHBoxLayout()
        path_label = QLabel("當前路徑:")
        path_label.setObjectName("fieldLabel")
        path_layout.addWidget(path_label)

        self.path_edit = QLineEdit()
        self.path_edit.setReadOnly(True)
        path_layout.addWidget(self.path_edit)

        change_btn = QPushButton("變更...")
        change_btn.clicked.connect(self.change_database_path)
        path_layout.addWidget(change_btn)

        layout.addLayout(path_layout)

        return group

    def create_info_group(self) -> QGroupBox:
        """建立資料庫資訊區塊"""
        group = QGroupBox("資料庫資訊")
        layout = QVBoxLayout(group)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        self.info_text = QTextEdit()
        self.info_text.setReadOnly(True)
        self.info_text.setMinimumHeight(300)
        self.info_text.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        layout.addWidget(self.info_text)

        return group

    def create_button_group(self) -> QWidget:
        """建立按鈕區塊"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)

        layout.addStretch()

        close_btn = QPushButton("關閉")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

        return widget

    def apply_modern_style(self):
        """應用現代化樣式，支援主題切換"""
        is_dark = self.is_dark_mode()

        if is_dark:
            self._apply_dark_theme()
        else:
            self._apply_light_theme()

    def is_dark_mode(self) -> bool:
        """檢查是否使用暗色模式"""
        # 遍歷父視窗鏈查找主題設定
        parent = self.parent()
        while parent:
            if hasattr(parent, "current_theme"):
                if parent.current_theme == "dark":
                    return True
                elif parent.current_theme == "system":
                    if hasattr(parent, "is_system_dark_mode"):
                        return parent.is_system_dark_mode()
                return False
            parent = parent.parent() if hasattr(parent, "parent") else None
        return False

    def _apply_light_theme(self):
        """套用亮色主題"""
        self.setStyleSheet("""
            QDialog {
                background-color: #f8f9fa;
            }

            QGroupBox {
                font-weight: bold;
                border: 2px solid #dee2e6;
                border-radius: 5px;
                margin-top: 1ex;
                background-color: white;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
                color: #495057;
            }

            QLabel#fieldLabel {
                color: #495057;
                font-weight: bold;
            }

            QLineEdit, QTextEdit {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 6px;
                background-color: white;
                color: #495057;
            }

            QLineEdit:focus, QTextEdit:focus {
                border-color: #80bdff;
                outline: none;
            }

            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }

            QPushButton:hover {
                background-color: #0056b3;
            }

            QPushButton:pressed {
                background-color: #004085;
            }

            QProgressBar {
                border: 1px solid #ced4da;
                border-radius: 4px;
                text-align: center;
            }

            QProgressBar::chunk {
                background-color: #007bff;
                border-radius: 3px;
            }
        """)

    def _apply_dark_theme(self):
        """套用暗色主題"""
        self.setStyleSheet("""
            QDialog {
                background-color: #2b2b2b;
            }

            QGroupBox {
                font-weight: bold;
                border: 2px solid #3d3d3d;
                border-radius: 5px;
                margin-top: 1ex;
                background-color: #363636;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
                color: #cccccc;
            }

            QLabel#fieldLabel {
                color: #cccccc;
                font-weight: bold;
            }

            QLineEdit, QTextEdit {
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 6px;
                background-color: #1e1e1e;
                color: #cccccc;
            }

            QLineEdit:focus, QTextEdit:focus {
                border-color: #0e639c;
                outline: none;
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

            QPushButton:pressed {
                background-color: #094771;
            }

            QProgressBar {
                border: 1px solid #555555;
                border-radius: 4px;
                text-align: center;
                background-color: #1e1e1e;
            }

            QProgressBar::chunk {
                background-color: #0e639c;
                border-radius: 3px;
            }
        """)

    def connect_signals(self):
        """連接訊號"""
        pass

    def load_current_settings(self):
        """載入當前設定"""
        if self.db_manager:
            self.path_edit.setText(str(self.db_manager.db_path))
            self.refresh_database_info()

    def change_database_path(self):
        """變更資料庫路徑"""
        current_path = self.path_edit.text()
        if not current_path:
            current_path = "./database/calendarua.db"

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "選擇資料庫檔案",
            current_path,
            "SQLite 資料庫 (*.db);;所有檔案 (*)",
        )

        if file_path:
            # 確認檔案副檔名
            if not file_path.endswith('.db'):
                file_path += '.db'

            # 檢查是否需要複製現有資料庫
            old_path = Path(self.path_edit.text())
            new_path = Path(file_path)

            if old_path.exists() and old_path != new_path:
                reply = QMessageBox.question(
                    self,
                    "複製資料庫",
                    f"是否要將現有資料庫複製到新位置？\n\n從: {old_path}\n到: {new_path}",
                    QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                )

                if reply == QMessageBox.Cancel:
                    return
                elif reply == QMessageBox.Yes:
                    try:
                        # 確保目標目錄存在
                        new_path.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(old_path, new_path)
                    except Exception as e:
                        QMessageBox.critical(
                            self,
                            "複製失敗",
                            f"無法複製資料庫檔案:\n{str(e)}"
                        )
                        return

            # 更新路徑
            self.path_edit.setText(str(new_path))

            # 發出變更訊號
            self.database_changed.emit(str(new_path))

            QMessageBox.information(
                self,
                "路徑變更",
                "資料庫路徑已更新。請重新啟動應用程式以使用新資料庫。"
            )

    def refresh_database_info(self):
        """重新整理資料庫資訊"""
        if not self.db_manager:
            self.info_text.setPlainText("資料庫管理器未初始化")
            return

        try:
            # 取得統計資訊
            total_schedules = len(self.db_manager.get_all_schedules())
            enabled_schedules = len(self.db_manager.get_all_schedules(enabled_only=True))

            # 取得資料庫檔案資訊
            db_path = Path(self.db_manager.db_path)
            file_size = 0
            if db_path.exists():
                file_size = db_path.stat().st_size

            # 格式化資訊
            info = f"""資料庫統計資訊:

總排程數量: {total_schedules}
啟用排程數量: {enabled_schedules}
停用排程數量: {total_schedules - enabled_schedules}

檔案資訊:
路徑: {db_path}
大小: {self.format_file_size(file_size)}
修改時間: {datetime.fromtimestamp(db_path.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S') if db_path.exists() else '檔案不存在'}
"""

            self.info_text.setPlainText(info)

        except Exception as e:
            self.info_text.setPlainText(f"取得資料庫資訊時發生錯誤:\n{str(e)}")

    def format_file_size(self, size_bytes: int) -> str:
        """格式化檔案大小"""
        if size_bytes == 0:
            return "0 B"

        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        size = size_bytes

        while size >= 1024 and i < len(size_names) - 1:
            size /= 1024.0
            i += 1

        return f"{size:.1f} {size_names[i]}"
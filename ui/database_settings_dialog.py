"""
資料庫設定對話框
提供資料庫路徑設定、備份、還原等功能
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
    QProgressBar,
    QFrame,
)
from PySide6.QtCore import Qt, Signal, QThread, QTimer
from PySide6.QtGui import QFont, QIcon
import os
import shutil
from pathlib import Path
from datetime import datetime
from database.sqlite_manager import SQLiteManager


class DatabaseSettingsDialog(QDialog):
    """資料庫設定對話框"""

    database_changed = Signal(str)  # 當資料庫路徑改變時發出訊號

    def __init__(self, parent=None, db_manager: SQLiteManager = None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.setWindowTitle("資料庫設定")
        self.setMinimumWidth(500)
        self.setMinimumHeight(400)
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

        # 資料庫操作區塊
        main_layout.addWidget(self.create_operations_group())

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

    def create_operations_group(self) -> QGroupBox:
        """建立資料庫操作區塊"""
        group = QGroupBox("資料庫操作")
        layout = QVBoxLayout(group)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        # 操作按鈕
        operations_layout = QHBoxLayout()

        self.backup_btn = QPushButton("備份資料庫")
        # self.backup_btn.setIcon(self.style().standardIcon(self.style().SP_DialogSaveButton))
        self.backup_btn.clicked.connect(self.backup_database)
        operations_layout.addWidget(self.backup_btn)

        self.restore_btn = QPushButton("還原資料庫")
        # self.restore_btn.setIcon(self.style().standardIcon(self.style().SP_DialogOpenButton))
        self.restore_btn.clicked.connect(self.restore_database)
        operations_layout.addWidget(self.restore_btn)

        self.clear_btn = QPushButton("清除所有資料")
        # self.clear_btn.setIcon(self.style().standardIcon(self.style().SP_TrashIcon))
        self.clear_btn.setStyleSheet("QPushButton { color: red; }")
        self.clear_btn.clicked.connect(self.clear_database)
        operations_layout.addWidget(self.clear_btn)

        layout.addLayout(operations_layout)

        # 進度條（用於長時間操作）
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        return group

    def create_info_group(self) -> QGroupBox:
        """建立資料庫資訊區塊"""
        group = QGroupBox("資料庫資訊")
        layout = QVBoxLayout(group)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        self.info_text = QTextEdit()
        self.info_text.setReadOnly(True)
        self.info_text.setMaximumHeight(120)
        layout.addWidget(self.info_text)

        refresh_btn = QPushButton("重新整理資訊")
        refresh_btn.clicked.connect(self.refresh_database_info)
        layout.addWidget(refresh_btn)

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

    def backup_database(self):
        """備份資料庫"""
        if not self.db_manager:
            QMessageBox.warning(self, "錯誤", "資料庫管理器未初始化")
            return

        current_path = Path(self.path_edit.text())
        if not current_path.exists():
            QMessageBox.warning(self, "錯誤", "資料庫檔案不存在")
            return

        # 產生備份檔案名稱
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"calendarua_backup_{timestamp}.db"

        backup_path, _ = QFileDialog.getSaveFileName(
            self,
            "選擇備份位置",
            backup_name,
            "SQLite 資料庫 (*.db);;所有檔案 (*)",
        )

        if backup_path:
            try:
                # 確保副檔名
                if not backup_path.endswith('.db'):
                    backup_path += '.db'

                shutil.copy2(current_path, backup_path)

                QMessageBox.information(
                    self,
                    "備份成功",
                    f"資料庫已成功備份到:\n{backup_path}"
                )

            except Exception as e:
                QMessageBox.critical(
                    self,
                    "備份失敗",
                    f"無法備份資料庫:\n{str(e)}"
                )

    def restore_database(self):
        """還原資料庫"""
        if not self.db_manager:
            QMessageBox.warning(self, "錯誤", "資料庫管理器未初始化")
            return

        backup_path, _ = QFileDialog.getOpenFileName(
            self,
            "選擇備份檔案",
            "",
            "SQLite 資料庫 (*.db);;所有檔案 (*)",
        )

        if backup_path:
            reply = QMessageBox.question(
                self,
                "確認還原",
                "還原資料庫將覆蓋現有資料，此操作無法復原。\n\n確定要繼續嗎？",
                QMessageBox.Yes | QMessageBox.No,
            )

            if reply == QMessageBox.Yes:
                try:
                    current_path = Path(self.path_edit.text())
                    # 確保目標目錄存在
                    current_path.parent.mkdir(parents=True, exist_ok=True)

                    shutil.copy2(backup_path, current_path)

                    QMessageBox.information(
                        self,
                        "還原成功",
                        "資料庫已成功還原。請重新啟動應用程式以載入還原的資料。"
                    )

                    # 重新整理資訊
                    self.refresh_database_info()

                except Exception as e:
                    QMessageBox.critical(
                        self,
                        "還原失敗",
                        f"無法還原資料庫:\n{str(e)}"
                    )

    def clear_database(self):
        """清除所有資料"""
        if not self.db_manager:
            QMessageBox.warning(self, "錯誤", "資料庫管理器未初始化")
            return

        reply = QMessageBox.question(
            self,
            "危險操作",
            "清除所有資料將刪除所有排程記錄，此操作無法復原。\n\n確定要繼續嗎？",
            QMessageBox.Yes | QMessageBox.No,
        )

        if reply == QMessageBox.Yes:
            # 再次確認
            reply2 = QMessageBox.question(
                self,
                "最後確認",
                "這是最後的確認：\n\n所有排程資料將被永久刪除！\n\n確定要清除所有資料嗎？",
                QMessageBox.Yes | QMessageBox.No,
            )

            if reply2 == QMessageBox.Yes:
                try:
                    success = self.db_manager.clear_all_schedules()
                    if success:
                        QMessageBox.information(
                            self,
                            "清除完成",
                            "所有排程資料已清除。"
                        )
                        self.refresh_database_info()
                    else:
                        QMessageBox.critical(
                            self,
                            "清除失敗",
                            "清除資料時發生錯誤。"
                        )

                except Exception as e:
                    QMessageBox.critical(
                        self,
                        "清除失敗",
                        f"清除資料時發生錯誤:\n{str(e)}"
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
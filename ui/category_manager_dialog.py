"""
Category Manager Dialog - 類別管理對話框
提供顏色類別的新增、編輯、刪除功能
"""

from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QMessageBox,
    QLabel,
    QLineEdit,
    QFormLayout,
    QWidget,
    QColorDialog,
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from typing import Optional, Dict, Any, List
import logging

logger = logging.getLogger(__name__)


class CategoryEditorDialog(QDialog):
    """類別編輯對話框"""

    def __init__(self, parent=None, category: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.category = category
        self.bg_color = category.get("bg_color", "#FF0000") if category else "#FF0000"
        self.fg_color = category.get("fg_color", "#FFFFFF") if category else "#FFFFFF"
        
        self.setup_ui()
        
        if category:
            self.setWindowTitle("編輯類別")
            self.load_category()
        else:
            self.setWindowTitle("新增類別")

    def setup_ui(self):
        """設定UI"""
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # 表單
        form_layout = QFormLayout()
        
        # 名稱
        self.name_edit = QLineEdit()
        form_layout.addRow("名稱:", self.name_edit)
        
        # 顏色選擇
        color_layout = QHBoxLayout()
        
        # 背景顏色
        bg_layout = QVBoxLayout()
        self.bg_color_label = QLabel("背景顏色")
        self.bg_color_button = QPushButton()
        self.bg_color_button.setFixedSize(100, 30)
        self.bg_color_button.clicked.connect(self.choose_bg_color)
        self.update_bg_color_button()
        bg_layout.addWidget(self.bg_color_label)
        bg_layout.addWidget(self.bg_color_button)
        
        # 前景顏色
        fg_layout = QVBoxLayout()
        self.fg_color_label = QLabel("前景顏色")
        self.fg_color_button = QPushButton()
        self.fg_color_button.setFixedSize(100, 30)
        self.fg_color_button.clicked.connect(self.choose_fg_color)
        self.update_fg_color_button()
        fg_layout.addWidget(self.fg_color_label)
        fg_layout.addWidget(self.fg_color_button)
        
        # 預覽
        preview_layout = QVBoxLayout()
        self.preview_label = QLabel("預覽")
        self.preview_button = QPushButton("範例文字")
        self.preview_button.setFixedSize(100, 30)
        self.preview_button.setEnabled(False)
        self.update_preview()
        preview_layout.addWidget(self.preview_label)
        preview_layout.addWidget(self.preview_button)
        
        color_layout.addLayout(bg_layout)
        color_layout.addLayout(fg_layout)
        color_layout.addLayout(preview_layout)
        color_layout.addStretch()
        
        form_layout.addRow("", color_layout)
        
        layout.addLayout(form_layout)
        
        # 按鈕
        button_layout = QHBoxLayout()
        
        self.btn_ok = QPushButton("確定")
        self.btn_ok.clicked.connect(self.accept)
        
        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(self.btn_ok)
        button_layout.addWidget(self.btn_cancel)
        
        layout.addLayout(button_layout)

    def load_category(self):
        """載入類別資料"""
        if not self.category:
            return
        
        self.name_edit.setText(self.category.get("name", ""))
        self.bg_color = self.category.get("bg_color", "#FF0000")
        self.fg_color = self.category.get("fg_color", "#FFFFFF")
        
        self.update_bg_color_button()
        self.update_fg_color_button()
        self.update_preview()

    def choose_bg_color(self):
        """選擇背景顏色"""
        color = QColorDialog.getColor(QColor(self.bg_color), self, "選擇背景顏色")
        if color.isValid():
            self.bg_color = color.name()
            self.update_bg_color_button()
            self.update_preview()

    def choose_fg_color(self):
        """選擇前景顏色"""
        color = QColorDialog.getColor(QColor(self.fg_color), self, "選擇前景顏色")
        if color.isValid():
            self.fg_color = color.name()
            self.update_fg_color_button()
            self.update_preview()

    def update_bg_color_button(self):
        """更新背景顏色按鈕"""
        self.bg_color_button.setStyleSheet(
            f"background-color: {self.bg_color}; border: 1px solid #ccc;"
        )

    def update_fg_color_button(self):
        """更新前景顏色按鈕"""
        self.fg_color_button.setStyleSheet(
            f"background-color: {self.fg_color}; border: 1px solid #ccc;"
        )

    def update_preview(self):
        """更新預覽"""
        self.preview_button.setStyleSheet(
            f"background-color: {self.bg_color}; color: {self.fg_color}; border: 1px solid #ccc;"
        )

    def get_data(self) -> Dict[str, Any]:
        """取得資料"""
        return {
            "name": self.name_edit.text().strip(),
            "bg_color": self.bg_color,
            "fg_color": self.fg_color,
        }


class CategoryManagerDialog(QDialog):
    """類別管理對話框"""

    category_changed = Signal()  # 類別變更信號

    def __init__(self, parent=None, db_manager=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.setWindowTitle("類別管理")
        self.setMinimumSize(700, 500)
        
        self.setup_ui()
        self.load_categories()

    def setup_ui(self):
        """設定UI"""
        layout = QVBoxLayout(self)
        
        # 說明
        info_label = QLabel(
            "系統類別 (標記為 ★) 不可刪除或重新命名。\n"
            "請勿刪除正在使用的類別。"
        )
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # 表格
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(
            ["ID", "名稱", "背景顏色", "前景顏色", "系統"]
        )
        
        # 設定表格樣式
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Fixed)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        
        self.table.setColumnWidth(2, 100)
        self.table.setColumnWidth(3, 100)
        
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setAlternatingRowColors(True)
        
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        self.table.doubleClicked.connect(self.edit_category)
        
        layout.addWidget(self.table)
        
        # 按鈕
        button_layout = QHBoxLayout()
        
        self.btn_add = QPushButton("+ 新增類別")
        self.btn_add.clicked.connect(self.add_category)
        
        self.btn_edit = QPushButton("✎ 編輯")
        self.btn_edit.setEnabled(False)
        self.btn_edit.clicked.connect(self.edit_category)
        
        self.btn_delete = QPushButton("✕ 刪除")
        self.btn_delete.setEnabled(False)
        self.btn_delete.clicked.connect(self.delete_category)
        
        self.btn_close = QPushButton("關閉")
        self.btn_close.clicked.connect(self.accept)
        
        button_layout.addWidget(self.btn_add)
        button_layout.addWidget(self.btn_edit)
        button_layout.addWidget(self.btn_delete)
        button_layout.addStretch()
        button_layout.addWidget(self.btn_close)
        
        layout.addLayout(button_layout)

    def load_categories(self):
        """載入類別列表"""
        if not self.db_manager:
            return
        
        categories = self.db_manager.get_all_categories()
        
        self.table.setRowCount(len(categories))
        
        for row, category in enumerate(categories):
            # ID
            id_item = QTableWidgetItem(str(category["id"]))
            id_item.setTextAlignment(Qt.AlignCenter)
            id_item.setData(Qt.UserRole, category)
            self.table.setItem(row, 0, id_item)
            
            # 名稱
            name = category["name"]
            if category["is_system"]:
                name = "★ " + name
            name_item = QTableWidgetItem(name)
            self.table.setItem(row, 1, name_item)
            
            # 背景顏色
            bg_item = QTableWidgetItem()
            bg_item.setBackground(QColor(category["bg_color"]))
            bg_item.setText(category["bg_color"])
            bg_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 2, bg_item)
            
            # 前景顏色
            fg_item = QTableWidgetItem()
            fg_item.setBackground(QColor(category["fg_color"]))
            fg_item.setText(category["fg_color"])
            fg_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 3, fg_item)
            
            # 系統
            system_item = QTableWidgetItem("是" if category["is_system"] else "")
            system_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 4, system_item)

    def on_selection_changed(self):
        """選擇變更"""
        has_selection = self.table.currentRow() >= 0
        self.btn_edit.setEnabled(has_selection)
        self.btn_delete.setEnabled(has_selection)

    def add_category(self):
        """新增類別"""
        dialog = CategoryEditorDialog(self)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()
            
            if not data["name"]:
                QMessageBox.warning(self, "警告", "請輸入類別名稱")
                return
            
            if self.db_manager:
                # 計算下一個排序順序
                categories = self.db_manager.get_all_categories()
                max_sort = max([c.get("sort_order", 0) for c in categories], default=0)
                
                category_id = self.db_manager.add_category(
                    name=data["name"],
                    bg_color=data["bg_color"],
                    fg_color=data["fg_color"],
                    sort_order=max_sort + 1
                )
                
                if category_id:
                    self.load_categories()
                    self.category_changed.emit()
                    QMessageBox.information(self, "成功", "類別已新增")
                else:
                    QMessageBox.critical(self, "錯誤", "新增類別失敗（名稱可能已存在）")

    def edit_category(self):
        """編輯類別"""
        current_row = self.table.currentRow()
        if current_row < 0:
            return
        
        category = self.table.item(current_row, 0).data(Qt.UserRole)
        
        if category["is_system"]:
            # 系統類別只能編輯顏色
            dialog = CategoryEditorDialog(self, category)
            dialog.name_edit.setEnabled(False)
            
            if dialog.exec() == QDialog.Accepted:
                data = dialog.get_data()
                
                if self.db_manager:
                    success = self.db_manager.update_category(
                        category_id=category["id"],
                        bg_color=data["bg_color"],
                        fg_color=data["fg_color"]
                    )
                    
                    if success:
                        self.load_categories()
                        self.category_changed.emit()
                        QMessageBox.information(self, "成功", "類別已更新")
                    else:
                        QMessageBox.critical(self, "錯誤", "更新類別失敗")
        else:
            # 使用者類別可以完全編輯
            dialog = CategoryEditorDialog(self, category)
            
            if dialog.exec() == QDialog.Accepted:
                data = dialog.get_data()
                
                if not data["name"]:
                    QMessageBox.warning(self, "警告", "請輸入類別名稱")
                    return
                
                if self.db_manager:
                    success = self.db_manager.update_category(
                        category_id=category["id"],
                        name=data["name"],
                        bg_color=data["bg_color"],
                        fg_color=data["fg_color"]
                    )
                    
                    if success:
                        self.load_categories()
                        self.category_changed.emit()
                        QMessageBox.information(self, "成功", "類別已更新")
                    else:
                        QMessageBox.critical(self, "錯誤", "更新類別失敗（名稱可能已存在）")

    def delete_category(self):
        """刪除類別"""
        current_row = self.table.currentRow()
        if current_row < 0:
            return
        
        category = self.table.item(current_row, 0).data(Qt.UserRole)
        
        if category["is_system"]:
            QMessageBox.warning(self, "警告", "無法刪除系統類別")
            return
        
        reply = QMessageBox.question(
            self,
            "確認刪除",
            f"確定要刪除類別 '{category['name']}' 嗎？\n\n"
            "注意：如果有排程正在使用此類別，將無法刪除。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            if self.db_manager:
                success = self.db_manager.delete_category(category["id"])
                
                if success:
                    self.load_categories()
                    self.category_changed.emit()
                    QMessageBox.information(self, "成功", "類別已刪除")
                else:
                    QMessageBox.critical(
                        self, 
                        "錯誤", 
                        "刪除類別失敗\n\n可能原因：\n- 有排程正在使用此類別\n- 系統類別不可刪除"
                    )

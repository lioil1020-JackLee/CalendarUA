"""測試完整 Category 功能流程"""
import sys
from PySide6.QtWidgets import QApplication
from CalendarUA import CalendarUA

if __name__ == "__main__":
    print("=== 測試 Category 完整功能 ===\n")
    print("請按照以下步驟測試:\n")
    print("1. 檢查 Preview Tab 右上方是否有「管理 Category」按鈕")
    print("2. 點擊「管理 Category」開啟 Category 管理對話框")
    print("3. 確認已有 8 個系統 Categories (Red, Pink, Light Purple, Green, Blue, Yellow, Orange, Gray)")
    print("4. 嘗試新增一個自訂 Category")
    print("5. 嘗試編輯 Category 顏色")
    print("6. 切換到 Weekly Tab,新增或編輯排程,確認 Category 下拉選單可正常使用")
    print("7. 切換回 Preview Tab,確認排程以正確的 Category 顏色顯示")
    print("\n程式啟動中...\n")
    
    app = QApplication(sys.argv)
    window = CalendarUA()
    window.show()
    sys.exit(app.exec())

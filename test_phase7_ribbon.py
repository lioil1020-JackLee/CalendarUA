"""測試 Phase 7 Ribbon 功能"""
import sys
from PySide6.QtWidgets import QApplication
from CalendarUA import CalendarUA

if __name__ == "__main__":
    print("=== Phase 7 Ribbon 功能測試 ===\n")
    print("請測試以下功能:\n")
    print("【Menu Bar】")
    print("  File 選單:")
    print("    - New Schedule (Ctrl+N)")
    print("    - Refresh (F5)")
    print("    - Load Profile...")
    print("    - Save Profile...")
    print("    - Exit (Ctrl+Q)")
    print("\n  Edit 選單:")
    print("    - Edit Schedule (Ctrl+E) - 目前禁用")
    print("    - Delete Schedule (Del) - 目前禁用")
    print("    - Manage Categories...")
    print("\n  View 選單:")
    print("    - Day View (Ctrl+1)")
    print("    - Week View (Ctrl+2)")
    print("    - Month View (Ctrl+3)")
    print("    - Go to Today (Ctrl+T)")
    print("\n  Tools 選單:")
    print("    - Database Settings...")
    print("\n  Help 選單:")
    print("    - About CalendarUA")
    
    print("\n【Toolbar】")
    print("  - New 按鈕")
    print("  - Edit 按鈕 (目前禁用)")
    print("  - Delete 按鈕 (目前禁用)")
    print("  - Refresh 按鈕")
    print("  - Categories 按鈕")
    print("  - Scheduler 狀態顯示")
    
    print("\n【快速鍵】")
    print("  - Ctrl+N: 新增排程")
    print("  - Ctrl+E: 編輯排程")
    print("  - Del: 刪除排程")
    print("  - F5: 重新載入")
    print("  - Ctrl+1/2/3: 切換視圖")
    print("  - Ctrl+T: 跳到今天")
    print("  - Ctrl+Q: 離開")
    
    print("\n【Profile 功能】")
    print("  1. File > Save Profile 儲存設定檔")
    print("  2. File > Load Profile 載入設定檔")
    print("  3. 確認 General Panel 設定正確載入/儲存")
    
    print("\n程式啟動中...\n")
    
    app = QApplication(sys.argv)
    window = CalendarUA()
    window.show()
    sys.exit(app.exec())

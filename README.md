CalendarUA
1. 專案簡介
CalendarUA 是一款專為工業自動化設計的排程管理系統。本專案採用 Python 3.12 開發，結合了 Office/Outlook 風格的行事曆介面，讓使用者能直觀地設定複雜的週期性任務（如：定期啟停設備、定時更改生產參數），並透過 OPC UA 協定自動執行。

2. 技術棧 (Tech Stack)
程式語言: Python 3.12

UI 框架: PySide6 (Qt for Python) - 打造現代化 Office/Windows 11 視覺風格。

排程邏輯: python-dateutil (遵循 RFC 5545 RRULE 國際標準)。

通訊協定: asyncua (非同步 OPC UA 用戶端)。

**資料庫**: SQLite (內建於 Python，無需額外安裝伺服器，資料儲存在 `./database/calendarua.db`)

3. 專案目錄架構
CalendarUA/
├── CalendarUA.py          # 程式進入點：包含主視窗與所有對話框
├── requirements.txt       # 專案依賴庫清單
├── README.md              # 專案說明文檔
├── ui/
│   └── recurrence_dialog.py # 週期性設定對話框 (仿 Office 週期規則視窗)
├── core/
│   ├── opc_handler.py         # OPC UA 讀寫邏輯與例外處理
│   ├── opc_security_config.py # OPC UA 安全配置定義
│   └── rrule_parser.py        # 邏輯轉換：將 UI 設定轉為 RRULE 字串
└── database/
    └── sqlite_manager.py   # SQLite 資料庫 CRUD 操作

## 4. 資料庫設計 (SQLite Schema)
表格名稱：schedules
欄位名稱 | 型態 | 說明
--- | --- | ---
id | INTEGER PRIMARY KEY AUTOINCREMENT | 唯一識別碼
task_name | TEXT | 任務名稱
opc_url | TEXT | OPC UA 伺服器位址
node_id | TEXT | OPC UA Tag NodeID
target_value | TEXT | 要寫入的數值
rrule_str | TEXT | RRULE 規則字串
opc_security_policy | TEXT | OPC 安全策略 (None/Basic256Sha256 等)
opc_security_mode | TEXT | OPC 安全模式 (None/Sign/SignAndEncrypt)
opc_username | TEXT | OPC 使用者名稱
opc_password | TEXT | OPC 密碼
opc_timeout | INTEGER | 連線超時 (秒)
is_enabled | INTEGER | 是否啟用 (1: 啟用, 0: 停用)
created_at | TIMESTAMP | 建立時間
updated_at | TIMESTAMP | 更新時間

## 5. 主要功能

- 📅 **Office 風格行事曆**：直觀的日期選擇和排程查看
- 🔄 **RRULE 週期設定**：支持每日、每週、每月、每年等複雜週期規則
- 🔒 **OPC UA 安全連線**：支持多種安全策略和驗證方式
- 💾 **SQLite 資料庫**：輕量級本地儲存，無需配置伺服器
- ⚙️ **自動伺服器檢測**：自動偵測 OPC UA 伺服器支持的安全模式
CalendarUA
1. 專案簡介
CalendarUA 是一款專為工業自動化設計的排程管理系統。本專案採用 Python 3.12 開發，結合了 Office/Outlook 風格的行事曆介面，讓使用者能直觀地設定複雜的週期性任務（如：定期啟停設備、定時更改生產參數），並透過 OPC UA 協定自動執行。

2. 技術棧 (Tech Stack)
程式語言: Python 3.12

UI 框架: PySide6 (Qt for Python) - 打造現代化 Office/Windows 11 視覺風格。

排程邏輯: python-dateutil (遵循 RFC 5545 RRULE 國際標準)。

通訊協定: asyncua (非同步 OPC UA 用戶端)。

資料庫: MySQL 8.0+ (儲存排程規則、Tag 配置及執行日誌)。

3. 專案目錄架構
CalendarUA/
├── main.py                # 程式進入點
├── requirements.txt       # 專案依賴庫清單
├── ui/
│   ├── main_window.py     # 主介面：包含日曆檢視與排程清單
│   ├── recurrence_dialog.py # 週期性設定對話框 (仿 Office 週期規則視窗)
│   └── tag_config.py      # OPC UA 連線與 MySQL 參數設定介面
├── core/
│   ├── scheduler.py       # 背景排程引擎：負責檢查時間並觸發任務
│   ├── opc_handler.py     # OPC UA 讀寫邏輯與例外處理
│   └── rrule_parser.py    # 邏輯轉換：將 UI 設定轉為 RRULE 字串
└── database/
    └── mysql_manager.py   # MySQL 資料庫 CRUD 操作與連線池管理

4. 資料庫設計 (MySQL Schema)
表格名稱：schedules
欄位名稱,型態,說明
id,"INT (PK, AI)",唯一識別碼
task_name,VARCHAR(100),任務名稱 (例如：每日早班開機)
opc_url,VARCHAR(255),OPC UA 伺服器位址
node_id,VARCHAR(255),OPC UA Tag NodeID
target_value,VARCHAR(50),要寫入的數值
rrule_str,VARCHAR(500),"RRULE 規則字串 (例如：FREQ=WEEKLY;BYDAY=TU,TH)"
is_enabled,TINYINT(1),"是否啟用該排程 (1: 啟用, 0: 停用)"

5. 開發指令建議 (Copilot Prompts)
在 VS Code 中開發時，建議使用以下指令引導 Copilot：

初始化資料庫: 「請參考 README.md，使用 mysql-connector-python 撰寫 database/mysql_manager.py，需包含建立 schedules 表格的功能。」

建立 UI: 「請使用 PySide6 建立 ui/recurrence_dialog.py，介面風格需模仿 Outlook 的週期性設定，並能回傳 RRULE 字串。」

處理排程: 「請寫一段程式碼使用 python-dateutil 解析 rrule_str，並計算下一次觸發的日期時間。」
## CalendarUA - 工業級 Outlook 風格行事曆＋OPC UA 排程

CalendarUA 是一個以 **Outlook 行事曆** 為靈感、專為 **OPC UA 工業寫值排程** 設計的桌面應用程式，使用 **PySide6** 開發。

核心目標：

- **完整行事曆體驗**：支援 Day / Week / Month 視圖、例外與假日、分類顏色。
- **工業排程整合**：每個排程都是一條 OPC UA 動作（Server URL、NodeId、Value、Datatype）。
- **可擴充農民曆 / 農曆排程**：預留 `core/lunar_calendar.py` 介面，可安裝農曆套件後擴充。

---

### 專案架構總覽

- `CalendarUA.py`  
  主視窗與應用程式入口，負責：
  - 建立主 UI（General / Holidays / Exceptions Tabs，以及其他排程相關視圖）
  - 管理資料庫連線與背景排程執行緒
  - 把資料從 `SQLiteManager` 傳遞到各個 UI Panel

- `core/`
  - `schedule_resolver.py`  
    排程引擎，將：
    - `schedules`（系列排程）
    - `schedule_exceptions`（例外）
    - `holiday_entries`（假日條目）  
    在指定時間範圍內解析成 **實際發生的 Occurrence 列表**（`ResolvedOccurrence`）。
  - `rrule_parser.py`  
    封裝 `dateutil.rrule` 的 RRULE 解析與計算工具，提供：
    - `get_next_trigger(...)`
    - `get_trigger_between(...)`
    - 常用 RRULE 產生器（`create_daily_rule` / `create_weekly_rule` / `create_monthly_rule`）。
  - `schedule_models.py`  
    行事曆 Domain Model dataclass 集中定義，例如：
    - `ScheduleSeries`（一條 OPC UA 排程系列）
    - `ScheduleException`（單次例外）
    - `HolidayCalendar` / `HolidayEntry`（假日日曆與條目）
    - `RuntimeOverride`（Runtime 覆寫狀態）。
  - `opc_handler.py`  
    OPC UA 非同步處理類別，使用 `asyncua`：
    - 建立安全連線（無加密 / 憑證 / 使用者帳密）
    - `write_node(node_id, value, data_type)`：寫值並讀回驗證
    - `read_node(node_id)`、`read_node_data_type(node_id)`。
  - `lunar_calendar.py`  
    **預留**的農曆 / 農民曆工具：
    - 若安裝了農曆套件，可用來做西曆 <-> 農曆轉換、查詢宜/忌等。
    - 若沒有安裝，函式會回傳 `None` / 空資料，不會影響主程式運作。

- `database/`
  - `sqlite_manager.py`  
    SQLite 資料庫存取層，負責：
    - 初始化/遷移資料表（schedules / schedule_exceptions / holiday_* / general_settings / runtime_override / schedule_categories）
    - 所有排程、例外、假日、General 設定、Runtime Override 的 CRUD。

- `ui/`
  - `general_panel.py`：General Tab，管理 Profile 與全域設定。
  - `holidays_panel.py`：Holidays Tab，管理假日日曆與假日條目。
  - `exceptions_panel.py`：Exceptions Tab，類似 Outlook/ScheduleWorX 的例外管理＋日曆視圖。
  - `schedule_canvas.py`：Day / Week 時間格視圖，用 `ResolvedOccurrence` 畫出全天 24h 排程。
  - `month_grid.py`：Month 視圖，每格顯示日期 ＋ 最多 3 個事件「小膠囊」。
  - `weekly_panel.py` / `weekly_event_dialog.py`：週間（FREQ=WEEKLY）班表編輯器。
  - `runtime_panel.py`：Runtime Override 面板，手動覆寫輸出值並顯示下一事件資訊。
  - `recurrence_dialog.py`：排程重複規則對話框。
  - `database_settings_dialog.py`：資料庫設定對話框。
  - `category_manager_dialog.py`：排程 Category 顏色管理。

---

### 資料庫結構概要

詳細欄位說明請見 `docs/db_schema.md`，以下為重點：

- `schedules`
  - 一條記錄 = 一個 **OPC UA 排程系列**：
    - `task_name`: 排程名稱
    - `opc_url`, `node_id`, `target_value`, `data_type`
    - `rrule_str`: RRULE 字串（頻率、時間、重複規則）
    - `category_id`: 類別顏色
    - `is_enabled`: 是否啟用
    - 及一組 OPC UA 安全 / 認證相關欄位（security_policy / username / password / timeout ...）。

- `schedule_exceptions`
  - 管理「取消」或「覆寫」單次 occurrence：
    - `schedule_id`, `occurrence_date`, `action`("cancel"/"override")
    - `override_start`, `override_end`
    - `override_task_name`, `override_target_value`
    - `override_category_id`, `note`。

- `holiday_calendars` / `holiday_entries`
  - 多個假日日曆，每個日曆底下有多個假日：
    - `holiday_date` + `is_full_day` + 可選時段 (`start_time` / `end_time`)
    - 可選覆寫 (`override_category_id` / `override_target_value`)。

- `schedule_categories`
  - 類似 Outlook 類別顏色標籤，定義：
    - `name`, `bg_color`, `fg_color`, `sort_order`, `is_system`。

- `general_settings`
  - 全域 Profile 設定，僅一筆：
    - `profile_name`, `description`
    - `enable_schedule`, `scan_rate`, `refresh_rate`
    - `use_active_period`, `active_from`, `active_to`
    - `output_type`, `refresh_output`, `generate_events`。

- `runtime_override`
  - Runtime Override（最多一筆），優先於所有排程：
    - `override_value`, `override_until`。

---

### 執行方式

1. 建議使用虛擬環境安裝相依套件，例如：

```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
pip install -r requirements.txt
```

2. 啟動主程式：

```bash
python CalendarUA.py
```

首次執行時會在 `database/calendarua.db` 自動建立資料表與預設 Category。

---

### 主要使用流程

- **新增排程**
  - 透過主視窗的「New」或 Weekly Tab、其他排程編輯入口。
  - 填寫：
    - 排程名稱、OPC UA 伺服器 URL、NodeId
    - 目標值（Target Value）與資料型別
    - RRULE 重複規則（每日、每週、每月等）
  - 儲存後即會出現在 Day / Week / Month 視圖與 Exceptions Panel。

- **設定假日與例外**
  - Holidays Tab：管理假日日曆與假日條目，可指定全天或特定時段。
  - Exceptions Tab：為某天的單次 occurrence 新增取消 / 覆寫記錄。

- **Runtime Override**
  - 在 Runtime Panel 輸入覆寫值，選擇有效期間：
    - 永久（手動清除前）、或 30 秒～1 週等選項。
  - 在有效期內，系統輸出將以 Runtime Override 為準。

- **OPC UA 寫值排程**
  - 背景排程執行緒定期檢查：
    - 目前時間附近哪些 occurrence 應觸發
    - 呼叫 `OPCHandler` 進行寫值與驗證。

---

### 農民曆 / 農曆排程擴充

目前專案已在 `core/lunar_calendar.py` 中預留：

- `to_lunar(gregorian: date) -> Optional[LunarDateInfo]`
- `from_lunar(year, month, day, leap=False) -> Optional[date]`
- `get_almanac_info(gregorian: date) -> Dict[str, Any]`

你可以：

1. 安裝適合的農曆 / 農民曆套件（例如 `lunarcalendar` 等）。
2. 根據實際套件 API，補齊 `lunar_calendar.py` 的 TODO 區塊。
3. 在排程編輯對話框中，新增「以農曆設定」的模式，內部再轉換成西曆 RRULE 或多個 RDATE。

---

### 專案設計原則

- **不在 UI 層寫排程邏輯**：  
  所有「什麼時候發生」的邏輯集中在 `core/rrule_parser.py` 與 `core/schedule_resolver.py`。

- **不在資料庫層寫商業邏輯**：  
  `database/sqlite_manager.py` 只做 CRUD 與 Schema 管理，沒有 UI 或排程判斷。

- **模組分界清楚、方便未來擴充**：  
  例如：未來要增加「多行事曆」或「Web/REST API」，只需要：
  - 重用 `schedule_models` 與 `schedule_resolver`
  - 換一層新的 UI 或 API 即可。

# CalendarUA - 工業自動化排程管理系統

[![Python Version](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![PySide6](https://img.shields.io/badge/PySide6-6.7.2-orange.svg)](https://pypi.org/project/PySide6/)

CalendarUA 是一款專為工業自動化領域設計的智慧排程管理系統，採用現代化 GUI 介面，整合 OPC UA 通訊協定與 RRULE 週期規則，提供企業級的自動化任務排程解決方案。

> 目前主介面分頁為：`General`、`Holidays`、`Exceptions`（已整併舊版 Weekly / Preview / Runtime 入口）。

## 📋 專案概述

本系統專為工業環境設計，能夠精確控制生產設備的啟停時序、參數調整等關鍵操作。系統採用模組化架構，支援多種安全認證方式，並提供直觀的 Office 風格使用者介面。

### 🎯 核心特性

- **📅 智慧行事曆介面**：採用類似 Outlook 的視覺化設計，支援多種檢視模式
- **🔄 靈活週期規則**：完整支援 RFC 5545 RRULE 標準，涵蓋每日、每週、每月、每年等複雜排程
- **🔒 企業級安全**：支援 OPC UA 多種安全策略，包括無加密、基本加密、憑證認證等
- **⚡ 非同步處理**：採用 asyncua 實現高效能 OPC UA 通訊
- **💾 輕量級儲存**：內建 SQLite 資料庫，無需額外伺服器配置
- **🎨 現代化 UI**：支援亮色/暗色主題切換，適配 Windows 11 設計語言
- **🔧 系統整合**：支援系統匣圖示、最小化至托盤等桌面應用功能
- **🔄 智慧重試機制**：支援連線逾時重試和寫值失敗重試，可自訂重試延遲時間

## 🏗️ 系統架構

```
CalendarUA/
├── 📁 core/                 # 核心業務邏輯
│   ├── opc_handler.py       # OPC UA 通訊處理器
│   ├── opc_security_config.py # 安全配置管理
│   └── rrule_parser.py      # RRULE 規則解析器
├── 📁 database/             # 資料持久化層
│   └── sqlite_manager.py    # SQLite 資料庫管理器
├── 📁 ui/                   # 使用者介面層
│   ├── recurrence_dialog.py # 週期設定對話框
│   └── database_settings_dialog.py # 資料庫設定介面
├── 📄 CalendarUA.py         # 主應用程式入口
├── 📄 requirements.txt      # Python 依賴清單
└── 📄 pyproject.toml        # 專案配置檔案
```

### 架構說明

- **展示層 (UI Layer)**：基於 PySide6 的現代化圖形介面
- **業務邏輯層 (Business Layer)**：排程計算、規則解析、安全處理
- **資料存取層 (Data Layer)**：SQLite 資料庫操作與資料模型
- **通訊層 (Communication Layer)**：OPC UA 協定實現與網路處理

## 🚀 快速開始

### 系統需求

- **作業系統**：Windows 10/11, Linux, macOS
- **Python 版本**：3.9 或更新版本
- **記憶體**：至少 512MB RAM
- **儲存空間**：100MB 可用磁碟空間

### 安裝步驟

1. **複製專案**
   ```bash
   git clone https://github.com/lioil1020-JackLee/CalendarUA.git
   cd CalendarUA
   ```

2. **建立虛擬環境**
   ```bash
   python -m venv .venv
   # Windows
   .venv\Scripts\activate
   # Linux/macOS
   source .venv/bin/activate
   ```

3. **安裝依賴**
   ```bash
   pip install -r requirements.txt
   ```

4. **啟動應用程式**
   ```bash
   python CalendarUA.py
   ```

### 首次設定

1. 啟動後，系統會自動建立預設資料庫
2. 點擊「資料庫設定」配置資料庫路徑（如需要）
3. 新增您的第一個排程任務

## 🎨 介面預覽

### 主視窗介面
![主視窗](image/主視窗.png)

*主視窗展示行事曆視圖和任務列表，支援拖拽操作和右鍵選單*

### 資料庫設定
![資料庫設定](image/資料庫設定.png)

*資料庫設定介面支援自訂資料庫路徑和統計資訊查看*

## 📖 使用指南

### 建立排程任務

#### 1. 新增排程任務
![新增排程](image/新增排程.png)

*新增排程對話框：設定基本任務資訊*

#### 2. OPC UA 連線設定
![OPC UA 連線設定](image/OPC_UA連線設定.png)

*OPC UA 安全配置：設定連線參數和認證資訊*

#### 3. 瀏覽 OPC UA 節點
![瀏覽 OPC UA 節點](image/瀏覽OPC_UA節點.png)

*OPC UA 節點瀏覽器：選擇要控制的節點*

#### 4. 編輯排程任務
![編輯排程](image/編輯排程.png)

*編輯排程介面：修改現有任務設定*

### 管理現有任務

- **啟用/停用**：右鍵選單快速切換任務狀態
- **編輯任務**：雙擊任務或使用編輯按鈕
- **刪除任務**：選取任務後點擊刪除按鈕
- **檢視日曆**：在左側日曆中查看任務分佈

#### 任務管理介面
![主視窗](image/主視窗.png)

*任務管理介面：支援右鍵選單操作、任務狀態切換和日曆視圖*

### 資料庫管理

系統提供完整的資料庫管理功能：

- **路徑設定**：變更資料庫儲存位置
- **資料備份**：匯出資料庫檔案
- **資料還原**：從備份檔案恢復資料
- **統計資訊**：查看任務數量、執行記錄等

## 🔧 技術規格

### 支援的 RRULE 參數

| 參數 | 說明 | 支援狀態 |
|------|------|----------|
| FREQ | 重複頻率 (DAILY/WEEKLY/MONTHLY/YEARLY) | ✅ 完整支援 |
| INTERVAL | 重複間隔 | ✅ 支援 |
| BYDAY | 星期幾指定 | ✅ 支援 |
| BYMONTHDAY | 月份日期 | ✅ 支援 |
| BYMONTH | 月份 | ✅ 支援 |
| BYSETPOS | 集合位置 | ✅ 支援 |
| BYHOUR/BYMINUTE | 時間設定 | ✅ 支援 |
| COUNT | 重複次數 | ✅ 支援 |
| UNTIL | 結束日期 | ✅ 支援 |
| DTSTART | 開始時間 | ✅ 支援 |
| DURATION | 持續時間 (自訂) | ✅ 支援 |

### 週期規則介面預覽

系統提供直觀的週期設定介面，支援多種重複模式：

#### 每日排程設定
![每日排程](image/週期性約會-每天.png)

#### 每週排程設定
![每週排程](image/週期性約會-每週.png)

#### 每月排程設定
![每月排程](image/週期性約會-每月.png)

#### 每年排程設定
![每年排程](image/週期性約會-每年.png)

### OPC UA 安全支援

| 安全策略 | 說明 | 支援狀態 |
|----------|------|----------|
| None | 無加密 | ✅ 支援 |
| Basic256Sha256 | 256位元加密 | ✅ 支援 |
| Basic128Rsa15 | 128位元加密 | ❌ 不支援 |
| Basic256 | 256位元加密 | ❌ 不支援 |

### 認證方式

- **匿名認證**：適用於測試環境
- **使用者名稱/密碼**：標準認證方式
- **X.509 憑證**：企業級安全認證

### 重試機制

| 參數 | 說明 | 預設值 |
|------|------|--------|
| 連線超時 | OPC UA 連線逾時時間 | 5 秒 |
| 寫值重試延遲 | 寫值失敗後重試間隔 | 3 秒 |
| 重試策略 | 根據持續時間決定重試行為 | 持續時間 > 0 分鐘時重試 |

## 📊 資料庫結構

### schedules 表格

| 欄位名稱 | 資料型態 | 說明 | 約束 |
|----------|----------|------|------|
| id | INTEGER | 主鍵 | PRIMARY KEY AUTOINCREMENT |
| task_name | TEXT | 任務名稱 | NOT NULL |
| opc_url | TEXT | OPC UA 伺服器位址 | NOT NULL |
| node_id | TEXT | OPC UA 節點 ID | NOT NULL |
| target_value | TEXT | 目標數值 | NOT NULL |
| rrule_str | TEXT | RRULE 規則字串 | NOT NULL |
| opc_security_policy | TEXT | 安全策略 | DEFAULT 'None' |
| opc_security_mode | TEXT | 安全模式 | DEFAULT 'None' |
| opc_username | TEXT | 使用者名稱 |  |
| opc_password | TEXT | 密碼 |  |
| opc_timeout | INTEGER | 連線超時(秒) | DEFAULT 5 |
| opc_write_timeout | INTEGER | 寫值重試延遲(秒) | DEFAULT 3 |
| is_enabled | INTEGER | 啟用狀態 | DEFAULT 1 |
| created_at | TIMESTAMP | 建立時間 | DEFAULT CURRENT_TIMESTAMP |
| updated_at | TIMESTAMP | 更新時間 | DEFAULT CURRENT_TIMESTAMP |

## 🔍 故障排除

### 常見問題

**Q: 應用程式啟動失敗**
A: 確認 Python 版本 ≥ 3.9，並已安裝所有依賴套件

**Q: OPC UA 連線失敗**
A: 檢查伺服器位址、網路連線和安全設定。連線逾時預設為 5 秒，可在任務設定中調整

**Q: 排程未執行**
A: 確認任務已啟用，且 RRULE 規則設定正確

**Q: 寫值失敗**
A: 檢查 Node ID 是否正確，寫值重試延遲預設為 3 秒，可在任務設定中調整

**Q: 資料庫錯誤**
A: 檢查資料庫檔案權限，或嘗試重新建立資料庫

### 記錄與除錯

系統預設記錄等級為 WARNING。如需詳細記錄，請修改 `CalendarUA.py` 中的記錄設定：

```python
logging.basicConfig(level=logging.DEBUG)
```

## 🤝 貢獻指南

歡迎參與專案開發！請遵循以下步驟：

1. Fork 此專案
2. 建立功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交變更 (`git commit -m 'Add some AmazingFeature'`)
4. 推送至分支 (`git push origin feature/AmazingFeature`)
5. 開啟 Pull Request

### 開發環境設定

```bash
# 安裝開發依賴
pip install -e ".[dev]"

# 執行測試
python -m pytest

# 建置可執行檔案
pyinstaller CalendarUA.spec
```

## 📄 授權條款

本專案採用 MIT 授權條款 - 詳見 [LICENSE](LICENSE) 檔案

## 📞 聯絡資訊

- **專案維護者**: [lioil1020-JackLee](https://github.com/lioil1020-JackLee)
- **問題回報**: [GitHub Issues](https://github.com/lioil1020-JackLee/CalendarUA/issues)
- **專案首頁**: [GitHub Repository](https://github.com/lioil1020-JackLee/CalendarUA)

---

**注意**: 本系統適用於工業自動化環境，請在專業人員指導下使用。如有特殊需求或客製化需求，請聯絡開發團隊。
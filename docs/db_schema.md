## CalendarUA 資料庫結構說明（SQLite）

資料庫檔案預設為 `database/calendarua.db`。  
本文件描述主要資料表的結構與用途，方便你之後維護或擴充（例如新增欄位、寫報表、外部系統整合）。

---

### 1. `schedules` 表 - 排程系列（含 OPC UA 寫值設定）

**用途**：一列代表一條排程「系列」：  
何時（RRULE）、對哪台 OPC UA Server、寫什麼值到哪個 NodeId。

**主要欄位**

- `id INTEGER PRIMARY KEY AUTOINCREMENT`  
  排程 ID。
- `task_name TEXT NOT NULL`  
  排程名稱（顯示在行事曆上的標題）。
- `opc_url TEXT NOT NULL`  
  OPC UA 伺服器 URL（例如 `opc.tcp://localhost:4840`）。
- `node_id TEXT NOT NULL`  
  要寫值的 NodeId（如 `ns=2;i=1001` 或 `ns=2;s=TagName`）。
- `target_value TEXT NOT NULL`  
  寫入的目標值（以文字儲存，實際寫值時再轉型）。
- `data_type TEXT DEFAULT 'auto'`  
  目標值型別：
  - `auto` / `int` / `float` / `string` / `bool`
- `rrule_str TEXT NOT NULL`  
  RRULE 字串，描述重複規則（FREQ, BYHOUR, BYMINUTE, BYDAY...）。
- `opc_security_policy TEXT DEFAULT 'None'`  
  OPC UA 安全策略。
- `opc_security_mode TEXT DEFAULT 'None'`  
  安全模式（None / Sign / SignAndEncrypt）。
- `opc_username TEXT DEFAULT ''` / `opc_password TEXT DEFAULT ''`  
  OPC UA 使用者／密碼。
- `opc_timeout INTEGER DEFAULT 10`  
  連線超時秒數。
- `opc_write_timeout INTEGER DEFAULT 3`  
  寫值重試延遲秒數。
- `is_enabled INTEGER DEFAULT 1`  
  是否啟用（1=啟用, 0=停用）。
- `category_id INTEGER DEFAULT 1`  
  類別顏色 ID（對應 `schedule_categories.id`）。
- `priority INTEGER DEFAULT 1`  
  優先權（暫未大量使用，可做排序或權重）。
- `location TEXT DEFAULT ''` / `description TEXT DEFAULT ''`  
  類似 Outlook 的地點與描述欄位。
- `last_execution_status TEXT` / `last_execution_time TIMESTAMP`  
  最近一次執行結果與時間。
- `next_execution_time TIMESTAMP`  
  預先計算出的下一次觸發時間（可選，用於加速查詢）。
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`  
  建立時間。
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`  
  最近更新時間。

**索引**

- `idx_schedules_enabled`：`(is_enabled)`
- `idx_schedules_node_id`：`(node_id)`
- `idx_schedules_category`：`(category_id)`
- `idx_schedules_next_time`：`(next_execution_time)`，優化依「下一次執行時間」掃描排程的查詢效能

---

### 2. `schedule_exceptions` 表 - 排程例外

**用途**：管理每一條排程在「某個日期」上的例外行為：

- `action = 'cancel'`：取消這一天的 occurrence。
- `action = 'override'`：覆寫這一天的時間 / 標題 / 目標值 / 類別。

**主要欄位**

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `schedule_id INTEGER NOT NULL`  
  對應 `schedules.id`，`ON DELETE CASCADE`。
- `occurrence_date TEXT NOT NULL`  
  例外發生日（格式：`YYYY-MM-DD`）。
- `action TEXT NOT NULL DEFAULT 'override'`  
  `"cancel"` 或 `"override"`。
- `override_start TEXT` / `override_end TEXT`  
  覆寫後的開始 / 結束時間（ISO 8601 字串）。  
  只有 `action='override'` 時有效。
- `override_task_name TEXT`  
  覆寫後的標題（若為空則沿用原本 `task_name`）。
- `override_target_value TEXT`  
  覆寫後的目標值。
- `override_category_id INTEGER`  
  覆寫後的 Category 顏色 ID。
- `note TEXT DEFAULT ''`  
  備註說明。
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

**索引**

- `idx_schedule_exceptions_schedule_date (schedule_id, occurrence_date)`
- `idx_exceptions_category (override_category_id)`

---

### 3. `holiday_calendars` 表 - 假日日曆清單

**用途**：管理多個假日日曆（例如「台灣國定假日」、「公司休假」）。

**主要欄位**

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `name TEXT UNIQUE NOT NULL`  
  日曆名稱。
- `description TEXT`  
  說明。
- `is_default INTEGER DEFAULT 0`  
  是否為預設假日日曆。
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

---

### 4. `holiday_entries` 表 - 假日條目

**用途**：實際的假日資料，一列對應一個日期（可全天或時段）。

**主要欄位**

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `calendar_id INTEGER NOT NULL`  
  對應 `holiday_calendars.id`，`ON DELETE CASCADE`。
- `holiday_date TEXT NOT NULL`  
  假日日期（`YYYY-MM-DD`）。
- `name TEXT NOT NULL`  
  假日名稱（例如：春節、國慶日）。
- `is_full_day INTEGER DEFAULT 1`  
  是否全天假日（1=全天, 0=僅時段）。
- `start_time TEXT` / `end_time TEXT`  
  時段假日的開始/結束時間（`HH:MM[:SS]`）。
- `override_category_id INTEGER`  
  若有設定，該日的排程會以此 Category 顏色顯示。
- `override_target_value TEXT`  
  若有設定，可覆寫該日排程的目標值。
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

**索引**

- `idx_holiday_entries_calendar_date (calendar_id, holiday_date)`
- `idx_holiday_entries_category (override_category_id)`

---

### 5. `schedule_categories` 表 - 類別顏色

**用途**：類似 Outlook 的分類顏色，主要用於排程與假日的視覺標示。

**主要欄位**

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `name TEXT UNIQUE NOT NULL`  
  類別名稱（例如：`Red (關閉)`、`Pink (自動)`）。
- `bg_color TEXT NOT NULL`  
  背景色（Hex，如 `#FF0000`）。
- `fg_color TEXT NOT NULL`  
  前景色（文字顏色）。
- `sort_order INTEGER DEFAULT 0`  
  排序用。
- `is_system INTEGER DEFAULT 0`  
  是否為系統預設類別（不可刪除）。
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

---

### 6. `general_settings` 表 - 全域設定 / Profile

**用途**：一般只會有一筆，描述系統層級設定與 Profile。

**主要欄位**

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `profile_name TEXT DEFAULT '預設 Profile'`
- `description TEXT`
- `enable_schedule INTEGER DEFAULT 1`  
  是否啟用排程執行緒。
- `scan_rate INTEGER DEFAULT 1`  
  排程掃描間隔（秒）。
- `refresh_rate INTEGER DEFAULT 5`  
  UI 資料更新頻率（秒）。
- `use_active_period INTEGER DEFAULT 0`  
  是否限制排程只在某段期間有效。
- `active_from TEXT` / `active_to TEXT`  
  有效期間（ISO 8601 字串）。
- `output_type TEXT DEFAULT 'OPC UA Write'`  
  預設輸出類型。
- `refresh_output INTEGER DEFAULT 1`
- `generate_events INTEGER DEFAULT 1`
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

---

### 7. `runtime_override` 表 - Runtime 覆寫

**用途**：儲存目前生效中的 Runtime Override（最多一筆）。

**主要欄位**

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `override_value TEXT NOT NULL`  
  覆寫輸出的值（字串，實際使用時再轉型）。
- `override_until TEXT`  
  覆寫有效期限（ISO 8601 字串），為 `NULL` 表示永久，需手動清除。
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

---

### 遷移與版本相容性

`database/sqlite_manager.py` 中的 `_migrate_db()` 會在啟動時自動：

- 檢查並補上新欄位（例如 security / category / description 等）。
- 建立缺少的表與索引。

因此：

- **直接升級程式碼**：舊版 DB 仍可被自動「遷移」到最新結構。
- 若你自行修改 Schema，建議同步更新 `_migrate_db()`，或建立新的遷移函式。


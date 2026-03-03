## CalendarUA 資料庫結構（SQLite）

預設資料庫檔案：`database/calendarua.db`

本文件描述目前主要資料表、重要欄位，以及與 UI 行為的對應關係。

---

### 1) schedules（排程主表）

一筆代表一條排程系列。

主要欄位：

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `task_name TEXT NOT NULL`
- `opc_url TEXT NOT NULL`
- `node_id TEXT NOT NULL`
- `target_value TEXT NOT NULL`
- `data_type TEXT DEFAULT 'auto'`
- `rrule_str TEXT NOT NULL`
- `opc_security_policy TEXT DEFAULT 'None'`
- `opc_security_mode TEXT DEFAULT 'None'`
- `opc_username TEXT DEFAULT ''`
- `opc_password TEXT DEFAULT ''`
- `opc_timeout INTEGER DEFAULT 5`
- `opc_write_timeout INTEGER DEFAULT 3`
- `lock_enabled INTEGER DEFAULT 0`
- `is_enabled INTEGER DEFAULT 1`
- `ignore_holiday INTEGER DEFAULT 0`
- `category_id INTEGER DEFAULT 1`
- `priority INTEGER DEFAULT 1`
- `location TEXT DEFAULT ''`
- `description TEXT DEFAULT ''`
- `last_execution_status TEXT`
- `last_execution_time TIMESTAMP`
- `next_execution_time TIMESTAMP`
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

常用索引：

- `idx_schedules_enabled (is_enabled)`
- `idx_schedules_node_id (node_id)`
- `idx_schedules_next_time (next_execution_time)`

執行期注意事項（v3.0.0）：

- 執行器於每次輪詢都會重新讀取該筆 `schedules`，因此此表的核心欄位更新可立即生效（不需重啟程式）。
- 若 `is_enabled` 於執行中被改為 `0`，當前任務會在下一次輪詢時停止。

---

### 2) schedule_exceptions（單次例外）

用於覆寫或取消特定日期 occurrence。

主要欄位：

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `schedule_id INTEGER NOT NULL`（FK -> `schedules.id`，`ON DELETE CASCADE`）
- `occurrence_date TEXT NOT NULL`（`YYYY-MM-DD`）
- `action TEXT NOT NULL DEFAULT 'override'`（`cancel` / `override`）
- `override_start TEXT`
- `override_end TEXT`
- `override_task_name TEXT`
- `override_target_value TEXT`
- `override_category_id INTEGER`
- `note TEXT DEFAULT ''`
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

常用索引：

- `idx_schedule_exceptions_schedule_date (schedule_id, occurrence_date)`

---

### 3) holidays（假日規則單表）

整合每週假日與固定月日假日。

主要欄位：

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `entry_type TEXT NOT NULL`（`weekday` / `date`）
- `calendar_type TEXT`（`solar` / `lunar`，僅 `date` 使用）
- `month INTEGER`
- `day INTEGER`
- `weekday INTEGER`（1=週一 ... 7=週日）
- `name TEXT DEFAULT ''`
- `override_target_value TEXT`
- `is_enabled INTEGER DEFAULT 1`
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

---

執行期注意事項（v3.0.0）：

- 執行器每次輪詢會即時查詢是否命中假日規則。
- 若命中規則且 `override_target_value` 非空，該次 OPC 寫值會使用覆寫值。
- 排程 `ignore_holiday = 1` 時，忽略假日覆寫。

---

### 4) schedule_categories（分類顏色）

提供排程顯示的分類與顏色配置。

主要欄位：

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `name TEXT UNIQUE NOT NULL`
- `bg_color TEXT NOT NULL`
- `fg_color TEXT NOT NULL`
- `sort_order INTEGER DEFAULT 0`
- `is_system INTEGER DEFAULT 0`
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

---

### 5) general_settings（全域設定）

通常只有一筆，用於保存 UI/系統層級設定。

主要欄位：

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `profile_name TEXT DEFAULT '預設 Profile'`
- `description TEXT`
- `enable_schedule INTEGER DEFAULT 1`
- `scan_rate INTEGER DEFAULT 1`
- `refresh_rate INTEGER DEFAULT 5`
- `use_active_period INTEGER DEFAULT 0`
- `active_from TEXT`
- `active_to TEXT`
- `output_type TEXT DEFAULT 'OPC UA Write'`
- `refresh_output INTEGER DEFAULT 1`
- `generate_events INTEGER DEFAULT 1`
- `last_opc_url TEXT DEFAULT ''`
- `last_opc_security_policy TEXT DEFAULT 'None'`
- `last_opc_security_mode TEXT DEFAULT 'None'`
- `last_opc_username TEXT DEFAULT ''`
- `last_opc_password TEXT DEFAULT ''`
- `last_opc_timeout INTEGER DEFAULT 5`
- `last_opc_write_timeout INTEGER DEFAULT 3`
- `time_scale_minutes INTEGER DEFAULT 60`  ← 新增
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

`time_scale_minutes` 用途：

- 保存日/週視圖左側時間軸的 Time Scale 設定
- 允許值：`5, 6, 10, 15, 30, 60`
- 啟動或切換專案資料庫時會載入並套用

---

### 6) runtime_override（執行期覆寫）

保存目前生效中的 runtime override。

主要欄位：

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `override_value TEXT NOT NULL`
- `override_until TEXT`
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

---

## 遷移策略

`database/sqlite_manager.py` 的 `_migrate_db()` 會在啟動時自動補齊缺欄位與缺表。

本次更新已納入：

- 若 `general_settings` 缺少 `time_scale_minutes`，會自動 `ALTER TABLE` 補上（預設 `60`）。

另補充執行行為（非 schema migration）：

- Scheduler 在排程或假日設定變更後可被重啟，以確保新規則立即接手。

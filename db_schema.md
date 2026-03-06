## CalendarUA 資料庫結構（SQLite）

預設資料庫檔案：`database/calendarua.db`

本文件以 `database/sqlite_manager.py` 的實際建表與遷移邏輯為準。

## 1. `schedules`（排程主表）

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
- `opc_timeout INTEGER DEFAULT 10`（遷移邏輯可能補成 5，依既有 DB 而定）
- `opc_write_timeout INTEGER DEFAULT 3`
- `lock_enabled INTEGER DEFAULT 0`
- `is_enabled INTEGER DEFAULT 1`
- `ignore_holiday INTEGER DEFAULT 0`
- `priority INTEGER DEFAULT 1`
- `location TEXT DEFAULT ''`
- `description TEXT DEFAULT ''`
- `last_execution_status TEXT DEFAULT ''`
- `last_execution_time TIMESTAMP`
- `next_execution_time TIMESTAMP`
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

索引：

- `idx_schedules_enabled (is_enabled)`
- `idx_schedules_node_id (node_id)`
- `idx_schedules_next_time (next_execution_time)`

## 2. `schedule_exceptions`（單次例外）

覆寫或取消特定 occurrence。

主要欄位：

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `schedule_id INTEGER NOT NULL`（FK -> `schedules.id`，`ON DELETE CASCADE`）
- `occurrence_date TEXT NOT NULL`（`YYYY-MM-DD`）
- `action TEXT NOT NULL DEFAULT 'override'`（`cancel` / `override`）
- `override_start TEXT`
- `override_end TEXT`
- `override_task_name TEXT`
- `override_target_value TEXT`
- `note TEXT DEFAULT ''`
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

索引：

- `idx_schedule_exceptions_schedule_date (schedule_id, occurrence_date)`

## 3. `holidays`（假日規則單表）

整合每週假日與固定月日假日。

主要欄位：

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `entry_type TEXT NOT NULL`（`weekday` / `date`）
- `calendar_type TEXT`（`solar` / `lunar`，僅 `entry_type='date'` 使用）
- `month INTEGER`
- `day INTEGER`
- `weekday INTEGER`（1=週一 ... 7=週日）
- `name TEXT DEFAULT ''`
- `override_target_value TEXT`
- `is_enabled INTEGER DEFAULT 1`
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

唯一索引：

- `idx_holidays_unique_weekday (entry_type, weekday) WHERE entry_type='weekday'`
- `idx_holidays_unique_date (entry_type, calendar_type, month, day) WHERE entry_type='date'`

一般索引：

- `idx_holidays_enabled (is_enabled)`

## 4. `general_settings`（全域設定）

通常只有一筆，保存 UI/系統層設定。

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
- `time_scale_minutes INTEGER DEFAULT 60`
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

`time_scale_minutes` 允許值：`5, 6, 10, 15, 30, 60`

## 5. `runtime_override`（執行期覆寫）

主要欄位：

- `id INTEGER PRIMARY KEY AUTOINCREMENT`
- `override_value TEXT NOT NULL`
- `override_until TEXT`
- `created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`
- `updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP`

## 6. 遷移說明

`SQLiteManager.init_db()` 會先建表，再呼叫 `_migrate_db()`：

- 補齊舊版缺欄位（例如 `priority`、`location`、`description`、`note`）。
- 將舊版 `holiday_entries` 資料轉入新 `holidays` 單表。
- 若假日規則為空，補入預設週假日與常見國/農曆固定日期。

建議：升級版本時不要手動刪表，讓啟動遷移自動補齊即可。

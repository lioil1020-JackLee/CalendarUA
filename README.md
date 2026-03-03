# CalendarUA

![主視窗 月視圖](image/main_month.png)
![主視窗 週視圖](image/main_week.png)
![主視窗 日視圖](image/main_day.png)
![假日設定](image/holidays.png)
![排程設定頁](image/schedules.png)
![OPC UA 連線設定頁](image/opc_connect.png)
![OPC UA 瀏覽頁](image/opc_browser.png)
![資料庫結構](image/db_structure.png)

CalendarUA 是以 `PySide6` 開發的 OPC UA 排程桌面工具，提供日/週/月視圖、RRULE 週期排程、例外與假日規則、以及 SQLite 專案資料庫。

## 主要功能

- 日 / 週 / 月視圖與右鍵操作
- 日視圖 / 週視圖左側時間軸支援 Time Scale 切換（`5 / 6 / 10 / 15 / 30 / 60 min`）
- Time Scale 設定寫入資料庫（`general_settings.time_scale_minutes`），下次開啟專案會保留
- RRULE 週期排程（含 `DURATION`）
- 排程例外（取消或覆寫單次 occurrence）
- 假日規則（每週假日 + 國曆/農曆月日）
- OPC UA 連線與寫值（安全模式、帳密、逾時）
- 排程 `忽略假日` 開關
- 執行中動態套用排程設定（不需重啟程式）
- 假日規則動態套用到執行流程（含 override target value）
- 啟動時可回補期間內舊觸發點（Lock 期間）

## 專案結構

```text
CalendarUA/
├─ CalendarUA.py
├─ core/
│  ├─ lunar_calendar.py
│  ├─ opc_handler.py
│  ├─ rrule_parser.py
│  ├─ schedule_models.py
│  └─ schedule_resolver.py
├─ database/
│  └─ sqlite_manager.py
├─ ui/
│  ├─ database_settings_dialog.py
│  ├─ holiday_settings_dialog.py
│  ├─ month_grid.py
│  ├─ recurrence_dialog.py
│  └─ schedule_canvas.py
├─ db_schema.md
├─ rrule.md
├─ requirements.txt
├─ pyproject.toml
├─ CalendarUA-onedir.spec
└─ CalendarUA-onefile.spec
```

## 環境需求

- Python `>= 3.10`
- 建議使用 `uv`
- Windows（目前主要開發/測試環境）

## 安裝與執行

### 使用 uv（建議）

```bash
uv sync --dev
uv run python CalendarUA.py
```

### 使用 pip

```bash
pip install -r requirements.txt
python CalendarUA.py
```

## 行事曆視圖補充

### Time Scale（新增）

- 在日視圖或週視圖，對左側時間軸（垂直時間欄）按右鍵
- 可選：`5min`、`6min`、`10min`、`15min`、`30min`、`60min`
- 設定後會立即重繪時間刻度
- 日/週視圖會同步刻度
- 刻度會保存到資料庫，重啟或重開專案後維持原設定

### 右鍵新增/編輯的時間精度

- 右鍵時間格新增/編輯時，會帶入「小時 + 分鐘」預設值
- 分鐘值會跟隨目前 Time Scale（例如 15 分刻度可帶入 `08:15`）

## 排程設定頁變更

- 「排程時間 → 期間」下拉選單已移除 `0 分`
- 目前最短期間為 `5 分`

## 執行期行為（v3.0.0）

### 1) 排程設定立即生效（Hot Reload）

排程進入執行迴圈後，會於每次輪詢重新讀取資料庫中的該筆排程設定，因此下列欄位修改後可立即影響執行：

- `lock_enabled`
- `target_value`
- `data_type`
- `rrule_str`（影響期間長度）
- `opc_write_timeout`
- `node_id`
- `is_enabled`
- `opc_url` / `opc_security_policy` / `opc_username` / `opc_password` / `opc_timeout`

其中 OPC 連線相關欄位若改變，執行器會自動重新建立連線後繼續執行。

### 2) 假日規則立即生效

- 開啟主視窗 `假日` 設定並儲存後，會立即重啟 scheduler worker。
- 執行中的任務於每次輪詢會重新判定當天是否為假日。
- 若命中假日規則且該規則有 `override_target_value`，本次實際寫值會改用覆寫值。
- 若排程勾選 `ignore_holiday`，則跳過假日覆寫，維持排程原始 `target_value`。

### 3) 啟動回補（Catch-up）

- 程式啟動後，若排程具有 `DURATION` 且目前時間落在期間內，會回補最近一次觸發點。
- 避免同一 occurrence 重複觸發：以 occurrence 錨點時間去重。

### 4) 優雅結束

- `main()` 已加入 `KeyboardInterrupt` 處理。
- 在終端按 `Ctrl+C` 時，程式會優雅退出，不再顯示大量 traceback。

## 假日設定

主視窗右上 `假日` 按鈕可開啟假日設定。

- 每週假日：週一到週日 checkbox
- 日期假日：表格右鍵新增 / 編輯 / 刪除
- 日期型規則：
  - 國曆（月/日）
  - 農曆（月/日）
- CSV 匯入/匯出欄位：
  - `entry_type,calendar_type,month,day,weekday`

## 資料庫

- 預設檔案：`database/calendarua.db`
- 主要資料表：
  - `schedules`
  - `schedule_exceptions`
  - `holidays`
  - `general_settings`
  - `runtime_override`

詳細請見 `db_schema.md`。

## 打包（PyInstaller）

### onedir

```bash
uv run pyinstaller --clean --noconfirm CalendarUA-onedir.spec
```

輸出資料夾：`dist/CalendarUA-onedir/`

### onefile

```bash
uv run pyinstaller --clean --noconfirm CalendarUA-onefile.spec
```

輸出檔案：`dist/CalendarUA-onefile.exe`

## 常用開發指令

```bash
# 語法檢查
uv run python -m py_compile CalendarUA.py core/*.py database/*.py ui/*.py

# 更新依賴
uv sync --dev
```

## 相關文件

- RRULE 說明：`rrule.md`
- 資料庫結構：`db_schema.md`

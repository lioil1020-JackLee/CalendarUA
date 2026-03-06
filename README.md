# CalendarUA

CalendarUA 是一套以 PySide6 開發的 OPC UA 排程桌面工具，提供日/週/月行事曆視圖、RRULE 週期排程、假日規則與排程例外，並以 SQLite 儲存設定與歷程資料。

![主視窗 月視圖](image/main_month.png)
![主視窗 週視圖](image/main_week.png)
![主視窗 日視圖](image/main_day.png)
![假日設定](image/holidays.png)
![排程設定頁](image/schedules.png)
![OPC UA 連線設定頁](image/opc_connect.png)
![OPC UA 瀏覽頁](image/opc_browser.png)

## 1. 核心功能

- 三種視圖：日視圖、週視圖、月視圖，支援快速切換與右鍵操作。
- 週期排程：支援 `DAILY/WEEKLY/MONTHLY/YEARLY`，可搭配 `UNTIL/COUNT` 與 `DURATION`。
- 排程例外：可取消單次 occurrence，或覆寫單次時間/標題/目標值。
- 假日規則：支援每週假日、國曆固定月日、農曆固定月日。
- 農曆顯示：主月曆與排程 UI 可顯示農曆日期文字。
- OPC UA 寫值：支援 `None/Basic256Sha256/Aes128Sha256RsaOaep/Aes256Sha256RsaPss`、帳密、憑證與連線逾時設定。
- 安全模式偵測：可針對伺服器自動檢測可用安全策略與模式，降低配置錯誤。
- Runtime 覆寫：可設定短期覆寫值與有效期限。
- 主題支援：亮色、暗色、跟隨系統。

## 2. 專案結構

```text
CalendarUA/
|- CalendarUA.py                  # 主程式與主視窗
|- core/
|  |- opc_handler.py              # OPC UA 連線、讀寫、型別處理
|  |- rrule_parser.py             # RRULE 解析與觸發計算（含農曆模式）
|  |- schedule_resolver.py        # occurrence 展開、例外與假日套用
|  |- lunar_calendar.py           # 農曆轉換與顯示文字
|- database/
|  |- sqlite_manager.py           # SQLite 初始化、遷移、CRUD
|- ui/
|  |- schedule_canvas.py          # 日/週視圖
|  |- month_grid.py               # 月視圖
|  |- recurrence_dialog.py        # 週期規則編輯器
|  |- holiday_settings_dialog.py  # 假日規則設定
|  |- database_settings_dialog.py # 資料庫設定
|  |- app_icon.py                 # 共用程式圖示 helper
|- db_schema.md                   # 資料庫結構文件
|- rrule.md                       # RRULE 規則文件
|- requirements.txt
|- pyproject.toml
|- CalendarUA-onedir.spec
`- CalendarUA-onefile.spec
```

## 3. 執行需求

- Python `>= 3.10`
- 主要開發/測試平台：Windows
- 依賴套件：`PySide6`、`asyncua`、`python-dateutil`、`qasync`、`lunardate`

## 4. 安裝與啟動

### 使用 `uv`（建議）

```bash
uv sync --dev
uv run python CalendarUA.py
```

### 使用 `pip`

```bash
pip install -r requirements.txt
python CalendarUA.py
```

## 5. 操作說明

### 5.1 首次啟動

1. 啟動後系統會初始化 `database/calendarua.db`。
2. 若舊版資料庫缺欄位，會自動執行遷移。
3. 會自動建立預設假日規則（週六、週日與常見國/農曆日期）。

### 5.2 新增排程

1. 在主畫面新增排程。
2. 設定 OPC UA 連線資訊與 `node_id`、`target_value`、`data_type`。
3. 開啟循環設定視窗，配置 RRULE（頻率、間隔、週期條件、結束條件）。
4. 儲存後，排程會由背景掃描程序自動判斷觸發並寫入 OPC UA。

### 5.3 編輯/刪除排程

- 可直接編輯排程主要欄位，或停用排程（`is_enabled=0`）。
- 刪除排程時，對應 `schedule_exceptions` 會一併清除（FK cascade）。

### 5.4 例外與假日

- 例外（`schedule_exceptions`）：
  - `cancel`：取消單次觸發
  - `override`：覆寫該次時間/名稱/目標值
- 假日（`holidays`）：
  - `weekday`：每週固定星期
  - `date + solar`：固定國曆月日
  - `date + lunar`：固定農曆月日
- 若排程勾選 `ignore_holiday=1`，該排程不受假日規則影響。

### 5.5 OPC UA 安全連線

- 可設定安全策略、安全模式、帳密、憑證路徑。
- 可使用「僅顯示伺服器支援模式」降低設定失敗機率。
- 連線後會做 session ready 握手檢查，避免剛連線即寫值造成錯誤。

### 5.6 視圖與時間刻度

- 日/週視圖支援 `5/6/10/15/30/60` 分鐘時間刻度。
- 月視圖、導覽月曆支援農曆文字顯示。
- UI 可在亮色/暗色/系統模式間切換。

## 6. 打包

### onedir

```bash
uv run pyinstaller --clean --noconfirm CalendarUA-onedir.spec
```

輸出：`dist/CalendarUA-onedir/`

### onefile

```bash
uv run pyinstaller --clean --noconfirm CalendarUA-onefile.spec
```

輸出：`dist/CalendarUA-onefile.exe`

## 7. 常用開發指令

```bash
# 語法檢查
python -m py_compile CalendarUA.py core/*.py database/*.py ui/*.py

# 重新安裝依賴
pip install -r requirements.txt
```

## 8. 疑難排解

- 無法連到 OPC UA：
  - 檢查 `opc_url`、憑證、帳密、安全策略是否與伺服器一致。
  - 降低安全級別或先以 `None` 驗證連通性。
- 排程未觸發：
  - 檢查 `is_enabled`、`RRULE`、系統時間、`ignore_holiday` 與假日規則。
- 農曆顯示為空：
  - 確認 `lunardate` 已安裝。

## 9. 相關文件

- 資料庫結構：`db_schema.md`
- RRULE 與自訂參數：`rrule.md`

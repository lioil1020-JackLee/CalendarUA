# CalendarUA

![主視窗 月視圖](image/main_month.png)
![主視窗 週視圖](image/main_week.png)
![主視窗 日視圖](image/main_day.png)
![假日設定](image/holidays.png)
![排程設定頁](image/schedules.png)
![OPC UA 連線設定頁](image/opc_connect.png)
![OPC UA 瀏覽頁](image/opc_browser.png)

CalendarUA 是以 PySide6 開發的 OPC UA 排程桌面工具，支援日/週/月視圖、RRULE 週期排程、排程例外與假日規則，以及 SQLite 專案資料庫。

## v4.0.0 重點

- 下拉元件滾輪行為全面整理：
  - 主視窗左上「年/月」採選單項目移動
  - 循環設定中的數值下拉（自訂元件）改為與年下拉一致的選單視窗平移邏輯
  - 每月模式中兩個 1~12 欄位改為固定下拉清單（不再依賴滾輪遞增）
- 內嵌週期設定 Esc 行為修正：在「編輯排程」中按一次 Esc 可直接完整關閉視窗
- UI 細節與主題樣式一致性調整

## 功能摘要

- 日 / 週 / 月視圖與右鍵操作
- 日視圖 / 週視圖 Time Scale 切換（5 / 6 / 10 / 15 / 30 / 60 分）
- RRULE 週期排程（含 DURATION）
- 排程例外（取消或覆寫單次 occurrence）
- 假日規則（每週假日 + 國曆/農曆月日）
- OPC UA 連線與寫值（安全模式、帳密、逾時）
- 排程「忽略假日」開關
- 執行中動態套用排程設定（不需重啟）

## 專案結構

```text
CalendarUA/
├─ CalendarUA.py
├─ core/
├─ database/
├─ ui/
├─ db_schema.md
├─ rrule.md
├─ requirements.txt
├─ pyproject.toml
├─ CalendarUA-onedir.spec
└─ CalendarUA-onefile.spec
```

## 環境需求

- Python >= 3.10
- 建議使用 uv
- Windows（主要開發/測試平台）

## 快速開始

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

## 打包

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

## 常用開發指令

```bash
# 語法檢查
uv run python -m py_compile CalendarUA.py core/*.py database/*.py ui/*.py

# 更新依賴
uv sync --dev
```

## 文件

- 資料庫結構：`db_schema.md`
- RRULE 說明：`rrule.md`

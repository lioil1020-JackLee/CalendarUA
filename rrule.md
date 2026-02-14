# RRULE 參數參考手冊

## 概述

CalendarUA 系統採用 RFC 5545 iCalendar 標準的 RRULE (Recurrence Rule) 規範，用於定義週期性任務的執行規則。本文件詳細說明系統支援的所有 RRULE 參數及其使用方法。

## RRULE 基本語法

```
RRULE:FREQ=<頻率>;[其他參數];DTSTART:<開始時間>;DURATION=<持續時間>
```

### 參數結構說明

- **必需參數**: `FREQ` (重複頻率)
- **時間參數**: `DTSTART`, `BYHOUR`, `BYMINUTE`
- **重複控制**: `INTERVAL`, `COUNT`, `UNTIL`
- **日期/時間指定**: `BYDAY`, `BYMONTHDAY`, `BYMONTH`, `BYSETPOS`
- **自訂參數**: `DURATION` (CalendarUA 專用)

## 詳細參數說明

### 1. FREQ (重複頻率) - 必需參數

定義任務重複執行的基本頻率。

| 參數值 | 說明 | 使用情境 |
|--------|------|----------|
| `DAILY` | 每日重複 | 日常維護、定期監控 |
| `WEEKLY` | 每週重複 | 週期性檢查、週報生成 |
| `MONTHLY` | 每月重複 | 月結作業、月度報表 |
| `YEARLY` | 每年重複 | 年終結算、年度檢查 |

**範例**:
```
FREQ=WEEKLY    # 每週重複
FREQ=MONTHLY   # 每月重複
```

### 2. INTERVAL (可選)
定義重複的間隔。

| 範例 | 說明 |
|------|------|
| `INTERVAL=2` | 每2天/週/月/年 |  (預設 1)

### 3. BYDAY (可選)
指定星期幾。支援單個星期或多個星期的組合。

| 值 | 說明 | 範例 |
|----|------|------|
| SU | 星期日 | `BYDAY=SU` |
| MO | 星期一 | `BYDAY=MO` |
| TU | 星期二 | `BYDAY=TU` |
| WE | 星期三 | `BYDAY=WE` |
| TH | 星期四 | `BYDAY=TH` |
| FR | 星期五 | `BYDAY=FR` |
| SA | 星期六 | `BYDAY=SA` |
| MO,TU,WE,TH,FR | 週一到週五 | `BYDAY=MO,TU,WE,TH,FR` | (預設 週一到週五)

### 4. BYMONTHDAY (可選)
指定月份中的日期 (1-31)。

| 範例 | 說明 |
|------|------|
| `BYMONTHDAY=1` | 每月1日 |  (預設 1)
| `BYMONTHDAY=15` | 每月15日 |

### 5. BYMONTH (可選)
指定月份 (1-12)。

| 範例 | 說明 |
|------|------|
| `BYMONTH=1` | 1月 |  (預設 1)
| `BYMONTH=6` | 6月 |

### 6. BYSETPOS (可選)
指定在集合中的位置。支援正數和負數。

| 範例 | 說明 |
|------|------|
| `BYSETPOS=1` | 第一個 |  (預設 第1個)
| `BYSETPOS=2` | 第二個 |
| `BYSETPOS=-1` | 最後一個 |
BYMONTH=6,12        # 半年結算
```

### 6. BYSETPOS (集合位置) - 可選參數

指定在符合條件的日期集合中的第幾個位置執行。

| 參數值 | 說明 |
|--------|------|
| `1` | 第一個 |
| `2` | 第二個 |
| `-1` | 最後一個 |
| `-2` | 倒數第二個 |

**使用情境**:
```
FREQ=MONTHLY;BYDAY=FR;BYSETPOS=-1    # 每月最後一個星期五
FREQ=MONTHLY;BYDAY=MO;BYSETPOS=2     # 每月第二個星期一
```

### 7. BYHOUR & BYMINUTE (時間指定) - 可選參數

精確指定任務執行的時分。

| 參數 | 範圍 | 說明 |
|------|------|------|
| `BYHOUR` | 0-23 | 小時 (24小時制) |
| `BYMINUTE` | 0-59 | 分鐘 |

**範例**:
```
BYHOUR=8;BYMINUTE=30     # 上午 8:30
BYHOUR=14;BYMINUTE=0     # 下午 2:00
BYHOUR=9;BYMINUTE=15     # 上午 9:15
```

### 8. COUNT (重複次數限制) - 可選參數

指定任務總共執行多少次後停止。

| 範例 | 說明 |
|------|------|
| `COUNT=10` | 執行 10 次後停止 |
| `COUNT=1` | 只執行 1 次 |

### 9. UNTIL (結束日期限制) - 可選參數

指定任務執行到哪個日期為止。

| 格式 | 範例 | 說明 |
|------|------|------|
| `YYYYMMDDTHHMMSS` | `UNTIL=20261231T235959` | 到 2026年12月31日 23:59:59 停止 |

### 10. DTSTART (開始日期時間) - 可選參數

指定重複規則的起始日期時間。

| 格式 | 範例 | 說明 |
|------|------|------|
| `YYYYMMDDTHHMMSS` | `DTSTART:20260214T080000` | 2026年2月14日 08:00:00 開始 |

### 11. DURATION (持續時間) - 自訂參數

CalendarUA 專用的參數，指定任務執行的持續時間。同時控制任務的重試行為：

- **DURATION = PT0M (0 分鐘)**: 單次執行，任務失敗後不重試
- **DURATION > PT0M (大於 0 分鐘)**: 持續執行模式，任務失敗後會重試直到成功或持續時間結束

| 格式 | 範例 | 說明 |
|------|------|------|
| `PT0M` | `DURATION=PT0M` | 單次執行，失敗不重試 |
| `PT{n}M` | `DURATION=PT30M` | 持續 30 分鐘，重試直到成功或時間結束 |
| `PT{n}H` | `DURATION=PT2H` | 持續 2 小時，重試直到成功或時間結束 |
| `PT{n}H{n}M` | `DURATION=PT1H30M` | 持續 1 小時 30 分鐘，重試直到成功或時間結束 |

## 完整範例

### 1. 單次執行排程
```
RRULE:FREQ=DAILY;BYHOUR=8;BYMINUTE=0;DTSTART:20260214T080000;DURATION=PT0M
```
**說明**: 每天上午 8:00 執行一次，失敗不重試

### 2. 持續執行排程
```
RRULE:FREQ=DAILY;BYHOUR=8;BYMINUTE=0;DTSTART:20260214T080000;DURATION=PT30M
```
**說明**: 每天上午 8:00 開始，持續 30 分鐘，重試直到成功或時間結束

### 3. 工作日排程
```
RRULE:FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR;BYHOUR=9;BYMINUTE=0;DTSTART:20260214T090000;DURATION=PT1H
```
**說明**: 每週一至週五上午 9:00 開始，持續 1 小時，重試直到成功或時間結束

### 4. 月度排程
```
RRULE:FREQ=MONTHLY;BYMONTHDAY=1;BYHOUR=10;BYMINUTE=0;DTSTART:20260201T100000;DURATION=PT2H
```
**說明**: 每月 1 日上午 10:00 開始，持續 2 小時

### 4. 季度排程
```
RRULE:FREQ=MONTHLY;INTERVAL=3;BYMONTHDAY=1;BYHOUR=14;BYMINUTE=30;DTSTART:20260301T143000;DURATION=PT45M
```
**說明**: 每 3 個月 1 日下午 2:30 開始，持續 45 分鐘

### 5. 年終排程
```
RRULE:FREQ=YEARLY;BYMONTH=12;BYMONTHDAY=31;BYHOUR=23;BYMINUTE=59;DTSTART:20261231T235900;DURATION=PT1M
```
**說明**: 每年 12 月 31 日 23:59 開始，持續 1 分鐘

### 6. 有限次數排程
```
RRULE:FREQ=WEEKLY;BYDAY=MO;BYHOUR=8;BYMINUTE=0;COUNT=10;DTSTART:20260214T080000;DURATION=PT30M
```
**說明**: 每週一上午 8:00 開始，持續 30 分鐘，總共執行 10 次

### 7. 指定結束日期排程
```
RRULE:FREQ=DAILY;BYHOUR=6;BYMINUTE=0;UNTIL=20261231T235959;DTSTART:20260214T060000;DURATION=PT15M
```
**說明**: 每天上午 6:00 開始，持續 15 分鐘，直到 2026 年 12 月 31 日

### 8. 複雜條件排程
```
RRULE:FREQ=MONTHLY;BYDAY=FR;BYSETPOS=-1;BYHOUR=17;BYMINUTE=0;DTSTART:20260228T170000;DURATION=PT1H
```
**說明**: 每月最後一個星期五下午 5:00 開始，持續 1 小時

## 參數優先順序與預設值

### 參數優先順序
1. **FREQ**: 決定基本重複單位
2. **INTERVAL**: 決定重複間隔
3. **BYxxx**: 決定具體的日期/時間條件
4. **COUNT/UNTIL**: 決定重複終止條件

### 預設值參考表

| 參數 | 預設值 | 備註 |
|------|--------|------|
| `INTERVAL` | `1` |  |
| `BYDAY` | `MO,TU,WE,TH,FR` | 工作日 |
| `BYMONTHDAY` | `1` | 每月 1 日 |
| `BYMONTH` | `1` | 1 月 |
| `BYSETPOS` | `1` | 第一個 |
| `BYHOUR` | 系統時間 | 最接近的合理時間 |
| `BYMINUTE` | `0` | 整點 |
| `COUNT` | 無限制 |  |
| `UNTIL` | 無限制 |  |
| `DURATION` | `PT0M` | 0 分鐘 (單次執行) |

## 資料儲存與解析

### RRULE 字串儲存內容
- ✅ **開始時間**: `DTSTART` + `BYHOUR` + `BYMINUTE`
- ✅ **持續時間**: `DURATION` (自訂參數)
- ✅ **重複規則**: `FREQ`, `INTERVAL`, `BYDAY` 等
- ✅ **終止條件**: `COUNT` 或 `UNTIL`

### 時間計算邏輯
```
任務結束時間 = 開始時間 + 持續時間
```

**計算範例**:
- 開始時間: `DTSTART:20260214T080000` (2026-02-14 08:00:00)
- 持續時間: `DURATION=PT30M` (30 分鐘)
- 結束時間: 2026-02-14 08:30:00 (動態計算，不儲存)

## 系統整合說明

### 與 UI 介面的對應
- **FREQ**: 對應「重複」下拉選單
- **BYDAY**: 對應星期選擇核取方塊
- **BYHOUR/BYMINUTE**: 對應時間選擇控制項
- **COUNT/UNTIL**: 對應「結束」選項
- **DURATION**: 對應「持續時間」設定

### 與資料庫的互動
- RRULE 字串完整儲存在 `schedules` 表格的 `rrule_str` 欄位
- 系統解析 RRULE 字串以計算下次執行時間
- 支援動態修改 RRULE 規則

## 故障排除

### 常見 RRULE 錯誤
1. **缺少 FREQ**: 所有 RRULE 必須包含 FREQ 參數
2. **無效的日期**: UNTIL 日期必須晚於 DTSTART
3. **衝突的參數**: 某些參數組合可能導致無效規則

### 除錯建議
- 使用系統的 RRULE 測試功能驗證規則
- 檢查日曆視圖中的任務分佈是否正確
- 檢視應用程式記錄中的解析錯誤訊息

- [RFC 5545 - Internet Calendaring and Scheduling Core Object Specification](https://tools.ietf.org/html/rfc5545)
- [python-dateutil RRULE 文件](https://dateutil.readthedocs.io/en/stable/rrule.html)
- CalendarUA 原始碼:
  - `core/rrule_parser.py` - RRULE 解析邏輯
  - `ui/recurrence_dialog.py` - RRULE 設定介面
  - `database/sqlite_manager.py` - RRULE 資料儲存
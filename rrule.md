# RRULE 參數清單

本文檔列出 CalendarUA 專案中支援的所有 RRULE (iCalendar Recurrence Rule) 參數。

## RRULE 基本格式

```
RRULE:FREQ=DAILY;BYHOUR=8;BYMINUTE=0;DTSTART:20260214T080000;DURATION=PT30M
```

## 支援的參數

### 1. FREQ (必需)
定義重複的頻率類型。

| 值 | 說明 | 範例 |
|----|------|------|
| DAILY | 每天 | `FREQ=DAILY` | (預設 每天)
| WEEKLY | 每週 | `FREQ=WEEKLY` |
| MONTHLY | 每月 | `FREQ=MONTHLY` |
| YEARLY | 每年 | `FREQ=YEARLY` |

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

### 7. 開始時間 (BYHOUR + BYMINUTE，可選)
指定重複規則中每次觸發的具體時間點。這是用來設定重複事件在一天中的哪個時間開始觸發。

| 參數組合 | 說明 | 範例 |
|----------|------|------|
| `BYHOUR=8;BYMINUTE=0` | 每天上午8點觸發 | (預設 以目前系統時間最接近的，但必須超過，例如目前08:56 則預設09:00)
| `BYHOUR=14;BYMINUTE=30` | 每天下午2點30分觸發 |
| `BYHOUR=9;BYMINUTE=15` | 每天上午9點15分觸發 |

**注意**: 開始時間參數決定了重複事件的觸發時間點，不是指事件的開始與結束時間。事件的持續時間是由 DURATION 參數決定。

### 8. COUNT (可選)
指定重複的總次數。

| 範例 | 說明 |
|------|------|
| `COUNT=10` | 重複10次後結束 |  (預設 10)

### 9. UNTIL (可選)
指定重複的結束日期時間。

| 範例 | 說明 |
|------|------|
| `UNTIL=20261231T235959` | 到2026年12月31日結束 |  (預設 UNTIL，所以當沒有結束日期選擇時，這個結束日期應該是反白不可設定的)

**注意**: 如果不指定 COUNT 或 UNTIL，RRULE 預設為無限重複（即「沒有結束日期」）。

### 10. DTSTART (可選)
指定重複規則的開始日期時間。

| 格式 | 範例 | 說明 |
|------|------|------|
| YYYYMMDDTHHMMSS | `DTSTART:20260214T080000` | 2026年2月14日08:00:00 |  (預設 系統今天日期)

### 11. DURATION (自訂參數)
指定任務的持續時間 (CalendarUA 專用參數)。

| 格式 | 範例 | 說明 |
|------|------|------|
| PT{n}M | `DURATION=PT30M` | 持續30分鐘 | (預設 5分鐘)
| PT{n}H | `DURATION=PT2H` | 持續2小時 |

## 常見 RRULE 範例

### 1. 每天早上8點，持續30分鐘
```
RRULE:FREQ=DAILY;BYHOUR=8;BYMINUTE=0;DTSTART:20260214T080000;DURATION=PT30M
```

### 2. 每週一早上8點，持續15分鐘
```
RRULE:FREQ=WEEKLY;BYDAY=MO;BYHOUR=8;BYMINUTE=0;DTSTART:20260214T080000;DURATION=PT15M
```

### 3. 每月第一天早上8點，持續1小時
```
RRULE:FREQ=MONTHLY;BYMONTHDAY=1;BYHOUR=8;BYMINUTE=0;DTSTART:20260214T080000;DURATION=PT1H
```

### 4. 每週一到週五早上8點，持續30分鐘
```
RRULE:FREQ=DAILY;BYDAY=MO,TU,WE,TH,FR;BYHOUR=8;BYMINUTE=0;DTSTART:20260214T080000;DURATION=PT30M
```

### 5. 每年1月1日早上8點，持續2小時
```
RRULE:FREQ=YEARLY;BYMONTH=1;BYMONTHDAY=1;BYHOUR=8;BYMINUTE=0;DTSTART:20260214T080000;DURATION=PT2H
```

### 6. 每月最後一個星期五早上8點，持續45分鐘
```
RRULE:FREQ=MONTHLY;BYDAY=FR;BYSETPOS=-1;BYHOUR=8;BYMINUTE=0;DTSTART:20260214T080000;DURATION=PT45M
```

## 資料儲存說明

### RRULE 儲存的資訊
- ✅ **開始時間**：DTSTART + BYHOUR + BYMINUTE
- ✅ **持續時間**：DURATION (自訂參數)
- ✅ **重複規則**：FREQ, INTERVAL, BYDAY 等
- ✅ **重複結束條件**：COUNT 或 UNTIL
- ❌ **結束時間**：不直接儲存，由開始時間 + 持續時間計算得出

### 結束時間的計算方式
```
結束時間 = 開始時間 + 持續時間
```

**範例**：
- 開始時間：`DTSTART:20260214T080000` (2026年2月14日 08:00:00)
- 持續時間：`DURATION=PT30M` (30分鐘)
- 結束時間：2026年2月14日 08:30:00（計算得出，不儲存）

### 與標準 RRULE 的差異
標準 RRULE 規範中沒有 DURATION 參數的概念。CalendarUA 使用自訂的 DURATION 參數來處理事件的持續時間，而不直接儲存結束時間。

### 實作注意事項

1. **必需參數**: 只有 `FREQ` 是必需的，其他都是可選的
2. **預設值**: 未指定的參數會使用合理的預設值
3. **DURATION**: 這是 CalendarUA 專用的自訂參數，不符合標準 RRULE 規範
4. **解析**: 程式會自動忽略不支援的參數（如 DURATION）
5. **編碼**: 所有設定都會編碼成 RRULE 字串儲存在資料庫中

## 相關檔案

- `core/rrule_parser.py` - RRULE 解析和處理邏輯
- `ui/recurrence_dialog.py` - RRULE 建構和UI設定
- `database/sqlite_manager.py` - RRULE 儲存到資料庫
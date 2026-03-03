# CalendarUA RRULE 說明

CalendarUA 使用 RFC 5545 RRULE 作為週期規則核心，並額外支援自訂參數 `DURATION`。

## 格式

實務上 `rrule_str` 以分號字串儲存，例如：

```text
FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR;BYHOUR=8;BYMINUTE=0;DTSTART=20260301T080000;DURATION=PT30M
```

## 主要參數

### 必填

- `FREQ`：重複頻率，支援 `DAILY` / `WEEKLY` / `MONTHLY` / `YEARLY`。

### 常用可選

- `INTERVAL`：間隔，預設 `1`。
- `BYDAY`：星期條件（`MO,TU,...`）。
- `BYMONTHDAY`：每月第幾日（`1-31`）。
- `BYMONTH`：月份（`1-12`）。
- `BYSETPOS`：例如第一個、最後一個（可負數）。
- `BYHOUR`：小時（`0-23`）。
- `BYMINUTE`：分鐘（`0-59`）。
- `COUNT`：最多觸發次數。
- `UNTIL`：結束時間（格式 `YYYYMMDDTHHMMSS`）。
- `DTSTART`：規則起始時間（格式 `YYYYMMDDTHHMMSS`）。

### CalendarUA 自訂

- `DURATION`：持續時間（ISO-8601 duration，例：`PT0M`、`PT30M`、`PT1H`）。

## 行為重點

- `DURATION=PT0M`：視為瞬時事件。
- `DURATION>PT0M`：事件有時間區間，會在日/週/月視圖以區間呈現。
- 假日覆寫與例外規則會在 resolver 階段套用到 occurrence。

## 範例

### 每天 08:00，持續 30 分鐘

```text
FREQ=DAILY;BYHOUR=8;BYMINUTE=0;DTSTART=20260301T080000;DURATION=PT30M
```

### 每週一到週五 09:30

```text
FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR;BYHOUR=9;BYMINUTE=30;DTSTART=20260301T093000;DURATION=PT15M
```

### 每月最後一個星期五 17:00

```text
FREQ=MONTHLY;BYDAY=FR;BYSETPOS=-1;BYHOUR=17;BYMINUTE=0;DTSTART=20260301T170000;DURATION=PT1H
```

### 有截止日期

```text
FREQ=DAILY;BYHOUR=6;BYMINUTE=0;UNTIL=20261231T235959;DTSTART=20260301T060000;DURATION=PT10M
```

## 常見錯誤

- 缺少 `FREQ`。
- `UNTIL` 早於 `DTSTART`。
- `BYDAY` / `BYMONTHDAY` / `BYSETPOS` 組合不合理，導致無 occurrence。

## 參考

- RFC 5545: https://datatracker.ietf.org/doc/html/rfc5545
- python-dateutil rrule: https://dateutil.readthedocs.io/en/stable/rrule.html
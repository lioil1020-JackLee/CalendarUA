# CalendarUA RRULE 指南

CalendarUA 以 RFC 5545 RRULE 為基礎，並擴充少量自訂欄位，供排程與 UI 使用。

## 1. 基本格式

`rrule_str` 以 `;` 分隔，例如：

```text
FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR;BYHOUR=8;BYMINUTE=0;DTSTART:20260301T080000;DURATION=PT30M
```

> 注意：本專案同時相容 `DTSTART:` 與部分 `X-*` 自訂參數。

## 2. 常用標準參數

- `FREQ`：`DAILY | WEEKLY | MONTHLY | YEARLY`
- `INTERVAL`：間隔，預設 `1`
- `BYDAY`：星期條件（如 `MO,TU,WE`）
- `BYMONTHDAY`：每月第幾天（`1~31`）
- `BYMONTH`：月份（`1~12`）
- `BYSETPOS`：位置（例如 `1`, `-1`）
- `BYHOUR` / `BYMINUTE` / `BYSECOND`
- `COUNT`：總次數上限
- `UNTIL`：截止時間（`YYYYMMDDTHHMMSS`）
- `DTSTART`：起始時間（`YYYYMMDDTHHMMSS`）

## 3. CalendarUA 自訂參數

- `DURATION=PT...`：事件持續時間（例如 `PT0M`, `PT15M`, `PT1H`）
- `X-LUNAR=1`：標記規則為農曆語意（由 UI 生成/解析）
- `X-RANGE-START=...`：範圍起始補充資訊（resolver 會讀取）

## 4. 解析行為（本專案）

- occurrence 先由 RRULE 觸發，再套用：
  1. `schedule_exceptions`（cancel/override）
  2. `holidays` 假日規則（若排程未勾 `ignore_holiday`）
- `DURATION` 決定事件區間長度
- 已停用排程會標記為關閉狀態顯示

## 5. 範例

### 每天 08:00，持續 30 分鐘

```text
FREQ=DAILY;BYHOUR=8;BYMINUTE=0;DTSTART:20260301T080000;DURATION=PT30M
```

### 每週一到週五 09:30，15 分鐘

```text
FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR;BYHOUR=9;BYMINUTE=30;DTSTART:20260301T093000;DURATION=PT15M
```

### 每月最後一個星期五 17:00，1 小時

```text
FREQ=MONTHLY;BYDAY=FR;BYSETPOS=-1;BYHOUR=17;BYMINUTE=0;DTSTART:20260301T170000;DURATION=PT1H
```

### 有截止日期

```text
FREQ=DAILY;BYHOUR=6;BYMINUTE=0;UNTIL=20261231T235959;DTSTART:20260301T060000;DURATION=PT10M
```

## 6. 常見錯誤

- 少 `FREQ`
- `UNTIL` 早於 `DTSTART`
- `BYDAY/BYMONTHDAY/BYSETPOS` 組合錯誤導致無 occurrence
- 時間欄位格式錯誤（非數字或超出範圍）

## 7. 參考

- RFC 5545: https://datatracker.ietf.org/doc/html/rfc5545
- python-dateutil rrule: https://dateutil.readthedocs.io/en/stable/rrule.html

# CalendarUA RRULE 指南

CalendarUA 以 RFC 5545 RRULE 為基礎，並加入專案自訂欄位給 UI 與 resolver 使用。

## 1) 基本格式

以 `;` 分隔，例如：

```text
FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR;BYHOUR=8;BYMINUTE=15;DTSTART:20260301T081500;DURATION=PT15M
```

## 2) 支援參數

- `FREQ`：`DAILY | WEEKLY | MONTHLY | YEARLY`
- `INTERVAL`：間隔（預設 `1`）
- `BYDAY`
- `BYMONTHDAY`
- `BYMONTH`
- `BYSETPOS`
- `BYHOUR` / `BYMINUTE` / `BYSECOND`
- `COUNT`
- `UNTIL`
- `DTSTART`

## 3) CalendarUA 自訂參數

- `DURATION=PT...`：事件持續時間
- `X-LUNAR=1`：農曆模式
- `X-RANGE-START=...`：範圍起始補充資訊

## 4) 與 UI 行為對應

- Time Scale（5/6/10/15/30/60 分）僅影響日/週視圖顯示與右鍵帶入時間，不改變 RRULE 本體
- 每月模式中兩個「每 N 個月」欄位現在是固定 `1~12` 下拉選單
- 排程解析流程：
  1. RRULE 生成 occurrence
  2. 套用 `schedule_exceptions`
  3. 若未勾 `ignore_holiday`，套用 `holidays`

## 5) 範例

### 每天 08:00，持續 30 分

```text
FREQ=DAILY;BYHOUR=8;BYMINUTE=0;DTSTART:20260301T080000;DURATION=PT30M
```

### 每週一到週五 09:30，持續 15 分

```text
FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR;BYHOUR=9;BYMINUTE=30;DTSTART:20260301T093000;DURATION=PT15M
```

### 每月最後一個星期五 17:00，持續 1 小時

```text
FREQ=MONTHLY;BYDAY=FR;BYSETPOS=-1;BYHOUR=17;BYMINUTE=0;DTSTART:20260301T170000;DURATION=PT1H
```

### 設定截止日期

```text
FREQ=DAILY;BYHOUR=6;BYMINUTE=0;UNTIL=20261231T235959;DTSTART:20260301T060000;DURATION=PT10M
```

## 6) 常見錯誤

- 缺少 `FREQ`
- `UNTIL` 早於 `DTSTART`
- `BYDAY/BYMONTHDAY/BYSETPOS` 組合不合法
- `BYHOUR/BYMINUTE` 超出範圍
- `DURATION` 格式錯誤

## 7) 參考

- RFC 5545: https://datatracker.ietf.org/doc/html/rfc5545
- python-dateutil rrule: https://dateutil.readthedocs.io/en/stable/rrule.html

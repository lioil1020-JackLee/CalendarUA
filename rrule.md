# CalendarUA RRULE 指南

CalendarUA 以 RFC 5545 RRULE 為基礎，並加入少量專案自訂欄位給 UI 與 resolver 使用。

## 1. 基本格式

`rrule_str` 以 `;` 分隔，例如：

```text
FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR;BYHOUR=8;BYMINUTE=15;DTSTART:20260301T081500;DURATION=PT15M
```

---

## 2. 支援的常用標準參數

- `FREQ`：`DAILY | WEEKLY | MONTHLY | YEARLY`
- `INTERVAL`：間隔，預設 `1`
- `BYDAY`：星期條件（例如 `MO,TU,WE`）
- `BYMONTHDAY`：每月第幾天（`1~31`）
- `BYMONTH`：月份（`1~12`）
- `BYSETPOS`：位置（例如 `1`, `-1`）
- `BYHOUR` / `BYMINUTE` / `BYSECOND`
- `COUNT`：總次數上限
- `UNTIL`：截止（`YYYYMMDDTHHMMSS`）
- `DTSTART`：起始（`YYYYMMDDTHHMMSS`）

---

## 3. CalendarUA 自訂參數

- `DURATION=PT...`：事件持續時間
- `X-LUNAR=1`：農曆模式標記
- `X-RANGE-START=...`：範圍起始補充資訊

說明：

- UI 目前「期間」最短為 `5 分`（已移除 `0 分`）
- 已存在資料若仍含 `PT0M`，解析層仍可讀取，但 UI 不再提供此選項

---

## 4. 與目前 UI 行為的關係

### 4.1 Time Scale 與 RRULE 的關係

- 日/週視圖 Time Scale（`5/6/10/15/30/60 min`）只影響視圖格線與右鍵點選時間
- RRULE 本身仍以 `BYHOUR/BYMINUTE` 與 `DURATION` 決定實際發生時間
- 右鍵新增/編輯時，會把目前格線對應的「小時+分鐘」帶入 `RecurrenceDialog`

### 4.2 Occurrence 套用順序

1. RRULE 先產生 occurrence
2. 套用 `schedule_exceptions`（`cancel/override`）
3. 若排程未勾 `ignore_holiday`，再套用 `holidays`

### 4.3 啟動回補與執行中熱更新（v3.0.0）

- 啟動回補：
	- 若排程包含 `DURATION`，且目前時間落在期間內，排程器可回補最近一次觸發點。
	- 用 occurrence 起始時間去重，避免同一觸發點重複執行。

- 執行中熱更新：
	- 任務執行迴圈會動態重讀排程設定，`lock_enabled`、`target_value`、`data_type`、`node_id`、`is_enabled`、OPC 連線參數可即時生效。
	- 若連線參數改變（如 `opc_url` / 帳密 / 安全策略 / timeout），會自動重連。

- 假日覆寫動態套用：
	- 每次輪詢會即時判定是否命中假日規則。
	- 命中且規則有 `override_target_value` 時，該次寫值使用覆寫值。

---

## 5. 範例

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

### 有截止日期

```text
FREQ=DAILY;BYHOUR=6;BYMINUTE=0;UNTIL=20261231T235959;DTSTART:20260301T060000;DURATION=PT10M
```

---

## 6. 常見錯誤

- 少 `FREQ`
- `UNTIL` 早於 `DTSTART`
- `BYDAY/BYMONTHDAY/BYSETPOS` 組合不合法，導致沒有 occurrence
- `BYHOUR/BYMINUTE/BYSECOND` 值超出範圍
- `DURATION` 格式不合法（建議使用 `PT{N}M` 或 `PT{N}H`）

---

## 7. 參考

- RFC 5545：<https://datatracker.ietf.org/doc/html/rfc5545>
- python-dateutil rrule：<https://dateutil.readthedocs.io/en/stable/rrule.html>

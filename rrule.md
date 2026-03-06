# CalendarUA RRULE 指南

CalendarUA 以 RFC 5545 RRULE 為基礎，並加入專案自訂欄位供排程引擎與 UI 使用。

## 1. 基本格式

以 `;` 分隔，例如：

```text
FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR;BYHOUR=8;BYMINUTE=15;DTSTART:20260301T081500;DURATION=PT15M
```

## 2. 支援參數（RFC 5545 / dateutil）

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

## 3. CalendarUA 自訂參數

- `DURATION=PT...`：事件持續時間（`schedule_resolver` 會解析，預設 60 分鐘）。
- `X-LUNAR=1`：啟用農曆規則模式。
- `X-RANGE-START=...`：觸發範圍起點（早於此時間的觸發會被忽略）。

## 4. 農曆模式說明（`X-LUNAR=1`）

農曆模式由 `core/rrule_parser.py` 自行計算，不走 dateutil 標準 rrule iterator。

- 目前支援頻率：`DAILY/WEEKLY/MONTHLY/YEARLY`
- `WEEKLY`：支援 `BYDAY`
- `MONTHLY`：支援 `BYMONTHDAY`（對應農曆日）
- `YEARLY`：支援 `BYMONTH` + `BYMONTHDAY`（對應農曆月/日）
- `COUNT`、`UNTIL` 在農曆模式同樣生效

注意：若未安裝 `lunardate`，農曆計算會失敗，觸發結果可能為空。

## 5. 解析流程（執行時）

1. `RRuleParser.get_trigger_between()` 先展開指定區間 occurrence。
2. `schedule_resolver.resolve_occurrences_for_range()` 套用 `schedule_exceptions`。
3. 若排程未勾 `ignore_holiday`，再套用 `holidays` 規則。
4. 最終結果交由排程器執行 OPC UA 寫值。

## 6. 範例

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

### 有截止日

```text
FREQ=DAILY;BYHOUR=6;BYMINUTE=0;UNTIL=20261231T235959;DTSTART:20260301T060000;DURATION=PT10M
```

### 農曆每年八月十五 09:00

```text
FREQ=YEARLY;BYMONTH=8;BYMONTHDAY=15;BYHOUR=9;BYMINUTE=0;X-LUNAR=1;DTSTART:20260301T090000;DURATION=PT30M
```

## 7. 常見錯誤

- 缺少 `FREQ`
- `UNTIL` 早於 `DTSTART`
- `BYDAY/BYMONTHDAY/BYSETPOS` 組合不合法
- `BYHOUR/BYMINUTE` 超出範圍
- `DURATION` 格式錯誤
- 農曆模式未安裝 `lunardate`

## 8. 參考

- RFC 5545: https://datatracker.ietf.org/doc/html/rfc5545
- python-dateutil rrule: https://dateutil.readthedocs.io/en/stable/rrule.html

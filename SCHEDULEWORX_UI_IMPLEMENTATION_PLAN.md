# CalendarUA 對齊 ScheduleWorX UI/功能 — 詳細實作與修改計畫

> 目標：以 **CalendarUA 現有架構**為基礎，逐步做出「接近 ScheduleWorX 操作體驗」的介面與功能。
> 原則：**先可用、再完整、最後優化**；每一步都可執行、可驗收、可回滾。

---

## 0. 專案現況（已具備）

### 已有核心能力
- RRULE 解析、下一次觸發、範圍查詢：`core/rrule_parser.py`
- 排程資料存取（SQLite）：`database/sqlite_manager.py`
- 排程主視窗與任務表格：`CalendarUA.py`
- 週期編輯對話框（Daily/Weekly/Monthly/Yearly）：`ui/recurrence_dialog.py`
- 排程新增/編輯對話框：`ScheduleEditDialog`（在 `CalendarUA.py` 內）

### 已完成的主要功能 (2026-02-24 更新)
✅ **6 個核心 Tab 系統完成**:
- General Tab: 全局設定（profile, scan rate, active period, output type）
- Weekly Tab: 7×24 網格排程編輯
- Holidays Tab: 假日日曆管理（full-day/partial-day支持）
- Exceptions Tab: 例外記錄管理（cancel/override actions）
- Preview Tab: Day/Week/Month 三視圖預覽
- Runtime Tab: 運行時覆寫與狀態監控

✅ **資料庫模型完成**:
- `schedules`: 排程主資料
- `schedule_exceptions`: 例外記錄（cancel/override）
- `holiday_calendars`: 假日日曆
- `holiday_entries`: 假日時間條目
- `general_settings`: 全局配置（單行表）
- `runtime_override`: 運行時覆寫（單行表，支持臨時過期）

✅ **核心功能實現**:
- Occurrence vs Series 編輯分流（OccurrenceChoiceDialog + OccurrenceEditDialog）
- 排程解析器（core/schedule_resolver.py）：series + exceptions 合併
- Day/Week/Month 三視圖（schedule_canvas.py, month_grid.py）
- 右鍵選單（Open/Delete/New/Copy/Cut/Paste）
- 例外系統（單次occurrence編輯寫入schedule_exceptions）
- 假日覆寫（holiday_entries with override_target_value）
- Runtime Tab 自動更新（QTimer 1秒間隔，countdown calculation）

✅ **UI 架構優化（2026-02-24）**:
- 移除冗餘的左側日曆面板（QCalendarWidget）
- 移除冗餘的"當天排程摘要"（已由 Runtime Tab 取代）
- 移除冗餘的"排程管理表格"（已由 Weekly Tab 取代）
- 主視窗改為全寬單面板佈局，6 個 Tab 獲得完整視窗空間
- 所有日期選擇統一使用 Preview Tab 的 QDateEdit

### 與目標的剩餘差距
1. ❌ 缺少 category（顏色分類）系統與對應顯示（計畫中的 Phase 4）。
2. ❌ 缺少完整 Ribbon 行為（Create/Edit/Delete/Refresh/Apply/Load/Save）的對應操作流。
3. ❌ Holiday resolver 尚未整合into schedule_resolver.py。
4. ❌ Preview Tab 的某些進階互動（如拖曳調整時間）尚未實現。
5. ❌ 效能優化：大量排程時的渲染優化、視圖快取機制。

---

## 1. 最終目標（完成定義）

### UI 目標
- 主畫面上方有 `General / Weekly / Holidays / Exceptions / Preview / Runtime`。
- 中央視圖支援 `Day / Week / Month` 切換。
- Week/Day 視圖有 24 小時時間軸、區塊事件塗色。
- Month 視圖可顯示每天多筆事件（色條）。
- 有日期導航（上一期 / 今天 / 下一期 / 快速選日曆）。

### 功能目標
- 可新增 recurring series。
- 點選 recurring occurrence 時，先問：
  - Open this occurrence（僅編輯本次）
  - Open the series（編輯整個系列）
- 支援 holiday calendar 與 schedule 覆寫（例如假日特殊班表）。
- 支援 category（名稱 + 顏色）並映射到事件色塊。
- Preview 分頁為唯讀（不可新增/編輯事件）。
- Runtime 分頁支援手動 Override、Temporary Override（秒/分/時/日/週）、Clear Override。
- Current Status 與 Next Event 需即時顯示（值、主題、類型、優先權、忙碌區間、下一事件時間）。

### 非目標（本階段不做）
- 不追求 100% 複刻第三方產品的像素級一致。
- 不做複雜拖曳（跨天拖移、resize handle）直到 MVP 穩定。
- 不先做多人協作/權限系統。

---

## 2. 架構重整策略（先抽離，再擴充）

### 2.1 新增模組建議

1. `core/schedule_resolver.py`
   - 職責：把 series、exceptions、holidays 合成「視圖可畫」的 occurrence 清單。
   - 輸入：時間範圍 + active profile。
   - 輸出：`ResolvedOccurrence[]`（含 start/end/title/category/source）。

2. `ui/schedule_canvas.py`
   - 職責：Day/Week 視圖畫布（時間軸 + 色塊）。
   - 元件：`DayViewWidget`、`WeekViewWidget`。

3. `ui/month_grid.py`
   - 職責：Month 視圖（每格顯示當日事件條）。

4. `ui/scheduleworx_shell.py`（可選）
   - 職責：把 `Exceptions/General/Holidays + Day/Week/Month + 導航列` 統一組合。
   - 若不另建，可直接在 `CalendarUA.py` 增量建立。

5. `ui/occurrence_choice_dialog.py`
   - 職責：彈窗 `Open this occurrence / Open the series`。

6. `ui/category_manager_dialog.py`（第二階段）
   - 職責：管理類別名稱與顏色。

7. `ui/weekly_panel.py`
  - 職責：Weekly Tab 事件建立/編輯（每週 recurring 為主）。

8. `ui/preview_panel.py`
  - 職責：Preview Tab 唯讀彙總視圖（weekly + holidays + exceptions 合成結果）。

9. `ui/runtime_panel.py`
  - 職責：Runtime Tab（Override / Clear Override / Current Status / Next Event）。

10. `core/runtime_state.py`
   - 職責：管理覆寫狀態、到期時間、目前值與下一事件快取。

### 2.2 現有檔案調整方向
- `CalendarUA.py`
  - 把目前右側表格 + 左側小月曆架構改成「上方六分頁 + 中央多視圖 + 下方狀態」。
  - 對齊 Runtime Ribbon 操作：Create/Edit/Delete/Refresh/Apply/Load/Save。
  - 保留原本排程表格可作為 debug/管理子頁，避免一次砍掉。
- `ui/recurrence_dialog.py`
  - 補齊與 ScheduleWorX 使用習慣接近的文案/欄位映射。
- `database/sqlite_manager.py`
  - 增加 categories、exceptions、holidays 相關 CRUD 與 migration。

---

## 3. 資料庫模型設計（重點）

> 原則：維持 `schedules` 為 series 主資料；另用附表表示例外與假日，避免破壞既有資料。

### 3.0 當前資料庫現況（2026-02-24）

#### 已實作的資料表

**A. `schedules` - 排程主資料表** ✅
```sql
CREATE TABLE schedules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_name TEXT NOT NULL,
    opc_url TEXT NOT NULL,
    node_id TEXT NOT NULL,
    target_value TEXT NOT NULL,
    data_type TEXT DEFAULT 'auto',
    rrule_str TEXT NOT NULL,
    opc_security_policy TEXT DEFAULT 'None',
    opc_security_mode TEXT DEFAULT 'None',
    opc_username TEXT DEFAULT '',
    opc_password TEXT DEFAULT '',
    opc_timeout INTEGER DEFAULT 10,
    is_enabled INTEGER DEFAULT 1,
    last_execution_status TEXT DEFAULT '',
    last_execution_time TIMESTAMP,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
```
**缺少欄位**: `category_id`, `priority`, `location`, `description`

**B. `schedule_exceptions` - 例外記錄表** ✅
```sql
CREATE TABLE schedule_exceptions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    schedule_id INTEGER NOT NULL,
    occurrence_date TEXT NOT NULL,
    action TEXT NOT NULL DEFAULT 'override',
    override_start TEXT,
    override_end TEXT,
    override_task_name TEXT,
    override_target_value TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(schedule_id) REFERENCES schedules(id) ON DELETE CASCADE
)
```
**缺少欄位**: `override_category_id`, `note`

**C. `holiday_calendars` - 假日日曆表** ✅
```sql
CREATE TABLE holiday_calendars (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    description TEXT,
    is_default INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
```
**狀態**: 完整

**D. `holiday_entries` - 假日時間條目表** ✅
```sql
CREATE TABLE holiday_entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    calendar_id INTEGER NOT NULL,
    holiday_date TEXT NOT NULL,
    name TEXT NOT NULL,
    is_full_day INTEGER DEFAULT 1,
    start_time TEXT,
    end_time TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(calendar_id) REFERENCES holiday_calendars(id) ON DELETE CASCADE
)
```
**缺少欄位**: `override_category_id`, `override_target_value`

**E. `general_settings` - 全局設定表（單行表）** ✅
```sql
CREATE TABLE general_settings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    profile_name TEXT DEFAULT '預設 Profile',
    description TEXT,
    enable_schedule INTEGER DEFAULT 1,
    scan_rate INTEGER DEFAULT 1,
    refresh_rate INTEGER DEFAULT 5,
    use_active_period INTEGER DEFAULT 0,
    active_from TEXT,
    active_to TEXT,
    output_type TEXT DEFAULT 'OPC UA Write',
    refresh_output INTEGER DEFAULT 1,
    generate_events INTEGER DEFAULT 1,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
```
**狀態**: 完整（替代了原計畫的 schedule_profiles 表）

**F. `runtime_override` - 運行時覆寫表（單行表）** ✅
```sql
CREATE TABLE runtime_override (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    override_value TEXT NOT NULL,
    override_until TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
```
**狀態**: 完整

### 3.1 需新增的資料表

#### A. `schedule_categories` - 類別管理表 ❌ **待實作**
#### A. `schedule_categories` - 類別管理表 ❌ **待實作**
```sql
CREATE TABLE IF NOT EXISTS schedule_categories (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    bg_color TEXT NOT NULL,  -- hex 格式如 '#FF0000'
    fg_color TEXT NOT NULL,  -- hex 格式如 '#FFFFFF'
    sort_order INTEGER DEFAULT 0,
    is_system INTEGER DEFAULT 0,  -- 系統預設類別不可刪除
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
```
**用途**: 
- 儲存顏色類別（Red, Pink, Light Purple...）
- 系統啟動時自動插入預設類別
- 支援使用者自訂類別

**預設系統類別**（參考 ScheduleWorX）:
| Name | bg_color | fg_color | is_system |
|------|----------|----------|-----------|
| Red (關閉) | #FF0000 | #FFFFFF | 1 |
| Pink (自動) | #FF69B4 | #FFFFFF | 1 |
| Light Purple (休假手動台) | #DDA0DD | #000000 | 1 |
| Green | #00FF00 | #000000 | 1 |
| Blue | #0000FF | #FFFFFF | 1 |
| Yellow | #FFFF00 | #000000 | 1 |
| Orange | #FFA500 | #000000 | 1 |
| Gray | #808080 | #FFFFFF | 1 |

### 3.2 既有表格需要的欄位擴充

#### schedules 表擴充 ❌ **待實作**
```sql
ALTER TABLE schedules ADD COLUMN category_id INTEGER DEFAULT 1;  -- 預設指向 Red
ALTER TABLE schedules ADD COLUMN priority INTEGER DEFAULT 1;
ALTER TABLE schedules ADD COLUMN location TEXT DEFAULT '';
ALTER TABLE schedules ADD COLUMN description TEXT DEFAULT '';
-- 建立外鍵索引
CREATE INDEX IF NOT EXISTS idx_schedules_category ON schedules(category_id);
```

#### schedule_exceptions 表擴充 ❌ **待實作**
```sql
ALTER TABLE schedule_exceptions ADD COLUMN override_category_id INTEGER;
ALTER TABLE schedule_exceptions ADD COLUMN note TEXT DEFAULT '';
-- 建立外鍵索引
CREATE INDEX IF NOT EXISTS idx_exceptions_category ON schedule_exceptions(override_category_id);
```

#### holiday_entries 表擴充 ❌ **待實作**
```sql
ALTER TABLE holiday_entries ADD COLUMN override_category_id INTEGER;
ALTER TABLE holiday_entries ADD COLUMN override_target_value TEXT;
-- 建立外鍵索引
CREATE INDEX IF NOT EXISTS idx_holiday_entries_category ON holiday_entries(override_category_id);
```

### 3.3 Migration 策略

**A. 安全遷移步驟**:
1. 使用 `PRAGMA table_info(table_name)` 檢查欄位是否已存在
2. 若不存在才執行 `ALTER TABLE ADD COLUMN`
3. 避免重複執行造成錯誤

**B. 預設資料插入**:
1. 檢查 `schedule_categories` 是否為空
2. 若為空則插入 8 個系統預設類別
3. 系統類別標記 `is_system=1`，防止使用者刪除

**C. 既有資料相容性**:
1. 新增欄位均使用 `DEFAULT` 值
2. 既有排程自動指向 `category_id=1` (Red)
3. 不影響現有功能運作

### 3.4 資料庫版本管理（建議）

雖然目前未實作完整的 migration 框架，但建議記錄版本:
```sql
CREATE TABLE IF NOT EXISTS schema_version (
    version INTEGER PRIMARY KEY,
    description TEXT,
    applied_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
-- 當前版本: 2 (新增 category 支援)
-- 版本 1: 初始 6 表系統
```

---

## 4. Domain 與運算邏輯（關鍵）

### 4.1 合併順序（Resolver）- 更新版

1. 先展開 weekly/base series（RRULE -> occurrence）
   - 每個 occurrence 帶有 `schedule.category_id` 的顏色
2. 套用 holiday 規則（若該日有 holiday time setting，覆蓋/補入該日事件）
   - 若 `holiday_entry.override_category_id` 存在，使用該顏色
   - 若 `holiday_entry.override_target_value` 存在，覆寫目標值
3. 套用 exception：
   - `cancel`：刪除該 occurrence
   - `override`：替換該 occurrence 的時間/值
   - 若 `exception.override_category_id` 存在，使用該顏色
4. 套用 runtime override（最高優先權，可暫時或永久）
   - 統一覆寫所有當前應該執行的事件
5. 依 priority 排序，輸出給 UI

### 4.2 優先權規則（對齊文件）- 含顏色
- **值的優先權**: `Runtime Override` > `Exception Override` > `Exception` > `Holiday Override` > `Weekly`
- **顏色的優先權**: `Exception Category` > `Holiday Category` > `Schedule Category`
- 文件明確指出：Exception 會蓋過 Weekly/Holiday；只有 Override 會再蓋過 Exception

### 4.3 建議資料結構 - 含 Category
### 4.3 建議資料結構 - 含 Category
```python
@dataclass
class ResolvedOccurrence:
    schedule_id: int
    source: str          # weekly / holiday / exception / override
    title: str
    start: datetime
    end: datetime
    category_id: int     # 類別 ID
    category_bg: str     # 背景顏色 (hex)
    category_fg: str     # 前景顏色 (hex)
    target_value: str
    priority: int = 1
    is_exception: bool = False
    is_holiday: bool = False
    is_override: bool = False
    occurrence_key: str = ""  # f"{schedule_id}:{start.isoformat()}"
    
    def __post_init__(self):
        if not self.occurrence_key:
            self.occurrence_key = f"{self.schedule_id}:{self.start.isoformat()}"
```

### 4.4 一致性規則
- 任何 occurrence 若 `end <= start`，視為無效，不進 UI。
- 跨日事件先拆分成每日片段再繪製（避免 week/day 畫布錯位）。
- Preview tab 一律唯讀，不提供 create/edit/delete 動作。
- `Refresh Schedule` 需還原未儲存變更；`Apply Schedule` 才寫入持久層。

---

## 5. UI 實作藍圖（逐步）

## Phase 1（MVP，先可用）

### 5.1 Shell 與導航
- 在主畫面上方加入：
  - Tab: `General | Weekly | Holidays | Exceptions | Preview | Runtime`
  - View Switch: `Day | Week | Month`
  - Date nav: `<< < Today > >>` + 快速小月曆 popup
- 先保留現有 schedule table 於下方或側邊「管理模式」。

**修改檔案**
- `CalendarUA.py`（主框架改造）

**驗收**
- 可切 Day/Week/Month。
- 日期切換時，三視圖同步跳轉。

### 5.2 Day/Week 畫布（只讀）
- 建立 `ui/schedule_canvas.py`：
  - 左側時間軸（00:00~23:00）
  - Day: 單日欄
  - Week: 7 欄（週日~週六）
  - 事件色塊（title 顯示）
- 先做 hover tooltip（標題、起迄、目標值）。

### 5.2.1 右鍵選單（Context Menu）
- 於 Day/Week/Month 視圖支援滑鼠右鍵選單（先以 Day/Week 完整落地，Month 次階段補齊）。
- 第一版功能項目對齊圖示需求：
  - `Open`
  - `Delete`
  - `New Event`
  - `Copy`
  - `Cut`
  - `Paste`
  - `Time Scale`（可先提供選單骨架）
  - `Refresh Schedule`
  - `Apply Schedule`
- 行為規範：
  - 右鍵點在事件上：Open/Delete/Copy/Cut 啟用。
  - 右鍵點在空白格：New Event/Paste 啟用。
  - `Copy`：保留來源事件資訊於剪貼簿緩存，不改原資料。
  - `Cut`：標記來源事件待搬移；`Paste` 成功後移除來源或改寫來源時間。
  - `Paste`：貼到目標日期/時間（優先使用右鍵位置）。
  - `Refresh Schedule`：重新載入資料並還原未套用的暫存顯示。
  - `Apply Schedule`：將暫存變更寫入資料庫（若目前採即時寫入，需在 UI 明確標示）。

**修改檔案**
- `ui/schedule_canvas.py`（新增右鍵事件與 action signal）
- `CalendarUA.py`（接收 action、執行 CRUD/剪貼簿流程）

**驗收**
- 可在 Week/Day 視圖右鍵開啟選單。
- Open/Delete/New/Copy/Cut/Paste 可觸發並作用於正確事件或時間格。
- Refresh/Apply 在目前實作策略下有一致且可理解的行為。

**修改檔案**
- `ui/schedule_canvas.py`（新）
- `CalendarUA.py`（嵌入 widget）

**驗收**
- 顏色事件塊能正確畫在時間軸。
- 週視圖可看到同一 series 於不同日的 occurrence。

### 5.3 Month 格子視圖（只讀）
- 建立 `ui/month_grid.py`：
  - 每格顯示日期 + 最多 N 條 event chip。
  - 超過顯示 `+X more`。

**修改檔案**
- `ui/month_grid.py`（新）
- `CalendarUA.py`（嵌入 widget）

**驗收**
- 月視圖可看到彩色事件條。
- 點日期可切換到 day 視圖（可選但建議）。

### 5.4 ResolvedOccurrence 來源
- 新建 resolver，先只串 series（不含例外與假日）。

**修改檔案**
- `core/schedule_resolver.py`（新）
- `CalendarUA.py`（改為從 resolver 取資料）

**驗收**
- 三視圖都不再直接吃 `schedules` 原始資料，而是吃 resolver 結果。

---

## Phase 2（功能對齊）

### 5.5 Weekly Tab
- 以週為主建立 recurring event（Show As / Subject / Location / Start/End / All day / Priority / ValueSet Value）。
- 可雙擊日曆建立事件，也可對既有事件編輯。

**修改檔案**
- `ui/weekly_panel.py`（新）
- `CalendarUA.py`（tab 嵌入）
- `database/sqlite_manager.py`（weekly event CRUD）

**驗收**
- 週事件可新增/編輯/刪除並即時反映到 Preview。

### 5.6 occurrence / series 編輯分流

### 5.5 occurrence / series 編輯分流
- 點擊 recurring occurrence 時先開選擇對話框：
  - Open this occurrence
  - Open the series
- `this occurrence` 寫入 `schedule_exceptions`。

**修改檔案**
- `ui/occurrence_choice_dialog.py`（新）
- `CalendarUA.py`（事件點擊流程）
- `database/sqlite_manager.py`（exception CRUD）

**驗收**
- 修改單次 occurrence 後，只有該次改變。
- 修改 series 後，未被 exception 覆寫的 occurrence 全部更新。

### 5.7 Exceptions Tab
- 顯示某 profile 的 exception 清單：
  - 新增 / 編輯 / 刪除
  - 指定日期、action(cancel/override)、覆寫值/顏色/時間
- 對齊規則：exception 優先於 weekly/holiday；但低於 runtime override。

**修改檔案**
- `ui/exceptions_panel.py`（新）
- `CalendarUA.py`（tab 嵌入）
- `database/sqlite_manager.py`（exception API）

**驗收**
- exception 變更後，週/月視圖即時反映。

### 5.8 Holidays Tab
- 左側：holiday calendar list
- 右側：當日時間軸或列表
- 支援新增假日條目、設為預設
- 支援 Holiday Time Settings + Holiday List 指派模式。

**修改檔案**
- `ui/holidays_panel.py`（新）
- `database/sqlite_manager.py`（holiday CRUD）
- `core/schedule_resolver.py`（holiday 合併）

**驗收**
- 指定假日可覆寫當日班表（例如紫色「休假手動台」）。

### 5.9 General Tab（Profile 設定）
- 目標對齊你圖中：
  - Name / Description
  - Enable Schedule
  - Scan Rate
  - Refresh Output / Refresh Rate
  - Active From/To
  - Output Type / ValueSet
  - Generate Events

**修改檔案**
- `ui/general_panel.py`（新）
- `database/sqlite_manager.py`（profile 欄位）
- `CalendarUA.py`（tab 嵌入）

**驗收**
- 變更後可保存並在啟動後恢復。

### 5.10 Preview Tab（唯讀彙整）
- 只顯示結果，不允許新增/編輯/刪除事件。
- 顯示 weekly + holiday + exception + override 合成結果。

**修改檔案**
- `ui/preview_panel.py`（新）
- `CalendarUA.py`（tab 嵌入與行為限制）

**驗收**
- Preview 只能瀏覽與導航，不可修改資料。

### 5.11 Runtime Tab（覆寫與狀態）
- Override：可選 ValueSet Value。
- Temporary Override：秒/分/時/日/週有效期。
- Clear Override：清除既有覆寫。
- Current Status：Value/Subject/Type/Busy period/Priority/Override Value/Override Until。
- Next Event：Next Event Value/Subject/Date。

**修改檔案**
- `ui/runtime_panel.py`（新）
- `core/runtime_state.py`（新）
- `database/sqlite_manager.py`（override/status CRUD）
- `CalendarUA.py`（tab 嵌入）

**驗收**
- 設定 temporary override 後，畫面與計算結果立刻反映，過期後自動回到排程值。
- Clear override 後，立即恢復以 resolver 結果為主。

### 5.12 Runtime Ribbon 對齊
- Tools: Create Event / Delete Event / Edit Event / Refresh Schedule / Apply Schedule。
- Configuration: Load / Save。
- 規則：Refresh 還原未儲存變更；Apply 才落盤。

**修改檔案**
- `CalendarUA.py`（toolbar 動作與交易流程）
- `database/sqlite_manager.py`（load/save/transaction support）

**驗收**
- 可明確區分「暫存變更」與「已套用變更」。

### 5.13 Category 管理
- 新增 category 下拉、色塊預覽與管理。
- series / occurrence / holiday 均可指定 category。
- runtime override 顯示也要套用 category（可用固定 override 色或指定類別）。

**修改檔案**
- `ui/category_manager_dialog.py`（新）
- `database/sqlite_manager.py`
- `CalendarUA.py`, `ScheduleEditDialog`

**驗收**
- 類別顏色改變後，所有相依事件立即更新。

---

## 6. 具體開發順序（建議照做）

1. 建立資料模型（migration + CRUD stub）。
2. 建立 resolver（先 series）。
3. 建立 Day/Week/Month 三視圖（只讀）。
4. 接上導航與 tab 殼層。
5. 補 Weekly tab。
6. 實作 occurrence vs series 選擇流程。
7. 補 exceptions。
8. 補 holidays。
9. 補 general profile。
10. 補 preview（唯讀）與 runtime（覆寫+狀態）。
11. 補 category 管理。
12. 最後做視覺細節、快捷鍵、效能優化。

---

## 7. 驗收清單（每次提交都跑）

## A. 功能驗收
- [ ] Day/Week/Month 切換正常。
- [ ] 日期導航正常（今天、前後切換、小月曆跳轉）。
- [ ] recurring item 點擊可選 occurrence / series。
- [ ] exception 只影響單次 occurrence。
- [ ] holiday 可覆寫同日排程。
- [ ] Preview 分頁不可修改事件（唯讀）。
- [ ] Runtime override/temporary override/clear override 行為正確。
- [ ] Current Status / Next Event 資訊正確更新。
- [ ] category 顏色與標題顯示正確。

## B. 資料一致性
- [ ] migration 可在舊 DB 上安全執行。
- [ ] 未設定新欄位的舊資料可正常顯示。
- [ ] 刪除 series 時，關聯 exception/holiday profile 不產生孤兒資料（或有防呆）。

## C. UI/效能
- [ ] 1,000 筆排程載入在可接受時間內（建議 < 1 秒級顯示，< 3 秒完整渲染）。
- [ ] 切換 Week/Month 無明顯卡頓。

---

## 8. 測試策略

### 單元測試（建議新增）
- `tests/test_rrule_parser.py`：確保基礎 RRULE 與邊界案例。
- `tests/test_schedule_resolver.py`：
  - series 展開
  - exception cancel/override
  - holiday 覆寫優先順序

### 整合測試
- 建立 demo 資料：
  - 綠色（關閉）
  - 紅色（自動）
  - 紫色（休假手動台）
- 驗證與你提供畫面一致的呈現情境。

---

## 9. 風險與對策

1. **RRULE + Exception 合併複雜**
   - 對策：先做 series-only，逐步加入 exception。
2. **UI 大改動風險高**
   - 對策：保留舊表格作 fallback，功能分段切換。
3. **舊資料相容性**
   - 對策：migration 全採「檢查後新增」，禁止 destructive 變更。
4. **渲染效能**
   - 對策：視圖只渲染當前可視範圍，使用快取。

---

## 10. 工作拆分（可直接變成 issue）

### Milestone 1：UI 框架與三視圖（1~2 週）
- [ ] 建立 shell + navigation
- [ ] 建立 Day/Week/Month read-only
- [ ] 建立 resolver(series-only)
- [ ] 建立右鍵選單與剪貼簿互動（Open/Delete/New/Copy/Cut/Paste）

### Milestone 2：Weekly + Occurrence/Series + Exceptions（1~2 週）
- [ ] weekly tab 與事件編輯流程
- [ ] occurrence choice dialog
- [ ] exception CRUD + UI
- [ ] resolver 合併 exception

### Milestone 3：Holidays + General + Preview + Runtime（1~2 週）
- [ ] holidays panel + CRUD
- [ ] general panel + profile
- [ ] preview 唯讀分頁
- [ ] runtime override/status/next-event

### Milestone 4：Categories + Ribbon + 回歸（3~7 天）
- [ ] category 管理與全域套用
- [ ] create/edit/delete/refresh/apply/load/save 對齊
- [ ] 測試、效能、UI 微調
- [ ] 文件更新（README / 使用手冊）

---

## 11. 第一個實作步驟（下一步就做這個）

> 下一個 commit 建議只做「**MVP-1：三視圖殼層 + 日期導航 + resolver(series-only)**」。

### 具體任務
1. 新增 `core/schedule_resolver.py`（先支援 series 展開）。
2. 新增 `ui/schedule_canvas.py`（Day/Week 畫布）。
3. 新增 `ui/month_grid.py`（Month 格子）。
4. 修改 `CalendarUA.py`：加入 `Day/Week/Month + Today + 前後切換`。
5. 保留原 `schedule_table` 作為暫時 fallback（可收合或次要區塊）。

### 完成標準
- 你附圖中的 Day/Week/Month 切換流程可跑起來。
- 同一資料在三種視圖都可看到（至少標題與顏色正確）。

---

## 13. 下一階段工作（優先順序）

### Phase 5: 清理與重構（✅ 已完成 2026-02-24）
- [x] 移除冗餘的左側日曆面板
- [x] 移除冗餘的當天排程摘要
- [x] 移除冗餘的排程管理表格（相容模式）
- [x] 調整主視窗為全寬單面板佈局
- [x] 統一所有日期選擇使用 Preview Tab 的 QDateEdit
- [x] 更新實施計畫文檔記錄完成狀態

### Phase 6: Category 管理系統（下一步）
- [ ] 建立 `ui/category_manager_dialog.py`
- [ ] 新增 category 下拉、色塊預覽與管理UI
- [ ] series / occurrence / holiday 均可指定 category
- [ ] runtime override 顯示也要套用 category顏色
- [ ] 類別顏色改變後，所有相依事件立即更新

### Phase 7: Ribbon 行為完整對齊
- [ ] Create Event: 從工具列/選單觸發新增
- [ ] Edit Event: 支援從各 Tab 觸發編輯流程
- [ ] Delete Event: 支援從各 Tab 觸發刪除流程
- [ ] Refresh Schedule: 重新載入資料並還原未套用的暫存顯示
- [ ] Apply Schedule: 將暫存變更寫入資料庫
- [ ] Load / Save: Profile 載入與儲存功能

### Phase 8: Holiday Resolver 整合
- [ ] 將 holiday_entries 整合進 schedule_resolver.py
- [ ] 覆寫邏輯：Holiday > Weekly（但 < Exception < Override）
- [ ] 測試假日覆寫優先權規則

### Phase 9: 效能與視覺優化
- [ ] 視圖只渲染當前可視範圍
- [ ] occurrence 快取機制
- [ ] 1,000+ 排程載入測試與優化
- [ ] 視覺主題細節調整（間距、字型、圖示）

---

## 14. 備註（授權與實務）
- 可以做出「相同操作邏輯與資訊架構」，但請避免直接複製第三方圖示、字型、素材檔。
- UI 文案可用通用詞（Edit Recurrence、Open this occurrence、Open the series）。

---

如果你同意，下一步就依本文件第 11 節開始：我先實作 `MVP-1`，完成後再按 Phase 2 往下推進。
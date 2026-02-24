# Phase 6 - Category 管理系統實作完成

## 完成時間
2026-02-24

## 實作內容

### 1. 資料庫層 (database/sqlite_manager.py)

#### 新增表格
- `schedule_categories` - Category 主表
  - id, name, bg_color, fg_color, sort_order, is_system, created_at, updated_at
  - 8 個預設系統 Categories (Red, Pink, Light Purple, Green, Blue, Yellow, Orange, Gray)

#### 欄位擴充 (Migration)
- `schedules` 表: category_id (DEFAULT 1), priority, location, description
- `schedule_exceptions` 表: override_category_id, note
- `holiday_entries` 表: override_category_id, override_target_value

#### 新增方法
- `_init_default_categories()` - 初始化 8 個預設 Categories
- `get_all_categories()` - 取得所有 Categories
- `get_category_by_id(category_id)` - 根據 ID 取得 Category
- `add_category(name, bg_color, fg_color, sort_order)` - 新增使用者 Category
- `update_category(category_id, ...)` - 更新 Category (保護系統 Category 名稱)
- `delete_category(category_id)` - 刪除 Category (保護系統 Category 與使用中 Category)

#### 修改方法
- `add_schedule()` - 加入 category_id 參數 (預設 1 = Red)
- `update_schedule()` - allowed_fields 加入 "category_id"

### 2. UI 層 (ui/category_manager_dialog.py)

#### CategoryEditorDialog
- 名稱編輯 (系統 Category 禁止編輯名稱)
- 背景顏色選擇器 (QColorDialog)
- 前景顏色選擇器 (QColorDialog)
- 即時預覽按鈕 (顯示顏色組合效果)

#### CategoryManagerDialog
- QTableWidget 顯示所有 Categories
- 5 欄位: ID, Name, BG Color, FG Color, System
- 系統 Category 標記 ★ 符號
- Add/Edit/Delete 按鈕與保護機制
- category_changed Signal 通知父視窗

### 3. 排程編輯整合 (CalendarUA.py - ScheduleEditDialog)

#### 新增欄位
- Category 下拉選單 (QComboBox)
- `_load_categories()` - 載入 Category 清單
- `load_data()` - 載入排程時設定 Category
- `get_data()` - 儲存時包含 category_id

#### 主視窗修改
- `add_schedule()` - 傳遞 category_id 到資料庫
- `edit_schedule()` - 更新時包含 category_id
- `manage_categories()` - 開啟 Category 管理對話框
- `on_category_changed()` - Category 變更後重新載入視圖

### 4. 顏色解析邏輯 (core/schedule_resolver.py)

#### 新增函數
- `_get_category_colors(db_manager, category_id, fallback_target_value)` 
  - 從資料庫查詢 Category 顏色
  - 失敗時使用 _pick_color() 作為 fallback

#### 修改函數
- `resolve_occurrences_for_range()` - 加入 db_manager 參數
  - 從排程的 category_id 取得顏色
  - Exception 可以覆寫 category (override_category_id)
  - 優先順序: Exception Category > Schedule Category

#### 主視窗整合
- `_resolve_day_occurrences()` - 傳遞 db_manager
- `_resolve_week_occurrences()` - 傳遞 db_manager
- `_resolve_month_occurrences()` - 傳遞 db_manager

### 5. UI 整合 (CalendarUA.py - Preview Tab)

#### 新增按鈕
- "管理 Category" 按鈕 (Preview Tab 右上方)
- 點擊開啟 CategoryManagerDialog

## 測試結果

### 資料庫測試
- ✅ 8 個系統 Categories 正確建立
- ✅ 排程與 Category 關聯正確
- ✅ Migration 安全執行 (PRAGMA table_info 檢查)

### 功能測試
- ✅ Category CRUD 全部正常 (Create/Read/Update/Delete)
- ✅ 系統 Category 保護機制有效 (無法刪除/重新命名)
- ✅ 使用中 Category 無法刪除
- ✅ Resolver 使用資料庫 Category 顏色 (不再依賴 target_value)

### 整合測試
- ✅ 排程編輯對話框顯示 Category 下拉選單
- ✅ 新增/編輯排程時可選擇 Category
- ✅ Category 顏色正確顯示在 Day/Week/Month 視圖
- ✅ Category 管理按鈕功能正常

## 檔案變更統計

### 新增檔案
- `ui/category_manager_dialog.py` (434 行)

### 修改檔案
- `database/sqlite_manager.py` (+241 行)
  - 新增表格建立
  - 新增 Migration 邏輯
  - 新增 6 個 Category CRUD 方法
  - 修改 add_schedule/update_schedule

- `CalendarUA.py` (+45 行)
  - ScheduleEditDialog 加入 Category 選單
  - 加入 manage_categories() 方法
  - 加入 _load_categories() 方法
  - 修改 add_schedule/edit_schedule 傳遞 category_id
  - Preview Tab 加入管理按鈕

- `core/schedule_resolver.py` (+33 行)
  - 加入 _get_category_colors() 函數
  - 修改 resolve_occurrences_for_range() 加入 db_manager 參數
  - 使用資料庫 Category 顏色取代 _pick_color()

## 預設 Category 列表

| ID | Name                          | BG Color | FG Color | 用途說明           |
|----|-------------------------------|----------|----------|-------------------|
| 1  | Red (關閉)                    | #FF0000  | #FFFFFF  | 關閉操作           |
| 2  | Pink (自動)                   | #FF69B4  | #FFFFFF  | 自動操作           |
| 3  | Light Purple (休假手動台)     | #DDA0DD  | #000000  | 休假/手動台操作    |
| 4  | Green                         | #00FF00  | #000000  | 一般操作 1         |
| 5  | Blue                          | #0000FF  | #FFFFFF  | 一般操作 2         |
| 6  | Yellow                        | #FFFF00  | #000000  | 一般操作 3         |
| 7  | Orange                        | #FFA500  | #000000  | 一般操作 4         |
| 8  | Gray                          | #808080  | #FFFFFF  | 其他               |

## 後續工作建議

### Phase 7 - Ribbon 功能完整化
- Create/Edit/Delete/Refresh 按鈕整合
- Apply/Load/Save 設定功能
- 快速鍵綁定

### Phase 8 - Category 進階功能
- Category 拖曳排序
- Category 圖示支援
- Category 搜尋/過濾
- Holiday Calendar 的 Category 整合

### Phase 9 - 視圖優化
- Category 圖例顯示
- Category 過濾器 (Preview Tab)
- Category 統計報表
- Batch Category 操作 (批次修改排程 Category)

## 已知限制

1. Category 顏色變更後需手動重新載入視圖 (已有 Signal 通知)
2. 目前未實作 Holiday Calendar 的 Category 整合
3. Exception 的 override_category_id 尚未在 UI 中開放編輯
4. Category 排序目前使用 sort_order + name,未提供 UI 拖曳

## 相關文件

- 實作計畫: SCHEDULEWORX_UI_IMPLEMENTATION_PLAN.md
- 資料庫 Schema: database/sqlite_manager.py (init_db 方法)
- UI 設計: ui/category_manager_dialog.py

"""驗證 Category 系統整合"""
from database.sqlite_manager import SQLiteManager
from core.schedule_resolver import resolve_occurrences_for_range
from datetime import datetime, timedelta

db = SQLiteManager()
db.init_db()

print("=== 1. 驗證 Category 資料庫 ===")
categories = db.get_all_categories()
print(f"總共 {len(categories)} 個 Categories:")
for cat in categories:
    print(f"  {cat['id']:2d}. {cat['name']:30s} - BG:{cat['bg_color']} FG:{cat['fg_color']}")

print("\n=== 2. 驗證排程-Category 關聯 ===")
schedules = db.get_all_schedules()
print(f"總共 {len(schedules)} 筆排程:")
for s in schedules:
    cat_id = s.get('category_id', 'N/A')
    cat = db.get_category_by_id(cat_id) if isinstance(cat_id, int) else None
    cat_name = cat['name'] if cat else 'Unknown'
    print(f"  ID:{s['id']:2d} {s['task_name']:30s} - Category: {cat_id} ({cat_name})")

print("\n=== 3. 驗證 Resolver 顏色邏輯 ===")
if schedules:
    # 測試解析一週的 occurrences
    now = datetime.now()
    week_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    week_end = week_start + timedelta(days=7)
    
    occurrences = resolve_occurrences_for_range(
        schedules, week_start, week_end, None, None, db
    )
    
    print(f"本週有 {len(occurrences)} 個排程觸發:")
    for occ in occurrences[:5]:  # 只顯示前 5 個
        print(f"  {occ.start.strftime('%Y-%m-%d %H:%M')} - {occ.title:30s} - BG:{occ.category_bg} FG:{occ.category_fg}")
    
    if len(occurrences) > 5:
        print(f"  ... 還有 {len(occurrences) - 5} 個")
else:
    print("沒有排程資料")

print("\n=== 4. 驗證 Category CRUD 操作 ===")
# 測試新增
new_id = db.add_category("Test Category", "#123456", "#ABCDEF", 100)
if new_id:
    print(f"✓ 新增測試 Category 成功 (ID: {new_id})")
    
    # 測試讀取
    test_cat = db.get_category_by_id(new_id)
    if test_cat and test_cat['name'] == "Test Category":
        print(f"✓ 讀取測試 Category 成功: {test_cat['name']}")
    
    # 測試更新
    if db.update_category(new_id, bg_color="#FEDCBA"):
        updated = db.get_category_by_id(new_id)
        if updated['bg_color'] == "#FEDCBA":
            print(f"✓ 更新測試 Category 成功: {updated['bg_color']}")
    
    # 測試刪除
    if db.delete_category(new_id):
        print(f"✓ 刪除測試 Category 成功")
    else:
        print(f"✗ 刪除測試 Category 失敗")
else:
    print("✗ 新增測試 Category 失敗")

print("\n=== 測試完成 ===")

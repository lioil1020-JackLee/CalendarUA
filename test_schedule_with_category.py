"""測試排程與 Category 整合"""
from database.sqlite_manager import SQLiteManager

db = SQLiteManager()
db.init_db()

# 新增測試排程,使用不同的 category_id
test_schedules = [
    {
        "task_name": "測試任務 - Red",
        "opc_url": "opc.tcp://localhost:4840",
        "node_id": "ns=2;i=1001",
        "target_value": "1",
        "rrule_str": "FREQ=DAILY",
        "category_id": 1,  # Red
    },
    {
        "task_name": "測試任務 - Pink",
        "opc_url": "opc.tcp://localhost:4840",
        "node_id": "ns=2;i=1002",
        "target_value": "1",
        "rrule_str": "FREQ=DAILY",
        "category_id": 2,  # Pink
    },
    {
        "task_name": "測試任務 - Light Purple",
        "opc_url": "opc.tcp://localhost:4840",
        "node_id": "ns=2;i=1003",
        "target_value": "1",
        "rrule_str": "FREQ=DAILY",
        "category_id": 3,  # Light Purple
    },
]

print("=== 新增測試排程 ===")
for schedule_data in test_schedules:
    schedule_id = db.add_schedule(**schedule_data)
    if schedule_id:
        print(f"✓ 新增成功: {schedule_data['task_name']} (ID: {schedule_id}, Category: {schedule_data['category_id']})")
    else:
        print(f"✗ 新增失敗: {schedule_data['task_name']}")

print("\n=== 查詢所有排程 ===")
schedules = db.get_all_schedules()
for s in schedules:
    cat_id = s.get('category_id', 'N/A')
    cat = db.get_category_by_id(cat_id) if cat_id != 'N/A' else None
    cat_name = cat['name'] if cat else 'Unknown'
    print(f"ID: {s['id']:2d} | {s['task_name']:30s} | Category: {cat_id} ({cat_name})")

print(f"\n總共 {len(schedules)} 筆排程")

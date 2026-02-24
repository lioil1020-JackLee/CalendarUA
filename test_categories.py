"""測試 Category 功能"""
from database.sqlite_manager import SQLiteManager

db = SQLiteManager()
print('初始化資料庫:', db.init_db())

cats = db.get_all_categories()
print(f'\n總共有 {len(cats)} 個 Categories:')
for c in cats:
    system_mark = '★' if c['is_system'] else '  '
    print(f"{system_mark} {c['id']:2d}: {c['name']:30s} - BG:{c['bg_color']} FG:{c['fg_color']}")

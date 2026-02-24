"""
SQLite 資料庫管理模組
使用 Python 內建 sqlite3 模組進行排程資料的 CRUD 操作
"""

import sqlite3
import logging
from pathlib import Path
import os
from typing import List, Dict, Optional, Any
from contextlib import contextmanager


# 設定日誌記錄
logging.basicConfig(
    level=logging.WARNING, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


class SQLiteManager:
    """
    SQLite 資料庫管理類別
    負責排程資料的建立、讀取、更新與刪除操作
    """

    def __init__(self, db_path: str = "./database/calendarua.db"):
        """
        初始化 SQLite 管理器

        Args:
            db_path: 資料庫檔案路徑，預設為 ./database/calendarua.db
        """
        self.db_path = Path(db_path)
        # 確保資料庫目錄存在
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        logger.info(f"SQLite 管理器初始化完成，資料庫路徑: {self.db_path}")

    @contextmanager
    def _get_connection(self):
        """
        取得資料庫連線的上下文管理器
        自動處理連線的開啟與關閉

        Yields:
            sqlite3.Connection: SQLite 連線物件
        """
        conn = None
        try:
            # 建立連線，啟用外鍵支援
            conn = sqlite3.connect(self.db_path)
            conn.execute("PRAGMA foreign_keys = ON")
            # 設定回傳結果為字典格式
            conn.row_factory = sqlite3.Row
            yield conn
        except sqlite3.Error as e:
            logger.error(f"資料庫連線錯誤: {e}")
            raise
        finally:
            if conn:
                conn.close()

    def init_db(self) -> bool:
        """
        初始化資料庫
        若資料庫檔案不存在則自動建立，並初始化 schedules 表格

        Returns:
            bool: 初始化成功回傳 True，否則回傳 False
        """
        # 建立表格的 SQL 語句
        create_table_sql = """
        CREATE TABLE IF NOT EXISTS schedules (
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
        """

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                # 建立表格
                cursor.execute(create_table_sql)

                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS schedule_exceptions (
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
                    """
                )

                # 創建 holiday_calendars 表
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS holiday_calendars (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        name TEXT UNIQUE NOT NULL,
                        description TEXT,
                        is_default INTEGER DEFAULT 0,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                    """
                )

                # 創建 holiday_entries 表
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS holiday_entries (
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
                    """
                )

                # 創建 general_settings 表
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS general_settings (
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
                    """
                )

                # 創建 runtime_override 表
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS runtime_override (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        override_value TEXT NOT NULL,
                        override_until TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                    """
                )

                # 創建 schedule_categories 表
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS schedule_categories (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        name TEXT UNIQUE NOT NULL,
                        bg_color TEXT NOT NULL,
                        fg_color TEXT NOT NULL,
                        sort_order INTEGER DEFAULT 0,
                        is_system INTEGER DEFAULT 0,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                    """
                )

                # 檢查並添加缺少的欄位
                cursor.execute("PRAGMA table_info(schedules)")
                columns = [column[1] for column in cursor.fetchall()]
                
                if 'last_execution_status' not in columns:
                    cursor.execute("ALTER TABLE schedules ADD COLUMN last_execution_status TEXT DEFAULT ''")
                    logger.info("已添加 last_execution_status 欄位")
                
                if 'last_execution_time' not in columns:
                    cursor.execute("ALTER TABLE schedules ADD COLUMN last_execution_time TIMESTAMP")
                    logger.info("已添加 last_execution_time 欄位")

                # 建立索引以提升查詢效能
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_schedules_enabled ON schedules(is_enabled)"
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_schedules_node_id ON schedules(node_id)"
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_schedule_exceptions_schedule_date ON schedule_exceptions(schedule_id, occurrence_date)"
                )

                conn.commit()
                logger.info("資料庫初始化成功，schedules 表格已建立")
                
                # 執行遷移以添加缺失的欄位
                self._migrate_db()
                
                # 初始化預設 categories
                self._init_default_categories()
                
                return True

        except sqlite3.Error as e:
            logger.error(f"資料庫初始化失敗: {e}")
            return False

    def _migrate_db(self) -> None:
        """
        遷移資料庫以添加新的欄位
        檢查並添加缺失的欄位（用於舊版本升級）
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                
                # 檢查 opc_security_mode 欄位是否存在
                cursor.execute("PRAGMA table_info(schedules)")
                columns = [column[1] for column in cursor.fetchall()]
                
                # 添加缺失的欄位
                if "opc_security_mode" not in columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN opc_security_mode TEXT DEFAULT 'None'"
                    )
                    logger.info("已添加 opc_security_mode 欄位")
                
                if "opc_security_policy" not in columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN opc_security_policy TEXT DEFAULT 'None'"
                    )
                    logger.info("已添加 opc_security_policy 欄位")
                
                if "opc_username" not in columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN opc_username TEXT DEFAULT ''"
                    )
                    logger.info("已添加 opc_username 欄位")
                
                if "opc_password" not in columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN opc_password TEXT DEFAULT ''"
                    )
                    logger.info("已添加 opc_password 欄位")
                
                if "opc_timeout" not in columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN opc_timeout INTEGER DEFAULT 5"
                    )
                    logger.info("已添加 opc_timeout 欄位")
                
                if "opc_write_timeout" not in columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN opc_write_timeout INTEGER DEFAULT 3"
                    )
                    logger.info("已添加 opc_write_timeout 欄位")
                
                if "last_execution_status" not in columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN last_execution_status TEXT DEFAULT '尚未執行'"
                    )
                    logger.info("已添加 last_execution_status 欄位")
                
                if "last_execution_time" not in columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN last_execution_time TIMESTAMP"
                    )
                    logger.info("已添加 last_execution_time 欄位")
                
                if "next_execution_time" not in columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN next_execution_time TIMESTAMP"
                    )
                    logger.info("已添加 next_execution_time 欄位")

                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS schedule_exceptions (
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
                    """
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_schedule_exceptions_schedule_date ON schedule_exceptions(schedule_id, occurrence_date)"
                )

                # 創建 holiday_calendars 表
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS holiday_calendars (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        name TEXT UNIQUE NOT NULL,
                        description TEXT,
                        is_default INTEGER DEFAULT 0,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                    """
                )

                # 創建 holiday_entries 表
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS holiday_entries (
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
                    """
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_holiday_entries_calendar_date ON holiday_entries(calendar_id, holiday_date)"
                )
                
                # ===== Category 相關欄位遷移 (Phase 6) =====
                
                # 為 schedules 表添加 category 相關欄位
                cursor.execute("PRAGMA table_info(schedules)")
                schedules_columns = [column[1] for column in cursor.fetchall()]
                
                if "category_id" not in schedules_columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN category_id INTEGER DEFAULT 1"
                    )
                    logger.info("已添加 schedules.category_id 欄位")
                
                if "priority" not in schedules_columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN priority INTEGER DEFAULT 1"
                    )
                    logger.info("已添加 schedules.priority 欄位")
                
                if "location" not in schedules_columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN location TEXT DEFAULT ''"
                    )
                    logger.info("已添加 schedules.location 欄位")
                
                if "description" not in schedules_columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN description TEXT DEFAULT ''"
                    )
                    logger.info("已添加 schedules.description 欄位")
                
                # 為 schedule_exceptions 表添加 category 相關欄位
                cursor.execute("PRAGMA table_info(schedule_exceptions)")
                exceptions_columns = [column[1] for column in cursor.fetchall()]
                
                if "override_category_id" not in exceptions_columns:
                    cursor.execute(
                        "ALTER TABLE schedule_exceptions ADD COLUMN override_category_id INTEGER"
                    )
                    logger.info("已添加 schedule_exceptions.override_category_id 欄位")
                
                if "note" not in exceptions_columns:
                    cursor.execute(
                        "ALTER TABLE schedule_exceptions ADD COLUMN note TEXT DEFAULT ''"
                    )
                    logger.info("已添加 schedule_exceptions.note 欄位")
                
                # 為 holiday_entries 表添加 override 相關欄位
                cursor.execute("PRAGMA table_info(holiday_entries)")
                holiday_entries_columns = [column[1] for column in cursor.fetchall()]
                
                if "override_category_id" not in holiday_entries_columns:
                    cursor.execute(
                        "ALTER TABLE holiday_entries ADD COLUMN override_category_id INTEGER"
                    )
                    logger.info("已添加 holiday_entries.override_category_id 欄位")
                
                if "override_target_value" not in holiday_entries_columns:
                    cursor.execute(
                        "ALTER TABLE holiday_entries ADD COLUMN override_target_value TEXT"
                    )
                    logger.info("已添加 holiday_entries.override_target_value 欄位")
                
                # 建立索引
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_schedules_category ON schedules(category_id)"
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_exceptions_category ON schedule_exceptions(override_category_id)"
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_holiday_entries_category ON holiday_entries(override_category_id)"
                )
                
                conn.commit()
                
        except sqlite3.Error as e:
            logger.error(f"資料庫遷移失敗: {e}")

    def _init_default_categories(self) -> None:
        """
        初始化預設 categories
        如果 schedule_categories 表為空，插入系統預設類別
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                
                # 檢查是否已有 categories
                cursor.execute("SELECT COUNT(*) as count FROM schedule_categories")
                count = cursor.fetchone()['count']
                
                if count == 0:
                    # 插入系統預設類別
                    default_categories = [
                        ("Red (關閉)", "#FF0000", "#FFFFFF", 1, 1),
                        ("Pink (自動)", "#FF69B4", "#FFFFFF", 2, 1),
                        ("Light Purple (休假手動台)", "#DDA0DD", "#000000", 3, 1),
                        ("Green", "#00FF00", "#000000", 4, 1),
                        ("Blue", "#0000FF", "#FFFFFF", 5, 1),
                        ("Yellow", "#FFFF00", "#000000", 6, 1),
                        ("Orange", "#FFA500", "#000000", 7, 1),
                        ("Gray", "#808080", "#FFFFFF", 8, 1),
                    ]
                    
                    cursor.executemany(
                        """
                        INSERT INTO schedule_categories 
                        (name, bg_color, fg_color, sort_order, is_system)
                        VALUES (?, ?, ?, ?, ?)
                        """,
                        default_categories
                    )
                    
                    conn.commit()
                    logger.info(f"已插入 {len(default_categories)} 個預設 categories")
                
        except sqlite3.Error as e:
            logger.error(f"初始化預設 categories 失敗: {e}")

    # ===== Category CRUD 方法 =====

    def get_all_categories(self) -> List[Dict[str, Any]]:
        """
        取得所有 categories
        
        Returns:
            List[Dict]: categories 列表
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    SELECT id, name, bg_color, fg_color, sort_order, is_system,
                           created_at, updated_at
                    FROM schedule_categories
                    ORDER BY sort_order, name
                    """
                )
                categories = [dict(row) for row in cursor.fetchall()]
                return categories
        except sqlite3.Error as e:
            logger.error(f"取得 categories 失敗: {e}")
            return []

    def get_category_by_id(self, category_id: int) -> Optional[Dict[str, Any]]:
        """
        根據 ID 取得 category
        
        Args:
            category_id: Category ID
            
        Returns:
            Optional[Dict]: Category 資料或 None
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    SELECT id, name, bg_color, fg_color, sort_order, is_system,
                           created_at, updated_at
                    FROM schedule_categories
                    WHERE id = ?
                    """,
                    (category_id,)
                )
                row = cursor.fetchone()
                return dict(row) if row else None
        except sqlite3.Error as e:
            logger.error(f"取得 category 失敗: {e}")
            return None

    def add_category(
        self,
        name: str,
        bg_color: str,
        fg_color: str,
        sort_order: int = 0
    ) -> Optional[int]:
        """
        新增 category
        
        Args:
            name: Category 名稱
            bg_color: 背景顏色 (hex)
            fg_color: 前景顏色 (hex)
            sort_order: 排序順序
            
        Returns:
            Optional[int]: 新增的 category ID 或 None
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    INSERT INTO schedule_categories 
                    (name, bg_color, fg_color, sort_order, is_system)
                    VALUES (?, ?, ?, ?, 0)
                    """,
                    (name, bg_color, fg_color, sort_order)
                )
                conn.commit()
                return cursor.lastrowid
        except sqlite3.IntegrityError:
            logger.error(f"Category 名稱已存在: {name}")
            return None
        except sqlite3.Error as e:
            logger.error(f"新增 category 失敗: {e}")
            return None

    def update_category(
        self,
        category_id: int,
        name: str = None,
        bg_color: str = None,
        fg_color: str = None,
        sort_order: int = None
    ) -> bool:
        """
        更新 category
        
        Args:
            category_id: Category ID
            name: 新名稱（可選）
            bg_color: 新背景顏色（可選）
            fg_color: 新前景顏色（可選）
            sort_order: 新排序順序（可選）
            
        Returns:
            bool: 是否成功
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                
                # 先檢查是否為系統類別
                cursor.execute(
                    "SELECT is_system FROM schedule_categories WHERE id = ?",
                    (category_id,)
                )
                row = cursor.fetchone()
                if not row:
                    logger.error(f"Category 不存在: {category_id}")
                    return False
                
                # 建立更新語句
                updates = []
                params = []
                
                if name is not None:
                    updates.append("name = ?")
                    params.append(name)
                if bg_color is not None:
                    updates.append("bg_color = ?")
                    params.append(bg_color)
                if fg_color is not None:
                    updates.append("fg_color = ?")
                    params.append(fg_color)
                if sort_order is not None:
                    updates.append("sort_order = ?")
                    params.append(sort_order)
                
                if not updates:
                    return True
                
                updates.append("updated_at = CURRENT_TIMESTAMP")
                params.append(category_id)
                
                cursor.execute(
                    f"UPDATE schedule_categories SET {', '.join(updates)} WHERE id = ?",
                    params
                )
                conn.commit()
                return cursor.rowcount > 0
                
        except sqlite3.IntegrityError:
            logger.error(f"Category 名稱已存在: {name}")
            return False
        except sqlite3.Error as e:
            logger.error(f"更新 category 失敗: {e}")
            return False

    def delete_category(self, category_id: int) -> bool:
        """
        刪除 category（系統類別不可刪除）
        
        Args:
            category_id: Category ID
            
        Returns:
            bool: 是否成功
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                
                # 檢查是否為系統類別
                cursor.execute(
                    "SELECT is_system FROM schedule_categories WHERE id = ?",
                    (category_id,)
                )
                row = cursor.fetchone()
                if not row:
                    logger.error(f"Category 不存在: {category_id}")
                    return False
                
                if row['is_system']:
                    logger.error(f"無法刪除系統 category: {category_id}")
                    return False
                
                # 檢查是否有排程使用此 category
                cursor.execute(
                    "SELECT COUNT(*) as count FROM schedules WHERE category_id = ?",
                    (category_id,)
                )
                count = cursor.fetchone()['count']
                if count > 0:
                    logger.error(f"無法刪除 category，有 {count} 個排程正在使用: {category_id}")
                    return False
                
                cursor.execute(
                    "DELETE FROM schedule_categories WHERE id = ?",
                    (category_id,)
                )
                conn.commit()
                return cursor.rowcount > 0
                
        except sqlite3.Error as e:
            logger.error(f"刪除 category 失敗: {e}")
            return False

    def add_schedule(
        self,
        task_name: str,
        opc_url: str,
        node_id: str,
        target_value: str,
        rrule_str: str,
        data_type: str = "auto",
        category_id: int = 1,
        opc_security_policy: str = "None",
        opc_security_mode: str = "None",
        opc_username: str = "",
        opc_password: str = "",
        opc_timeout: int = 5,
        opc_write_timeout: int = 3,
        is_enabled: int = 1,
    ) -> Optional[int]:
        """
        新增排程資料

        Args:
            task_name: 任務名稱
            opc_url: OPC UA 伺服器位址
            node_id: OPC UA Tag NodeID
            target_value: 要寫入的數值
            rrule_str: RRULE 規則字串
            data_type: 資料型別 (auto/int/float/string/bool)
            category_id: Category ID (預設 1 = Red)
            opc_security_policy: OPC安全策略
            opc_security_mode: OPC安全模式 (None/Sign/SignAndEncrypt)
            opc_username: OPC使用者名稱
            opc_password: OPC密碼
            opc_timeout: 連線超時秒數
            opc_write_timeout: 寫值重試延遲秒數
            is_enabled: 是否啟用 (1: 啟用, 0: 停用)，預設為 1

        Returns:
            Optional[int]: 新增排程的 ID，失敗時回傳 None
        """
        insert_sql = """
        INSERT INTO schedules (task_name, opc_url, node_id, target_value, data_type, rrule_str,
                              category_id, opc_security_policy, opc_security_mode, opc_username, opc_password, opc_timeout, opc_write_timeout, is_enabled)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    insert_sql,
                    (task_name, opc_url, node_id, target_value, data_type, rrule_str,
                     category_id, opc_security_policy, opc_security_mode, opc_username, opc_password, opc_timeout, opc_write_timeout, is_enabled),
                )
                conn.commit()
                new_id = cursor.lastrowid
                logger.info(f"排程新增成功，ID: {new_id}")
                return new_id

        except sqlite3.Error as e:
            logger.error(f"新增排程失敗: {e}")
            return None

    def get_all_schedules(self, enabled_only: bool = False) -> List[Dict[str, Any]]:
        """
        查詢所有排程資料

        Args:
            enabled_only: 是否只查詢啟用的排程，預設為 False

        Returns:
            List[Dict[str, Any]]: 排程資料列表，每個項目為字典格式
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()

                if enabled_only:
                    # 只查詢啟用的排程
                    cursor.execute(
                        "SELECT * FROM schedules WHERE is_enabled = 1 ORDER BY created_at DESC"
                    )
                else:
                    # 查詢所有排程
                    cursor.execute("SELECT * FROM schedules ORDER BY created_at DESC")

                # 將查詢結果轉換為字典列表
                rows = cursor.fetchall()
                schedules = [dict(row) for row in rows]

                logger.info(f"查詢到 {len(schedules)} 筆排程資料")
                return schedules

        except sqlite3.Error as e:
            logger.error(f"查詢排程失敗: {e}")
            return []

    def delete_schedule(self, schedule_id: int) -> bool:
        """
        刪除指定排程

        Args:
            schedule_id: 要刪除的排程 ID

        Returns:
            bool: 刪除成功回傳 True，否則回傳 False
        """
        delete_sql = "DELETE FROM schedules WHERE id = ?"

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(delete_sql, (schedule_id,))
                conn.commit()

                # 檢查是否有資料被刪除
                if cursor.rowcount > 0:
                    logger.info(f"排程 {schedule_id} 刪除成功")
                    return True
                else:
                    logger.warning(f"排程 {schedule_id} 不存在，無法刪除")
                    return False

        except sqlite3.Error as e:
            logger.error(f"刪除排程失敗: {e}")
            return False

    # 以下為相容性方法，與舊版 MySQLManager 介面保持一致

    def create_schedule(
        self,
        task_name: str,
        opc_url: str,
        node_id: str,
        target_value: str,
        rrule_str: str,
        data_type: str = "auto",
        opc_security_policy: str = "None",
        opc_security_mode: str = "None",
        opc_username: str = "",
        opc_password: str = "",
        opc_timeout: int = 5,
        opc_write_timeout: int = 3,
        is_enabled: int = 1,
    ) -> Optional[int]:
        """
        新增排程（add_schedule 的別名，與舊版介面相容）

        Args:
            task_name: 任務名稱
            opc_url: OPC UA 伺服器位址
            node_id: OPC UA Tag NodeID
            target_value: 要寫入的數值
            rrule_str: RRULE 規則字串
            opc_security_policy: OPC安全策略
            opc_security_mode: OPC安全模式 (None/Sign/SignAndEncrypt)
            opc_username: OPC使用者名稱
            opc_password: OPC密碼
            opc_timeout: 連線超時秒數
            opc_write_timeout: 寫值重試延遲秒數
            is_enabled: 是否啟用 (1: 啟用, 0: 停用)，預設為 1

        Returns:
            Optional[int]: 新增排程的 ID，失敗時回傳 None
        """
        return self.add_schedule(
            task_name=task_name,
            opc_url=opc_url,
            node_id=node_id,
            target_value=target_value,
            data_type=data_type,
            rrule_str=rrule_str,
            opc_security_policy=opc_security_policy,
            opc_security_mode=opc_security_mode,
            opc_username=opc_username,
            opc_password=opc_password,
            opc_timeout=opc_timeout,
            opc_write_timeout=opc_write_timeout,
            is_enabled=is_enabled,
        )

    def get_schedule(self, schedule_id: int) -> Optional[Dict[str, Any]]:
        """
        查詢單一排程

        Args:
            schedule_id: 排程 ID

        Returns:
            Optional[Dict[str, Any]]: 排程資料字典，找不到時回傳 None
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM schedules WHERE id = ?", (schedule_id,))
                row = cursor.fetchone()

                if row:
                    logger.info(f"查詢到排程 {schedule_id}")
                    return dict(row)
                else:
                    logger.warning(f"排程 {schedule_id} 不存在")
                    return None

        except sqlite3.Error as e:
            logger.error(f"查詢排程失敗: {e}")
            return None

    def update_schedule(self, schedule_id: int, **kwargs) -> bool:
        """
        更新排程資料

        Args:
            schedule_id: 要更新的排程 ID
            **kwargs: 要更新的欄位（task_name, opc_url, node_id, target_value, rrule_str, category_id, is_enabled）

        Returns:
            bool: 更新成功回傳 True，否則回傳 False
        """
        # 定義允許更新的欄位
        allowed_fields = {
            "task_name",
            "opc_url",
            "node_id",
            "target_value",
            "data_type",
            "rrule_str",
            "category_id",
            "opc_security_policy",
            "opc_security_mode",
            "opc_username",
            "opc_password",
            "opc_timeout",
            "opc_write_timeout",
            "is_enabled",
            "last_execution_status",
            "last_execution_time",
        }

        # 過濾出有效的更新欄位
        updates = {k: v for k, v in kwargs.items() if k in allowed_fields}

        if not updates:
            logger.warning("沒有提供要更新的欄位")
            return False

        # 建立 UPDATE 語句
        set_clause = ", ".join([f"{k} = ?" for k in updates.keys()])
        update_sql = f"UPDATE schedules SET {set_clause} WHERE id = ?"
        values = list(updates.values()) + [schedule_id]

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(update_sql, values)
                conn.commit()

                if cursor.rowcount > 0:
                    logger.info(f"排程 {schedule_id} 更新成功")
                    return True
                else:
                    logger.warning(f"排程 {schedule_id} 不存在或無變更")
                    return False

        except sqlite3.Error as e:
            logger.error(f"更新排程失敗: {e}")
            return False

    def toggle_schedule(self, schedule_id: int, is_enabled: int) -> bool:
        """
        切換排程啟用狀態

        Args:
            schedule_id: 排程 ID
            is_enabled: 1 表示啟用，0 表示停用

        Returns:
            bool: 操作成功回傳 True，否則回傳 False
        """
        return self.update_schedule(schedule_id, is_enabled=is_enabled)

    def update_execution_status(self, schedule_id: int, status: str) -> bool:
        """
        更新排程的執行狀態

        Args:
            schedule_id: 排程 ID
            status: 執行狀態描述

        Returns:
            bool: 更新成功回傳 True，否則回傳 False
        """
        from datetime import datetime
        return self.update_schedule(
            schedule_id,
            last_execution_status=status,
            last_execution_time=datetime.now()
        )

    def clear_all_schedules(self) -> bool:
        """
        清除所有排程資料

        Returns:
            bool: 清除成功回傳 True，否則回傳 False
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM schedules")
                conn.commit()

                deleted_count = cursor.rowcount
                logger.info(f"已清除 {deleted_count} 筆排程資料")
                return True

        except sqlite3.Error as e:
            logger.error(f"清除排程資料失敗: {e}")
            return False

    def get_next_task_name(self) -> str:
        """
        獲取下一個預設任務名稱（任務1、任務2、任務3...）

        Returns:
            str: 下一個可用的任務名稱
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                
                # 查詢所有以"任務"開頭的任務名稱
                cursor.execute("SELECT task_name FROM schedules WHERE task_name LIKE '任務%'")
                existing_names = cursor.fetchall()
                
                # 提取數字部分
                numbers = []
                for (name,) in existing_names:
                    if name.startswith('任務') and len(name) > 2:
                        try:
                            num = int(name[2:])  # 去掉"任務"取數字
                            numbers.append(num)
                        except ValueError:
                            continue
                
                # 找出下一個可用的數字
                next_num = 1
                if numbers:
                    numbers.sort()
                    # 找到第一個缺失的數字
                    for i, num in enumerate(numbers, 1):
                        if num != i:
                            next_num = i
                            break
                    else:
                        next_num = len(numbers) + 1
                
                return f"任務{next_num}"

        except sqlite3.Error as e:
            logger.error(f"獲取下一個任務名稱失敗: {e}")
            return "任務1"

    def add_schedule_exception_override(
        self,
        schedule_id: int,
        occurrence_date: str,
        override_start: str,
        override_end: str,
        override_task_name: str,
        override_target_value: str,
    ) -> Optional[int]:
        """新增或覆寫單次 occurrence 的 exception"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()

                cursor.execute(
                    "DELETE FROM schedule_exceptions WHERE schedule_id = ? AND occurrence_date = ?",
                    (schedule_id, occurrence_date),
                )

                cursor.execute(
                    """
                    INSERT INTO schedule_exceptions
                    (schedule_id, occurrence_date, action, override_start, override_end, override_task_name, override_target_value)
                    VALUES (?, ?, 'override', ?, ?, ?, ?)
                    """,
                    (
                        schedule_id,
                        occurrence_date,
                        override_start,
                        override_end,
                        override_task_name,
                        override_target_value,
                    ),
                )
                conn.commit()
                return cursor.lastrowid
        except sqlite3.Error as e:
            logger.error(f"新增 exception 失敗: {e}")
            return None

    def add_schedule_exception_cancel(self, schedule_id: int, occurrence_date: str) -> Optional[int]:
        """新增單次 occurrence 的取消 exception"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "DELETE FROM schedule_exceptions WHERE schedule_id = ? AND occurrence_date = ?",
                    (schedule_id, occurrence_date),
                )
                cursor.execute(
                    """
                    INSERT INTO schedule_exceptions (schedule_id, occurrence_date, action)
                    VALUES (?, ?, 'cancel')
                    """,
                    (schedule_id, occurrence_date),
                )
                conn.commit()
                return cursor.lastrowid
        except sqlite3.Error as e:
            logger.error(f"新增 cancel exception 失敗: {e}")
            return None

    def get_all_schedule_exceptions(self) -> List[Dict[str, Any]]:
        """查詢所有 exception 記錄"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM schedule_exceptions ORDER BY occurrence_date DESC, id DESC")
                return [dict(row) for row in cursor.fetchall()]
        except sqlite3.Error as e:
            logger.error(f"查詢 exceptions 失敗: {e}")
            return []

    def delete_schedule_exception(self, exception_id: int) -> bool:
        """刪除 exception 記錄（按 ID）"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM schedule_exceptions WHERE id = ?", (exception_id,))
                conn.commit()
                return cursor.rowcount > 0
        except sqlite3.Error as e:
            logger.error(f"刪除 exception 失敗: {e}")
            return False

    # ==================== Holiday Calendars CRUD ====================

    def add_holiday_calendar(self, name: str, description: str = "", is_default: int = 0) -> Optional[int]:
        """新增假日日曆"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    INSERT INTO holiday_calendars (name, description, is_default)
                    VALUES (?, ?, ?)
                    """,
                    (name, description, is_default),
                )
                conn.commit()
                return cursor.lastrowid
        except sqlite3.Error as e:
            logger.error(f"新增 holiday calendar 失敗: {e}")
            return None

    def get_all_holiday_calendars(self) -> List[Dict[str, Any]]:
        """查詢所有假日日曆"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM holiday_calendars ORDER BY is_default DESC, name")
                return [dict(row) for row in cursor.fetchall()]
        except sqlite3.Error as e:
            logger.error(f"查詢 holiday calendars 失敗: {e}")
            return []

    def update_holiday_calendar(self, calendar_id: int, name: str, description: str = "", is_default: int = 0) -> bool:
        """更新假日日曆"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    UPDATE holiday_calendars
                    SET name = ?, description = ?, is_default = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                    """,
                    (name, description, is_default, calendar_id),
                )
                conn.commit()
                return cursor.rowcount > 0
        except sqlite3.Error as e:
            logger.error(f"更新 holiday calendar 失敗: {e}")
            return False

    def delete_holiday_calendar(self, calendar_id: int) -> bool:
        """刪除假日日曆（CASCADE 會自動刪除關聯的 entries）"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM holiday_calendars WHERE id = ?", (calendar_id,))
                conn.commit()
                return cursor.rowcount > 0
        except sqlite3.Error as e:
            logger.error(f"刪除 holiday calendar 失敗: {e}")
            return False

    # ==================== Holiday Entries CRUD ====================

    def add_holiday_entry(
        self,
        calendar_id: int,
        holiday_date: str,
        name: str,
        is_full_day: int = 1,
        start_time: str = None,
        end_time: str = None,
    ) -> Optional[int]:
        """新增假日條目"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    INSERT INTO holiday_entries (calendar_id, holiday_date, name, is_full_day, start_time, end_time)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (calendar_id, holiday_date, name, is_full_day, start_time, end_time),
                )
                conn.commit()
                return cursor.lastrowid
        except sqlite3.Error as e:
            logger.error(f"新增 holiday entry 失敗: {e}")
            return None

    def get_holiday_entries_by_calendar(self, calendar_id: int) -> List[Dict[str, Any]]:
        """查詢指定日曆的所有假日條目"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT * FROM holiday_entries WHERE calendar_id = ? ORDER BY holiday_date",
                    (calendar_id,),
                )
                return [dict(row) for row in cursor.fetchall()]
        except sqlite3.Error as e:
            logger.error(f"查詢 holiday entries 失敗: {e}")
            return []

    def get_all_holiday_entries(self) -> List[Dict[str, Any]]:
        """查詢所有假日條目"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM holiday_entries ORDER BY holiday_date DESC")
                return [dict(row) for row in cursor.fetchall()]
        except sqlite3.Error as e:
            logger.error(f"查詢所有 holiday entries 失敗: {e}")
            return []

    def update_holiday_entry(
        self,
        entry_id: int,
        holiday_date: str,
        name: str,
        is_full_day: int = 1,
        start_time: str = None,
        end_time: str = None,
    ) -> bool:
        """更新假日條目"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    UPDATE holiday_entries
                    SET holiday_date = ?, name = ?, is_full_day = ?, start_time = ?, end_time = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                    """,
                    (holiday_date, name, is_full_day, start_time, end_time, entry_id),
                )
                conn.commit()
                return cursor.rowcount > 0
        except sqlite3.Error as e:
            logger.error(f"更新 holiday entry 失敗: {e}")
            return False

    def delete_holiday_entry(self, entry_id: int) -> bool:
        """刪除假日條目"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM holiday_entries WHERE id = ?", (entry_id,))
                conn.commit()
                return cursor.rowcount > 0
        except sqlite3.Error as e:
            logger.error(f"刪除 holiday entry 失敗: {e}")
            return False

    # ==================== General Settings ====================

    def get_general_settings(self) -> Dict[str, Any]:
        """查詢全局設定（只有一筆記錄）"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM general_settings LIMIT 1")
                row = cursor.fetchone()
                if row:
                    return dict(row)
                else:
                    # 首次使用，創建預設設定
                    self._create_default_settings()
                    cursor.execute("SELECT * FROM general_settings LIMIT 1")
                    row = cursor.fetchone()
                    return dict(row) if row else {}
        except sqlite3.Error as e:
            logger.error(f"查詢 general settings 失敗: {e}")
            return {}

    def _create_default_settings(self):
        """創建預設全局設定"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    INSERT INTO general_settings (profile_name, description, enable_schedule, scan_rate, refresh_rate)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    ("預設 Profile", "CalendarUA 排程系統", 1, 1, 5),
                )
                conn.commit()
        except sqlite3.Error as e:
            logger.error(f"創建預設設定失敗: {e}")

    def save_general_settings(self, settings: Dict[str, Any]) -> bool:
        """儲存全局設定"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                
                # 檢查是否已有記錄
                cursor.execute("SELECT id FROM general_settings LIMIT 1")
                existing = cursor.fetchone()
                
                if existing:
                    # 更新現有記錄
                    cursor.execute(
                        """
                        UPDATE general_settings
                        SET profile_name = ?, description = ?, enable_schedule = ?, 
                            scan_rate = ?, refresh_rate = ?, use_active_period = ?,
                            active_from = ?, active_to = ?, output_type = ?,
                            refresh_output = ?, generate_events = ?, updated_at = CURRENT_TIMESTAMP
                        WHERE id = ?
                        """,
                        (
                            settings.get("profile_name", "預設 Profile"),
                            settings.get("description", ""),
                            settings.get("enable_schedule", 1),
                            settings.get("scan_rate", 1),
                            settings.get("refresh_rate", 5),
                            settings.get("use_active_period", 0),
                            settings.get("active_from"),
                            settings.get("active_to"),
                            settings.get("output_type", "OPC UA Write"),
                            settings.get("refresh_output", 1),
                            settings.get("generate_events", 1),
                            existing["id"],
                        ),
                    )
                else:
                    # 插入新記錄
                    cursor.execute(
                        """
                        INSERT INTO general_settings 
                        (profile_name, description, enable_schedule, scan_rate, refresh_rate,
                         use_active_period, active_from, active_to, output_type, refresh_output, generate_events)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            settings.get("profile_name", "預設 Profile"),
                            settings.get("description", ""),
                            settings.get("enable_schedule", 1),
                            settings.get("scan_rate", 1),
                            settings.get("refresh_rate", 5),
                            settings.get("use_active_period", 0),
                            settings.get("active_from"),
                            settings.get("active_to"),
                            settings.get("output_type", "OPC UA Write"),
                            settings.get("refresh_output", 1),
                            settings.get("generate_events", 1),
                        ),
                    )
                
                conn.commit()
                return True
        except sqlite3.Error as e:
            logger.error(f"儲存 general settings 失敗: {e}")
            return False

    # ==================== Runtime Override ====================

    def get_runtime_override(self) -> Optional[Dict[str, Any]]:
        """查詢 runtime override（只有一筆有效記錄）"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM runtime_override LIMIT 1")
                row = cursor.fetchone()
                return dict(row) if row else None
        except sqlite3.Error as e:
            logger.error(f"查詢 runtime override 失敗: {e}")
            return None

    def set_runtime_override(self, override_value: str, override_until: Optional[str] = None) -> bool:
        """設定 runtime override"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                
                # 清除舊記錄
                cursor.execute("DELETE FROM runtime_override")
                
                # 插入新記錄
                cursor.execute(
                    """
                    INSERT INTO runtime_override (override_value, override_until)
                    VALUES (?, ?)
                    """,
                    (override_value, override_until),
                )
                
                conn.commit()
                return True
        except sqlite3.Error as e:
            logger.error(f"設定 runtime override 失敗: {e}")
            return False

    def clear_runtime_override(self) -> bool:
        """清除 runtime override"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM runtime_override")
                conn.commit()
                return True
        except sqlite3.Error as e:
            logger.error(f"清除 runtime override 失敗: {e}")
            return False




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
from datetime import datetime

from core.lunar_calendar import to_lunar


# 設定日誌記錄
logging.basicConfig(
    level=logging.WARNING, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


DEFAULT_SOLAR_HOLIDAYS = [(1, 1), (2, 28), (4, 4), (4, 5), (5, 1), (10, 10)]
DEFAULT_LUNAR_HOLIDAYS = [(12, 31), (1, 1), (1, 2), (1, 3), (1, 4), (1, 5), (5, 5), (8, 15)]
DEFAULT_WEEKDAY_HOLIDAYS = [6, 7]


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

    def _create_holidays_table(self, cursor: sqlite3.Cursor) -> None:
        """建立單一假日設定表（整合週別與國/農曆日期）。"""
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS holidays (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entry_type TEXT NOT NULL,
                calendar_type TEXT,
                month INTEGER,
                day INTEGER,
                weekday INTEGER,
                name TEXT DEFAULT '',
                override_target_value TEXT,
                is_enabled INTEGER DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        cursor.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_holidays_unique_weekday ON holidays(entry_type, weekday) WHERE entry_type = 'weekday'"
        )
        cursor.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_holidays_unique_date ON holidays(entry_type, calendar_type, month, day) WHERE entry_type = 'date'"
        )
        cursor.execute(
            "CREATE INDEX IF NOT EXISTS idx_holidays_enabled ON holidays(is_enabled)"
        )

    def _migrate_holiday_data(self, cursor: sqlite3.Cursor) -> None:
        """將舊版 holiday_entries 資料（YYYY-MM-DD）遷移到新 holidays 單表。"""
        cursor.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='holiday_entries'"
        )
        if not cursor.fetchone():
            return

        cursor.execute("SELECT COUNT(1) AS cnt FROM holidays")
        row = cursor.fetchone()
        if row and int(row["cnt"] or 0) > 0:
            return

        cursor.execute(
            "SELECT holiday_date, name, override_target_value FROM holiday_entries ORDER BY id"
        )
        legacy_rows = cursor.fetchall()
        for legacy in legacy_rows:
            holiday_date = str(legacy["holiday_date"] or "").strip()
            try:
                dt = datetime.strptime(holiday_date, "%Y-%m-%d")
            except ValueError:
                continue

            month = dt.month
            day = dt.day
            name = str(legacy["name"] or "")
            override_target_value = legacy["override_target_value"]

            cursor.execute(
                """
                INSERT OR IGNORE INTO holidays
                (entry_type, calendar_type, month, day, name, override_target_value, is_enabled)
                VALUES ('date', 'solar', ?, ?, ?, ?, 1)
                """,
                (month, day, name, override_target_value),
            )

    def _ensure_default_holiday_rules(self, cursor: sqlite3.Cursor) -> None:
        """若假日規則為空，補上預設週六/週日與國定/農曆日期。"""
        cursor.execute("SELECT COUNT(1) AS cnt FROM holidays WHERE is_enabled = 1")
        row = cursor.fetchone()
        if row and int(row["cnt"] or 0) > 0:
            return

        for weekday in DEFAULT_WEEKDAY_HOLIDAYS:
            cursor.execute(
                """
                INSERT OR IGNORE INTO holidays
                (entry_type, weekday, name, is_enabled)
                VALUES ('weekday', ?, ?, 1)
                """,
                (weekday, f"週{weekday}假日"),
            )

        for month, day in DEFAULT_SOLAR_HOLIDAYS:
            cursor.execute(
                """
                INSERT OR IGNORE INTO holidays
                (entry_type, calendar_type, month, day, name, is_enabled)
                VALUES ('date', 'solar', ?, ?, '', 1)
                """,
                (month, day),
            )

        for month, day in DEFAULT_LUNAR_HOLIDAYS:
            cursor.execute(
                """
                INSERT OR IGNORE INTO holidays
                (entry_type, calendar_type, month, day, name, is_enabled)
                VALUES ('date', 'lunar', ?, ?, '', 1)
                """,
                (month, day),
            )

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
            lock_enabled INTEGER DEFAULT 0,
            is_enabled INTEGER DEFAULT 1,
            ignore_holiday INTEGER DEFAULT 0,
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

                # 假日統一採單一 holidays 表
                self._create_holidays_table(cursor)

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
                        last_opc_url TEXT DEFAULT '',
                        last_opc_security_policy TEXT DEFAULT 'None',
                        last_opc_security_mode TEXT DEFAULT 'None',
                        last_opc_username TEXT DEFAULT '',
                        last_opc_password TEXT DEFAULT '',
                        last_opc_timeout INTEGER DEFAULT 5,
                        last_opc_write_timeout INTEGER DEFAULT 3,
                        time_scale_minutes INTEGER DEFAULT 60,
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

                # 檢查並添加缺少的欄位
                cursor.execute("PRAGMA table_info(schedules)")
                columns = [column[1] for column in cursor.fetchall()]
                
                if 'last_execution_status' not in columns:
                    cursor.execute("ALTER TABLE schedules ADD COLUMN last_execution_status TEXT DEFAULT ''")
                    logger.info("已添加 last_execution_status 欄位")
                
                if 'last_execution_time' not in columns:
                    cursor.execute("ALTER TABLE schedules ADD COLUMN last_execution_time TIMESTAMP")
                    logger.info("已添加 last_execution_time 欄位")

                if 'ignore_holiday' not in columns:
                    cursor.execute("ALTER TABLE schedules ADD COLUMN ignore_holiday INTEGER DEFAULT 0")
                    logger.info("已添加 ignore_holiday 欄位")

                if 'lock_enabled' not in columns:
                    cursor.execute("ALTER TABLE schedules ADD COLUMN lock_enabled INTEGER DEFAULT 0")
                    logger.info("已添加 lock_enabled 欄位")

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

                # 假日資料遷移與預設資料
                self._migrate_holiday_data(cursor)
                self._ensure_default_holiday_rules(cursor)

                conn.commit()
                logger.info("資料庫初始化成功，schedules 表格已建立")
                
                # 執行遷移以添加缺失的欄位
                self._migrate_db()
                
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

                if "ignore_holiday" not in columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN ignore_holiday INTEGER DEFAULT 0"
                    )
                    logger.info("已添加 ignore_holiday 欄位")

                if "lock_enabled" not in columns:
                    cursor.execute(
                        "ALTER TABLE schedules ADD COLUMN lock_enabled INTEGER DEFAULT 0"
                    )
                    logger.info("已添加 lock_enabled 欄位")

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

                # 假日統一採單一 holidays 表
                self._create_holidays_table(cursor)
                
                cursor.execute("PRAGMA table_info(schedules)")
                schedules_columns = [column[1] for column in cursor.fetchall()]

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

                cursor.execute("PRAGMA table_info(schedule_exceptions)")
                exceptions_columns = [column[1] for column in cursor.fetchall()]
                if "note" not in exceptions_columns:
                    cursor.execute(
                        "ALTER TABLE schedule_exceptions ADD COLUMN note TEXT DEFAULT ''"
                    )
                    logger.info("已添加 schedule_exceptions.note 欄位")

                cursor.execute("PRAGMA table_info(holidays)")
                holiday_columns = [column[1] for column in cursor.fetchall()]
                if "override_target_value" not in holiday_columns:
                    cursor.execute("ALTER TABLE holidays ADD COLUMN override_target_value TEXT")
                    logger.info("已添加 holidays.override_target_value 欄位")
                if "is_enabled" not in holiday_columns:
                    cursor.execute("ALTER TABLE holidays ADD COLUMN is_enabled INTEGER DEFAULT 1")
                    logger.info("已添加 holidays.is_enabled 欄位")

                self._migrate_holiday_data(cursor)
                self._ensure_default_holiday_rules(cursor)

                cursor.execute("PRAGMA table_info(general_settings)")
                general_columns = [column[1] for column in cursor.fetchall()]
                if "last_opc_url" not in general_columns:
                    cursor.execute("ALTER TABLE general_settings ADD COLUMN last_opc_url TEXT DEFAULT ''")
                    logger.info("已添加 general_settings.last_opc_url 欄位")
                if "last_opc_security_policy" not in general_columns:
                    cursor.execute("ALTER TABLE general_settings ADD COLUMN last_opc_security_policy TEXT DEFAULT 'None'")
                    logger.info("已添加 general_settings.last_opc_security_policy 欄位")
                if "last_opc_security_mode" not in general_columns:
                    cursor.execute("ALTER TABLE general_settings ADD COLUMN last_opc_security_mode TEXT DEFAULT 'None'")
                    logger.info("已添加 general_settings.last_opc_security_mode 欄位")
                if "last_opc_username" not in general_columns:
                    cursor.execute("ALTER TABLE general_settings ADD COLUMN last_opc_username TEXT DEFAULT ''")
                    logger.info("已添加 general_settings.last_opc_username 欄位")
                if "last_opc_password" not in general_columns:
                    cursor.execute("ALTER TABLE general_settings ADD COLUMN last_opc_password TEXT DEFAULT ''")
                    logger.info("已添加 general_settings.last_opc_password 欄位")
                if "last_opc_timeout" not in general_columns:
                    cursor.execute("ALTER TABLE general_settings ADD COLUMN last_opc_timeout INTEGER DEFAULT 5")
                    logger.info("已添加 general_settings.last_opc_timeout 欄位")
                if "last_opc_write_timeout" not in general_columns:
                    cursor.execute("ALTER TABLE general_settings ADD COLUMN last_opc_write_timeout INTEGER DEFAULT 3")
                    logger.info("已添加 general_settings.last_opc_write_timeout 欄位")
                if "time_scale_minutes" not in general_columns:
                    cursor.execute("ALTER TABLE general_settings ADD COLUMN time_scale_minutes INTEGER DEFAULT 60")
                    logger.info("已添加 general_settings.time_scale_minutes 欄位")

                # 依照下一次執行時間優化查詢效能（排程掃描常用）
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_schedules_next_time ON schedules(next_execution_time)"
                )
                
                conn.commit()
                
        except sqlite3.Error as e:
            logger.error(f"資料庫遷移失敗: {e}")

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
        lock_enabled: int = 0,
        is_enabled: int = 1,
        ignore_holiday: int = 0,
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
            opc_security_policy: OPC安全策略
            opc_security_mode: OPC安全模式 (None/Sign/SignAndEncrypt)
            opc_username: OPC使用者名稱
            opc_password: OPC密碼
            opc_timeout: 連線超時秒數
            opc_write_timeout: 寫值重試延遲秒數
            lock_enabled: 是否啟用鎖定模式 (1: 鎖定, 0: 不鎖定)
            is_enabled: 是否啟用 (1: 啟用, 0: 停用)，預設為 1
            ignore_holiday: 是否忽略假日規則 (1: 忽略, 0: 不忽略)

        Returns:
            Optional[int]: 新增排程的 ID，失敗時回傳 None
        """
        insert_sql = """
        INSERT INTO schedules (task_name, opc_url, node_id, target_value, data_type, rrule_str,
                              opc_security_policy, opc_security_mode, opc_username, opc_password, opc_timeout, opc_write_timeout, lock_enabled, is_enabled, ignore_holiday)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    insert_sql,
                    (task_name, opc_url, node_id, target_value, data_type, rrule_str,
                     opc_security_policy, opc_security_mode, opc_username, opc_password, opc_timeout, opc_write_timeout, lock_enabled, is_enabled, ignore_holiday),
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
        lock_enabled: int = 0,
        is_enabled: int = 1,
        ignore_holiday: int = 0,
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
            lock_enabled: 是否啟用鎖定模式 (1: 鎖定, 0: 不鎖定)
            is_enabled: 是否啟用 (1: 啟用, 0: 停用)，預設為 1
            ignore_holiday: 是否忽略假日規則 (1: 忽略, 0: 不忽略)

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
            lock_enabled=lock_enabled,
            is_enabled=is_enabled,
            ignore_holiday=ignore_holiday,
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
            **kwargs: 要更新的欄位（task_name, opc_url, node_id, target_value, rrule_str, is_enabled）

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
            "opc_security_policy",
            "opc_security_mode",
            "opc_username",
            "opc_password",
            "opc_timeout",
            "opc_write_timeout",
            "lock_enabled",
            "is_enabled",
            "ignore_holiday",
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
        """相容舊介面：單表模式不再使用 calendar，回傳固定 ID。"""
        return 1

    def get_all_holiday_calendars(self) -> List[Dict[str, Any]]:
        """相容舊介面：回傳單一內建假日日曆。"""
        return [{"id": 1, "name": "預設假日", "description": "單表模式", "is_default": 1}]

    def update_holiday_calendar(self, calendar_id: int, name: str, description: str = "", is_default: int = 0) -> bool:
        """相容舊介面：單表模式無此操作。"""
        return True

    def delete_holiday_calendar(self, calendar_id: int) -> bool:
        """相容舊介面：單表模式無此操作。"""
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
        """新增假日條目（舊介面：接受 YYYY-MM-DD，轉存成國曆月/日規則）。"""
        try:
            dt = datetime.strptime(holiday_date, "%Y-%m-%d")
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    INSERT INTO holidays (entry_type, calendar_type, month, day, name, is_enabled)
                    VALUES ('date', 'solar', ?, ?, ?, 1)
                    """,
                    (dt.month, dt.day, name),
                )
                conn.commit()
                return cursor.lastrowid
        except (ValueError, sqlite3.Error) as e:
            logger.error(f"新增 holiday entry 失敗: {e}")
            return None

    def get_holiday_entries_by_calendar(self, calendar_id: int) -> List[Dict[str, Any]]:
        """相容舊介面：單表模式忽略 calendar_id，回傳全部規則。"""
        return self.get_all_holiday_entries()

    def get_all_holiday_entries(self) -> List[Dict[str, Any]]:
        """查詢所有假日規則（單表）。"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    SELECT *
                    FROM holidays
                    WHERE is_enabled = 1
                    ORDER BY CASE entry_type WHEN 'weekday' THEN 0 ELSE 1 END, calendar_type, month, day, weekday
                    """
                )
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
        """更新假日條目（舊介面：更新為國曆月/日規則）。"""
        try:
            dt = datetime.strptime(holiday_date, "%Y-%m-%d")
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    UPDATE holidays
                    SET entry_type = 'date', calendar_type = 'solar', month = ?, day = ?, name = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                    """,
                    (dt.month, dt.day, name, entry_id),
                )
                conn.commit()
                return cursor.rowcount > 0
        except (ValueError, sqlite3.Error) as e:
            logger.error(f"更新 holiday entry 失敗: {e}")
            return False

    def delete_holiday_entry(self, entry_id: int) -> bool:
        """刪除假日條目"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM holidays WHERE id = ?", (entry_id,))
                conn.commit()
                return cursor.rowcount > 0
        except sqlite3.Error as e:
            logger.error(f"刪除 holiday entry 失敗: {e}")
            return False

    def set_weekday_holidays(self, weekdays: List[int]) -> bool:
        """設定週幾為假日（1=週一 ... 7=週日）。"""
        normalized = sorted({int(w) for w in weekdays if 1 <= int(w) <= 7})
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM holidays WHERE entry_type = 'weekday'")
                for weekday in normalized:
                    cursor.execute(
                        """
                        INSERT INTO holidays (entry_type, weekday, name, is_enabled)
                        VALUES ('weekday', ?, ?, 1)
                        """,
                        (weekday, f"週{weekday}假日"),
                    )
                conn.commit()
                return True
        except sqlite3.Error as e:
            logger.error(f"設定 weekday 假日失敗: {e}")
            return False

    def add_holiday_rule(self, calendar_type: str, month: int, day: int, name: str = "") -> Optional[int]:
        """新增日期型假日規則（calendar_type: solar/lunar）。"""
        if calendar_type not in {"solar", "lunar"}:
            return None
        if month < 1 or month > 12 or day < 1 or day > 31:
            return None
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    INSERT INTO holidays (entry_type, calendar_type, month, day, name, is_enabled)
                    VALUES ('date', ?, ?, ?, ?, 1)
                    """,
                    (calendar_type, month, day, name),
                )
                conn.commit()
                return cursor.lastrowid
        except sqlite3.IntegrityError:
            return None
        except sqlite3.Error as e:
            logger.error(f"新增 holiday rule 失敗: {e}")
            return None

    def update_holiday_rule(self, rule_id: int, calendar_type: str, month: int, day: int, name: str = "") -> bool:
        """更新日期型假日規則。"""
        if calendar_type not in {"solar", "lunar"}:
            return False
        if month < 1 or month > 12 or day < 1 or day > 31:
            return False
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    UPDATE holidays
                    SET entry_type = 'date', calendar_type = ?, month = ?, day = ?, name = ?, is_enabled = 1, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                    """,
                    (calendar_type, month, day, name, rule_id),
                )
                conn.commit()
                return cursor.rowcount > 0
        except sqlite3.IntegrityError:
            return False
        except sqlite3.Error as e:
            logger.error(f"更新 holiday rule 失敗: {e}")
            return False

    def replace_holiday_rules(self, weekdays: List[int], date_rules: List[Dict[str, Any]]) -> bool:
        """以整包資料覆蓋假日規則。"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM holidays")

                normalized = sorted({int(w) for w in weekdays if 1 <= int(w) <= 7})
                for weekday in normalized:
                    cursor.execute(
                        "INSERT INTO holidays (entry_type, weekday, name, is_enabled) VALUES ('weekday', ?, ?, 1)",
                        (weekday, f"週{weekday}假日"),
                    )

                for rule in date_rules:
                    calendar_type = str(rule.get("calendar_type", "")).strip().lower()
                    month = int(rule.get("month", 0) or 0)
                    day = int(rule.get("day", 0) or 0)
                    name = str(rule.get("name", "") or "")
                    if calendar_type not in {"solar", "lunar"}:
                        continue
                    if month < 1 or month > 12 or day < 1 or day > 31:
                        continue
                    cursor.execute(
                        """
                        INSERT OR IGNORE INTO holidays (entry_type, calendar_type, month, day, name, is_enabled)
                        VALUES ('date', ?, ?, ?, ?, 1)
                        """,
                        (calendar_type, month, day, name),
                    )

                self._ensure_default_holiday_rules(cursor)
                conn.commit()
                return True
        except sqlite3.Error as e:
            logger.error(f"覆蓋 holiday rules 失敗: {e}")
            return False

    def get_holiday_rules_payload(self) -> Dict[str, Any]:
        """輸出匯入/匯出的假日設定資料。"""
        entries = self.get_all_holiday_entries()
        weekdays: List[int] = []
        dates: List[Dict[str, Any]] = []
        for entry in entries:
            if entry.get("entry_type") == "weekday":
                weekday = int(entry.get("weekday", 0) or 0)
                if 1 <= weekday <= 7:
                    weekdays.append(weekday)
            elif entry.get("entry_type") == "date":
                calendar_type = str(entry.get("calendar_type", "") or "").strip().lower()
                month = int(entry.get("month", 0) or 0)
                day = int(entry.get("day", 0) or 0)
                if calendar_type in {"solar", "lunar"} and 1 <= month <= 12 and 1 <= day <= 31:
                    dates.append(
                        {
                            "id": entry.get("id"),
                            "calendar_type": calendar_type,
                            "month": month,
                            "day": day,
                            "name": str(entry.get("name", "") or ""),
                        }
                    )
        return {
            "weekdays": sorted(set(weekdays)),
            "dates": dates,
        }

    def is_holiday_on_date(self, date_obj) -> Optional[Dict[str, Any]]:
        """判斷指定日期是否命中任一假日規則，命中時回傳規則。"""
        try:
            rules = self.get_all_holiday_entries()
            weekday = date_obj.isoweekday()
            for rule in rules:
                if rule.get("entry_type") == "weekday" and int(rule.get("weekday", 0) or 0) == weekday:
                    return rule

            for rule in rules:
                if rule.get("entry_type") != "date":
                    continue
                month = int(rule.get("month", 0) or 0)
                day = int(rule.get("day", 0) or 0)
                calendar_type = str(rule.get("calendar_type", "") or "").strip().lower()
                if calendar_type == "solar" and date_obj.month == month and date_obj.day == day:
                    return rule
                if calendar_type == "lunar":
                    lunar_info = to_lunar(date_obj)
                    if lunar_info and lunar_info.lunar_month == month and lunar_info.lunar_day == day:
                        return rule
            return None
        except Exception:
            return None

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

    def get_last_opc_defaults(self) -> Dict[str, Any]:
        """取得上一次使用的 OPC 設定（給新增排程預設帶入）。"""
        settings = self.get_general_settings() or {}
        return {
            "opc_url": settings.get("last_opc_url", "") or "",
            "opc_security_policy": settings.get("last_opc_security_policy", "None") or "None",
            "opc_security_mode": settings.get("last_opc_security_mode", "None") or "None",
            "opc_username": settings.get("last_opc_username", "") or "",
            "opc_password": settings.get("last_opc_password", "") or "",
            "opc_timeout": int(settings.get("last_opc_timeout", 5) or 5),
            "opc_write_timeout": int(settings.get("last_opc_write_timeout", 3) or 3),
        }

    def save_last_opc_defaults(self, defaults: Dict[str, Any]) -> bool:
        """儲存上一次使用的 OPC 設定。"""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM general_settings LIMIT 1")
                existing = cursor.fetchone()

                if existing:
                    cursor.execute(
                        """
                        UPDATE general_settings
                        SET last_opc_url = ?,
                            last_opc_security_policy = ?,
                            last_opc_security_mode = ?,
                            last_opc_username = ?,
                            last_opc_password = ?,
                            last_opc_timeout = ?,
                            last_opc_write_timeout = ?,
                            updated_at = CURRENT_TIMESTAMP
                        WHERE id = ?
                        """,
                        (
                            defaults.get("opc_url", ""),
                            defaults.get("opc_security_policy", "None"),
                            defaults.get("opc_security_mode", "None"),
                            defaults.get("opc_username", ""),
                            defaults.get("opc_password", ""),
                            int(defaults.get("opc_timeout", 5) or 5),
                            int(defaults.get("opc_write_timeout", 3) or 3),
                            existing["id"],
                        ),
                    )
                else:
                    cursor.execute(
                        """
                        INSERT INTO general_settings (
                            profile_name, description, enable_schedule, scan_rate, refresh_rate,
                            last_opc_url, last_opc_security_policy, last_opc_security_mode,
                            last_opc_username, last_opc_password, last_opc_timeout, last_opc_write_timeout
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            "預設 Profile",
                            "CalendarUA 排程系統",
                            1,
                            1,
                            5,
                            defaults.get("opc_url", ""),
                            defaults.get("opc_security_policy", "None"),
                            defaults.get("opc_security_mode", "None"),
                            defaults.get("opc_username", ""),
                            defaults.get("opc_password", ""),
                            int(defaults.get("opc_timeout", 5) or 5),
                            int(defaults.get("opc_write_timeout", 3) or 3),
                        ),
                    )

                conn.commit()
                return True
        except sqlite3.Error as e:
            logger.error(f"儲存最後 OPC 設定失敗: {e}")
            return False

    def get_time_scale_minutes(self) -> int:
        """取得日/週視圖時間刻度（分鐘）。"""
        settings = self.get_general_settings() or {}
        value = settings.get("time_scale_minutes", 60)
        try:
            minutes = int(value)
        except (TypeError, ValueError):
            minutes = 60

        return minutes if minutes in {5, 6, 10, 15, 30, 60} else 60

    def save_time_scale_minutes(self, minutes: int) -> bool:
        """儲存日/週視圖時間刻度（分鐘）。"""
        if minutes not in {5, 6, 10, 15, 30, 60}:
            return False

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM general_settings LIMIT 1")
                existing = cursor.fetchone()

                if existing:
                    cursor.execute(
                        """
                        UPDATE general_settings
                        SET time_scale_minutes = ?,
                            updated_at = CURRENT_TIMESTAMP
                        WHERE id = ?
                        """,
                        (minutes, existing["id"]),
                    )
                else:
                    cursor.execute(
                        """
                        INSERT INTO general_settings (
                            profile_name, description, enable_schedule, scan_rate, refresh_rate, time_scale_minutes
                        ) VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        ("預設 Profile", "CalendarUA 排程系統", 1, 1, 5, minutes),
                    )

                conn.commit()
                return True
        except sqlite3.Error as e:
            logger.error(f"儲存 time scale 失敗: {e}")
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




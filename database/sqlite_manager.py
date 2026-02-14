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
                        "ALTER TABLE schedules ADD COLUMN opc_timeout INTEGER DEFAULT 10"
                    )
                    logger.info("已添加 opc_timeout 欄位")
                
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
        opc_security_policy: str = "None",
        opc_security_mode: str = "None",
        opc_username: str = "",
        opc_password: str = "",
        opc_timeout: int = 10,
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
            opc_security_policy: OPC安全策略
            opc_security_mode: OPC安全模式 (None/Sign/SignAndEncrypt)
            opc_username: OPC使用者名稱
            opc_password: OPC密碼
            opc_timeout: 連線超時秒數
            is_enabled: 是否啟用 (1: 啟用, 0: 停用)，預設為 1

        Returns:
            Optional[int]: 新增排程的 ID，失敗時回傳 None
        """
        insert_sql = """
        INSERT INTO schedules (task_name, opc_url, node_id, target_value, data_type, rrule_str,
                              opc_security_policy, opc_security_mode, opc_username, opc_password, opc_timeout, is_enabled)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    insert_sql,
                    (task_name, opc_url, node_id, target_value, data_type, rrule_str,
                     opc_security_policy, opc_security_mode, opc_username, opc_password, opc_timeout, is_enabled),
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
        opc_timeout: int = 10,
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

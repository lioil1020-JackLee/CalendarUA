import mysql.connector
from mysql.connector import pooling
from typing import List, Dict, Optional, Any
from contextlib import contextmanager
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class MySQLManager:
    """MySQL 資料庫管理類別，負責排程資料的 CRUD 操作"""

    def __init__(
        self, host: str, user: str, password: str, database: str, port: int = 3306
    ):
        """
        初始化 MySQL 連線池

        Args:
            host: MySQL 伺服器位址
            user: 使用者名稱
            password: 密碼
            database: 資料庫名稱
            port: 連接埠 (預設 3306)
        """
        self.config = {
            "host": host,
            "user": user,
            "password": password,
            "database": database,
            "port": port,
            "charset": "utf8mb4",
            "collation": "utf8mb4_unicode_ci",
        }
        self.pool = None
        self._create_pool()

    def _create_pool(self):
        """建立連線池"""
        try:
            self.pool = pooling.MySQLConnectionPool(
                pool_name="calendarua_pool", pool_size=5, **self.config
            )
            logger.info("MySQL 連線池建立成功")
        except mysql.connector.Error as err:
            logger.error(f"建立連線池失敗: {err}")
            raise

    @contextmanager
    def _get_connection(self):
        """取得連線的上下文管理器"""
        conn = None
        cursor = None
        try:
            conn = self.pool.get_connection()
            cursor = conn.cursor(dictionary=True)
            yield cursor
            conn.commit()
        except mysql.connector.Error as err:
            if conn:
                conn.rollback()
            logger.error(f"資料庫操作錯誤: {err}")
            raise
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def create_table(self) -> bool:
        """
        建立 schedules 資料表

        Returns:
            bool: 建立成功回傳 True，否則回傳 False
        """
        create_table_sql = """
        CREATE TABLE IF NOT EXISTS schedules (
            id INT AUTO_INCREMENT PRIMARY KEY,
            task_name VARCHAR(100) NOT NULL COMMENT '任務名稱',
            opc_url VARCHAR(255) NOT NULL COMMENT 'OPC UA 伺服器位址',
            node_id VARCHAR(255) NOT NULL COMMENT 'OPC UA Tag NodeID',
            target_value VARCHAR(50) NOT NULL COMMENT '要寫入的數值',
            rrule_str VARCHAR(500) NOT NULL COMMENT 'RRULE 規則字串',
            is_enabled TINYINT(1) DEFAULT 1 COMMENT '是否啟用 (1: 啟用, 0: 停用)',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            INDEX idx_enabled (is_enabled),
            INDEX idx_node_id (node_id)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """

        try:
            with self._get_connection() as cursor:
                cursor.execute(create_table_sql)
                logger.info("schedules 資料表建立成功或已存在")
                return True
        except mysql.connector.Error as err:
            logger.error(f"建立資料表失敗: {err}")
            return False

    def create_schedule(
        self,
        task_name: str,
        opc_url: str,
        node_id: str,
        target_value: str,
        rrule_str: str,
        is_enabled: int = 1,
    ) -> Optional[int]:
        """
        新增排程

        Args:
            task_name: 任務名稱
            opc_url: OPC UA 伺服器位址
            node_id: OPC UA Tag NodeID
            target_value: 要寫入的數值
            rrule_str: RRULE 規則字串
            is_enabled: 是否啟用 (1: 啟用, 0: 停用)

        Returns:
            Optional[int]: 新排程的 ID，失敗回傳 None
        """
        insert_sql = """
        INSERT INTO schedules (task_name, opc_url, node_id, target_value, rrule_str, is_enabled)
        VALUES (%s, %s, %s, %s, %s, %s)
        """

        try:
            with self._get_connection() as cursor:
                cursor.execute(
                    insert_sql,
                    (task_name, opc_url, node_id, target_value, rrule_str, is_enabled),
                )
                new_id = cursor.lastrowid
                logger.info(f"排程新增成功，ID: {new_id}")
                return new_id
        except mysql.connector.Error as err:
            logger.error(f"新增排程失敗: {err}")
            return None

    def get_schedule(self, schedule_id: int) -> Optional[Dict[str, Any]]:
        """
        查詢單一排程

        Args:
            schedule_id: 排程 ID

        Returns:
            Optional[Dict]: 排程資料字典，找不到回傳 None
        """
        select_sql = "SELECT * FROM schedules WHERE id = %s"

        try:
            with self._get_connection() as cursor:
                cursor.execute(select_sql, (schedule_id,))
                result = cursor.fetchone()
                return result
        except mysql.connector.Error as err:
            logger.error(f"查詢排程失敗: {err}")
            return None

    def get_all_schedules(self, enabled_only: bool = False) -> List[Dict[str, Any]]:
        """
        查詢所有排程

        Args:
            enabled_only: 是否只查詢啟用的排程

        Returns:
            List[Dict]: 排程資料列表
        """
        if enabled_only:
            select_sql = (
                "SELECT * FROM schedules WHERE is_enabled = 1 ORDER BY created_at DESC"
            )
        else:
            select_sql = "SELECT * FROM schedules ORDER BY created_at DESC"

        try:
            with self._get_connection() as cursor:
                cursor.execute(select_sql)
                results = cursor.fetchall()
                return results
        except mysql.connector.Error as err:
            logger.error(f"查詢所有排程失敗: {err}")
            return []

    def update_schedule(self, schedule_id: int, **kwargs) -> bool:
        """
        更新排程

        Args:
            schedule_id: 排程 ID
            **kwargs: 要更新的欄位 (task_name, opc_url, node_id, target_value, rrule_str, is_enabled)

        Returns:
            bool: 更新成功回傳 True，否則回傳 False
        """
        allowed_fields = {
            "task_name",
            "opc_url",
            "node_id",
            "target_value",
            "rrule_str",
            "is_enabled",
        }
        updates = {k: v for k, v in kwargs.items() if k in allowed_fields}

        if not updates:
            logger.warning("沒有提供要更新的欄位")
            return False

        set_clause = ", ".join([f"{k} = %s" for k in updates.keys()])
        update_sql = f"UPDATE schedules SET {set_clause} WHERE id = %s"
        values = list(updates.values()) + [schedule_id]

        try:
            with self._get_connection() as cursor:
                cursor.execute(update_sql, values)
                if cursor.rowcount > 0:
                    logger.info(f"排程 {schedule_id} 更新成功")
                    return True
                else:
                    logger.warning(f"排程 {schedule_id} 不存在或無變更")
                    return False
        except mysql.connector.Error as err:
            logger.error(f"更新排程失敗: {err}")
            return False

    def delete_schedule(self, schedule_id: int) -> bool:
        """
        刪除排程

        Args:
            schedule_id: 排程 ID

        Returns:
            bool: 刪除成功回傳 True，否則回傳 False
        """
        delete_sql = "DELETE FROM schedules WHERE id = %s"

        try:
            with self._get_connection() as cursor:
                cursor.execute(delete_sql, (schedule_id,))
                if cursor.rowcount > 0:
                    logger.info(f"排程 {schedule_id} 刪除成功")
                    return True
                else:
                    logger.warning(f"排程 {schedule_id} 不存在")
                    return False
        except mysql.connector.Error as err:
            logger.error(f"刪除排程失敗: {err}")
            return False

    def toggle_schedule(self, schedule_id: int, is_enabled: int) -> bool:
        """
        啟用/停用排程

        Args:
            schedule_id: 排程 ID
            is_enabled: 1 啟用, 0 停用

        Returns:
            bool: 操作成功回傳 True，否則回傳 False
        """
        return self.update_schedule(schedule_id, is_enabled=is_enabled)

    def get_schedules_by_node(self, node_id: str) -> List[Dict[str, Any]]:
        """
        根據 Node ID 查詢排程

        Args:
            node_id: OPC UA Node ID

        Returns:
            List[Dict]: 排程資料列表
        """
        select_sql = (
            "SELECT * FROM schedules WHERE node_id = %s ORDER BY created_at DESC"
        )

        try:
            with self._get_connection() as cursor:
                cursor.execute(select_sql, (node_id,))
                results = cursor.fetchall()
                return results
        except mysql.connector.Error as err:
            logger.error(f"根據 Node ID 查詢排程失敗: {err}")
            return []

    def close(self):
        """關閉連線池"""
        if self.pool:
            logger.info("MySQL 連線池已關閉")


# 使用範例
if __name__ == "__main__":
    # 初始化資料庫管理器
    db = MySQLManager(
        host="localhost", user="root", password="your_password", database="calendarua"
    )

    # 建立資料表
    db.create_table()

    # 新增排程範例
    schedule_id = db.create_schedule(
        task_name="每日早班開機",
        opc_url="opc.tcp://localhost:4840",
        node_id="ns=2;i=1001",
        target_value="1",
        rrule_str="FREQ=DAILY;BYHOUR=8;BYMINUTE=0",
        is_enabled=1,
    )

    if schedule_id:
        # 查詢排程
        schedule = db.get_schedule(schedule_id)
        print(f"新增排程: {schedule}")

        # 更新排程
        db.update_schedule(schedule_id, target_value="2")

        # 查詢所有排程
        all_schedules = db.get_all_schedules()
        print(f"所有排程數量: {len(all_schedules)}")

        # 刪除排程
        # db.delete_schedule(schedule_id)

    db.close()

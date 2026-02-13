"""
SQLite 資料庫管理模組
使用 Python 內建 sqlite3 模組進行排程資料的 CRUD 操作
"""

import sqlite3
import logging
from pathlib import Path
from typing import List, Dict, Optional, Any
from contextlib import contextmanager


# 設定日誌記錄
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
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
            rrule_str TEXT NOT NULL,
            is_enabled INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                # 建立表格
                cursor.execute(create_table_sql)

                # 建立索引以提升查詢效能
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_schedules_enabled ON schedules(is_enabled)"
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_schedules_node_id ON schedules(node_id)"
                )

                conn.commit()
                logger.info("資料庫初始化成功，schedules 表格已建立")
                return True

        except sqlite3.Error as e:
            logger.error(f"資料庫初始化失敗: {e}")
            return False

    def add_schedule(
        self,
        task_name: str,
        opc_url: str,
        node_id: str,
        target_value: str,
        rrule_str: str,
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
            is_enabled: 是否啟用 (1: 啟用, 0: 停用)，預設為 1

        Returns:
            Optional[int]: 新增排程的 ID，失敗時回傳 None
        """
        insert_sql = """
        INSERT INTO schedules (task_name, opc_url, node_id, target_value, rrule_str, is_enabled)
        VALUES (?, ?, ?, ?, ?, ?)
        """

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    insert_sql,
                    (task_name, opc_url, node_id, target_value, rrule_str, is_enabled),
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
            is_enabled: 是否啟用 (1: 啟用, 0: 停用)，預設為 1

        Returns:
            Optional[int]: 新增排程的 ID，失敗時回傳 None
        """
        return self.add_schedule(
            task_name=task_name,
            opc_url=opc_url,
            node_id=node_id,
            target_value=target_value,
            rrule_str=rrule_str,
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
            "rrule_str",
            "is_enabled",
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


# 使用範例
if __name__ == "__main__":
    # 初始化資料庫管理器
    db = SQLiteManager()

    # 初始化資料庫（建立表格）
    if db.init_db():
        print("✓ 資料庫初始化成功")

        # 新增排程範例
        schedule_id = db.add_schedule(
            task_name="每日早班開機",
            opc_url="opc.tcp://localhost:4840",
            node_id="ns=2;i=1001",
            target_value="1",
            rrule_str="FREQ=DAILY;BYHOUR=8;BYMINUTE=0",
            is_enabled=1,
        )

        if schedule_id:
            print(f"✓ 排程新增成功，ID: {schedule_id}")

            # 查詢所有排程
            schedules = db.get_all_schedules()
            print(f"✓ 目前共有 {len(schedules)} 筆排程")

            # 刪除測試排程
            # if db.delete_schedule(schedule_id):
            #     print(f"✓ 排程 {schedule_id} 刪除成功")
    else:
        print("✗ 資料庫初始化失敗")

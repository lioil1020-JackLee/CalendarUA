import asyncio
from typing import Optional, Any, Union
from asyncua import Client, ua
from asyncua.common.node import Node
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class OPCHandler:
    """OPC UA 非同步處理類別，負責連線與讀寫操作"""

    def __init__(self, url: str, timeout: int = 10):
        """
        初始化 OPC UA 處理器

        Args:
            url: OPC UA 伺服器 URL (例如: opc.tcp://localhost:4840)
            timeout: 連線超時秒數
        """
        self.url = url
        self.timeout = timeout
        self.client: Optional[Client] = None
        self.is_connected = False

    async def connect(self) -> bool:
        """
        連線到 OPC UA 伺服器

        Returns:
            bool: 連線成功回傳 True
        """
        try:
            self.client = Client(self.url)
            await asyncio.wait_for(self.client.connect(), timeout=self.timeout)
            self.is_connected = True
            logger.info(f"成功連線到 OPC UA 伺服器: {self.url}")
            return True
        except asyncio.TimeoutError:
            logger.error(f"連線超時: {self.url}")
            self.is_connected = False
            return False
        except Exception as e:
            logger.error(f"連線失敗: {e}")
            self.is_connected = False
            return False

    async def disconnect(self):
        """中斷 OPC UA 連線"""
        if self.client and self.is_connected:
            try:
                await self.client.disconnect()
                logger.info(f"已中斷 OPC UA 連線: {self.url}")
            except Exception as e:
                logger.error(f"中斷連線時發生錯誤: {e}")
            finally:
                self.is_connected = False
                self.client = None

    async def __aenter__(self):
        """非同步上下文管理器進入"""
        await self.connect()
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        """非同步上下文管理器退出"""
        await self.disconnect()

    async def write_node(
        self, node_id: str, value: Union[str, int, float, bool]
    ) -> bool:
        """
        寫入數值到指定的 Node

        Args:
            node_id: OPC UA Node ID (例如: ns=2;i=1001 或 ns=2;s=MyTag)
            value: 要寫入的數值

        Returns:
            bool: 寫入成功回傳 True
        """
        if not self.is_connected or not self.client:
            logger.error("尚未連線到 OPC UA 伺服器")
            return False

        try:
            node = self.client.get_node(node_id)

            # 根據值類型轉換
            if isinstance(value, str):
                if value.lower() in ("true", "1"):
                    typed_value = True
                elif value.lower() in ("false", "0"):
                    typed_value = False
                else:
                    # 嘗試轉換為數字
                    try:
                        if "." in value:
                            typed_value = float(value)
                        else:
                            typed_value = int(value)
                    except ValueError:
                        typed_value = value
            else:
                typed_value = value

            await node.write_value(typed_value)
            logger.info(f"成功寫入 {node_id} = {typed_value}")
            return True

        except Exception as e:
            logger.error(f"寫入 Node {node_id} 失敗: {e}")
            return False

    async def read_node(self, node_id: str) -> Optional[Any]:
        """
        讀取指定 Node 的數值

        Args:
            node_id: OPC UA Node ID

        Returns:
            Optional[Any]: Node 數值，失敗回傳 None
        """
        if not self.is_connected or not self.client:
            logger.error("尚未連線到 OPC UA 伺服器")
            return None

        try:
            node = self.client.get_node(node_id)
            value = await node.read_value()
            logger.info(f"成功讀取 {node_id} = {value}")
            return value
        except Exception as e:
            logger.error(f"讀取 Node {node_id} 失敗: {e}")
            return None

    async def browse_nodes(self, node_id: str = None) -> list:
        """
        瀏覽 Node 的子節點

        Args:
            node_id: 父節點 ID (None 表示根節點)

        Returns:
            list: 子節點列表
        """
        if not self.is_connected or not self.client:
            logger.error("尚未連線到 OPC UA 伺服器")
            return []

        try:
            if node_id:
                node = self.client.get_node(node_id)
            else:
                node = self.client.get_objects_node()

            children = await node.get_children()
            nodes_info = []

            for child in children:
                try:
                    browse_name = await child.read_browse_name()
                    node_class = await child.read_node_class()
                    nodes_info.append(
                        {
                            "node_id": child.nodeid.to_string(),
                            "browse_name": browse_name.Name,
                            "node_class": node_class.name,
                        }
                    )
                except Exception as e:
                    logger.warning(f"讀取子節點資訊失敗: {e}")

            return nodes_info

        except Exception as e:
            logger.error(f"瀏覽節點失敗: {e}")
            return []


# 使用範例
async def main():
    """測試範例"""
    opc_url = "opc.tcp://localhost:4840"

    async with OPCHandler(opc_url) as handler:
        if handler.is_connected:
            # 寫入數值
            success = await handler.write_node("ns=2;i=1001", "1")
            print(f"寫入結果: {success}")

            # 讀取數值
            value = await handler.read_node("ns=2;i=1001")
            print(f"讀取數值: {value}")

            # 瀏覽節點
            nodes = await handler.browse_nodes()
            print(f"可用節點: {nodes}")


if __name__ == "__main__":
    asyncio.run(main())

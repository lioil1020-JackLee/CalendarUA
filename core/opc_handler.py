import asyncio
import os
from typing import Optional, Any, Union
from asyncua import Client, ua
from asyncua.common.node import Node
from asyncua.crypto.security_policies import SecurityPolicyBasic256Sha256
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class OPCHandler:
    """OPC UA 非同步處理類別，負責連線與讀寫操作

    支援的安全設定：
    - 無安全模式（僅測試使用）
    - 使用者名稱/密碼認證
    - X.509 憑證認證
    - 加密傳輸
    """

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

        # 安全配置
        self.security_policy: Optional[str] = None
        self.message_security_mode: Optional[str] = None
        self.username: Optional[str] = None
        self.password: Optional[str] = None
        self.client_cert_path: Optional[str] = None
        self.client_key_path: Optional[str] = None

    def set_security_policy(self, policy: str):
        """設定安全策略

        Args:
            policy: 安全策略名稱
                - None: 無加密
                - Basic256Sha256: 256位加密（推薦）
                - Aes128Sha256RsaOaep: AES-128 高安全
                - Aes256Sha256RsaPss: AES-256 最高安全
        """
        self.security_policy = policy
        logger.info(f"設定安全策略: {policy}")

    def set_user_credentials(self, username: str, password: str):
        """設定使用者名稱密碼認證

        Args:
            username: 使用者名稱
            password: 密碼
        """
        self.username = username
        self.password = password
        logger.info(f"設定使用者認證: {username}")

    def set_certificate(self, cert_path: str, key_path: str):
        """設定 X.509 憑證

        Args:
            cert_path: 憑證檔案路徑 (.der 或 .pem)
            key_path: 私鑰檔案路徑 (.pem)
        """
        self.client_cert_path = cert_path
        self.client_key_path = key_path
        logger.info(f"設定憑證認證: {cert_path}")

    async def connect(self) -> bool:
        """
        連線到 OPC UA 伺服器，根據配置自動選擇安全模式

        Returns:
            bool: 連線成功回傳 True
        """
        try:
            self.client = Client(self.url)

            # 設定安全策略
            if self.security_policy:
                await self._configure_security()

            # 設定使用者認證
            if self.username and self.password:
                self.client.set_user(self.username)
                self.client.set_password(self.password)

            # 連線到伺服器
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

    async def _configure_security(self):
        """配置安全設定"""
        try:
            # 載入憑證（如果有）
            if self.client_cert_path and self.client_key_path:
                await self.client.load_client_certificate(self.client_cert_path)
                await self.client.load_private_key(self.client_key_path)
                logger.info("已載入憑證和金鑰")

            # 設定安全策略
            if self.security_policy == "Basic256Sha256":
                from asyncua.crypto.security_policies import SecurityPolicyBasic256Sha256
                self.client.set_security_policy(SecurityPolicyBasic256Sha256)
                logger.info("已設定 Basic256Sha256 安全策略")
            elif self.security_policy == "Aes128Sha256RsaOaep":
                from asyncua.crypto.security_policies import SecurityPolicyAes128Sha256RsaOaep
                self.client.set_security_policy(SecurityPolicyAes128Sha256RsaOaep)
                logger.info("已設定 Aes128Sha256RsaOaep 安全策略")
            elif self.security_policy == "Aes256Sha256RsaPss":
                from asyncua.crypto.security_policies import SecurityPolicyAes256Sha256RsaPss
                self.client.set_security_policy(SecurityPolicyAes256Sha256RsaPss)
                logger.info("已設定 Aes256Sha256RsaPss 安全策略")
            # None 或其他值表示不設定安全策略（無加密）

        except Exception as e:
            logger.error(f"配置安全設定失敗: {e}")
            raise

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

    async def get_objects_node(self):
        """
        取得 Objects 節點

        Returns:
            Node: Objects 節點，如果未連線則回傳 None
        """
        if not self.is_connected or not self.client:
            logger.error("尚未連線到 OPC UA 伺服器")
            return None

        try:
            return self.client.get_objects_node()
        except Exception as e:
            logger.error(f"取得 Objects 節點失敗: {e}")
            return None


# 使用範例
async def main():
    """測試範例"""
    opc_url = os.environ.get("OPC_DEFAULT_URL", "opc.tcp://localhost:4840")

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

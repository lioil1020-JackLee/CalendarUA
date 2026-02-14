import asyncio
import os
from typing import Optional, Any, Union
from asyncua import Client, ua
from asyncua.common.node import Node
from asyncua.crypto.security_policies import SecurityPolicyBasic256Sha256
import logging

logging.basicConfig(level=logging.WARNING)
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
        self, node_id: str, value: Union[str, int, float, bool], data_type: str = "auto"
    ) -> bool:
        """
        寫入數值到指定的 Node，並進行讀取驗證

        Args:
            node_id: OPC UA Node ID (例如: ns=2;i=1001 或 ns=2;s=MyTag)
            value: 要寫入的數值
            data_type: 資料型別 (auto/int/float/string/bool)

        Returns:
            bool: 寫入並驗證成功回傳 True
        """
        if not self.is_connected or not self.client:
            logger.error("尚未連線到 OPC UA 伺服器")
            return False

        try:
            node = self.client.get_node(node_id)

            # 根據指定的 data_type 轉換
            if data_type == "auto":
                # 自動偵測模式（原有邏輯）
                typed_values_to_try = []

                if isinstance(value, str):
                    # 對於字串輸入，嘗試多種型別
                    original_str = value.strip()

                    # 1. 嘗試作為字串
                    typed_values_to_try.append(original_str)

                    # 2. 嘗試轉換為布林值
                    if original_str.lower() in ("true", "1", "on", "yes"):
                        typed_values_to_try.append(True)
                    elif original_str.lower() in ("false", "0", "off", "no"):
                        typed_values_to_try.append(False)

                    # 3. 嘗試轉換為數值
                    try:
                        if "." in original_str:
                            typed_values_to_try.append(float(original_str))
                        else:
                            # 同時嘗試int和float
                            int_val = int(original_str)
                            typed_values_to_try.append(int_val)
                            typed_values_to_try.append(float(original_str))
                    except ValueError:
                        pass  # 如果無法轉換為數值，跳過
                else:
                    # 非字串輸入直接使用
                    typed_values_to_try.append(value)
            else:
                # 指定型別模式
                typed_values_to_try = []
                original_str = str(value).strip() if isinstance(value, str) else str(value)

                if data_type == "string":
                    typed_values_to_try.append(original_str)
                elif data_type == "int":
                    try:
                        typed_values_to_try.append(int(float(original_str)))
                    except ValueError:
                        logger.error(f"無法將 '{original_str}' 轉換為整數")
                        return False
                elif data_type == "float":
                    try:
                        typed_values_to_try.append(float(original_str))
                    except ValueError:
                        logger.error(f"無法將 '{original_str}' 轉換為浮點數")
                        return False
                elif data_type == "bool":
                    if original_str.lower() in ("true", "1", "on", "yes"):
                        typed_values_to_try.append(True)
                    elif original_str.lower() in ("false", "0", "off", "no"):
                        typed_values_to_try.append(False)
                    else:
                        logger.error(f"無法將 '{original_str}' 轉換為布林值")
                        return False
                else:
                    logger.error(f"不支援的資料型別: {data_type}")
                    return False

            # 嘗試不同的型別
            last_exception = None
            for typed_value in typed_values_to_try:
                try:
                    # 根據資料型別設定正確的 VariantType
                    from asyncua import ua
                    
                    if data_type == "float":
                        # 對於 float 型別，明確使用 Float 而不是 Double
                        variant = ua.Variant(typed_value, ua.VariantType.Float)
                    else:
                        # 其他型別使用自動偵測
                        variant = ua.Variant(typed_value)
                    
                    # 寫入數值
                    await node.write_value(variant)
                    logger.info(f"成功寫入 {node_id} = {typed_value} (型別: {type(typed_value).__name__})")

                    # 讀取驗證
                    read_value = await node.read_value()
                    logger.info(f"驗證讀取 {node_id} = {read_value} (型別: {type(read_value).__name__})")

                    # 比較寫入值和讀取值
                    if self._values_equal(typed_value, read_value):
                        logger.info(f"驗證成功: 寫入值 {typed_value} 與讀取值 {read_value} 一致")
                        return True
                    else:
                        logger.warning(f"驗證失敗: 寫入值 {typed_value} 與讀取值 {read_value} 不一致，繼續嘗試其他型別")
                        continue

                except Exception as e:
                    last_exception = e
                    logger.warning(f"寫入型別 {type(typed_value).__name__} 失敗: {e}，繼續嘗試其他型別")
                    continue

            # 如果所有型別都失敗，記錄最後的錯誤
            logger.error(f"寫入 Node {node_id} 失敗，所有型別轉換都失敗: {last_exception}")
            return False

        except Exception as e:
            logger.error(f"寫入 Node {node_id} 失敗: {e}")
            return False

    def _values_equal(self, written_value: Any, read_value: Any) -> bool:
        """
        比較寫入值和讀取值是否相等，處理類型轉換和浮點數精度問題

        Args:
            written_value: 寫入的數值
            read_value: 讀取的數值

        Returns:
            bool: 值相等回傳 True
        """
        try:
            # 處理布林值比較
            if isinstance(written_value, bool) and isinstance(read_value, (int, bool)):
                return bool(written_value) == bool(read_value)
            elif isinstance(read_value, bool) and isinstance(written_value, (int, bool)):
                return bool(read_value) == bool(written_value)

            # 處理數值比較，允許小幅浮點誤差
            if isinstance(written_value, (int, float)) and isinstance(read_value, (int, float)):
                return abs(float(written_value) - float(read_value)) < 1e-6

            # 處理字串比較
            if isinstance(written_value, str) and isinstance(read_value, str):
                return str(written_value).strip() == str(read_value).strip()

            # 其他類型直接比較
            return written_value == read_value

        except Exception:
            # 如果比較過程中出現異常，視為不相等
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

    async def read_node_data_type(self, node_id: str) -> Optional[str]:
        """
        讀取指定 Node 的資料型別

        Args:
            node_id: OPC UA Node ID

        Returns:
            Optional[str]: 資料型別 (int/float/string/bool/auto)，失敗回傳 None
        """
        if not self.is_connected or not self.client:
            logger.error("尚未連線到 OPC UA 伺服器")
            return None

        try:
            node = self.client.get_node(node_id)
            
            # 讀取 VariantType
            variant_type = await node.read_data_type_as_variant_type()
            logger.info(f"成功讀取 {node_id} 資料型別: {variant_type}")
            
            # 映射到簡單型別系統
            return self._map_variant_type_to_simple_type(variant_type)
            
        except Exception as e:
            logger.error(f"讀取 Node {node_id} 資料型別失敗: {e}")
            return None

    def _map_variant_type_to_simple_type(self, variant_type: ua.VariantType) -> str:
        """
        將 OPC UA VariantType 映射到簡單型別系統

        Args:
            variant_type: OPC UA VariantType

        Returns:
            str: 簡單型別 (int/float/string/bool/auto)
        """
        # 整數型別映射到 "int"
        if variant_type in (
            ua.VariantType.SByte, ua.VariantType.Byte,
            ua.VariantType.Int16, ua.VariantType.UInt16,
            ua.VariantType.Int32, ua.VariantType.UInt32,
            ua.VariantType.Int64, ua.VariantType.UInt64
        ):
            return "int"
        
        # 浮點數型別映射到 "float"
        elif variant_type in (ua.VariantType.Float, ua.VariantType.Double):
            return "float"
        
        # 字串型別
        elif variant_type == ua.VariantType.String:
            return "string"
        
        # 布林型別
        elif variant_type == ua.VariantType.Boolean:
            return "bool"
        
        # 其他型別使用 "auto"
        else:
            return "auto"

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

    async def read_node_data_type(self, node_id: str) -> Optional[str]:
        """
        讀取指定 Node 的資料型別
        
        Args:
            node_id: OPC UA Node ID
            
        Returns:
            Optional[str]: 資料型別 (int/float/string/bool/auto)，失敗回傳 None
        """
        if not self.is_connected or not self.client:
            return None
        
        try:
            node = self.client.get_node(node_id)
            
            # 方法1：讀取 DataType 屬性並解析型別名稱
            try:
                data_type_nodeid = await node.read_data_type()
                data_type_node = self.client.get_node(data_type_nodeid)
                browse_name = await data_type_node.read_browse_name()
                type_name = browse_name.Name
                
                # 根據 OPC UA 標準型別名稱映射
                result = self._map_data_type_name_to_simple_type(type_name)
                if result != "auto":  # 如果不是複雜型別，返回結果
                    return result
            except Exception:
                pass
            
            # 方法2：使用 read_data_type_as_variant_type
            try:
                variant_type = await node.read_data_type_as_variant_type()
                return self._map_variant_type_to_simple_type(variant_type)
            except Exception:
                pass
                
            return "auto"  # 如果都失敗，使用auto
        except Exception as e:
            logger.error(f"讀取 Node {node_id} 資料型別失敗: {e}")
            return None

    def _map_data_type_name_to_simple_type(self, type_name: str) -> str:
        """
        將 OPC UA 資料型別名稱對應到簡單型別系統
        
        Args:
            type_name: OPC UA 資料型別名稱 (如 "Float", "Double", "Int32"等)
            
        Returns:
            str: 簡單型別 (int/float/string/bool/auto)
        """
        # 浮點數型別
        if type_name in ("Float", "Double"):
            return "float"
        
        # 整數型別
        elif type_name in ("SByte", "Byte", "Int16", "UInt16", "Int32", "UInt32", "Int64", "UInt64"):
            return "int"
        
        # 字串型別
        elif type_name in ("String", "LocalizedText"):
            return "string"
        
        # 布林型別
        elif type_name == "Boolean":
            return "bool"
        
        # 其他型別
        else:
            # 對於未知的複雜型別，使用auto讓系統自動處理
            return "auto"

    def _map_variant_type_to_simple_type(self, variant_type) -> str:
        """
        將 OPC UA VariantType 對應到簡單型別系統
        
        Args:
            variant_type: asyncua.ua.VariantType 列舉值
            
        Returns:
            str: 簡單型別 (int/float/string/bool/auto)
        """
        from asyncua.ua import VariantType
        
        # 浮點數型別
        if variant_type in (VariantType.Float, VariantType.Double):
            return "float"
        
        # 整數型別
        elif variant_type in (VariantType.SByte, VariantType.Byte, VariantType.Int16, 
                             VariantType.UInt16, VariantType.Int32, VariantType.UInt32, 
                             VariantType.Int64, VariantType.UInt64):
            return "int"
        
        # 字串型別
        elif variant_type in (VariantType.String, VariantType.LocalizedText):
            return "string"
        
        # 布林型別
        elif variant_type == VariantType.Boolean:
            return "bool"
        
        # 其他型別
        else:
            # 對於未知的複雜型別，使用auto讓系統自動處理
            return "auto"


# 使用範例
async def main():
    """測試範例"""
    opc_url = os.environ.get("OPC_DEFAULT_URL", "opc.tcp://localhost:4840")

    async with OPCHandler(opc_url) as handler:
        if handler.is_connected:
            # 寫入數值
            success = await handler.write_node("ns=2;i=1001", "1", "auto")
            print(f"寫入結果: {success}")

            # 讀取數值
            value = await handler.read_node("ns=2;i=1001")
            print(f"讀取數值: {value}")

            # 讀取資料型別
            data_type = await handler.read_node_data_type("ns=2;i=1001")
            print(f"讀取資料型別: {data_type}")

            # 瀏覽節點
            nodes = await handler.browse_nodes()
            print(f"可用節點: {nodes}")


if __name__ == "__main__":
    asyncio.run(main())

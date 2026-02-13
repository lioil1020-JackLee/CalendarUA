"""
OPC UA 安全設定配置檔案
提供不同安全等級的配置選項
"""

from dataclasses import dataclass
from typing import Optional, List
from enum import Enum


class SecurityLevel(Enum):
    """安全等級"""

    NONE = "none"  # 無安全（僅測試）
    BASIC = "basic"  # 基本（使用者名稱密碼）
    STANDARD = "standard"  # 標準（加密 + 憑證）
    HIGH = "high"  # 高安全（雙向憑證驗證）


@dataclass
class OPCSecurityConfig:
    """OPC UA 安全配置"""

    # 連線資訊
    url: str

    # 安全等級
    security_level: SecurityLevel = SecurityLevel.STANDARD

    # 安全策略
    security_policy: str = "Basic256Sha256"

    # 訊息安全模式
    # - None: 無安全
    # - Sign: 僅簽名
    # - SignAndEncrypt: 簽名+加密
    message_security_mode: str = "SignAndEncrypt"

    # 匿名認證
    use_anonymous: bool = False

    # 使用者名稱/密碼認證
    username: Optional[str] = None
    password: Optional[str] = None

    # 憑證檔案路徑
    client_certificate_path: Optional[str] = None
    client_private_key_path: Optional[str] = None
    server_certificate_path: Optional[str] = None

    # 憑證授權中心 (CA)
    trusted_ca_path: Optional[str] = None

    # 連線超時設定
    session_timeout: int = 60000  # 60秒
    secure_channel_timeout: int = 600000  # 10分鐘


class OPCSecurityPolicy:
    """OPC UA 安全策略定義"""

    POLICIES = {
        "None": {
            "uri": "http://opcfoundation.org/UA/SecurityPolicy#None",
            "description": "無加密（僅限測試環境）",
            "security_level": 0,
        },
        "Basic128Rsa15": {
            "uri": "http://opcfoundation.org/UA/SecurityPolicy#Basic128Rsa15",
            "description": "128位加密（已過時，不推薦）",
            "security_level": 1,
        },
        "Basic256": {
            "uri": "http://opcfoundation.org/UA/SecurityPolicy#Basic256",
            "description": "256位加密（過時，僅相容舊系統）",
            "security_level": 2,
        },
        "Basic256Sha256": {
            "uri": "http://opcfoundation.org/UA/SecurityPolicy#Basic256Sha256",
            "description": "SHA-256 + 256位加密（推薦）",
            "security_level": 3,
        },
        "Aes128Sha256RsaOaep": {
            "uri": "http://opcfoundation.org/UA/SecurityPolicy#Aes128_Sha256_RsaOaep",
            "description": "AES-128 + SHA-256 + RSA-OAEP（高安全）",
            "security_level": 4,
        },
        "Aes256Sha256RsaPss": {
            "uri": "http://opcfoundation.org/UA/SecurityPolicy#Aes256_Sha256_RsaPss",
            "description": "AES-256 + SHA-256 + RSA-PSS（最高安全）",
            "security_level": 5,
        },
    }

    @classmethod
    def get_recommended(cls) -> str:
        """取得推薦的安全策略"""
        return "Basic256Sha256"

    @classmethod
    def get_high_security(cls) -> str:
        """取得高安全策略"""
        return "Aes128Sha256RsaOaep"


class SecurityConfigFactory:
    """安全配置工廠"""

    @staticmethod
    def create_test_config(url: str) -> OPCSecurityConfig:
        """創建測試環境配置（無安全）"""
        return OPCSecurityConfig(
            url=url,
            security_level=SecurityLevel.NONE,
            security_policy="None",
            message_security_mode="None",
            use_anonymous=True,
        )

    @staticmethod
    def create_basic_config(
        url: str, username: str, password: str
    ) -> OPCSecurityConfig:
        """創建基本安全配置（使用者名稱+密碼+加密）"""
        return OPCSecurityConfig(
            url=url,
            security_level=SecurityLevel.BASIC,
            security_policy="Basic256Sha256",
            message_security_mode="SignAndEncrypt",
            username=username,
            password=password,
        )

    @staticmethod
    def create_standard_config(
        url: str,
        cert_path: str,
        key_path: str,
    ) -> OPCSecurityConfig:
        """創建標準安全配置（X.509 憑證）"""
        return OPCSecurityConfig(
            url=url,
            security_level=SecurityLevel.STANDARD,
            security_policy="Basic256Sha256",
            message_security_mode="SignAndEncrypt",
            client_certificate_path=cert_path,
            client_private_key_path=key_path,
        )

    @staticmethod
    def create_high_security_config(
        url: str,
        cert_path: str,
        key_path: str,
        ca_path: str,
    ) -> OPCSecurityConfig:
        """創建高安全配置（雙向憑證驗證）"""
        return OPCSecurityConfig(
            url=url,
            security_level=SecurityLevel.HIGH,
            security_policy="Aes128Sha256RsaOaep",
            message_security_mode="SignAndEncrypt",
            client_certificate_path=cert_path,
            client_private_key_path=key_path,
            trusted_ca_path=ca_path,
        )


# 使用範例
if __name__ == "__main__":
    # 範例 1: 測試環境
    test_config = SecurityConfigFactory.create_test_config("opc.tcp://localhost:4840")
    print(f"測試配置: {test_config}")

    # 範例 2: 基本安全（使用者名稱+密碼）
    basic_config = SecurityConfigFactory.create_basic_config(
        "opc.tcp://192.168.1.100:4840", username="operator", password="SecureP@ss123!"
    )
    print(f"基本安全配置: {basic_config}")

    # 範例 3: 標準安全（X.509 憑證）
    standard_config = SecurityConfigFactory.create_standard_config(
        "opc.tcp://192.168.1.100:4840",
        cert_path="./certs/client_cert.pem",
        key_path="./certs/client_key.pem",
    )
    print(f"標準安全配置: {standard_config}")

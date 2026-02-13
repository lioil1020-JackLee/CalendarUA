"""
OPC UA 安全設定配置檔案
提供不同安全等級的配置選項
"""

import os
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

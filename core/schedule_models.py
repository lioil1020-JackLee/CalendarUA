from __future__ import annotations

"""
核心資料模型 (Domain Models)

這個模組把行事曆與排程相關的概念，用 dataclass 統一定義出來，作為
UI、資料庫層與排程引擎之間的「語言」。

目前資料仍然儲存在 SQLite（database/sqlite_manager.py），這裡的模型
主要用於型別與結構說明，並不直接存取資料庫。
"""

from dataclasses import dataclass
from datetime import datetime, date, time
from typing import Optional


@dataclass
class CalendarProfile:
    """全域 Profile / 一般設定"""

    id: int
    profile_name: str
    description: str
    enable_schedule: bool
    scan_rate: int
    refresh_rate: int
    use_active_period: bool
    active_from: Optional[datetime]
    active_to: Optional[datetime]
    output_type: str
    refresh_output: bool
    generate_events: bool


@dataclass
class ScheduleSeries:
    """
    一條排程「系列」，對應資料庫 schedules 表的一列。

    這裡同時承載 OPC UA 寫值設定，讓一條排程就能描述：
    - 何時發生（RRULE）
    - 要對哪個 Node 寫入什麼值（OPC UA）
    """

    id: int
    task_name: str
    opc_url: str
    node_id: str
    target_value: str
    data_type: str
    rrule_str: str
    category_id: int
    is_enabled: bool

    # OPC UA 連線 / 安全設定
    opc_security_policy: str
    opc_security_mode: str
    opc_username: str
    opc_password: str
    opc_timeout: int
    opc_write_timeout: int


@dataclass
class ScheduleException:
    """
    排程例外：取消單次 occurrence，或覆寫單次 occurrence 的時間 / 標題 / 值。
    """

    id: int
    schedule_id: int
    occurrence_date: date
    action: str  # "cancel" or "override"
    override_start: Optional[datetime]
    override_end: Optional[datetime]
    override_task_name: Optional[str]
    override_target_value: Optional[str]


@dataclass
class HolidayCalendar:
    """假日日曆（僅描述 metadata，實際條目在 HolidayEntry）"""

    id: int
    name: str
    description: str
    is_default: bool


@dataclass
class HolidayEntry:
    """
    單一天的假日設定，可為全天或部分時段，並且可對排程做覆寫
    （例如節假日時改寫 OPC UA 目標值）。
    """

    id: int
    calendar_id: int
    holiday_date: date
    name: str
    is_full_day: bool
    start_time: Optional[time]
    end_time: Optional[time]
    override_category_id: Optional[int] = None
    override_target_value: Optional[str] = None


@dataclass
class RuntimeOverride:
    """
    Runtime Override：用戶在 RuntimePanel 設定的「立即覆寫值」，
    在有效期內，會優先於所有排程與假日的輸出。
    """

    override_value: str
    override_until: Optional[datetime]


"""
農曆 / 農民曆相關工具

目標：
- 提供西曆 <-> 農曆 轉換
- 提供指定日期的農民曆資訊（宜/忌、節氣等）
- 讓排程可以「以農曆規則」來設定，再轉換成實際西曆日期

目前實作採「可選依賴」策略：
- 若安裝了第三方套件 (例如 `lunarcalendar` 或其他農曆庫)，會優先使用
- 若沒有安裝，仍然可以匯入本模組，但只會回傳 None / 空結構

這樣可以讓整個專案在沒有安裝農曆套件時照常運作，
同時預留介面，方便未來擴充出完整的農民曆功能。
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Optional, Dict, Any


@dataclass
class LunarDateInfo:
    gregorian: date
    lunar_year: int
    lunar_month: int
    lunar_day: int
    is_leap_month: bool = False
    ganzhi_year: Optional[str] = None
    ganzhi_month: Optional[str] = None
    ganzhi_day: Optional[str] = None
    zodiac: Optional[str] = None
    solar_term: Optional[str] = None
    almanac_yi: Optional[str] = None
    almanac_ji: Optional[str] = None


def _load_backend():
    """
    嘗試載入實際的農曆計算套件。
    若找不到，回傳 (None, None)，呼叫者需自行處理降級行為。
    """
    try:
        import lunarcalendar  # type: ignore

        return lunarcalendar, "lunarcalendar"
    except Exception:
        return None, None


def to_lunar(gregorian: date) -> Optional[LunarDateInfo]:
    """
    將西曆日期轉換為農曆資訊。
    若系統尚未安裝農曆套件，回傳 None。
    """
    backend, name = _load_backend()
    if backend is None:
        return None

    # 以下為範例邏輯，實際需依照選用套件 API 實作
    try:
        # 假設 backend 提供 LunarDate.from_solar(year, month, day)
        lunar = backend.Converter.Solar2Lunar(gregorian)  # type: ignore[attr-defined]
        return LunarDateInfo(
            gregorian=gregorian,
            lunar_year=lunar.year,
            lunar_month=lunar.month,
            lunar_day=lunar.day,
            is_leap_month=getattr(lunar, "isleap", False),
        )
    except Exception:
        return None


def from_lunar(year: int, month: int, day: int, leap: bool = False) -> Optional[date]:
    """
    將農曆年月日轉為西曆日期。
    若無農曆套件或轉換失敗，回傳 None。
    """
    backend, name = _load_backend()
    if backend is None:
        return None

    try:
        lunar = backend.Lunar(year, month, day, isleap=leap)  # type: ignore[attr-defined]
        solar = backend.Converter.Lunar2Solar(lunar)
        return date(solar.year, solar.month, solar.day)
    except Exception:
        return None


def get_almanac_info(gregorian: date) -> Dict[str, Any]:
    """
    取得指定西曆日期的農民曆資訊（宜/忌、節氣等）。

    回傳格式範例：
    {
        "yi": "嫁娶 開市 ...",
        "ji": "動土 安葬 ...",
        "solar_term": "清明",
        "lunar_text": "農曆二月初三",
    }

    若無後端支援，回傳空 dict。
    """
    backend, name = _load_backend()
    if backend is None:
        return {}

    # 這裡僅示意，實際實作需依照選用的農民曆套件 API 來寫
    try:
        info: Dict[str, Any] = {}
        # 假設 backend 提供類似 API，可在此填入實作
        return info
    except Exception:
        return {}


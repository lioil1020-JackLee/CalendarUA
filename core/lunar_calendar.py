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
from typing import Optional


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
        import lunardate  # type: ignore

        return lunardate, "lunardate"
    except Exception:
        pass

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

    try:
        if name == "lunardate":
            lunar = backend.LunarDate.fromSolarDate(gregorian.year, gregorian.month, gregorian.day)
            return LunarDateInfo(
                gregorian=gregorian,
                lunar_year=lunar.year,
                lunar_month=lunar.month,
                lunar_day=lunar.day,
                is_leap_month=getattr(lunar, "isLeapMonth", False),
            )

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


def format_lunar_day_text(info: LunarDateInfo) -> str:
    """將農曆日轉換成 UI 顯示文字（如初一、十五、閏二月）。"""
    n = info.lunar_day
    if n <= 0 or n > 30:
        return ""

    if n == 1:
        month_names = {
            1: "元",
            2: "二",
            3: "三",
            4: "四",
            5: "五",
            6: "六",
            7: "七",
            8: "八",
            9: "九",
            10: "十",
            11: "十一",
            12: "十二",
        }
        month_text = month_names.get(info.lunar_month, str(info.lunar_month))
        leap_prefix = "閏" if info.is_leap_month else ""
        return f"{leap_prefix}{month_text}月"

    if n == 10:
        return "初十"
    if n == 20:
        return "二十"
    if n == 30:
        return "三十"

    chinese_ten = ["初", "十", "廿", "卅"]
    numerals = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
    return f"{chinese_ten[(n - 1) // 10]}{numerals[(n - 1) % 10]}"


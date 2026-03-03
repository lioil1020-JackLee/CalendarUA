from datetime import date, datetime, timedelta
from typing import Optional
from dateutil.rrule import rrulestr, rrule
import logging

from core.lunar_calendar import to_lunar

logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)


class RRuleParser:
    """RRULE 解析器，負責解析週期性規則並計算下一次觸發時間"""

    @staticmethod
    def _split_rrule_parts(rrule_str: str) -> tuple[dict[str, str], str]:
        params: dict[str, str] = {}
        dtstart_raw = ""
        for part in (rrule_str or "").split(";"):
            p = part.strip()
            if not p:
                continue
            if p.startswith("DTSTART:"):
                dtstart_raw = p.split(":", 1)[1]
                continue
            if "=" in p:
                key, value = p.split("=", 1)
                params[key.upper()] = value
        return params, dtstart_raw

    @staticmethod
    def _parse_dtstart(dtstart_raw: str, fallback: Optional[datetime] = None) -> datetime:
        if dtstart_raw:
            for fmt in ("%Y%m%dT%H%M%S", "%Y%m%dT%H%M", "%Y%m%d"):
                try:
                    parsed = datetime.strptime(dtstart_raw, fmt)
                    if fmt == "%Y%m%d":
                        return parsed.replace(hour=0, minute=0, second=0, microsecond=0)
                    return parsed.replace(microsecond=0)
                except ValueError:
                    continue
        if fallback is not None:
            return fallback.replace(microsecond=0)
        return datetime.now().replace(second=0, microsecond=0)

    @staticmethod
    def _to_int(value: Optional[str], default: int) -> int:
        try:
            if value is None:
                return default
            return int(value)
        except (TypeError, ValueError):
            return default

    @staticmethod
    def _is_lunar_mode(params: dict[str, str]) -> bool:
        return params.get("X-LUNAR", "0") == "1"

    @staticmethod
    def _supports_lunar_custom(params: dict[str, str]) -> bool:
        return params.get("FREQ", "").upper() in {"DAILY", "WEEKLY", "MONTHLY", "YEARLY"}

    @staticmethod
    def _build_solar_rrule_string(rrule_str: str) -> tuple[str, Optional[str]]:
        dtstart_str = None
        parts = rrule_str.split(";")
        rrule_parts = []
        for part in parts:
            p = part.strip()
            if not p:
                continue
            if p.startswith("DTSTART:"):
                dtstart_str = p.split(":", 1)[1]
            elif p.startswith("DURATION="):
                continue
            elif p.startswith("X-LUNAR="):
                continue
            elif p.startswith("X-RANGE-START="):
                continue
            else:
                rrule_parts.append(p)
        return ";".join(rrule_parts), dtstart_str

    @staticmethod
    def _parse_range_start(params: dict[str, str]) -> Optional[datetime]:
        raw = (params.get("X-RANGE-START") or "").strip()
        if not raw:
            return None
        try:
            if "T" in raw and len(raw) >= 15:
                return datetime.strptime(raw[:15], "%Y%m%dT%H%M%S")
            if len(raw) >= 8:
                d = datetime.strptime(raw[:8], "%Y%m%d")
                return d.replace(hour=0, minute=0, second=0, microsecond=0)
        except ValueError:
            return None
        return None

    @staticmethod
    def _parse_until(params: dict[str, str]) -> Optional[datetime]:
        until = params.get("UNTIL")
        if not until:
            return None

        for fmt in ("%Y%m%dT%H%M%S", "%Y%m%dT%H%M", "%Y%m%d"):
            try:
                parsed = datetime.strptime(until, fmt)
                if fmt == "%Y%m%d":
                    return parsed.replace(hour=23, minute=59, second=59)
                return parsed
            except ValueError:
                continue
        return None

    @staticmethod
    def _matches_lunar_rule(candidate: datetime, params: dict[str, str], dtstart: datetime) -> bool:
        freq = params.get("FREQ", "").upper()
        interval = max(1, RRuleParser._to_int(params.get("INTERVAL"), 1))

        if freq == "DAILY":
            day_diff = (candidate.date() - dtstart.date()).days
            return day_diff >= 0 and day_diff % interval == 0

        if freq == "WEEKLY":
            day_diff = (candidate.date() - dtstart.date()).days
            if day_diff < 0:
                return False
            week_diff = day_diff // 7
            if week_diff % interval != 0:
                return False

            byday = params.get("BYDAY", "")
            if not byday:
                return candidate.weekday() == dtstart.weekday()

            weekday_map = {
                "MO": 0,
                "TU": 1,
                "WE": 2,
                "TH": 3,
                "FR": 4,
                "SA": 5,
                "SU": 6,
            }
            allowed = {
                weekday_map[token.strip()]
                for token in byday.split(",")
                if token.strip() in weekday_map
            }
            if not allowed:
                return candidate.weekday() == dtstart.weekday()
            return candidate.weekday() in allowed

        start_lunar = to_lunar(dtstart.date())
        candidate_lunar = to_lunar(candidate.date())
        if not start_lunar or not candidate_lunar:
            return False

        bymonth = params.get("BYMONTH")
        bymonthday = params.get("BYMONTHDAY")

        if freq == "MONTHLY":
            day_target = RRuleParser._to_int(bymonthday, start_lunar.lunar_day)
            month_diff = ((candidate_lunar.lunar_year - start_lunar.lunar_year) * 12
                          + (candidate_lunar.lunar_month - start_lunar.lunar_month))
            if month_diff < 0 or month_diff % interval != 0:
                return False
            return candidate_lunar.lunar_day == day_target

        if freq == "YEARLY":
            month_target = RRuleParser._to_int(bymonth, start_lunar.lunar_month)
            day_target = RRuleParser._to_int(bymonthday, start_lunar.lunar_day)
            year_diff = candidate_lunar.lunar_year - start_lunar.lunar_year
            if year_diff < 0 or year_diff % interval != 0:
                return False
            return candidate_lunar.lunar_month == month_target and candidate_lunar.lunar_day == day_target

        return False

    @staticmethod
    def _generate_lunar_occurrences(
        rrule_str: str,
        horizon_end: datetime,
        lower_bound: Optional[datetime] = None,
        limit: Optional[int] = None,
    ) -> list[datetime]:
        params, dtstart_raw = RRuleParser._split_rrule_parts(rrule_str)
        dtstart = RRuleParser._parse_dtstart(dtstart_raw)

        hour = RRuleParser._to_int(params.get("BYHOUR"), dtstart.hour)
        minute = RRuleParser._to_int(params.get("BYMINUTE"), dtstart.minute)
        second = RRuleParser._to_int(params.get("BYSECOND"), dtstart.second)

        until = RRuleParser._parse_until(params)
        count = max(0, RRuleParser._to_int(params.get("COUNT"), 0))

        if horizon_end < dtstart:
            return []

        if until and until < dtstart:
            return []

        end_dt = horizon_end
        if until is not None and until < end_dt:
            end_dt = until

        occurrences: list[datetime] = []
        emitted_total = 0
        cursor = dtstart.date()
        end_date = end_dt.date()

        while cursor <= end_date:
            candidate = datetime(cursor.year, cursor.month, cursor.day, hour, minute, second)

            if candidate < dtstart:
                cursor += timedelta(days=1)
                continue

            if until is not None and candidate > until:
                break

            if RRuleParser._matches_lunar_rule(candidate, params, dtstart):
                emitted_total += 1
                if count > 0 and emitted_total > count:
                    break

                if lower_bound is None or candidate >= lower_bound:
                    occurrences.append(candidate)
                    if limit is not None and len(occurrences) >= limit:
                        break

            cursor += timedelta(days=1)

        return occurrences

    @staticmethod
    def parse_rrule(
        rrule_str: str, dtstart: Optional[datetime] = None
    ) -> Optional[rrule]:
        """
        解析 RRULE 字串為 rrule 物件

        Args:
            rrule_str: RRULE 規則字串 (例如: FREQ=DAILY;BYHOUR=8;BYMINUTE=0)
            dtstart: 開始時間 (預設為現在)

        Returns:
            Optional[rrule]: rrule 物件，解析失敗回傳 None
        """
        try:
            # 從 RRULE 字串中提取 DTSTART，並移除 dateutil 不支援參數
            rrule_str, dtstart_str = RRuleParser._build_solar_rrule_string(rrule_str)

            if dtstart is None:
                if dtstart_str:
                    dtstart = RRuleParser._parse_dtstart(dtstart_str)
                else:
                    dtstart = datetime.now().replace(second=0, microsecond=0)

            # 確保 RRULE 字串格式正確
            if not rrule_str.upper().startswith("RRULE:"):
                rrule_str = f"RRULE:{rrule_str}"

            rule = rrulestr(rrule_str, dtstart=dtstart)
            return rule

        except Exception as e:
            logger.error(f"解析 RRULE 失敗 '{rrule_str}': {e}")
            return None

    @staticmethod
    def get_next_trigger(
        rrule_str: str,
        dtstart: Optional[datetime] = None,
        after: Optional[datetime] = None,
    ) -> Optional[datetime]:
        """
        取得下一次觸發時間

        Args:
            rrule_str: RRULE 規則字串
            dtstart: 開始時間
            after: 在此時間之後的下一個觸發時間 (預設為現在)

        Returns:
            Optional[datetime]: 下一次觸發時間，無效回傳 None
        """
        try:
            params, _ = RRuleParser._split_rrule_parts(rrule_str)
            range_start = RRuleParser._parse_range_start(params)
            if RRuleParser._is_lunar_mode(params) and RRuleParser._supports_lunar_custom(params):
                if after is None:
                    after = datetime.now()
                if range_start is not None and after < range_start:
                    after = range_start - timedelta(seconds=1)
                search_from = after + timedelta(seconds=1)
                horizon = after + timedelta(days=365 * 20)
                occurrences = RRuleParser._generate_lunar_occurrences(
                    rrule_str,
                    horizon_end=horizon,
                    lower_bound=search_from,
                    limit=1,
                )
                if occurrences:
                    return occurrences[0]
                return None

            rule = RRuleParser.parse_rrule(rrule_str, dtstart)
            if rule is None:
                return None

            if after is None:
                after = datetime.now()
            if range_start is not None and after < range_start:
                after = range_start - timedelta(seconds=1)

            # 取得 after 之後的下一個觸發時間
            next_trigger = rule.after(after, inc=False)

            if next_trigger:
                logger.info(f"下一次觸發時間: {next_trigger}")
                return next_trigger
            else:
                logger.info("沒有下一次觸發時間 (可能已達結束條件)")
                return None

        except Exception as e:
            logger.error(f"計算下一次觸發時間失敗: {e}")
            return None

    @staticmethod
    def get_upcoming_triggers(
        rrule_str: str,
        count: int = 5,
        dtstart: Optional[datetime] = None,
        after: Optional[datetime] = None,
    ) -> list[datetime]:
        """
        取得接下來 N 次的觸發時間。

        注意：這個方法會使用 dateutil.rrule 提供的 iterator，
        不會修改或依賴任何 UI / 資料庫邏輯。
        """
        try:
            params, _ = RRuleParser._split_rrule_parts(rrule_str)
            range_start = RRuleParser._parse_range_start(params)
            if RRuleParser._is_lunar_mode(params) and RRuleParser._supports_lunar_custom(params):
                if after is None:
                    after = datetime.now()
                if range_start is not None and after < range_start:
                    after = range_start - timedelta(seconds=1)
                search_from = after + timedelta(seconds=1)
                horizon = after + timedelta(days=365 * 20)
                return RRuleParser._generate_lunar_occurrences(
                    rrule_str,
                    horizon_end=horizon,
                    lower_bound=search_from,
                    limit=max(0, count),
                )

            rule = RRuleParser.parse_rrule(rrule_str, dtstart)
            if rule is None:
                return []

            if after is None:
                after = datetime.now()
            if range_start is not None and after < range_start:
                after = range_start - timedelta(seconds=1)

            triggers: list[datetime] = []
            current = rule.after(after, inc=False)
            while current is not None and len(triggers) < count:
                triggers.append(current)
                current = rule.after(current, inc=False)

            return triggers

        except Exception as e:
            logger.error(f"計算觸發時間列表失敗: {e}")
            return []

    @staticmethod
    def get_trigger_between(
        rrule_str: str,
        start: datetime,
        end: datetime,
        dtstart: Optional[datetime] = None,
    ) -> list:
        """
        取得指定時間範圍內的所有觸發時間

        Args:
            rrule_str: RRULE 規則字串
            start: 開始時間
            end: 結束時間
            dtstart: RRULE 開始時間

        Returns:
            list: 觸發時間列表
        """
        try:
            params, _ = RRuleParser._split_rrule_parts(rrule_str)
            if RRuleParser._is_lunar_mode(params) and RRuleParser._supports_lunar_custom(params):
                occurrences = RRuleParser._generate_lunar_occurrences(
                    rrule_str,
                    horizon_end=end,
                    lower_bound=start,
                )
                return [occ for occ in occurrences if occ <= end]

            rule = RRuleParser.parse_rrule(rrule_str, dtstart)
            if rule is None:
                return []

            # 取得時間範圍內的所有觸發時間
            triggers = list(rule.between(start, end, inc=True))
            return triggers

        except Exception as e:
            logger.error(f"計算時間範圍觸發時間失敗: {e}")
            return []

    @staticmethod
    def is_trigger_time(
        rrule_str: str,
        check_time: Optional[datetime] = None,
        tolerance_seconds: int = 60,
    ) -> bool:
        """
        檢查指定時間是否為觸發時間 (考慮容許誤差)

        Args:
            rrule_str: RRULE 規則字串
            check_time: 要檢查的時間 (預設為現在)
            tolerance_seconds: 容許誤差秒數

        Returns:
            bool: 是否為觸發時間
        """
        try:
            if check_time is None:
                check_time = datetime.now()

            params, _ = RRuleParser._split_rrule_parts(rrule_str)
            range_start = RRuleParser._parse_range_start(params)
            if range_start is not None and check_time < range_start:
                return False
            if RRuleParser._is_lunar_mode(params) and RRuleParser._supports_lunar_custom(params):
                window_start = check_time - timedelta(seconds=tolerance_seconds)
                window_end = check_time
                if range_start is not None and window_start < range_start:
                    window_start = range_start
                triggers = RRuleParser.get_trigger_between(rrule_str, window_start, window_end)
                return any(0 <= (check_time - trigger).total_seconds() <= tolerance_seconds for trigger in triggers)

            # 取得上一次和下一次觸發時間
            rule = RRuleParser.parse_rrule(
                rrule_str, dtstart=check_time - timedelta(days=1)
            )
            if rule is None:
                return False

            # 只檢查 check_time 之前最近一次觸發，避免提早觸發
            prev_trigger = rule.before(check_time, inc=True)

            # 檢查是否接近觸發時間
            if prev_trigger:
                diff = (check_time - prev_trigger).total_seconds()
                if diff <= tolerance_seconds:
                    return True

            return False

        except Exception as e:
            logger.error(f"檢查觸發時間失敗: {e}")
            return False

    @staticmethod
    def validate_rrule(rrule_str: str) -> bool:
        """
        驗證 RRULE 字串是否有效

        Args:
            rrule_str: RRULE 規則字串

        Returns:
            bool: 有效回傳 True
        """
        try:
            params, _ = RRuleParser._split_rrule_parts(rrule_str)
            if RRuleParser._is_lunar_mode(params) and RRuleParser._supports_lunar_custom(params):
                dtstart = RRuleParser._parse_dtstart(RRuleParser._split_rrule_parts(rrule_str)[1])
                RRuleParser._generate_lunar_occurrences(
                    rrule_str,
                    horizon_end=dtstart + timedelta(days=370),
                    lower_bound=dtstart,
                    limit=1,
                )
                return True

            RRuleParser.parse_rrule(rrule_str)
            return True
        except Exception:
            return False

    @staticmethod
    def create_daily_rule(hour: int, minute: int, **kwargs) -> str:
        """
        建立每日觸發的 RRULE 字串。
        不依賴任何 UI / DB，方便在對話框或測試程式中直接呼叫。
        """
        rrule_str = f"FREQ=DAILY;BYHOUR={hour};BYMINUTE={minute}"
        for key, value in kwargs.items():
            rrule_str += f";{key.upper()}={value}"
        return rrule_str

    @staticmethod
    def create_weekly_rule(hour: int, minute: int, days: list[str], **kwargs) -> str:
        """建立每週觸發的 RRULE 字串。"""
        byday = ",".join(days)
        rrule_str = f"FREQ=WEEKLY;BYHOUR={hour};BYMINUTE={minute};BYDAY={byday}"
        for key, value in kwargs.items():
            rrule_str += f";{key.upper()}={value}"
        return rrule_str

    @staticmethod
    def create_monthly_rule(hour: int, minute: int, monthday: int, **kwargs) -> str:
        """建立每月觸發的 RRULE 字串。"""
        rrule_str = (
            f"FREQ=MONTHLY;BYHOUR={hour};BYMINUTE={minute};BYMONTHDAY={monthday}"
        )
        for key, value in kwargs.items():
            rrule_str += f";{key.upper()}={value}"
        return rrule_str


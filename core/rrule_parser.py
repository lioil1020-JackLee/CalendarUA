from datetime import datetime, timedelta
from typing import Optional, Iterator
from dateutil.rrule import rrulestr, rrule
from dateutil.parser import parse
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class RRuleParser:
    """RRULE 解析器，負責解析週期性規則並計算下一次觸發時間"""

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
            # 從 RRULE 字串中提取 DTSTART 和過濾不支援的參數
            dtstart_str = None
            if "DTSTART:" in rrule_str:
                parts = rrule_str.split(";")
                rrule_parts = []
                for part in parts:
                    if part.startswith("DTSTART:"):
                        dtstart_str = part.split(":", 1)[1]
                    elif part.startswith("DURATION="):
                        # 忽略 DURATION 參數，因為 dateutil.rrule 不支援
                        continue
                    else:
                        rrule_parts.append(part)
                rrule_str = ";".join(rrule_parts)

            if dtstart is None:
                if dtstart_str:
                    # 解析 DTSTART 字串
                    try:
                        dtstart = datetime.strptime(dtstart_str, "%Y%m%dT%H%M%S")
                    except ValueError:
                        dtstart = datetime.now().replace(second=0, microsecond=0)
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
            rule = RRuleParser.parse_rrule(rrule_str, dtstart)
            if rule is None:
                return None

            if after is None:
                after = datetime.now()

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
    ) -> list:
        """
        取得接下來 N 次的觸發時間

        Args:
            rrule_str: RRULE 規則字串
            count: 要取得的觸發次數
            dtstart: 開始時間
            after: 從此時間之後開始計算

        Returns:
            list: 觸發時間列表
        """
        try:
            rule = RRuleParser.parse_rrule(rrule_str, dtstart)
            if rule is None:
                return []

            if after is None:
                after = datetime.now()

            # 取得接下來的觸發時間
            triggers = list(rule.after(after, inc=False) for _ in range(count))
            # 過濾掉 None 值
            triggers = [t for t in triggers if t is not None][:count]

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

            # 取得上一次和下一次觸發時間
            rule = RRuleParser.parse_rrule(
                rrule_str, dtstart=check_time - timedelta(days=1)
            )
            if rule is None:
                return False

            # 取得最接近 check_time 的觸發時間
            prev_trigger = rule.before(check_time, inc=True)
            next_trigger = rule.after(check_time, inc=False)

            # 檢查是否接近觸發時間
            if prev_trigger:
                diff = abs((check_time - prev_trigger).total_seconds())
                if diff <= tolerance_seconds:
                    return True

            if next_trigger:
                diff = abs((next_trigger - check_time).total_seconds())
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
            RRuleParser.parse_rrule(rrule_str)
            return True
        except Exception:
            return False

    @staticmethod
    def create_daily_rule(hour: int, minute: int, **kwargs) -> str:
        """
        建立每日觸發的 RRULE 字串

        Args:
            hour: 小時 (0-23)
            minute: 分鐘 (0-59)
            **kwargs: 其他 RRULE 參數

        Returns:
            str: RRULE 字串
        """
        rrule_str = f"FREQ=DAILY;BYHOUR={hour};BYMINUTE={minute}"
        for key, value in kwargs.items():
            rrule_str += f";{key.upper()}={value}"
        return rrule_str

    @staticmethod
    def create_weekly_rule(hour: int, minute: int, days: list, **kwargs) -> str:
        """
        建立每週觸發的 RRULE 字串

        Args:
            hour: 小時 (0-23)
            minute: 分鐘 (0-59)
            days: 星期列表 ['MO', 'TU', 'WE', 'TH', 'FR', 'SA', 'SU']
            **kwargs: 其他 RRULE 參數

        Returns:
            str: RRULE 字串
        """
        byday = ",".join(days)
        rrule_str = f"FREQ=WEEKLY;BYHOUR={hour};BYMINUTE={minute};BYDAY={byday}"
        for key, value in kwargs.items():
            rrule_str += f";{key.upper()}={value}"
        return rrule_str

    @staticmethod
    def create_monthly_rule(hour: int, minute: int, monthday: int, **kwargs) -> str:
        """
        建立每月觸發的 RRULE 字串

        Args:
            hour: 小時 (0-23)
            minute: 分鐘 (0-59)
            monthday: 每月幾號 (1-31)
            **kwargs: 其他 RRULE 參數

        Returns:
            str: RRULE 字串
        """
        rrule_str = (
            f"FREQ=MONTHLY;BYHOUR={hour};BYMINUTE={minute};BYMONTHDAY={monthday}"
        )
        for key, value in kwargs.items():
            rrule_str += f";{key.upper()}={value}"
        return rrule_str


# 使用範例
if __name__ == "__main__":
    # 每日早上 8:00
    daily_rule = "FREQ=DAILY;BYHOUR=8;BYMINUTE=0"
    next_trigger = RRuleParser.get_next_trigger(daily_rule)
    print(f"每日觸發，下次: {next_trigger}")

    # 每週二、四下午 2:00
    weekly_rule = "FREQ=WEEKLY;BYHOUR=14;BYMINUTE=0;BYDAY=TU,TH"
    next_trigger = RRuleParser.get_next_trigger(weekly_rule)
    print(f"每週二四觸發，下次: {next_trigger}")

    # 取得接下來 5 次觸發
    triggers = RRuleParser.get_upcoming_triggers(weekly_rule, count=5)
    print(f"\n接下來 5 次觸發:")
    for i, t in enumerate(triggers, 1):
        print(f"  {i}. {t}")

    # 使用輔助方法建立規則
    print("\n使用輔助方法:")
    daily = RRuleParser.create_daily_rule(9, 30)
    print(f"每日 9:30: {daily}")

    weekly = RRuleParser.create_weekly_rule(14, 0, ["MO", "WE", "FR"])
    print(f"每週一三五 14:00: {weekly}")

    monthly = RRuleParser.create_monthly_rule(10, 0, 15)
    print(f"每月 15 日 10:00: {monthly}")

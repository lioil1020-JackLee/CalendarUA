from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Any, Dict, List, Optional
import re

from core.rrule_parser import RRuleParser


@dataclass
class ResolvedOccurrence:
    schedule_id: int
    source: str
    title: str
    start: datetime
    end: datetime
    category_bg: str
    category_fg: str
    target_value: str
    is_exception: bool
    is_holiday: bool
    occurrence_key: str


def _extract_duration_minutes(rrule_str: str) -> int:
    match = re.search(r"DURATION=PT(?:(\d+)H)?(?:(\d+)M)?", rrule_str.upper())
    if not match:
        return 60

    hours = int(match.group(1) or 0)
    minutes = int(match.group(2) or 0)
    total = hours * 60 + minutes
    return total if total > 0 else 60


def _extract_dtstart(rrule_str: str) -> Optional[datetime]:
    match = re.search(r"DTSTART:(\d{8}(?:T\d{6})?)", rrule_str.upper())
    if not match:
        return None

    raw = match.group(1)
    try:
        if "T" in raw:
            return datetime.strptime(raw, "%Y%m%dT%H%M%S")
        return datetime.strptime(raw, "%Y%m%d")
    except ValueError:
        return None


def _extract_range_start(rrule_str: str) -> Optional[datetime]:
    match = re.search(r"X-RANGE-START=(\d{8}(?:T\d{6})?)", rrule_str.upper())
    if not match:
        return None

    raw = match.group(1)
    try:
        if "T" in raw:
            return datetime.strptime(raw, "%Y%m%dT%H%M%S")
        return datetime.strptime(raw, "%Y%m%d")
    except ValueError:
        return None


def _extract_title(schedule: Dict[str, Any]) -> str:
    task_name = str(schedule.get("task_name", "")).strip()
    return task_name or f"任務{schedule.get('id', '')}"


def _extract_target_value(schedule: Dict[str, Any]) -> str:
    return str(schedule.get("target_value", "")).strip()


def _pick_color(target_value: str) -> tuple[str, str]:
    return "#ff1a1a", "#ffffff"


def _build_holiday_map(holiday_entries: Optional[List[Dict[str, Any]]]) -> Dict[str, List[Dict[str, Any]]]:
    holiday_map: Dict[str, List[Dict[str, Any]]] = {}
    if not holiday_entries:
        return holiday_map

    for entry in holiday_entries:
        date_key = str(entry.get("holiday_date", "")).strip()
        if not date_key:
            continue
        holiday_map.setdefault(date_key, []).append(entry)

    return holiday_map


def _parse_time_str(time_str: str) -> time:
    if not time_str:
        return time.min

    try:
        parts = [int(p) for p in time_str.split(":")]
        if len(parts) == 2:
            return time(parts[0], parts[1], 0)
        if len(parts) >= 3:
            return time(parts[0], parts[1], parts[2])
    except ValueError:
        return time.min

    return time.min


def _holiday_entry_overlaps(entry: Dict[str, Any], start_dt: datetime, end_dt: datetime) -> bool:
    if entry.get("is_full_day", 1):
        return True

    start_time = _parse_time_str(str(entry.get("start_time") or ""))
    end_time = _parse_time_str(str(entry.get("end_time") or ""))

    if end_time <= start_time:
        return False

    holiday_start = datetime.combine(start_dt.date(), start_time)
    holiday_end = datetime.combine(start_dt.date(), end_time)
    return not (end_dt <= holiday_start or start_dt >= holiday_end)


def _pick_holiday_entry(
    entries: List[Dict[str, Any]],
    start_dt: datetime,
    end_dt: datetime,
) -> Optional[Dict[str, Any]]:
    matched: List[Dict[str, Any]] = []
    for entry in entries:
        if _holiday_entry_overlaps(entry, start_dt, end_dt):
            matched.append(entry)

    if not matched:
        return None

    for entry in matched:
        if entry.get("override_target_value"):
            return entry

    return matched[0]


def resolve_occurrences_for_range(
    schedules: List[Dict[str, Any]],
    range_start: datetime,
    range_end: datetime,
    schedule_exceptions: Optional[List[Dict[str, Any]]] = None,
    holiday_entries: Optional[List[Dict[str, Any]]] = None,
    db_manager = None,
) -> List[ResolvedOccurrence]:
    occurrences: List[ResolvedOccurrence] = []
    exception_map: Dict[tuple[int, str], Dict[str, Any]] = {}

    if schedule_exceptions:
        for exception in schedule_exceptions:
            try:
                sid = int(exception.get("schedule_id", 0) or 0)
            except (TypeError, ValueError):
                continue
            date_key = str(exception.get("occurrence_date", "")).strip()
            if sid > 0 and date_key:
                exception_map[(sid, date_key)] = exception

    holiday_entries_list = holiday_entries
    if holiday_entries_list is None and db_manager:
        try:
            holiday_entries_list = db_manager.get_all_holiday_entries()
        except Exception:
            holiday_entries_list = []

    holiday_map = _build_holiday_map(holiday_entries_list)

    for schedule in schedules:
        if not schedule.get("is_enabled"):
            continue

        rrule_str = str(schedule.get("rrule_str", "")).strip()
        if not rrule_str:
            continue

        triggers = RRuleParser.get_trigger_between(rrule_str, range_start, range_end)
        configured_range_start = _extract_range_start(rrule_str)
        duration_minutes = _extract_duration_minutes(rrule_str)
        title = _extract_title(schedule)
        target_value = _extract_target_value(schedule)

        schedule_bg, schedule_fg = _pick_color(target_value)

        if not triggers:
            dtstart = _extract_dtstart(rrule_str)
            upper_rrule = rrule_str.upper()
            has_expire_condition = ("UNTIL=" in upper_rrule) or ("COUNT=" in upper_rrule)
            next_trigger = RRuleParser.get_next_trigger(rrule_str, after=datetime.now())
            is_truly_expired = has_expire_condition and next_trigger is None

            if (
                is_truly_expired
                and dtstart
                and range_start <= dtstart < range_end
                and dtstart <= datetime.now()
            ):
                expired_occurrence = ResolvedOccurrence(
                    schedule_id=int(schedule.get("id", 0) or 0),
                    source="expired",
                    title=f"{title} (過期)",
                    start=dtstart,
                    end=dtstart + timedelta(minutes=duration_minutes),
                    category_bg="#000000",
                    category_fg="#ffffff",
                    target_value=target_value,
                    is_exception=False,
                    is_holiday=False,
                    occurrence_key=f"{int(schedule.get('id', 0) or 0)}:{dtstart.isoformat()}:expired",
                )
                occurrences.append(expired_occurrence)
            continue

        for trigger in triggers:
            start = trigger
            end = trigger + timedelta(minutes=duration_minutes)

            if end <= range_start or start >= range_end:
                continue

            schedule_id = int(schedule.get("id", 0) or 0)
            occurrence_date = start.date().isoformat()
            exception = exception_map.get((schedule_id, occurrence_date))

            if exception and str(exception.get("action", "")).lower() == "cancel":
                continue

            source = "weekly"
            is_exception = False
            is_holiday = False
            resolved_title = title
            resolved_target = target_value
            resolved_start = start
            resolved_end = end

            holiday_entry = None
            if holiday_map:
                holiday_entry = _pick_holiday_entry(holiday_map.get(occurrence_date, []), start, end)

            if holiday_entry:
                is_holiday = True
                source = "holiday"

                if holiday_entry.get("override_target_value"):
                    resolved_target = str(holiday_entry.get("override_target_value"))

                bg_color, fg_color = _pick_color(resolved_target)
            else:
                bg_color, fg_color = schedule_bg, schedule_fg

            if exception and str(exception.get("action", "")).lower() == "override":
                try:
                    override_start = exception.get("override_start")
                    override_end = exception.get("override_end")
                    if override_start:
                        resolved_start = datetime.fromisoformat(str(override_start))
                    if override_end:
                        resolved_end = datetime.fromisoformat(str(override_end))
                except ValueError:
                    resolved_start = start
                    resolved_end = end

                if exception.get("override_task_name"):
                    resolved_title = str(exception.get("override_task_name"))
                if exception.get("override_target_value"):
                    resolved_target = str(exception.get("override_target_value"))

                source = "exception"
                is_exception = True
                bg_color, fg_color = _pick_color(resolved_target)
            else:
                # 沒有 exception 覆寫,保留 holiday 或 schedule 顏色
                pass

            if configured_range_start and resolved_start < configured_range_start:
                if not resolved_title.endswith("(過期)"):
                    resolved_title = f"{resolved_title} (過期)"
                source = "expired"
                bg_color, fg_color = "#000000", "#ffffff"

            if resolved_end <= resolved_start:
                continue

            occurrence = ResolvedOccurrence(
                schedule_id=schedule_id,
                source=source,
                title=resolved_title,
                start=resolved_start,
                end=resolved_end,
                category_bg=bg_color,
                category_fg=fg_color,
                target_value=resolved_target,
                is_exception=is_exception,
                is_holiday=is_holiday,
                occurrence_key=f"{schedule_id}:{resolved_start.isoformat()}",
            )
            occurrences.append(occurrence)

    occurrences.sort(key=lambda item: (item.start, item.end, item.schedule_id))
    return occurrences

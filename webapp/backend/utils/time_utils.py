from datetime import datetime, timedelta, date, time


def parse_excel_datetime(value):
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, time(0, 0))
    return None


def parse_excel_time(value):
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, time):
        return value
    return None


def parse_excel_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def is_time_match(mail_time: datetime, target_date, target_time, tolerance_minutes: int = 1) -> bool:
    if isinstance(target_date, datetime):
        target_date = target_date.date()
    if isinstance(target_time, datetime):
        target_time = target_time.time()

    target_dt = datetime.combine(target_date, target_time)
    time_from = target_dt - timedelta(minutes=tolerance_minutes)
    time_to = target_dt + timedelta(minutes=tolerance_minutes)

    return time_from <= mail_time <= time_to

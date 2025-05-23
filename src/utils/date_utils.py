import pandas as pd
from datetime import datetime

def date_to_week(date_str: str) -> int:
    """
    Convert a YYYYMMDD string (or int) into an ISO week number.
    """
    dt = pd.to_datetime(str(date_str), format='%Y%m%d', errors='coerce')
    if pd.isna(dt):
        raise ValueError(f"Invalid date format: {date_str!r}")
    return dt.isocalendar()[1]

def date_to_year(date_str: str) -> int:
    """
    Convert a YYYYMMDD string (or int) into a four-digit year.
    """
    dt = pd.to_datetime(str(date_str), format='%Y%m%d', errors='coerce')
    if pd.isna(dt):
        raise ValueError(f"Invalid date format: {date_str!r}")
    return dt.year

def generate_weeks_range_monday(start_date: str, end_date: str) -> list[str]:
    """
    Generate a list of week-labels between two YYYYMMDD dates,
    using Mondays as week starts. Labels look like "YY-MM-DD\n(Www)".
    """
    start_dt = pd.to_datetime(str(start_date), format='%Y%m%d', errors='coerce')
    end_dt = pd.to_datetime(str(end_date), format='%Y%m%d', errors='coerce')
    if pd.isna(start_dt) or pd.isna(end_dt):
        return []
    # Roll back to the Monday of the start week
    start_monday = start_dt - pd.Timedelta(days=start_dt.weekday())
    weeks = []
    current = start_monday
    while current <= end_dt:
        label = f"{current.strftime('%y-%m-%d')}\n(W{current.isocalendar()[1]})"
        weeks.append(label)
        current += pd.Timedelta(weeks=1)
    return weeks

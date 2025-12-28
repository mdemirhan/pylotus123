"""Date and time functions for the formula engine.

Implements Lotus 1-2-3 compatible date/time functions:
@DATE, @DATEVALUE, @DAY, @MONTH, @YEAR
@TIME, @TIMEVALUE, @HOUR, @MINUTE, @SECOND
@NOW, @TODAY, @WEEKDAY
"""
from __future__ import annotations

import datetime
from typing import Any


# Lotus 1-2-3 date epoch: Day 1 = January 1, 1900
# Note: Lotus has the infamous leap year bug (treats 1900 as leap year)
EPOCH = datetime.date(1899, 12, 31)


def _serial_to_date(serial: float) -> datetime.date:
    """Convert Lotus serial number to date."""
    # Handle the 1900 leap year bug
    if serial >= 60:
        serial -= 1
    return EPOCH + datetime.timedelta(days=int(serial))


def _date_to_serial(date: datetime.date) -> int:
    """Convert date to Lotus serial number."""
    delta = date - EPOCH
    serial = delta.days
    # Account for leap year bug
    if serial >= 60:
        serial += 1
    return serial


def _serial_to_time(serial: float) -> datetime.time:
    """Convert fractional serial to time."""
    fraction = serial % 1
    total_seconds = int(fraction * 86400)
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return datetime.time(hours, minutes, seconds)


def _time_to_serial(time: datetime.time) -> float:
    """Convert time to fractional serial."""
    total_seconds = time.hour * 3600 + time.minute * 60 + time.second
    return total_seconds / 86400


def _parse_date_string(text: str) -> datetime.date | None:
    """Try to parse a date string."""
    formats = [
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%d-%b-%Y",
        "%d-%b-%y",
        "%B %d, %Y",
        "%d.%m.%Y",
    ]
    for fmt in formats:
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _to_number(value: Any) -> float:
    """Convert value to number."""
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", ""))
    except ValueError:
        return 0.0


# === Date Functions ===

def fn_date(year: Any, month: Any, day: Any) -> int:
    """@DATE - Create date serial number.

    Usage: @DATE(year, month, day)
    Returns serial number representing the date.
    """
    y = int(_to_number(year))
    m = int(_to_number(month))
    d = int(_to_number(day))

    # Handle 2-digit years
    if y < 100:
        y += 1900 if y >= 30 else 2000

    try:
        # Handle month overflow/underflow
        while m < 1:
            y -= 1
            m += 12
        while m > 12:
            y += 1
            m -= 12

        date = datetime.date(y, m, d)
        return _date_to_serial(date)
    except ValueError:
        return 0


def fn_datevalue(date_text: Any) -> int:
    """@DATEVALUE - Convert text to date serial.

    Usage: @DATEVALUE("2023-06-15")
    """
    text = str(date_text).strip()
    date = _parse_date_string(text)
    if date:
        return _date_to_serial(date)
    return 0


def fn_day(serial_number: Any) -> int:
    """@DAY - Day of month from date serial.

    Usage: @DAY(serial_number) or @DAY(@NOW())
    Returns 1-31.
    """
    try:
        date = _serial_to_date(_to_number(serial_number))
        return date.day
    except (ValueError, OverflowError):
        return 0


def fn_month(serial_number: Any) -> int:
    """@MONTH - Month from date serial.

    Usage: @MONTH(serial_number)
    Returns 1-12.
    """
    try:
        date = _serial_to_date(_to_number(serial_number))
        return date.month
    except (ValueError, OverflowError):
        return 0


def fn_year(serial_number: Any) -> int:
    """@YEAR - Year from date serial.

    Usage: @YEAR(serial_number)
    Returns 4-digit year.
    """
    try:
        date = _serial_to_date(_to_number(serial_number))
        return date.year
    except (ValueError, OverflowError):
        return 0


def fn_weekday(serial_number: Any, return_type: Any = 1) -> int:
    """@WEEKDAY - Day of week from date serial.

    Usage: @WEEKDAY(serial_number, return_type)
    return_type: 1 = Sunday=1 to Saturday=7
                 2 = Monday=1 to Sunday=7
                 3 = Monday=0 to Sunday=6
    """
    try:
        date = _serial_to_date(_to_number(serial_number))
        weekday = date.weekday()  # Monday = 0

        rtype = int(_to_number(return_type))
        if rtype == 1:
            # Sunday = 1, Saturday = 7
            return (weekday + 2) % 7 or 7
        elif rtype == 2:
            # Monday = 1, Sunday = 7
            return weekday + 1
        else:  # rtype == 3
            # Monday = 0, Sunday = 6
            return weekday
    except (ValueError, OverflowError):
        return 0


def fn_today() -> int:
    """@TODAY - Current date as serial number."""
    return _date_to_serial(datetime.date.today())


def fn_now() -> float:
    """@NOW - Current date and time as serial number."""
    now = datetime.datetime.now()
    date_serial = _date_to_serial(now.date())
    time_serial = _time_to_serial(now.time())
    return date_serial + time_serial


# === Time Functions ===

def fn_time(hour: Any, minute: Any, second: Any) -> float:
    """@TIME - Create time serial number.

    Usage: @TIME(hour, minute, second)
    Returns fractional serial representing the time.
    """
    h = int(_to_number(hour))
    m = int(_to_number(minute))
    s = int(_to_number(second))

    # Handle overflow
    m += s // 60
    s = s % 60
    h += m // 60
    m = m % 60
    h = h % 24

    time = datetime.time(h, m, s)
    return _time_to_serial(time)


def fn_timevalue(time_text: Any) -> float:
    """@TIMEVALUE - Convert text to time serial.

    Usage: @TIMEVALUE("14:30:00")
    """
    text = str(time_text).strip()

    formats = [
        "%H:%M:%S",
        "%H:%M",
        "%I:%M:%S %p",
        "%I:%M %p",
    ]
    for fmt in formats:
        try:
            time = datetime.datetime.strptime(text, fmt).time()
            return _time_to_serial(time)
        except ValueError:
            continue
    return 0.0


def fn_hour(serial_number: Any) -> int:
    """@HOUR - Hour from time serial.

    Usage: @HOUR(serial_number) or @HOUR(@NOW())
    Returns 0-23.
    """
    try:
        time = _serial_to_time(_to_number(serial_number))
        return time.hour
    except (ValueError, OverflowError):
        return 0


def fn_minute(serial_number: Any) -> int:
    """@MINUTE - Minute from time serial.

    Usage: @MINUTE(serial_number)
    Returns 0-59.
    """
    try:
        time = _serial_to_time(_to_number(serial_number))
        return time.minute
    except (ValueError, OverflowError):
        return 0


def fn_second(serial_number: Any) -> int:
    """@SECOND - Second from time serial.

    Usage: @SECOND(serial_number)
    Returns 0-59.
    """
    try:
        time = _serial_to_time(_to_number(serial_number))
        return time.second
    except (ValueError, OverflowError):
        return 0


# === Date Calculation Functions ===

def fn_days(end_date: Any, start_date: Any) -> int:
    """@DAYS - Days between two dates.

    Usage: @DAYS(end_date, start_date)
    """
    return int(_to_number(end_date) - _to_number(start_date))


def fn_edate(start_date: Any, months: Any) -> int:
    """@EDATE - Date offset by months.

    Usage: @EDATE(start_date, months)
    """
    try:
        date = _serial_to_date(_to_number(start_date))
        m = int(_to_number(months))

        new_month = date.month + m
        new_year = date.year + (new_month - 1) // 12
        new_month = ((new_month - 1) % 12) + 1

        # Handle day overflow
        import calendar
        max_day = calendar.monthrange(new_year, new_month)[1]
        new_day = min(date.day, max_day)

        new_date = datetime.date(new_year, new_month, new_day)
        return _date_to_serial(new_date)
    except (ValueError, OverflowError):
        return 0


def fn_eomonth(start_date: Any, months: Any) -> int:
    """@EOMONTH - Last day of month offset by months.

    Usage: @EOMONTH(start_date, months)
    """
    try:
        date = _serial_to_date(_to_number(start_date))
        m = int(_to_number(months))

        new_month = date.month + m
        new_year = date.year + (new_month - 1) // 12
        new_month = ((new_month - 1) % 12) + 1

        import calendar
        last_day = calendar.monthrange(new_year, new_month)[1]
        new_date = datetime.date(new_year, new_month, last_day)
        return _date_to_serial(new_date)
    except (ValueError, OverflowError):
        return 0


def fn_yearfrac(start_date: Any, end_date: Any, basis: Any = 0) -> float:
    """@YEARFRAC - Fraction of year between dates.

    basis: 0 = 30/360 US, 1 = actual/actual, etc.
    """
    days = abs(int(_to_number(end_date) - _to_number(start_date)))
    b = int(_to_number(basis))

    if b == 0:  # 30/360 US
        return days / 360
    elif b == 1:  # Actual/actual
        return days / 365.25
    elif b == 2:  # Actual/360
        return days / 360
    elif b == 3:  # Actual/365
        return days / 365
    else:  # 30/360 European
        return days / 360


# Function registry for this module
DATETIME_FUNCTIONS = {
    # Date construction
    "DATE": fn_date,
    "DATEVALUE": fn_datevalue,

    # Date extraction
    "DAY": fn_day,
    "MONTH": fn_month,
    "YEAR": fn_year,
    "WEEKDAY": fn_weekday,

    # Current date/time
    "TODAY": fn_today,
    "NOW": fn_now,

    # Time construction
    "TIME": fn_time,
    "TIMEVALUE": fn_timevalue,

    # Time extraction
    "HOUR": fn_hour,
    "MINUTE": fn_minute,
    "SECOND": fn_second,

    # Date calculations
    "DAYS": fn_days,
    "EDATE": fn_edate,
    "EOMONTH": fn_eomonth,
    "YEARFRAC": fn_yearfrac,
}

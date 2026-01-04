"""Formatting system for numeric, date, and time values."""

import datetime
import math
from dataclasses import dataclass
from enum import Enum
from functools import lru_cache
from typing import Any


class FormatCode(Enum):
    """Standard Lotus 1-2-3 format codes."""

    GENERAL = "G"  # Automatic formatting
    FIXED = "F"  # Fixed decimal places (F0-F15)
    SCIENTIFIC = "S"  # Scientific notation (S0-S15)
    CURRENCY = "C"  # Currency format (C0-C15)
    COMMA = ","  # Thousands separator (,0-,15)
    PERCENT = "P"  # Percentage (P0-P15)
    DATE = "D"  # Date formats (D1-D9)
    TIME = "T"  # Time formats (T1-T4)
    HIDDEN = "H"  # Hidden (suppressed display)
    PLUSMINUS = "+"  # Horizontal bar graph
    LABEL = "L"  # Label prefix handling


class DateFormat(Enum):
    """Lotus 1-2-3 date format variants."""

    D1 = "DD-MMM-YY"  # 05-Jun-23
    D2 = "DD-MMM"  # 05-Jun
    D3 = "MMM-YY"  # Jun-23
    D4 = "MM/DD/YY"  # 06/05/23
    D5 = "MM/DD"  # 06/05
    D6 = "DD-MMM-YYYY"  # 05-Jun-2023 (extended)
    D7 = "YYYY-MM-DD"  # 2023-06-05 (ISO)
    D8 = "DD/MM/YY"  # 05/06/23 (European)
    D9 = "DD.MM.YYYY"  # 05.06.2023 (German)


class TimeFormat(Enum):
    """Lotus 1-2-3 time format variants."""

    T1 = "HH:MM:SS AM/PM"  # 02:30:45 PM
    T2 = "HH:MM AM/PM"  # 02:30 PM
    T3 = "HH:MM:SS"  # 14:30:45 (24-hour)
    T4 = "HH:MM"  # 14:30 (24-hour)


# Lotus 1-2-3 date epoch: January 1, 1900
# Note: Lotus has the infamous "1900 leap year bug" where it treats 1900 as leap year
LOTUS_EPOCH = datetime.date(1899, 12, 31)  # Day 0 is actually Dec 31, 1899


@dataclass
class FormatSpec:
    """Parsed format specification.

    Attributes:
        format_type: The base format type
        decimals: Number of decimal places (0-15)
        date_variant: Specific date format if DATE type
        time_variant: Specific time format if TIME type
        currency_symbol: Currency symbol to use
        negative_format: How to display negatives (parentheses, red, etc.)
    """

    format_type: FormatCode
    decimals: int = 2
    date_variant: DateFormat | None = None
    time_variant: TimeFormat | None = None
    currency_symbol: str = "$"
    negative_format: str = "-"  # "-", "()", or "red"


def normalize_format_code(code: str) -> str | None:
    """Normalize and validate a format code.

    Args:
        code: User-entered format code

    Returns:
        Normalized format code, or None if invalid
    """
    code = code.strip().upper()
    if not code:
        return None

    # Single character formats
    if code in ("G", "H", "+"):
        return code

    # Formats with decimal places: F, S, C, P (0-15)
    if code[0] in ("F", "S", "C", "P"):
        if len(code) == 1:
            return code + "2"  # Default to 2 decimal places
        try:
            decimals = int(code[1:])
            if 0 <= decimals <= 15:
                return f"{code[0]}{decimals}"
        except ValueError:
            pass
        return None

    # Comma format (,0-,15)
    if code.startswith(","):
        if len(code) == 1:
            return ",2"  # Default to 2 decimal places
        try:
            decimals = int(code[1:])
            if 0 <= decimals <= 15:
                return f",{decimals}"
        except ValueError:
            pass
        return None

    # Date formats (D1-D9)
    if code[0] == "D":
        if len(code) == 1:
            return "D1"  # Default to D1
        try:
            variant = int(code[1:])
            if 1 <= variant <= 9:
                return f"D{variant}"
        except ValueError:
            pass
        return None

    # Time formats (T1-T4)
    if code[0] == "T":
        if len(code) == 1:
            return "T1"  # Default to T1
        try:
            variant = int(code[1:])
            if 1 <= variant <= 4:
                return f"T{variant}"
        except ValueError:
            pass
        return None

    return None


@lru_cache(maxsize=1024)
def parse_format_code(code: str) -> FormatSpec:
    """Parse a format code string into a FormatSpec.

    Format codes:
        G           - General
        F0-F15      - Fixed with 0-15 decimal places
        S0-S15      - Scientific with 0-15 decimal places
        C0-C15      - Currency with 0-15 decimal places
        ,0-,15      - Comma format with 0-15 decimal places
        P0-P15      - Percent with 0-15 decimal places
        D1-D9       - Date formats
        T1-T4       - Time formats
        H           - Hidden
        +           - Plus/minus bar graph

    Args:
        code: Format code string

    Returns:
        FormatSpec with parsed settings
    """
    code = code.strip().upper()

    if not code or code == "G":
        return FormatSpec(FormatCode.GENERAL)

    if code == "H":
        return FormatSpec(FormatCode.HIDDEN)

    if code == "+":
        return FormatSpec(FormatCode.PLUSMINUS)

    # Parse format with optional decimal places
    if len(code) >= 1:
        fmt_char = code[0]
        decimals = 2  # default

        if len(code) > 1:
            try:
                decimals = int(code[1:])
                decimals = max(0, min(15, decimals))  # Clamp to 0-15
            except ValueError:
                pass

        if fmt_char == "F":
            return FormatSpec(FormatCode.FIXED, decimals=decimals)
        elif fmt_char == "S":
            return FormatSpec(FormatCode.SCIENTIFIC, decimals=decimals)
        elif fmt_char == "C":
            return FormatSpec(FormatCode.CURRENCY, decimals=decimals)
        elif fmt_char == ",":
            return FormatSpec(FormatCode.COMMA, decimals=decimals)
        elif fmt_char == "P":
            return FormatSpec(FormatCode.PERCENT, decimals=decimals)
        elif fmt_char == "D":
            variant = f"D{decimals}" if 1 <= decimals <= 9 else "D1"
            try:
                date_fmt = DateFormat[variant]
            except KeyError:
                date_fmt = DateFormat.D1
            return FormatSpec(FormatCode.DATE, date_variant=date_fmt)
        elif fmt_char == "T":
            variant = f"T{decimals}" if 1 <= decimals <= 4 else "T1"
            try:
                time_fmt = TimeFormat[variant]
            except KeyError:
                time_fmt = TimeFormat.T1
            return FormatSpec(FormatCode.TIME, time_variant=time_fmt)

    return FormatSpec(FormatCode.GENERAL)


def serial_to_date(serial: float) -> datetime.date:
    """Convert Lotus serial number to date.

    Lotus 1-2-3 uses serial numbers where 1 = Jan 1, 1900.

    Args:
        serial: Serial date number

    Returns:
        Python date object
    """
    # Handle the 1900 leap year bug (Lotus thinks Feb 29, 1900 existed)
    if serial >= 60:
        serial -= 1  # Adjust for the bug
    return LOTUS_EPOCH + datetime.timedelta(days=int(serial))


def date_to_serial(date: datetime.date) -> int:
    """Convert date to Lotus serial number.

    Args:
        date: Python date object

    Returns:
        Lotus serial number
    """
    delta = date - LOTUS_EPOCH
    serial = delta.days
    # Add back the leap year bug adjustment
    if serial >= 60:
        serial += 1
    return serial


def serial_to_time(serial: float) -> datetime.time:
    """Convert fractional serial to time.

    The fractional part represents the time of day.
    0.5 = 12:00:00 noon

    Args:
        serial: Serial number (fractional part used)

    Returns:
        Python time object
    """
    fraction = serial % 1
    total_seconds = int(fraction * 86400)
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return datetime.time(hours, minutes, seconds)


def time_to_serial(time: datetime.time) -> float:
    """Convert time to fractional serial.

    Args:
        time: Python time object

    Returns:
        Fractional serial number
    """
    total_seconds = time.hour * 3600 + time.minute * 60 + time.second
    return total_seconds / 86400


def format_value(value: Any, spec: FormatSpec, width: int = 10) -> str:
    """Format a value according to format specification.

    Args:
        value: Value to format
        spec: Format specification
        width: Available width for display

    Returns:
        Formatted string
    """
    if spec.format_type == FormatCode.HIDDEN:
        return ""

    if value is None or value == "":
        return ""

    # Handle error values
    if isinstance(value, str) and value.startswith("#"):
        return value

    # Handle non-numeric values for numeric formats
    numeric_value: float | None = None
    if spec.format_type in (
        FormatCode.FIXED,
        FormatCode.SCIENTIFIC,
        FormatCode.CURRENCY,
        FormatCode.COMMA,
        FormatCode.PERCENT,
        FormatCode.PLUSMINUS,
    ):
        if isinstance(value, (int, float)):
            numeric_value = float(value)
        else:
            try:
                numeric_value = float(value)
            except (ValueError, TypeError):
                return str(value)

    # Format by type
    if spec.format_type == FormatCode.GENERAL:
        return _format_general(value)

    elif spec.format_type == FormatCode.FIXED:
        assert numeric_value is not None
        return _format_fixed(numeric_value, spec.decimals)

    elif spec.format_type == FormatCode.SCIENTIFIC:
        assert numeric_value is not None
        return _format_scientific(numeric_value, spec.decimals)

    elif spec.format_type == FormatCode.CURRENCY:
        assert numeric_value is not None
        return _format_currency(numeric_value, spec.decimals, spec.currency_symbol)

    elif spec.format_type == FormatCode.COMMA:
        assert numeric_value is not None
        return _format_comma(numeric_value, spec.decimals)

    elif spec.format_type == FormatCode.PERCENT:
        assert numeric_value is not None
        return _format_percent(numeric_value, spec.decimals)

    elif spec.format_type == FormatCode.DATE:
        date_value = float(value) if isinstance(value, (int, float)) else 0.0
        return _format_date(date_value, spec.date_variant or DateFormat.D1)

    elif spec.format_type == FormatCode.TIME:
        time_value = float(value) if isinstance(value, (int, float)) else 0.0
        return _format_time(time_value, spec.time_variant or TimeFormat.T1)

    elif spec.format_type == FormatCode.PLUSMINUS:
        assert numeric_value is not None
        return _format_plusminus(numeric_value, width)

    return str(value)


def _format_general(value: Any) -> str:
    """Format value in general (automatic) format."""
    if isinstance(value, float):
        # Handle IEEE 754 special values (infinity, NaN)
        if not math.isfinite(value):
            return str(value)
        if value == int(value):
            return str(int(value))
        # Use up to 10 significant digits
        formatted = f"{value:.10g}"
        return formatted
    return str(value)


def _format_fixed(value: float, decimals: int) -> str:
    """Format as fixed decimal."""
    return f"{value:.{decimals}f}"


def _format_scientific(value: float, decimals: int) -> str:
    """Format in scientific notation."""
    return f"{value:.{decimals}E}"


def _format_currency(value: float, decimals: int, symbol: str) -> str:
    """Format as currency."""
    if value < 0:
        return f"({symbol}{abs(value):,.{decimals}f})"
    return f"{symbol}{value:,.{decimals}f}"


def _format_comma(value: float, decimals: int) -> str:
    """Format with thousands separator."""
    return f"{value:,.{decimals}f}"


def _format_percent(value: float, decimals: int) -> str:
    """Format as percentage."""
    return f"{value * 100:.{decimals}f}%"


def _format_date(value: float, variant: DateFormat) -> str:
    """Format serial number as date."""
    try:
        date = serial_to_date(value)

        if variant == DateFormat.D1:
            return date.strftime("%d-%b-%y").upper()
        elif variant == DateFormat.D2:
            return date.strftime("%d-%b").upper()
        elif variant == DateFormat.D3:
            return date.strftime("%b-%y").upper()
        elif variant == DateFormat.D4:
            return date.strftime("%m/%d/%y")
        elif variant == DateFormat.D5:
            return date.strftime("%m/%d")
        elif variant == DateFormat.D6:
            return date.strftime("%d-%b-%Y").upper()
        elif variant == DateFormat.D7:
            return date.strftime("%Y-%m-%d")
        elif variant == DateFormat.D8:
            return date.strftime("%d/%m/%y")
        elif variant == DateFormat.D9:
            return date.strftime("%d.%m.%Y")
        return str(date)
    except (ValueError, OverflowError):
        return str(value)


def _format_time(value: float, variant: TimeFormat) -> str:
    """Format fractional serial as time."""
    try:
        time = serial_to_time(value)

        if variant == TimeFormat.T1:
            return time.strftime("%I:%M:%S %p")
        elif variant == TimeFormat.T2:
            return time.strftime("%I:%M %p")
        elif variant == TimeFormat.T3:
            return time.strftime("%H:%M:%S")
        elif variant == TimeFormat.T4:
            return time.strftime("%H:%M")
        return str(time)
    except (ValueError, OverflowError):
        return str(value)


def _format_plusminus(value: float, width: int) -> str:
    """Format as horizontal bar graph (+/- characters)."""
    if not isinstance(value, (int, float)):
        return str(value)

    # Handle IEEE 754 special values (infinity, NaN)
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)

    # Scale: typically -10 to +10 maps to full width
    bar_width = width - 1  # Leave room for sign
    max_val = 10

    # Calculate bar length (proportional to value)
    bar_len = int(min(abs(value), max_val) / max_val * bar_width)

    if value >= 0:
        return "+" * bar_len
    else:
        return "-" * bar_len

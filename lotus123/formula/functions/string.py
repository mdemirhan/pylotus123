"""String functions for the formula engine.

Implements Lotus 1-2-3 compatible string functions:
@LEFT, @RIGHT, @MID, @LENGTH, @FIND, @REPLACE
@UPPER, @LOWER, @PROPER, @TRIM, @CLEAN
@VALUE, @STRING, @CHAR, @CODE, @REPEAT, @N, @S
"""

from __future__ import annotations

import re
from typing import Any


def _to_string(value: Any) -> str:
    """Convert value to string."""
    if value is None or value == "":
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value)


def _to_int(value: Any) -> int:
    """Convert value to integer."""
    if isinstance(value, (int, float)):
        return int(value)
    try:
        return int(float(str(value)))
    except (ValueError, TypeError):
        return 0


def fn_left(text: Any, num_chars: Any = 1) -> str:
    """@LEFT - Extract leftmost characters.

    Usage: @LEFT(text, num_chars)
    """
    s = _to_string(text)
    n = max(0, _to_int(num_chars))
    return s[:n]


def fn_right(text: Any, num_chars: Any = 1) -> str:
    """@RIGHT - Extract rightmost characters.

    Usage: @RIGHT(text, num_chars)
    """
    s = _to_string(text)
    n = max(0, _to_int(num_chars))
    if n == 0:
        return ""
    return s[-n:]


def fn_mid(text: Any, start_pos: Any, num_chars: Any) -> str:
    """@MID - Extract substring.

    Usage: @MID(text, start_pos, num_chars)
    Start position is 1-based.
    """
    s = _to_string(text)
    start = max(1, _to_int(start_pos)) - 1  # Convert to 0-based
    n = max(0, _to_int(num_chars))
    return s[start : start + n]


def fn_length(text: Any) -> int:
    """@LENGTH - Length of text.

    Also known as @LEN.
    """
    return len(_to_string(text))


def fn_len(text: Any) -> int:
    """@LEN - Alias for @LENGTH."""
    return fn_length(text)


def fn_find(find_text: Any, within_text: Any, start_pos: Any = 1) -> int:
    """@FIND - Find substring position (case-sensitive).

    Returns 1-based position or 0 if not found.
    Usage: @FIND(find_text, within_text, start_pos)
    """
    find_s = _to_string(find_text)
    within_s = _to_string(within_text)
    start = max(0, _to_int(start_pos) - 1)  # Convert to 0-based

    pos = within_s.find(find_s, start)
    return pos + 1 if pos >= 0 else 0


def fn_search(find_text: Any, within_text: Any, start_pos: Any = 1) -> int:
    """@SEARCH - Find substring position (case-insensitive).

    Returns 1-based position or 0 if not found.
    Supports wildcards: ? matches any single char, * matches any chars.
    """
    find_s = _to_string(find_text).lower()
    within_s = _to_string(within_text).lower()
    start = max(0, _to_int(start_pos) - 1)

    # Convert wildcards to regex
    pattern = find_s.replace("?", ".").replace("*", ".*")
    try:
        match = re.search(pattern, within_s[start:])
        if match:
            return match.start() + start + 1
    except re.error:
        pass
    return 0


def fn_replace(old_text: Any, start_pos: Any, num_chars: Any, new_text: Any) -> str:
    """@REPLACE - Replace characters by position.

    Usage: @REPLACE(old_text, start_pos, num_chars, new_text)
    """
    s = _to_string(old_text)
    start = max(1, _to_int(start_pos)) - 1  # Convert to 0-based
    n = max(0, _to_int(num_chars))
    new_s = _to_string(new_text)

    return s[:start] + new_s + s[start + n :]


def fn_substitute(text: Any, old_text: Any, new_text: Any, instance: Any = None) -> str:
    """@SUBSTITUTE - Replace text occurrences.

    Usage: @SUBSTITUTE(text, old_text, new_text, instance)
    instance: which occurrence to replace (all if omitted)
    """
    s = _to_string(text)
    old_s = _to_string(old_text)
    new_s = _to_string(new_text)

    if instance is None:
        return s.replace(old_s, new_s)

    inst = _to_int(instance)
    if inst < 1:
        return s

    # Replace specific instance
    count = 0
    result = []
    i = 0
    while i < len(s):
        if s[i : i + len(old_s)] == old_s:
            count += 1
            if count == inst:
                result.append(new_s)
                i += len(old_s)
                continue
        result.append(s[i])
        i += 1
    return "".join(result)


def fn_upper(text: Any) -> str:
    """@UPPER - Convert to uppercase."""
    return _to_string(text).upper()


def fn_lower(text: Any) -> str:
    """@LOWER - Convert to lowercase."""
    return _to_string(text).lower()


def fn_proper(text: Any) -> str:
    """@PROPER - Capitalize first letter of each word."""
    return _to_string(text).title()


def fn_trim(text: Any) -> str:
    """@TRIM - Remove leading/trailing spaces and reduce internal spaces."""
    s = _to_string(text)
    # Remove leading/trailing spaces and reduce multiple spaces to single
    return " ".join(s.split())


def fn_clean(text: Any) -> str:
    """@CLEAN - Remove non-printable characters."""
    s = _to_string(text)
    return "".join(c for c in s if c.isprintable() or c in "\t\n")


def fn_value(text: Any) -> float:
    """@VALUE - Convert text to number."""
    s = _to_string(text).strip()
    try:
        # Handle percentage
        if s.endswith("%"):
            return float(s[:-1].replace(",", "")) / 100
        # Handle currency symbols
        s = s.lstrip("$").replace(",", "")
        return float(s)
    except ValueError:
        return 0.0


def fn_string(number: Any, decimals: Any = 0) -> str:
    """@STRING - Convert number to text with fixed decimals.

    Also known as @TEXT in some versions.
    """
    n = float(number) if isinstance(number, (int, float)) else 0.0
    d = max(0, _to_int(decimals))
    return f"{n:.{d}f}"


def fn_text(value: Any, format_text: Any = "") -> str:
    """@TEXT - Format number as text.

    Simplified version - just converts to string.
    """
    return _to_string(value)


def fn_char(number: Any) -> str:
    """@CHAR - Character from ASCII/Unicode code."""
    n = _to_int(number)
    try:
        return chr(n)
    except (ValueError, OverflowError):
        return ""


def fn_code(text: Any) -> int:
    """@CODE - ASCII/Unicode code of first character."""
    s = _to_string(text)
    if s:
        return ord(s[0])
    return 0


def fn_repeat(text: Any, times: Any) -> str:
    """@REPEAT - Repeat text multiple times.

    Also known as @REPT.
    """
    s = _to_string(text)
    n = max(0, _to_int(times))
    return s * n


def fn_rept(text: Any, times: Any) -> str:
    """@REPT - Alias for @REPEAT."""
    return fn_repeat(text, times)


def fn_n(value: Any) -> float:
    """@N - Convert value to number.

    Returns 0 for text, the number for numbers.
    """
    if isinstance(value, (int, float)):
        return float(value)
    return 0.0


def fn_s(value: Any) -> str:
    """@S - Convert value to text.

    Returns empty string for numbers, the text for text.
    """
    if isinstance(value, str):
        return value
    return ""


def fn_exact(text1: Any, text2: Any) -> bool:
    """@EXACT - Case-sensitive comparison."""
    return _to_string(text1) == _to_string(text2)


def fn_concatenate(*args: Any) -> str:
    """@CONCATENATE - Join text values."""
    return "".join(_to_string(arg) for arg in args)


def fn_concat(*args: Any) -> str:
    """@CONCAT - Alias for @CONCATENATE."""
    return fn_concatenate(*args)


def fn_fixed(number: Any, decimals: Any = 2, no_commas: Any = False) -> str:
    """@FIXED - Format number with fixed decimals and optional commas."""
    n = float(number) if isinstance(number, (int, float)) else 0.0
    d = max(0, _to_int(decimals))

    if no_commas:
        return f"{n:.{d}f}"
    else:
        return f"{n:,.{d}f}"


def fn_dollar(number: Any, decimals: Any = 2) -> str:
    """@DOLLAR - Format as currency."""
    n = float(number) if isinstance(number, (int, float)) else 0.0
    d = max(0, _to_int(decimals))
    return f"${n:,.{d}f}"


def fn_t(value: Any) -> str:
    """@T - Return text if value is text, empty string otherwise."""
    if isinstance(value, str):
        return value
    return ""


# Function registry for this module
STRING_FUNCTIONS = {
    # Extraction
    "LEFT": fn_left,
    "RIGHT": fn_right,
    "MID": fn_mid,
    # Length
    "LENGTH": fn_length,
    "LEN": fn_len,
    # Search
    "FIND": fn_find,
    "SEARCH": fn_search,
    # Replacement
    "REPLACE": fn_replace,
    "SUBSTITUTE": fn_substitute,
    # Case conversion
    "UPPER": fn_upper,
    "LOWER": fn_lower,
    "PROPER": fn_proper,
    # Cleaning
    "TRIM": fn_trim,
    "CLEAN": fn_clean,
    # Conversion
    "VALUE": fn_value,
    "STRING": fn_string,
    "TEXT": fn_text,
    "CHAR": fn_char,
    "CODE": fn_code,
    "N": fn_n,
    "S": fn_s,
    "T": fn_t,
    # Repetition
    "REPEAT": fn_repeat,
    "REPT": fn_rept,
    # Comparison
    "EXACT": fn_exact,
    # Concatenation
    "CONCATENATE": fn_concatenate,
    "CONCAT": fn_concat,
    # Formatting
    "FIXED": fn_fixed,
    "DOLLAR": fn_dollar,
}

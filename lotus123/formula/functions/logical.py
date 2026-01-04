"""Logical functions for the formula engine.

Implements Lotus 1-2-3 compatible logical functions:
@IF, @TRUE, @FALSE, @AND, @OR, @NOT
@ISERR, @ISNA, @ISNUMBER, @ISSTRING
"""

from typing import Any

from ...core.errors import FormulaError


def _flatten_args(args: tuple) -> list:
    """Flatten nested lists in arguments."""
    result = []
    for arg in args:
        if isinstance(arg, list):
            result.extend(_flatten_args(tuple(arg)))
        else:
            result.append(arg)
    return result


def _to_bool(value: Any) -> bool:
    """Convert value to boolean.

    In Lotus 1-2-3:
    - 0 is False
    - Any non-zero number is True
    - Empty string is False
    - Non-empty string is True (in some contexts)
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        if value.upper() == "TRUE":
            return True
        if value.upper() == "FALSE":
            return False
        # Try as number
        try:
            return float(value) != 0
        except ValueError:
            return bool(value)  # Non-empty string is truthy
    return False


def fn_if(condition: Any, true_value: Any, false_value: Any = "") -> Any:
    """@IF - Conditional logic.

    Usage: @IF(condition, true_value, false_value)
    Returns true_value if condition is true, false_value otherwise.
    """
    if _to_bool(condition):
        return true_value
    return false_value


def fn_true() -> bool:
    """@TRUE - Returns logical TRUE."""
    return True


def fn_false() -> bool:
    """@FALSE - Returns logical FALSE."""
    return False


def fn_and(*args) -> bool:
    """@AND - Logical AND.

    Returns TRUE if all arguments are true.
    """
    values = _flatten_args(args)
    if not values:
        return True
    return all(_to_bool(v) for v in values)


def fn_or(*args) -> bool:
    """@OR - Logical OR.

    Returns TRUE if any argument is true.
    """
    values = _flatten_args(args)
    if not values:
        return False
    return any(_to_bool(v) for v in values)


def fn_not(value: Any) -> bool:
    """@NOT - Logical NOT.

    Returns opposite of argument.
    """
    return not _to_bool(value)


def fn_xor(*args: Any) -> bool:
    """@XOR - Logical exclusive OR.

    Returns TRUE if an odd number of arguments are true.
    """
    values = _flatten_args(args)
    count = sum(1 for v in values if _to_bool(v))
    return count % 2 == 1


def fn_iserr(value: Any) -> bool:
    """@ISERR - Check if value is an error (except #N/A).

    Returns TRUE for errors like #DIV/0!, #ERR!, #CIRC!, etc.
    """
    if isinstance(value, str) and value.startswith("#"):
        return value != FormulaError.NA
    return False


def fn_iserror(value: Any) -> bool:
    """@ISERROR - Check if value is any error (including #N/A)."""
    return isinstance(value, str) and value.startswith("#")


def fn_isna(value: Any) -> bool:
    """@ISNA - Check if value is #N/A error."""
    return bool(value == FormulaError.NA)


def fn_isnumber(value: Any) -> bool:
    """@ISNUMBER - Check if value is numeric."""
    if isinstance(value, bool):
        return False  # Booleans are not numbers in this context
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        try:
            float(value.replace(",", ""))
            return True
        except ValueError:
            return False
    return False


def fn_isstring(value: Any) -> bool:
    """@ISSTRING - Check if value is text.

    Also known as @ISTEXT.
    """
    if isinstance(value, str):
        # Check it's not an error or number
        if value.startswith("#"):
            return False
        try:
            float(value.replace(",", ""))
            return False  # It's a number formatted as text
        except ValueError:
            return True
    return False


def fn_istext(value: Any) -> bool:
    """@ISTEXT - Alias for @ISSTRING."""
    return fn_isstring(value)


def fn_isblank(value: Any) -> bool:
    """@ISBLANK - Check if cell is empty."""
    return value is None or value == ""


def fn_islogical(value: Any) -> bool:
    """@ISLOGICAL - Check if value is a boolean."""
    return isinstance(value, bool)


def fn_iseven(value: Any) -> bool:
    """@ISEVEN - Check if number is even."""
    try:
        n = int(float(value))
        return n % 2 == 0
    except (ValueError, TypeError):
        return False


def fn_isodd(value: Any) -> bool:
    """@ISODD - Check if number is odd."""
    try:
        n = int(float(value))
        return n % 2 != 0
    except (ValueError, TypeError):
        return False


def fn_isref(value: Any) -> bool:
    """@ISREF - Check if value is a cell reference.

    Note: This is difficult to implement properly in the evaluator
    as references are resolved before reaching the function.
    """
    return False  # Placeholder


def fn_na() -> str:
    """@NA - Return #N/A error value."""
    return FormulaError.NA


def fn_err() -> str:
    """@ERR - Return #ERR! error value."""
    return FormulaError.ERR


def fn_iferror(value: Any, value_if_error: Any) -> Any:
    """@IFERROR - Return alternate value if error.

    Usage: @IFERROR(value, value_if_error)
    """
    if isinstance(value, str) and value.startswith("#"):
        return value_if_error
    return value


def fn_ifna(value: Any, value_if_na: Any) -> Any:
    """@IFNA - Return alternate value if #N/A.

    Usage: @IFNA(value, value_if_na)
    """
    if value == FormulaError.NA:
        return value_if_na
    return value


def fn_switch(expression: Any, *args: Any) -> Any:
    """@SWITCH - Match value against list of cases.

    Usage: @SWITCH(expression, value1, result1, value2, result2, ..., default)
    """
    pairs = list(args)

    # Check for default value (odd number of remaining args)
    default: Any = "" if len(pairs) % 2 == 0 else pairs.pop()

    # Check each value/result pair
    for i in range(0, len(pairs), 2):
        if expression == pairs[i]:
            return pairs[i + 1]

    return default


def fn_choose(index: Any, *values: Any) -> Any:
    """@CHOOSE - Select from list by index.

    Usage: @CHOOSE(index, value1, value2, ...)
    Index is 1-based.
    """
    try:
        idx = int(float(index))
        if 1 <= idx <= len(values):
            return values[idx - 1]
    except (ValueError, TypeError):
        pass
    return FormulaError.NA


# Function registry for this module
LOGICAL_FUNCTIONS = {
    # Core logical
    "IF": fn_if,
    "TRUE": fn_true,
    "FALSE": fn_false,
    "AND": fn_and,
    "OR": fn_or,
    "NOT": fn_not,
    "XOR": fn_xor,
    # Error checking
    "ISERR": fn_iserr,
    "ISERROR": fn_iserror,
    "ISNA": fn_isna,
    "NA": fn_na,
    "ERR": fn_err,
    # Type checking
    "ISNUMBER": fn_isnumber,
    "ISSTRING": fn_isstring,
    "ISTEXT": fn_istext,
    "ISBLANK": fn_isblank,
    "ISLOGICAL": fn_islogical,
    "ISEVEN": fn_iseven,
    "ISODD": fn_isodd,
    "ISREF": fn_isref,
    # Conditional
    "IFERROR": fn_iferror,
    "IFNA": fn_ifna,
    "SWITCH": fn_switch,
    "CHOOSE": fn_choose,
}

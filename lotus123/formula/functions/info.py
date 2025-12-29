"""Information functions for the formula engine.

Implements Lotus 1-2-3 compatible information functions:
@CELL, @CELLPOINTER, @ROWS, @COLS
@ISNUMBER, @ISSTRING, @ISERR, @ISNA, @TYPE
"""

from __future__ import annotations

from typing import Any


def fn_type(value: Any) -> int:
    """@TYPE - Return type number of value.

    Returns:
        1 = Number
        2 = Text
        4 = Logical (boolean)
        16 = Error
        64 = Array
    """
    if isinstance(value, bool):
        return 4
    if isinstance(value, (int, float)):
        return 1
    if isinstance(value, str):
        if value.startswith("#"):
            return 16  # Error
        return 2  # Text
    if isinstance(value, list):
        return 64
    return 1  # Default to number


def fn_cell(info_type: Any, reference: Any = None) -> Any:
    """@CELL - Return information about a cell.

    Info types:
        "address" - Cell address
        "col" - Column number
        "row" - Row number
        "contents" - Cell contents
        "type" - "b" (blank), "l" (label), "v" (value)
        "width" - Column width
        "format" - Format code
        "protect" - Protection status (0 or 1)
        "prefix" - Label prefix character

    Note: Without spreadsheet context, returns placeholder values.
    """
    info = str(info_type).lower().strip('"')

    if info == "address":
        return "$A$1"
    elif info == "col":
        return 1
    elif info == "row":
        return 1
    elif info == "contents":
        return reference if reference else ""
    elif info == "type":
        if reference is None or reference == "":
            return "b"  # blank
        elif isinstance(reference, (int, float)):
            return "v"  # value
        else:
            return "l"  # label
    elif info == "width":
        return 9  # Default width
    elif info == "format":
        return "G"  # General
    elif info == "protect":
        return 0
    elif info == "prefix":
        return "'"  # Default left-align
    else:
        return ""


def fn_cellpointer(attribute: Any = "contents") -> Any:
    """@CELLPOINTER - Return info about current cell.

    Similar to @CELL but for the cell containing the formula.
    Note: Without context, returns placeholder.
    """
    return fn_cell(attribute)


def fn_info(type_text: Any) -> Any:
    """@INFO - Return system information.

    Types:
        "directory" - Current directory
        "numfile" - Number of active worksheets
        "origin" - Top-left visible cell
        "osversion" - OS version
        "recalc" - Recalculation mode
        "release" - Lotus version
        "system" - Operating system
        "totmem" - Total memory
        "usedmem" - Used memory
    """
    import os
    import sys

    info = str(type_text).lower().strip('"')

    if info == "directory":
        return os.getcwd()
    elif info == "numfile":
        return 1
    elif info == "origin":
        return "$A:$A$1"
    elif info == "osversion":
        return sys.platform
    elif info == "recalc":
        return "Automatic"
    elif info == "release":
        return "1.0"
    elif info == "system":
        return sys.platform
    elif info == "totmem":
        return 1000000
    elif info == "usedmem":
        return 100000
    else:
        return ""


def fn_error_type(error_val: Any) -> int:
    """@ERROR.TYPE - Return error type number.

    Returns:
        1 = #NULL!
        2 = #DIV/0!
        3 = #VALUE!
        4 = #REF!
        5 = #NAME?
        6 = #NUM!
        7 = #N/A
        8 = #CIRC!
    """
    if not isinstance(error_val, str) or not error_val.startswith("#"):
        return 0

    error_map = {
        "#NULL!": 1,
        "#DIV/0!": 2,
        "#VALUE!": 3,
        "#REF!": 4,
        "#NAME?": 5,
        "#NUM!": 6,
        "#N/A": 7,
        "#CIRC!": 8,
        "#ERR!": 3,
    }

    # Check for partial matches
    for err, num in error_map.items():
        if error_val.startswith(err.rstrip("!")):
            return num
    return 0


def fn_sheet(value: Any = None) -> int:
    """@SHEET - Return sheet number.

    Note: Single-sheet implementation returns 1.
    """
    return 1


def fn_sheets(reference: Any = None) -> int:
    """@SHEETS - Return number of sheets.

    Note: Single-sheet implementation returns 1.
    """
    return 1


def fn_areas(reference: Any) -> int:
    """@AREAS - Return number of areas in reference.

    Note: Simple implementation returns 1.
    """
    return 1


def fn_isformula(reference: Any) -> bool:
    """@ISFORMULA - Check if cell contains formula.

    Note: Without cell context, checks if value looks like a formula.
    """
    if isinstance(reference, str):
        return reference.startswith("=") or reference.startswith("+")
    return False


def fn_n(value: Any) -> float:
    """@N - Convert to number.

    Returns 0 for text, error values.
    Returns number for numbers.
    Returns serial for dates.
    """
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    return 0.0


def fn_version() -> str:
    """@VERSION - Return version string."""
    return "Lotus123 Clone 1.0"


# Function registry for this module
INFO_FUNCTIONS = {
    # Type checking (also in logical.py, but included here for completeness)
    "TYPE": fn_type,
    # Cell information
    "CELL": fn_cell,
    "CELLPOINTER": fn_cellpointer,
    # System information
    "INFO": fn_info,
    "VERSION": fn_version,
    # Error handling
    "ERROR.TYPE": fn_error_type,
    # Sheet information
    "SHEET": fn_sheet,
    "SHEETS": fn_sheets,
    "AREAS": fn_areas,
    # Formula checking
    "ISFORMULA": fn_isformula,
    # Conversion
    "N": fn_n,
}

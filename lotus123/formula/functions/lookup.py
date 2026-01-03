"""Lookup and reference functions for the formula engine.

Implements Lotus 1-2-3 compatible lookup functions:
@VLOOKUP, @HLOOKUP, @INDEX, @CHOOSE
@CELL, @CELLPOINTER, @COLS, @ROWS
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from ...core.errors import FormulaError

if TYPE_CHECKING:
    pass


# These functions need access to the spreadsheet for range operations
# They're called with the spreadsheet as the first hidden argument


def _try_float(value: Any) -> float | None:
    """Try to parse a value as float."""
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def _compare_for_sort(a: Any, b: Any) -> int:
    """Compare two values for sorting (numeric if possible, else string)."""
    a_num = _try_float(a)
    b_num = _try_float(b)
    if a_num is not None and b_num is not None:
        return (a_num > b_num) - (a_num < b_num)
    a_str = str(a).upper()
    b_str = str(b).upper()
    return (a_str > b_str) - (a_str < b_str)


def _is_sorted(values: list[Any]) -> bool:
    """Check if values are sorted ascending for range lookups."""
    if len(values) < 2:
        return True
    last = values[0]
    for value in values[1:]:
        if _compare_for_sort(last, value) > 0:
            return False
        last = value
    return True


def _is_match(lookup_value: Any, table_value: Any, range_lookup: bool = True) -> bool:
    """Check if values match for lookup.

    Args:
        lookup_value: Value to find
        table_value: Value in table
        range_lookup: If True, use range matching (<=). If False, exact match.
    """
    if range_lookup:
        # Range lookup - find largest value <= lookup_value
        table_num = _try_float(table_value)
        lookup_num = _try_float(lookup_value)
        if table_num is not None and lookup_num is not None:
            return table_num <= lookup_num
        return str(table_value).upper() <= str(lookup_value).upper()
    else:
        # Exact match
        if isinstance(lookup_value, str) and isinstance(table_value, str):
            return lookup_value.upper() == table_value.upper()
        return bool(lookup_value == table_value)


def fn_vlookup(
    lookup_value: Any, table: list[Any], col_index: Any, range_lookup: Any = True
) -> Any:
    """@VLOOKUP - Vertical lookup.

    Searches first column of table for lookup_value, returns value from col_index.

    Usage: @VLOOKUP(lookup_value, table_range, col_index, range_lookup)

    Args:
        lookup_value: Value to search for
        table: 2D list of values (the table range)
        col_index: Column to return (1-based)
    range_lookup: True for approximate match (default), False for exact.
                  Approximate match requires a sorted first column.
    """
    if not isinstance(table, list) or not table:
        return FormulaError.NA

    col_idx = int(col_index) - 1  # Convert to 0-based
    range_match = range_lookup if isinstance(range_lookup, bool) else bool(range_lookup)

    # Ensure table is 2D
    if not isinstance(table[0], list):
        table = [[v] for v in table]

    if col_idx < 0 or (table and col_idx >= len(table[0])):
        return FormulaError.REF

    if range_match:
        first_col = [
            row[0] if isinstance(row, list) else row
            for row in table
            if row is not None
        ]
        if not _is_sorted(first_col):
            return FormulaError.NA

    last_match_row = None

    for row_idx, row in enumerate(table):
        if not row:
            continue

        cell_value = row[0] if isinstance(row, list) else row

        if range_match:
            if _is_match(lookup_value, cell_value, True):
                last_match_row = row_idx
        else:
            if _is_match(lookup_value, cell_value, False):
                if isinstance(row, list) and col_idx < len(row):
                    return row[col_idx]
                return FormulaError.REF

    if range_match and last_match_row is not None:
        row = table[last_match_row]
        if isinstance(row, list) and col_idx < len(row):
            return row[col_idx]
        return FormulaError.REF

    return FormulaError.NA


def fn_hlookup(
    lookup_value: Any, table: list[Any], row_index: Any, range_lookup: Any = True
) -> Any:
    """@HLOOKUP - Horizontal lookup.

    Searches first row of table for lookup_value, returns value from row_index.

    Usage: @HLOOKUP(lookup_value, table_range, row_index, range_lookup)
    Approximate match requires a sorted first row.
    """
    if not isinstance(table, list) or not table:
        return FormulaError.NA

    row_idx = int(row_index) - 1  # Convert to 0-based
    range_match = range_lookup if isinstance(range_lookup, bool) else bool(range_lookup)

    # Ensure table is 2D
    if not isinstance(table[0], list):
        table = [table]  # Make it a single row

    if row_idx < 0 or row_idx >= len(table):
        return FormulaError.REF

    first_row = table[0]
    if range_match and not _is_sorted(first_row):
        return FormulaError.NA
    last_match_col = None

    for col_idx, cell_value in enumerate(first_row):
        if range_match:
            if _is_match(lookup_value, cell_value, True):
                last_match_col = col_idx
        else:
            if _is_match(lookup_value, cell_value, False):
                if row_idx < len(table) and col_idx < len(table[row_idx]):
                    return table[row_idx][col_idx]
                return FormulaError.REF

    if range_match and last_match_col is not None:
        if row_idx < len(table) and last_match_col < len(table[row_idx]):
            return table[row_idx][last_match_col]
        return FormulaError.REF

    return FormulaError.NA


def fn_index(array: Any, row_num: Any, col_num: Any = None) -> Any:
    """@INDEX - Return value at position in array.

    Usage: @INDEX(array, row_num, col_num)
    row_num and col_num are 1-based.
    """
    if not isinstance(array, list):
        return array if row_num == 1 else FormulaError.REF

    row_idx = int(row_num) - 1

    # Handle 1D array
    if not array or not isinstance(array[0], list):
        if 0 <= row_idx < len(array):
            return array[row_idx]
        return FormulaError.REF

    # Handle 2D array
    if row_idx < 0 or row_idx >= len(array):
        return FormulaError.REF

    if col_num is None:
        return array[row_idx]

    col_idx = int(col_num) - 1
    row = array[row_idx]
    if col_idx < 0 or col_idx >= len(row):
        return FormulaError.REF

    return row[col_idx]


def fn_match(lookup_value: Any, lookup_array: list[Any], match_type: Any = 1) -> int:
    """@MATCH - Find position of value in array.

    Usage: @MATCH(lookup_value, lookup_array, match_type)
    match_type: 1 = largest value <= (sorted ascending)
                0 = exact match
               -1 = smallest value >= (sorted descending)

    Returns 1-based position.
    """
    if not isinstance(lookup_array, list):
        return 0

    # Flatten if nested
    flat = []
    for item in lookup_array:
        if isinstance(item, list):
            flat.extend(item)
        else:
            flat.append(item)

    mtype = int(match_type) if match_type is not None else 1

    if mtype == 0:
        # Exact match
        for i, val in enumerate(flat):
            if _is_match(lookup_value, val, False):
                return i + 1
        return 0

    elif mtype == 1:
        # Largest <= lookup_value (ascending order)
        last_match = None
        for i, val in enumerate(flat):
            try:
                if float(val) <= float(lookup_value):
                    last_match = i + 1
                else:
                    break
            except (ValueError, TypeError):
                if str(val).upper() <= str(lookup_value).upper():
                    last_match = i + 1
                else:
                    break
        return last_match or 0

    else:  # mtype == -1
        # Smallest >= lookup_value (descending order)
        last_match = None
        for i, val in enumerate(flat):
            try:
                if float(val) >= float(lookup_value):
                    last_match = i + 1
                else:
                    break
            except (ValueError, TypeError):
                if str(val).upper() >= str(lookup_value).upper():
                    last_match = i + 1
                else:
                    break
        return last_match or 0


def fn_lookup(
    lookup_value: Any, lookup_vector: list[Any], result_vector: list[Any] | None = None
) -> Any:
    """@LOOKUP - Simple lookup.

    Usage: @LOOKUP(lookup_value, lookup_vector, result_vector)
    If result_vector is omitted, returns value from lookup_vector.
    """
    if not isinstance(lookup_vector, list):
        return FormulaError.NA

    # Flatten vectors
    lv = []
    for item in lookup_vector:
        if isinstance(item, list):
            lv.extend(item)
        else:
            lv.append(item)

    if result_vector is None:
        rv = lv
    else:
        rv = []
        for item in result_vector:
            if isinstance(item, list):
                rv.extend(item)
            else:
                rv.append(item)

    # Find last value <= lookup_value
    last_match = None
    for i, val in enumerate(lv):
        try:
            if float(val) <= float(lookup_value):
                last_match = i
        except (ValueError, TypeError):
            if str(val).upper() <= str(lookup_value).upper():
                last_match = i

    if last_match is not None and last_match < len(rv):
        return rv[last_match]
    return FormulaError.NA


def fn_rows(array: Any) -> int:
    """@ROWS - Number of rows in a range."""
    if not isinstance(array, list):
        return 1
    return len(array)


def fn_cols(array: Any) -> int:
    """@COLS - Number of columns in a range.

    Also known as @COLUMNS.
    """
    if not isinstance(array, list) or not array:
        return 1
    first = array[0]
    if isinstance(first, list):
        return len(first)
    return len(array)  # 1D horizontal array


def fn_columns(array: Any) -> int:
    """@COLUMNS - Alias for @COLS."""
    return fn_cols(array)


def fn_transpose(array: Any) -> list[Any]:
    """@TRANSPOSE - Transpose rows and columns."""
    if not isinstance(array, list):
        return [[array]]

    if not array:
        return [[]]

    # Handle 1D array
    if not isinstance(array[0], list):
        return [[v] for v in array]

    # Transpose 2D array
    rows = len(array)
    cols = max(len(row) for row in array)

    result = []
    for c in range(cols):
        new_row = []
        for r in range(rows):
            if c < len(array[r]):
                new_row.append(array[r][c])
            else:
                new_row.append("")
        result.append(new_row)
    return result


def fn_offset(reference: Any, rows: Any, cols: Any, height: Any = None, width: Any = None) -> Any:
    """@OFFSET - Reference offset from starting point.

    Note: This is a reference function that requires spreadsheet context.
    Returns a placeholder in this simple implementation.
    """
    # This would need spreadsheet context to implement properly
    return FormulaError.REF


def fn_indirect(ref_text: Any) -> Any:
    """@INDIRECT - Reference from text string.

    Note: This is a reference function that requires spreadsheet context.
    """
    return FormulaError.REF


def fn_row(reference: Any = None) -> int:
    """@ROW - Row number of reference.

    Note: Without context, returns 1.
    """
    return 1


def fn_column(reference: Any = None) -> int:
    """@COLUMN - Column number of reference.

    Note: Without context, returns 1.
    """
    return 1


def fn_address(row: Any, col: Any, abs_type: Any = 1, a1_style: Any = True) -> str:
    """@ADDRESS - Create cell reference text.

    abs_type: 1 = $A$1, 2 = A$1, 3 = $A1, 4 = A1
    """
    from ...core.reference import index_to_col

    r = int(row)
    c = int(col) - 1  # Convert to 0-based for index_to_col

    col_letter = index_to_col(c)
    abs_t = int(abs_type) if abs_type else 1

    if abs_t == 1:
        return f"${col_letter}${r}"
    elif abs_t == 2:
        return f"{col_letter}${r}"
    elif abs_t == 3:
        return f"${col_letter}{r}"
    else:
        return f"{col_letter}{r}"


# Function registry for this module
LOOKUP_FUNCTIONS = {
    # Lookup
    "VLOOKUP": fn_vlookup,
    "HLOOKUP": fn_hlookup,
    "LOOKUP": fn_lookup,
    "MATCH": fn_match,
    # Reference
    "INDEX": fn_index,
    "OFFSET": fn_offset,
    "INDIRECT": fn_indirect,
    "ROW": fn_row,
    "COLUMN": fn_column,
    "ADDRESS": fn_address,
    # Array info
    "ROWS": fn_rows,
    "COLS": fn_cols,
    "COLUMNS": fn_columns,
    "TRANSPOSE": fn_transpose,
}

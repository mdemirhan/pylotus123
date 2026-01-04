"""Database statistical functions for the formula engine.

Implements Lotus 1-2-3 compatible database functions:
@DSUM, @DAVG, @DCOUNT, @DMIN, @DMAX, @DSTD, @DVAR

Database functions operate on a table with headers and apply criteria
to filter rows before calculating statistics.
"""

import math
from typing import Any

from ...core.errors import FormulaError


def _to_number(value: Any) -> float | None:
    """Convert value to number, returning None for non-numeric."""
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.replace(",", ""))
        except ValueError:
            return None
    return None


def _get_field_index(database: list[Any], field: Any) -> int | None:
    """Get the column index for a field.

    field can be:
    - A number (1-based column index)
    - A string (column header name)
    """
    if not database:
        return None

    headers = database[0] if isinstance(database[0], list) else [database[0]]

    if isinstance(field, (int, float)):
        idx = int(field) - 1  # Convert to 0-based
        if 0 <= idx < len(headers):
            return idx
        return None

    # String field name
    field_str = str(field).upper()
    for i, header in enumerate(headers):
        if str(header).upper() == field_str:
            return i
    return None


def _matches_criteria(row: list[Any], headers: list[Any], criteria: list[Any]) -> bool:
    """Check if a row matches the criteria.

    Criteria format:
    - First row is headers
    - Subsequent rows are OR conditions (any row matches)
    - Within a row, conditions are AND (all must match)
    """
    if not criteria or len(criteria) < 2:
        return True  # No criteria, all rows match

    criteria_headers = criteria[0] if isinstance(criteria[0], list) else [criteria[0]]

    # Check each criteria row (OR between rows)
    for crit_row in criteria[1:]:
        if not isinstance(crit_row, list):
            crit_row = [crit_row]

        row_matches = True
        has_condition = False

        # Check each condition in the row (AND within row)
        for i, crit_value in enumerate(crit_row):
            if crit_value == "" or crit_value is None:
                continue  # Empty criteria, skip

            has_condition = True

            if i >= len(criteria_headers):
                continue

            # Find matching column in database
            crit_header = str(criteria_headers[i]).upper()
            col_idx = None
            for j, header in enumerate(headers):
                if str(header).upper() == crit_header:
                    col_idx = j
                    break

            if col_idx is None or col_idx >= len(row):
                row_matches = False
                break

            cell_value = row[col_idx]
            crit_str = str(crit_value)

            # Handle comparison operators
            if crit_str.startswith(">="):
                try:
                    if float(cell_value) < float(crit_str[2:]):
                        row_matches = False
                        break
                except (ValueError, TypeError):
                    row_matches = False
                    break
            elif crit_str.startswith("<="):
                try:
                    if float(cell_value) > float(crit_str[2:]):
                        row_matches = False
                        break
                except (ValueError, TypeError):
                    row_matches = False
                    break
            elif crit_str.startswith("<>") or crit_str.startswith("!="):
                if str(cell_value).upper() == crit_str[2:].upper():
                    row_matches = False
                    break
            elif crit_str.startswith(">"):
                try:
                    if float(cell_value) <= float(crit_str[1:]):
                        row_matches = False
                        break
                except (ValueError, TypeError):
                    row_matches = False
                    break
            elif crit_str.startswith("<"):
                try:
                    if float(cell_value) >= float(crit_str[1:]):
                        row_matches = False
                        break
                except (ValueError, TypeError):
                    row_matches = False
                    break
            elif crit_str.startswith("="):
                if str(cell_value).upper() != crit_str[1:].upper():
                    row_matches = False
                    break
            else:
                # Exact match (case-insensitive) or wildcard
                if "*" in crit_str or "?" in crit_str:
                    # Wildcard matching
                    import fnmatch

                    if not fnmatch.fnmatch(str(cell_value).upper(), crit_str.upper()):
                        row_matches = False
                        break
                else:
                    if str(cell_value).upper() != crit_str.upper():
                        row_matches = False
                        break

        if has_condition and row_matches:
            return True

    return not any(
        any(c != "" and c is not None for c in (r if isinstance(r, list) else [r]))
        for r in criteria[1:]
    )


def _get_matching_values(database: list[Any], field: Any, criteria: list[Any]) -> list[float]:
    """Get numeric values from matching rows."""
    if not database or len(database) < 2:
        return []

    headers = database[0] if isinstance(database[0], list) else [database[0]]
    field_idx = _get_field_index(database, field)

    if field_idx is None:
        return []

    values = []
    for row in database[1:]:
        if not isinstance(row, list):
            row = [row]

        if _matches_criteria(row, headers, criteria):
            if field_idx < len(row):
                num = _to_number(row[field_idx])
                if num is not None:
                    values.append(num)

    return values


def fn_dsum(database: Any, field: Any, criteria: Any) -> float:
    """@DSUM - Sum of matching database records.

    Usage: @DSUM(database, field, criteria)
    """
    if not isinstance(database, list):
        return 0.0

    crit = criteria if isinstance(criteria, list) else []
    values = _get_matching_values(database, field, crit)
    return sum(values)


def fn_davg(database: Any, field: Any, criteria: Any) -> float:
    """@DAVG - Average of matching database records.

    Usage: @DAVG(database, field, criteria)
    Also known as @DAVERAGE.
    """
    if not isinstance(database, list):
        return 0.0

    crit = criteria if isinstance(criteria, list) else []
    values = _get_matching_values(database, field, crit)

    if not values:
        return 0.0
    return sum(values) / len(values)


def fn_dcount(database: Any, field: Any, criteria: Any) -> int:
    """@DCOUNT - Count of matching database records with numbers.

    Usage: @DCOUNT(database, field, criteria)
    """
    if not isinstance(database, list):
        return 0

    crit = criteria if isinstance(criteria, list) else []
    values = _get_matching_values(database, field, crit)
    return len(values)


def fn_dcounta(database: Any, field: Any, criteria: Any) -> int:
    """@DCOUNTA - Count of matching non-blank records.

    Usage: @DCOUNTA(database, field, criteria)
    """
    if not isinstance(database, list) or len(database) < 2:
        return 0

    headers = database[0] if isinstance(database[0], list) else [database[0]]
    field_idx = _get_field_index(database, field)

    if field_idx is None:
        return 0

    crit = criteria if isinstance(criteria, list) else []
    count = 0

    for row in database[1:]:
        if not isinstance(row, list):
            row = [row]

        if _matches_criteria(row, headers, crit):
            if field_idx < len(row):
                val = row[field_idx]
                if val != "" and val is not None:
                    count += 1

    return count


def fn_dmin(database: Any, field: Any, criteria: Any) -> float:
    """@DMIN - Minimum of matching database records.

    Usage: @DMIN(database, field, criteria)
    """
    if not isinstance(database, list):
        return 0.0

    crit = criteria if isinstance(criteria, list) else []
    values = _get_matching_values(database, field, crit)

    if not values:
        return 0.0
    return min(values)


def fn_dmax(database: Any, field: Any, criteria: Any) -> float:
    """@DMAX - Maximum of matching database records.

    Usage: @DMAX(database, field, criteria)
    """
    if not isinstance(database, list):
        return 0.0

    crit = criteria if isinstance(criteria, list) else []
    values = _get_matching_values(database, field, crit)

    if not values:
        return 0.0
    return max(values)


def fn_dstd(database: Any, field: Any, criteria: Any) -> float:
    """@DSTD - Sample standard deviation of matching records.

    Usage: @DSTD(database, field, criteria)
    Also known as @DSTDEV.
    """
    if not isinstance(database, list):
        return 0.0

    crit = criteria if isinstance(criteria, list) else []
    values = _get_matching_values(database, field, crit)

    if len(values) < 2:
        return 0.0

    mean = sum(values) / len(values)
    variance = sum((x - mean) ** 2 for x in values) / (len(values) - 1)
    return math.sqrt(variance)


def fn_dstdp(database: Any, field: Any, criteria: Any) -> float:
    """@DSTDP - Population standard deviation of matching records.

    Usage: @DSTDP(database, field, criteria)
    """
    if not isinstance(database, list):
        return 0.0

    crit = criteria if isinstance(criteria, list) else []
    values = _get_matching_values(database, field, crit)

    if not values:
        return 0.0

    mean = sum(values) / len(values)
    variance = sum((x - mean) ** 2 for x in values) / len(values)
    return math.sqrt(variance)


def fn_dvar(database: Any, field: Any, criteria: Any) -> float:
    """@DVAR - Sample variance of matching records.

    Usage: @DVAR(database, field, criteria)
    """
    if not isinstance(database, list):
        return 0.0

    crit = criteria if isinstance(criteria, list) else []
    values = _get_matching_values(database, field, crit)

    if len(values) < 2:
        return 0.0

    mean = sum(values) / len(values)
    return sum((x - mean) ** 2 for x in values) / (len(values) - 1)


def fn_dvarp(database: Any, field: Any, criteria: Any) -> float:
    """@DVARP - Population variance of matching records.

    Usage: @DVARP(database, field, criteria)
    """
    if not isinstance(database, list):
        return 0.0

    crit = criteria if isinstance(criteria, list) else []
    values = _get_matching_values(database, field, crit)

    if not values:
        return 0.0

    mean = sum(values) / len(values)
    return sum((x - mean) ** 2 for x in values) / len(values)


def fn_dget(database: Any, field: Any, criteria: Any) -> Any:
    """@DGET - Get single value from matching record.

    Usage: @DGET(database, field, criteria)
    Returns error if more than one record matches.
    """
    if not isinstance(database, list) or len(database) < 2:
        return FormulaError.VALUE

    headers = database[0] if isinstance(database[0], list) else [database[0]]
    field_idx = _get_field_index(database, field)

    if field_idx is None:
        return FormulaError.VALUE

    crit = criteria if isinstance(criteria, list) else []
    matches = []

    for row in database[1:]:
        if not isinstance(row, list):
            row = [row]

        if _matches_criteria(row, headers, crit):
            if field_idx < len(row):
                matches.append(row[field_idx])

    if len(matches) == 0:
        return FormulaError.VALUE
    if len(matches) > 1:
        return FormulaError.NUM  # Multiple matches

    return matches[0]


# Function registry for this module
DATABASE_FUNCTIONS = {
    "DSUM": fn_dsum,
    "DAVG": fn_davg,
    "DAVERAGE": fn_davg,
    "DCOUNT": fn_dcount,
    "DCOUNTA": fn_dcounta,
    "DMIN": fn_dmin,
    "DMAX": fn_dmax,
    "DSTD": fn_dstd,
    "DSTDEV": fn_dstd,
    "DSTDP": fn_dstdp,
    "DVAR": fn_dvar,
    "DVARP": fn_dvarp,
    "DGET": fn_dget,
}

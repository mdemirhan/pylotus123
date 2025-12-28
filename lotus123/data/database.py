"""Database operations: sort, query, extract.

Implements Lotus 1-2-3 /Data commands for treating ranges as databases.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum, auto
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


class SortOrder(Enum):
    """Sort order direction."""
    ASCENDING = auto()
    DESCENDING = auto()


@dataclass
class SortKey:
    """A sorting key specification."""
    column: int  # Column index within the range
    order: SortOrder = SortOrder.ASCENDING


class DatabaseOperations:
    """Database operations on spreadsheet ranges.

    Treats a range as a database where:
    - First row contains field names (headers)
    - Subsequent rows are records
    - Columns are fields
    """

    def __init__(self, spreadsheet: Spreadsheet) -> None:
        self.spreadsheet = spreadsheet

    def sort_range(self, start_row: int, start_col: int,
                   end_row: int, end_col: int,
                   keys: list[SortKey],
                   has_header: bool = True) -> None:
        """Sort a range of data.

        Args:
            start_row, start_col: Top-left of range
            end_row, end_col: Bottom-right of range
            keys: List of SortKey specifying sort columns and order
            has_header: If True, first row is header and not sorted
        """
        # Normalize range
        if start_row > end_row:
            start_row, end_row = end_row, start_row
        if start_col > end_col:
            start_col, end_col = end_col, start_col

        # Determine data rows
        data_start = start_row + 1 if has_header else start_row

        if data_start > end_row:
            return  # No data to sort

        # Extract rows as list of (row_idx, values)
        rows = []
        for r in range(data_start, end_row + 1):
            row_values = []
            for c in range(start_col, end_col + 1):
                row_values.append(self.spreadsheet.get_value(r, c))
            rows.append((r, row_values))

        # Sort using the keys
        def sort_key(item):
            row_idx, values = item
            key_values = []
            for sk in keys:
                if 0 <= sk.column < len(values):
                    val = values[sk.column]
                    # Handle sort order
                    if sk.order == SortOrder.DESCENDING:
                        # For descending, we negate numbers or reverse strings
                        if isinstance(val, (int, float)):
                            val = -val
                        elif isinstance(val, str):
                            # Use tuple with reverse flag
                            val = (1, val)  # Will be sorted in reverse
                    else:
                        if isinstance(val, str):
                            val = (0, val)
                    key_values.append(val)
                else:
                    key_values.append("")
            return tuple(key_values)

        # Sort rows
        sorted_rows = sorted(rows, key=sort_key)

        # Handle descending order for strings
        for sk in reversed(keys):
            if sk.order == SortOrder.DESCENDING:
                # Re-sort with this key in reverse
                def desc_key(item, col=sk.column):
                    row_idx, values = item
                    if 0 <= col < len(values):
                        val = values[col]
                        if isinstance(val, str):
                            return val
                    return ""
                sorted_rows = sorted(sorted_rows, key=desc_key, reverse=True)

        # Simple approach: collect all cell data, sort, put back
        cell_data = []
        for r in range(data_start, end_row + 1):
            row_data = []
            for c in range(start_col, end_col + 1):
                cell = self.spreadsheet.get_cell_if_exists(r, c)
                if cell:
                    row_data.append((cell.raw_value, cell.format_code))
                else:
                    row_data.append(("", "G"))
            cell_data.append(row_data)

        # Sort cell data
        def cell_sort_key(row_data):
            key_values = []
            for sk in keys:
                if 0 <= sk.column < len(row_data):
                    raw_val = row_data[sk.column][0]
                    # Parse value for comparison
                    try:
                        val = float(raw_val.replace(",", "")) if raw_val else 0
                    except ValueError:
                        val = raw_val.lower() if raw_val else ""

                    if sk.order == SortOrder.DESCENDING:
                        if isinstance(val, (int, float)):
                            val = -val
                    key_values.append(val)
            return tuple(key_values)

        sorted_data = sorted(cell_data, key=cell_sort_key)

        # For descending string columns, reverse sort
        for sk in reversed(keys):
            if sk.order == SortOrder.DESCENDING:
                # Check if this column has strings
                has_strings = any(
                    not row_data[sk.column][0].replace(",", "").replace(".", "").replace("-", "").isdigit()
                    for row_data in sorted_data
                    if sk.column < len(row_data) and row_data[sk.column][0]
                )
                if has_strings:
                    sorted_data = sorted(sorted_data,
                                        key=lambda rd, col=sk.column: rd[col][0].lower() if col < len(rd) else "",
                                        reverse=True)

        # Write back
        for i, row_data in enumerate(sorted_data):
            r = data_start + i
            for j, (raw_val, fmt) in enumerate(row_data):
                c = start_col + j
                cell = self.spreadsheet.get_cell(r, c)
                cell.set_value(raw_val)
                cell.format_code = fmt

        self.spreadsheet._invalidate_cache()

    def query(self, data_range: tuple[int, int, int, int],
              criteria_range: tuple[int, int, int, int] | None = None,
              criteria_func: callable = None) -> list[int]:
        """Query records matching criteria.

        Args:
            data_range: (start_row, start_col, end_row, end_col) of data
            criteria_range: Range containing criteria (Lotus-style)
            criteria_func: Alternative: Python function(row_values) -> bool

        Returns:
            List of row indices matching criteria
        """
        start_row, start_col, end_row, end_col = data_range

        # Skip header row
        data_start = start_row + 1
        matching_rows = []

        for r in range(data_start, end_row + 1):
            row_values = []
            for c in range(start_col, end_col + 1):
                row_values.append(self.spreadsheet.get_value(r, c))

            # Check criteria
            if criteria_func:
                if criteria_func(row_values):
                    matching_rows.append(r)
            else:
                # No criteria = match all
                matching_rows.append(r)

        return matching_rows

    def extract(self, data_range: tuple[int, int, int, int],
                output_start: tuple[int, int],
                matching_rows: list[int],
                columns: list[int] | None = None) -> int:
        """Extract matching records to output range.

        Args:
            data_range: Source data range
            output_start: (row, col) for output
            matching_rows: List of row indices to extract
            columns: Which columns to extract (None = all)

        Returns:
            Number of records extracted
        """
        start_row, start_col, end_row, end_col = data_range
        out_row, out_col = output_start

        if columns is None:
            columns = list(range(end_col - start_col + 1))

        # Copy header
        for i, col_idx in enumerate(columns):
            src_col = start_col + col_idx
            header_val = self.spreadsheet.get_value(start_row, src_col)
            self.spreadsheet.set_cell(out_row, out_col + i, str(header_val))

        # Copy matching rows
        for row_offset, src_row in enumerate(matching_rows, 1):
            for i, col_idx in enumerate(columns):
                src_col = start_col + col_idx
                cell = self.spreadsheet.get_cell_if_exists(src_row, src_col)
                if cell:
                    dest_cell = self.spreadsheet.get_cell(out_row + row_offset, out_col + i)
                    dest_cell.set_value(cell.raw_value)
                    dest_cell.format_code = cell.format_code

        self.spreadsheet._invalidate_cache()
        return len(matching_rows)

    def delete_matching(self, data_range: tuple[int, int, int, int],
                        matching_rows: list[int]) -> int:
        """Delete matching records from data range.

        Args:
            data_range: Data range
            matching_rows: Rows to delete

        Returns:
            Number of records deleted
        """
        # Delete from bottom to top to maintain indices
        for row in sorted(matching_rows, reverse=True):
            self.spreadsheet.delete_row(row)

        return len(matching_rows)

    def unique(self, data_range: tuple[int, int, int, int],
               key_columns: list[int]) -> list[int]:
        """Find unique records based on key columns.

        Args:
            data_range: Data range
            key_columns: Columns to use for uniqueness check

        Returns:
            List of row indices for first occurrence of each unique combination
        """
        start_row, start_col, end_row, end_col = data_range
        data_start = start_row + 1

        seen = set()
        unique_rows = []

        for r in range(data_start, end_row + 1):
            key = tuple(
                self.spreadsheet.get_value(r, start_col + c)
                for c in key_columns
            )
            if key not in seen:
                seen.add(key)
                unique_rows.append(r)

        return unique_rows

    def subtotal(self, data_range: tuple[int, int, int, int],
                 group_col: int, sum_cols: list[int]) -> dict:
        """Calculate subtotals by group.

        Args:
            data_range: Data range
            group_col: Column to group by
            sum_cols: Columns to sum

        Returns:
            Dict mapping group values to sum dicts
        """
        start_row, start_col, end_row, end_col = data_range
        data_start = start_row + 1

        totals = {}

        for r in range(data_start, end_row + 1):
            group_val = self.spreadsheet.get_value(r, start_col + group_col)

            if group_val not in totals:
                totals[group_val] = {c: 0 for c in sum_cols}

            for c in sum_cols:
                val = self.spreadsheet.get_value(r, start_col + c)
                if isinstance(val, (int, float)):
                    totals[group_val][c] += val

        return totals

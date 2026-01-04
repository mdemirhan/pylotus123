"""Database operations: sort, query, extract.

Implements Lotus 1-2-3 /Data commands for treating ranges as databases.
"""

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import TYPE_CHECKING, Any, Callable, Iterator

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


class SortOrder(Enum):
    """Sort order direction."""

    ASCENDING = auto()
    DESCENDING = auto()


@dataclass(frozen=True, slots=True)
class CellData:
    """Represents a cell's value and format for sorting operations."""

    raw_value: str
    format_code: str

    def __iter__(self) -> Iterator[str]:
        """Implemented to support tuple unpacking."""
        yield self.raw_value
        yield self.format_code


EMPTY_CELL = CellData("", "G")


@dataclass(frozen=True, slots=True)
class CellArray:
    """A row of cell data."""

    _items: list[CellData] = field(repr=False)

    def __getitem__(self, index: int) -> CellData:
        return self._items[index]

    def __len__(self) -> int:
        return len(self._items)

    def __iter__(self) -> Iterator[CellData]:
        return iter(self._items)


@dataclass(frozen=True, slots=True)
class CellMatrix:
    """A 2D grid of cell data for sorting operations."""

    _rows: list[CellArray] = field(repr=False)

    def __getitem__(self, index: int) -> CellArray:
        return self._rows[index]

    def __len__(self) -> int:
        return len(self._rows)

    def __iter__(self) -> Iterator[CellArray]:
        return iter(self._rows)


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

    @staticmethod
    def _parse_sort_value(raw_val: str) -> float | str:
        """Parse a raw string value for sorting comparison.

        Numbers (with commas) are converted to floats for numeric sorting.
        Non-numeric values are lowercased for case-insensitive string sorting.
        """
        if not raw_val:
            return 0.0
        try:
            return float(raw_val.replace(",", ""))
        except ValueError:
            return raw_val.lower()

    @staticmethod
    def _is_numeric_string(value: str) -> bool:
        """Check if a string represents a numeric value."""
        if not value:
            return False
        cleaned = value.replace(",", "").replace(".", "").replace("-", "")
        return cleaned.isdigit()

    def _extract_cell_data(
        self,
        data_start: int,
        end_row: int,
        start_col: int,
        end_col: int,
        values_only: bool,
    ) -> CellMatrix:
        """Extract cell data (raw_value, format_code) for sorting.

        When values_only=True, formulas are converted to their computed values.
        """
        rows: list[CellArray] = []
        for r in range(data_start, end_row + 1):
            cells: list[CellData] = []

            for c in range(start_col, end_col + 1):
                cell = self.spreadsheet.get_cell_if_exists(r, c)

                if cell is None:
                    cells.append(EMPTY_CELL)
                elif values_only and cell.is_formula:
                    computed = self.spreadsheet.get_value(r, c)
                    cells.append(CellData(str(computed) if computed else "", cell.format_code))
                else:
                    cells.append(CellData(cell.raw_value, cell.format_code))

            rows.append(CellArray(cells))
        return CellMatrix(rows)

    def _create_sort_key(self, keys: list[SortKey]) -> Callable[[CellArray], tuple[Any, ...]]:
        """Create a sort key function for the given sort keys.

        Handles numeric negation for descending order on numeric columns.
        """

        def sort_key(row_data: CellArray) -> tuple[Any, ...]:
            key_values: list[Any] = []
            for sk in keys:
                if 0 <= sk.column < len(row_data):
                    raw_val = row_data[sk.column].raw_value
                    sort_val = self._parse_sort_value(raw_val)

                    if sk.order == SortOrder.DESCENDING and isinstance(sort_val, (int, float)):
                        sort_val = -sort_val
                    key_values.append(sort_val)
            return tuple(key_values)

        return sort_key

    def _apply_descending_string_sorts(
        self, sorted_data: CellMatrix, keys: list[SortKey]
    ) -> CellMatrix:
        """Apply reverse sorting for descending string columns.

        Numeric columns are handled by negation in the primary sort,
        but string columns need a secondary reverse sort.
        """
        for sk in reversed(keys):
            if sk.order != SortOrder.DESCENDING:
                continue

            # Check if this column has any non-numeric values
            has_strings = any(
                not self._is_numeric_string(row_data[sk.column].raw_value)
                for row_data in sorted_data
                if sk.column < len(row_data) and row_data[sk.column].raw_value
            )

            if has_strings:
                col = sk.column

                def string_key(rd: CellArray, c: int = col) -> str:
                    return rd[c].raw_value.lower() if c < len(rd) else ""

                sorted_data = CellMatrix(sorted(sorted_data, key=string_key, reverse=True))

        return sorted_data

    def sort_range(
        self,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int,
        keys: list[SortKey],
        has_header: bool = True,
        values_only: bool = True,
    ) -> None:
        """Sort a range of data.

        Args:
            start_row, start_col: Top-left of range
            end_row, end_col: Bottom-right of range
            keys: List of SortKey specifying sort columns and order
            has_header: If True, first row is header and not sorted
            values_only: If True, formulas are converted to their computed values
                        before sorting. If False, formulas are moved as-is (but
                        their references won't be adjusted, which may break them).
        """
        # Normalize range (handle reversed indices)
        if start_row > end_row:
            start_row, end_row = end_row, start_row
        if start_col > end_col:
            start_col, end_col = end_col, start_col

        # Determine data rows (skip header if present)
        data_start = start_row + 1 if has_header else start_row
        if data_start > end_row:
            return  # No data to sort

        # Extract cell data for sorting
        cell_data: CellMatrix = self._extract_cell_data(
            data_start, end_row, start_col, end_col, values_only
        )

        # Primary sort with numeric descending handled by negation
        sort_key = self._create_sort_key(keys)

        # Apply reverse sorts for descending string columns
        sorted_data: CellMatrix = self._apply_descending_string_sorts(
            CellMatrix(sorted(cell_data, key=sort_key)), keys
        )

        # Write sorted data back to cells
        for i, row_data in enumerate(sorted_data):
            r = data_start + i
            for j, (raw_val, fmt) in enumerate(row_data):
                c = start_col + j
                cell = self.spreadsheet.get_cell(r, c)
                cell.set_value(raw_val)
                cell.format_code = fmt

        self.spreadsheet.invalidate_cache()
        self.spreadsheet.rebuild_dependency_graph()

    def sort_range_with_changes(
        self,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int,
        keys: list[SortKey],
        has_header: bool = True,
        values_only: bool = True,
    ) -> list[tuple[int, int, str, str]]:
        """Sort a range of data and return the changes made.

        This method is similar to sort_range but returns the changes for undo support.

        Args:
            start_row, start_col: Top-left of range
            end_row, end_col: Bottom-right of range
            keys: List of SortKey specifying sort columns and order
            has_header: If True, first row is header and not sorted
            values_only: If True, formulas are converted to their computed values

        Returns:
            List of (row, col, new_value, old_value) tuples for undo support.
        """
        # Normalize range (handle reversed indices)
        if start_row > end_row:
            start_row, end_row = end_row, start_row
        if start_col > end_col:
            start_col, end_col = end_col, start_col

        # Determine data rows (skip header if present)
        data_start = start_row + 1 if has_header else start_row
        if data_start > end_row:
            return []  # No data to sort

        # Extract cell data for sorting
        cell_data: CellMatrix = self._extract_cell_data(
            data_start, end_row, start_col, end_col, values_only
        )

        # Primary sort with numeric descending handled by negation
        sort_key = self._create_sort_key(keys)

        # Apply reverse sorts for descending string columns
        sorted_data: CellMatrix = self._apply_descending_string_sorts(
            CellMatrix(sorted(cell_data, key=sort_key)), keys
        )

        # Collect changes and write sorted data back to cells
        # Format: (row, col, new_value, old_value) - matches RangeChangeCommand
        changes: list[tuple[int, int, str, str]] = []
        for i, row_data in enumerate(sorted_data):
            r = data_start + i
            for j, (raw_val, fmt) in enumerate(row_data):
                c = start_col + j
                cell = self.spreadsheet.get_cell(r, c)
                old_value = cell.raw_value
                if raw_val != old_value:
                    changes.append((r, c, raw_val, old_value))
                cell.set_value(raw_val)
                cell.format_code = fmt

        self.spreadsheet.invalidate_cache()
        self.spreadsheet.rebuild_dependency_graph()

        return changes

    def query(
        self,
        data_range: tuple[int, int, int, int],
        criteria_range: tuple[int, int, int, int] | None = None,
        criteria_func: Callable[[list[Any]], bool] | None = None,
    ) -> list[int]:
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

    def extract(
        self,
        data_range: tuple[int, int, int, int],
        output_start: tuple[int, int],
        matching_rows: list[int],
        columns: list[int] | None = None,
    ) -> int:
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

        self.spreadsheet.invalidate_cache()
        return len(matching_rows)

    def delete_matching(
        self, data_range: tuple[int, int, int, int], matching_rows: list[int]
    ) -> int:
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

    def unique(self, data_range: tuple[int, int, int, int], key_columns: list[int]) -> list[int]:
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
            key = tuple(self.spreadsheet.get_value(r, start_col + c) for c in key_columns)
            if key not in seen:
                seen.add(key)
                unique_rows.append(r)

        return unique_rows

    def subtotal(
        self, data_range: tuple[int, int, int, int], group_col: int, sum_cols: list[int]
    ) -> dict[Any, dict[int, float]]:
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

        totals: dict[Any, dict[int, float]] = {}

        for r in range(data_start, end_row + 1):
            group_val = self.spreadsheet.get_value(r, start_col + group_col)

            if group_val not in totals:
                totals[group_val] = {c: 0.0 for c in sum_cols}

            for c in sum_cols:
                val = self.spreadsheet.get_value(r, start_col + c)
                if isinstance(val, (int, float)):
                    totals[group_val][c] += val

        return totals

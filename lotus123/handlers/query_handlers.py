"""Data Query handler methods for LotusApp."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable

from ..core import make_cell_ref, parse_cell_ref
from ..ui import CommandInput
from .base import BaseHandler

if TYPE_CHECKING:
    from .base import AppProtocol


class QueryHandler(BaseHandler):
    """Handler for data query operations."""

    def __init__(self, app: "AppProtocol") -> None:
        super().__init__(app)
        # Query state - owned by this handler
        self.input_range: tuple[int, int, int, int] | None = None
        self.criteria_range: tuple[int, int, int, int] | None = None
        self.output_range: tuple[int, int] | None = None
        self.find_results: list[int] | None = None
        self.find_index: int = 0

    def set_input(self) -> None:
        """Set the input (database) range for queries."""
        grid = self.get_grid()
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            self.input_range = (r1, c1, r2, c2)
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self.notify(f"Query Input range set to {range_str}")
        else:
            self._app.push_screen(
                CommandInput("Input range (e.g., A1:D100):"), self._do_set_input
            )

    def _do_set_input(self, result: str | None) -> None:
        if not result:
            return
        try:
            parts = result.upper().split(":")
            if len(parts) == 2:
                r1, c1 = parse_cell_ref(parts[0])
                r2, c2 = parse_cell_ref(parts[1])
                self.input_range = (r1, c1, r2, c2)
                self.notify(f"Query Input range set to {result.upper()}")
            else:
                self.notify("Invalid range format", severity="error")
        except ValueError as e:
            self.notify(f"Invalid range: {e}", severity="error")

    def set_criteria(self) -> None:
        """Set the criteria range for queries."""
        grid = self.get_grid()
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            self.criteria_range = (r1, c1, r2, c2)
            self.find_results = None
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self.notify(f"Query Criteria range set to {range_str}")
        else:
            self._app.push_screen(
                CommandInput("Criteria range (e.g., F1:G2):"),
                self._do_set_criteria,
            )

    def _do_set_criteria(self, result: str | None) -> None:
        if not result:
            return
        try:
            parts = result.upper().split(":")
            if len(parts) == 2:
                r1, c1 = parse_cell_ref(parts[0])
                r2, c2 = parse_cell_ref(parts[1])
                self.criteria_range = (r1, c1, r2, c2)
                self.find_results = None
                self.notify(f"Query Criteria range set to {result.upper()}")
            else:
                self.notify("Invalid range format", severity="error")
        except ValueError as e:
            self.notify(f"Invalid range: {e}", severity="error")

    def set_output(self) -> None:
        """Set the output range for query extraction."""
        grid = self.get_grid()
        if grid.has_selection:
            r1, c1, _, _ = grid.selection_range
            self.output_range = (r1, c1)
            self.notify(f"Query Output range set to {make_cell_ref(r1, c1)}")
        else:
            self._app.push_screen(
                CommandInput("Output start cell (e.g., H1):"), self._do_set_output
            )

    def _do_set_output(self, result: str | None) -> None:
        if not result:
            return
        try:
            row, col = parse_cell_ref(result.upper())
            self.output_range = (row, col)
            self.notify(f"Query Output range set to {result.upper()}")
        except ValueError as e:
            self.notify(f"Invalid cell reference: {e}", severity="error")

    def _build_criteria_filter(self) -> Callable[[list], bool] | None:
        """Build a filter function from the criteria range."""
        if not self.criteria_range or not self.input_range:
            return None

        from ..data.criteria import CriteriaParser

        cr1, cc1, cr2, cc2 = self.criteria_range
        ir1, ic1, ir2, ic2 = self.input_range

        # Get input headers
        input_headers = []
        for c in range(ic1, ic2 + 1):
            val = self.spreadsheet.get_value(ir1, c)
            input_headers.append(str(val).strip().upper() if val else "")

        # Get criteria headers and map to input column indices
        criteria_headers = []
        for c in range(cc1, cc2 + 1):
            val = self.spreadsheet.get_value(cr1, c)
            criteria_headers.append(str(val).strip().upper() if val else "")

        # Build column mapping: criteria col index -> input col index
        col_mapping: dict[int, int] = {}
        for ci, ch in enumerate(criteria_headers):
            if ch:
                for ii, ih in enumerate(input_headers):
                    if ch == ih:
                        col_mapping[ci] = ii
                        break

        # Get criteria values (rows after header)
        criteria_rows = []
        for r in range(cr1 + 1, cr2 + 1):
            row_vals = []
            for c in range(cc1, cc2 + 1):
                row_vals.append(self.spreadsheet.get_value(r, c))
            criteria_rows.append(row_vals)

        # Parse criteria with remapped columns
        parser = CriteriaParser()
        parser.parse_range(criteria_headers, criteria_rows)

        # Create filter that maps input row values to criteria columns
        def criteria_filter(row_values: list) -> bool:
            # Remap row values to match criteria column order
            mapped_values = [""] * len(criteria_headers)
            for ci, ii in col_mapping.items():
                if ii < len(row_values):
                    mapped_values[ci] = row_values[ii]
            return parser.matches(mapped_values)

        return criteria_filter

    def find(self) -> None:
        """Find records matching the criteria."""
        if not self.input_range:
            self.notify("Set Input range first (Data:Query:Input)", severity="warning")
            return

        grid = self.get_grid()

        # If we already have results, cycle to the next one
        if self.find_results:
            self.find_index = (
                self.find_index + 1
            ) % len(self.find_results)
            target_row = self.find_results[self.find_index]
            grid.cursor_row = target_row
            grid.cursor_col = self._get_first_criteria_col()
            pos = self.find_index + 1
            total = len(self.find_results)
            self.notify(f"Record {pos} of {total}")
            return

        # Run a new query
        from ..data.database import DatabaseOperations

        db = DatabaseOperations(self.spreadsheet)
        criteria_filter = self._build_criteria_filter()

        matching_rows = db.query(
            self.input_range, criteria_func=criteria_filter
        )

        if not matching_rows:
            self.notify("No matching records found")
            self.find_results = None
            return

        self.find_results = matching_rows
        self.find_index = 0

        # Navigate to first match
        first_row = matching_rows[0]
        grid.cursor_row = first_row
        grid.cursor_col = self._get_first_criteria_col()

        self.notify(
            f"Found {len(matching_rows)} matching record(s). Press Find again for next."
        )

    def _get_first_criteria_col(self) -> int:
        """Get the column index in the input range that matches the first criteria header."""
        if not self.criteria_range or not self.input_range:
            return self.input_range[1] if self.input_range else 0

        crit_start_row, crit_start_col, _, crit_end_col = self.criteria_range
        input_start_row, input_start_col, _, input_end_col = self.input_range

        # Get input headers
        input_headers: dict[str, int] = {}
        for col in range(input_start_col, input_end_col + 1):
            header = str(self.spreadsheet.get_value(input_start_row, col)).strip().upper()
            input_headers[header] = col

        # Find the first criteria header that matches an input header
        for col in range(crit_start_col, crit_end_col + 1):
            crit_header = str(self.spreadsheet.get_value(crit_start_row, col)).strip().upper()
            if crit_header in input_headers:
                return input_headers[crit_header]

        # Default to start of input range
        return input_start_col

    def extract(self) -> None:
        """Extract matching records to the output range."""
        if not self.input_range:
            self.notify("Set Input range first (Data:Query:Input)", severity="warning")
            return
        if not self.output_range:
            self.notify(
                "Set Output range first (Data:Query:Output)", severity="warning"
            )
            return

        from ..data.database import DatabaseOperations

        db = DatabaseOperations(self.spreadsheet)
        criteria_filter = self._build_criteria_filter()

        matching_rows = db.query(
            self.input_range, criteria_func=criteria_filter
        )

        if not matching_rows:
            self.notify("No matching records to extract")
            return

        count = db.extract(
            self.input_range, self.output_range, matching_rows
        )
        grid = self.get_grid()
        grid.refresh_grid()
        self._app._mark_dirty()
        self.notify(f"Extracted {count} record(s)")

    def unique(self) -> None:
        """Extract unique matching records to the output range."""
        if not self.input_range:
            self.notify("Set Input range first (Data:Query:Input)", severity="warning")
            return
        if not self.output_range:
            self.notify(
                "Set Output range first (Data:Query:Output)", severity="warning"
            )
            return

        from ..data.database import DatabaseOperations

        db = DatabaseOperations(self.spreadsheet)
        criteria_filter = self._build_criteria_filter()

        # First query matching rows
        matching_rows = db.query(
            self.input_range, criteria_func=criteria_filter
        )

        if not matching_rows:
            self.notify("No matching records found")
            return

        # Then find unique among matching
        ir1, ic1, ir2, ic2 = self.input_range
        all_columns = list(range(ic2 - ic1 + 1))
        unique_rows = db.unique(self.input_range, all_columns)

        # Intersect with matching
        unique_matching = [r for r in unique_rows if r in matching_rows]

        count = db.extract(
            self.input_range, self.output_range, unique_matching
        )
        grid = self.get_grid()
        grid.refresh_grid()
        self._app._mark_dirty()
        self.notify(f"Extracted {count} unique record(s)")

    def delete(self) -> None:
        """Delete records matching the criteria."""
        if not self.input_range:
            self.notify("Set Input range first (Data:Query:Input)", severity="warning")
            return

        from ..data.database import DatabaseOperations

        db = DatabaseOperations(self.spreadsheet)
        criteria_filter = self._build_criteria_filter()

        matching_rows = db.query(
            self.input_range, criteria_func=criteria_filter
        )

        if not matching_rows:
            self.notify("No matching records to delete")
            return

        count = db.delete_matching(self.input_range, matching_rows)
        grid = self.get_grid()
        grid.refresh_grid()
        self._app._mark_dirty()
        self.notify(f"Deleted {count} record(s)")

    def reset(self) -> None:
        """Reset all query settings."""
        self.input_range = None
        self.criteria_range = None
        self.output_range = None
        self.find_results = None
        self.find_index = 0
        self.notify("Query settings reset")

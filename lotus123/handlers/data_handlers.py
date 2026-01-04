"""Data operation handler methods for LotusApp (fill, sort)."""

from __future__ import annotations

from typing import TYPE_CHECKING

from ..core import col_to_index, index_to_col
from ..data.database import DatabaseOperations, SortKey, SortOrder
from ..ui import CommandInput
from ..utils.undo import RangeChangeCommand
from .base import BaseHandler

if TYPE_CHECKING:
    from .base import AppProtocol


class DataHandler(BaseHandler):
    """Handler for data operations (fill, sort)."""

    def __init__(self, app: "AppProtocol") -> None:
        super().__init__(app)

    def data_fill(self) -> None:
        """Start a fill operation on the selected range."""
        grid = self.get_grid()
        if not grid.has_selection:
            self.notify("Select a range first")
            return
        self._app.push_screen(
            CommandInput("Fill with (start,step,stop) or value:"), self._do_data_fill
        )

    def _do_data_fill(self, result: str | None) -> None:
        if not result:
            return
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        changes = []
        try:
            if "," in result:
                parts = [p.strip() for p in result.split(",")]
                start = float(parts[0])
                step = float(parts[1]) if len(parts) > 1 else 1
                val = start
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        cell = self.spreadsheet.get_cell(r, c)
                        old_value = cell.raw_value
                        new_value = str(int(val) if val == int(val) else val)
                        changes.append((r, c, new_value, old_value))
                        val += step
            else:
                fill_value = result
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        cell = self.spreadsheet.get_cell(r, c)
                        old_value = cell.raw_value
                        changes.append((r, c, fill_value, old_value))
            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid.refresh_grid()
                self.update_status()
                self.mark_dirty()
                self.notify(f"Filled {len(changes)} cell(s)")
        except ValueError as e:
            self.notify(f"Invalid fill value: {e}", severity="error")

    def data_sort(self) -> None:
        """Start a sort operation on the selected range."""
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        first_col = index_to_col(c1)
        last_col = index_to_col(c2)
        col_range = first_col if c1 == c2 else f"{first_col}-{last_col}"
        self._app.push_screen(
            CommandInput(f"Sort column [{col_range}] (add D for descending, e.g., 'A' or 'AD'):"),
            self._do_data_sort,
        )

    def _do_data_sort(self, result: str | None) -> None:
        if not result:
            return
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        try:
            # Parse user input: column letter + optional 'D' suffix for descending
            # e.g., "A" = col A ascending, "AD" = col A descending
            #       "D" = col D ascending, "DD" = col D descending
            result = result.strip().upper()
            if len(result) >= 2 and result.endswith("D"):
                descending = True
                sort_col_letter = result[:-1]
            else:
                descending = False
                sort_col_letter = result

            # Convert column letter to absolute index within selection
            sort_col_abs = col_to_index(sort_col_letter)
            if sort_col_abs < c1 or sort_col_abs > c2:
                # Try as relative column (A=first col in selection)
                sort_col_idx = ord(sort_col_letter) - ord("A")
                sort_col_abs = c1 + sort_col_idx
            if sort_col_abs < c1 or sort_col_abs > c2:
                self.notify(
                    f"Sort column must be within selection ({index_to_col(c1)}-{index_to_col(c2)})",
                    severity="error",
                )
                return

            # Create SortKey with column index relative to range start
            order = SortOrder.DESCENDING if descending else SortOrder.ASCENDING
            sort_key = SortKey(column=sort_col_abs - c1, order=order)

            # Use DatabaseOperations for sorting with change tracking
            db_ops = DatabaseOperations(self.spreadsheet)
            changes = db_ops.sort_range_with_changes(
                r1, c1, r2, c2, keys=[sort_key], has_header=False
            )

            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid.refresh_grid()
                self.update_status()
                self.mark_dirty()
                order_name = "descending" if descending else "ascending"
                self.notify(f"Sorted by column {sort_col_letter} ({order_name})")
            else:
                self.notify("Data already sorted")
        except Exception as e:
            self.notify(f"Sort error: {e}", severity="error")

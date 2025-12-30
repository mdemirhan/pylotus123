"""Data operation handler methods for LotusApp (fill, sort)."""

from __future__ import annotations

from typing import TYPE_CHECKING

from ..core import col_to_index, index_to_col
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
                self._app._update_status()
                self._app._mark_dirty()
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
            CommandInput(
                f"Sort column [{col_range}] (add D for descending, e.g., 'A' or 'AD'):"
            ),
            self._do_data_sort,
        )

    def _do_data_sort(self, result: str | None) -> None:
        if not result:
            return
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        try:
            result = result.strip().upper()
            reverse = result.endswith("D")
            sort_col_letter = result.rstrip("D").rstrip("A") or result[0]
            sort_col_abs = col_to_index(sort_col_letter)
            if sort_col_abs < c1 or sort_col_abs > c2:
                sort_col_idx = ord(sort_col_letter) - ord("A")
                sort_col_abs = c1 + sort_col_idx
            if sort_col_abs < c1 or sort_col_abs > c2:
                self.notify(
                    f"Sort column must be within selection ({index_to_col(c1)}-{index_to_col(c2)})",
                    severity="error",
                )
                return
            rows_data = []
            for r in range(r1, r2 + 1):
                row_values = []
                for c in range(c1, c2 + 1):
                    cell = self.spreadsheet.get_cell(r, c)
                    row_values.append(cell.raw_value)
                sort_val = self.spreadsheet.get_value(r, sort_col_abs)
                sort_key: tuple[int, str | int | float]
                if sort_val == "" or sort_val is None:
                    sort_key = (2, "")
                elif isinstance(sort_val, (int, float)):
                    sort_key = (0, sort_val)
                else:
                    sort_key = (1, str(sort_val).lower())
                rows_data.append((sort_key, row_values))
            rows_data.sort(key=lambda x: x[0], reverse=reverse)
            changes = []
            for row_idx, (_, row_values) in enumerate(rows_data):
                target_row = r1 + row_idx
                for col_idx, value in enumerate(row_values):
                    target_col = c1 + col_idx
                    cell = self.spreadsheet.get_cell(target_row, target_col)
                    old_value = cell.raw_value
                    if value != old_value:
                        changes.append((target_row, target_col, value, old_value))
            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid.refresh_grid()
                self._app._update_status()
                self._app._mark_dirty()
                order_name = "descending" if reverse else "ascending"
                self.notify(
                    f"Sorted {len(rows_data)} rows by column {sort_col_letter} ({order_name})"
                )
            else:
                self.notify("Data already sorted")
        except Exception as e:
            self.notify(f"Sort error: {e}", severity="error")

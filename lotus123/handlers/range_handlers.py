"""Range operation handler methods for LotusApp."""

from __future__ import annotations

from typing import TYPE_CHECKING

from ..core import make_cell_ref
from ..ui import CommandInput
from ..utils.undo import RangeChangeCommand
from .base import BaseHandler

if TYPE_CHECKING:
    from .base import AppProtocol


class RangeHandler(BaseHandler):
    """Handler for range operations (format, label, name, protect)."""

    def __init__(self, app: "AppProtocol") -> None:
        super().__init__(app)

    def range_format(self) -> None:
        """Set the format for the selected range."""
        self._app.push_screen(
            CommandInput(
                "Format (F=Fixed, S=Scientific, C=Currency, P=Percent, G=General):"
            ),
            self._do_range_format,
        )

    def _do_range_format(self, result: str | None) -> None:
        if not result:
            return
        format_char = result.upper()[0] if result else "G"
        format_map = {"F": "F2", "S": "S", "C": "C2", "P": "P2", "G": "G", ",": ",2"}
        format_code = format_map.get(format_char, "G")
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                cell.format_code = format_code
        grid.refresh_grid()
        self._app._update_status()
        self._app._mark_dirty()
        self.notify(f"Format set to {format_code}")

    def range_label(self) -> None:
        """Set the label alignment for the selected range."""
        self._app.push_screen(
            CommandInput("Label alignment (L=Left, R=Right, C=Center):"),
            self._do_range_label,
        )

    def _do_range_label(self, result: str | None) -> None:
        if not result:
            return
        align_char = result.upper()[0] if result else "L"
        prefix_map = {"L": "'", "R": '"', "C": "^"}
        prefix = prefix_map.get(align_char, "'")
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        changes = []
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                old_value = cell.raw_value
                if old_value and not cell.is_formula:
                    display = cell.display_value
                    new_value = prefix + display
                    if new_value != old_value:
                        changes.append((r, c, new_value, old_value))
        if changes:
            cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
            self.undo_manager.execute(cmd)
            grid.refresh_grid()
            self._app._update_status()
            self._app._mark_dirty()
        align_names = {"L": "Left", "R": "Right", "C": "Center"}
        self.notify(f"Label alignment set to {align_names.get(align_char, 'Left')}")

    def range_name(self) -> None:
        """Create a named range from the selection."""
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
        self._app._pending_range = range_str
        self._app.push_screen(
            CommandInput(f"Name for range {range_str}:"), self._do_range_name
        )

    def _do_range_name(self, result: str | None) -> None:
        if not result:
            return
        name = result.strip().upper()
        if not name:
            return
        self.spreadsheet.named_ranges.add_from_string(name, self._app._pending_range)
        self._app._mark_dirty()
        self.notify(f"Named range '{name}' created for {self._app._pending_range}")

    def range_protect(self) -> None:
        """Toggle protection for the selected range."""
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        protected_count = 0
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                cell.is_protected = not cell.is_protected
                if cell.is_protected:
                    protected_count += 1
        total_cells = (r2 - r1 + 1) * (c2 - c1 + 1)
        self._app._mark_dirty()
        if protected_count > 0:
            self.notify(f"Protected {protected_count} cell(s)")
        else:
            self.notify(f"Unprotected {total_cells} cell(s)")

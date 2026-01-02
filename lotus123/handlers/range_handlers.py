"""Range operation handler methods for LotusApp."""

from __future__ import annotations

from typing import TYPE_CHECKING

from ..core import make_cell_ref
from ..ui import CommandInput
from ..utils.undo import RangeChangeCommand, RangeFormatCommand
from .base import BaseHandler

if TYPE_CHECKING:
    from .base import AppProtocol


class RangeHandler(BaseHandler):
    """Handler for range operations (format, label, name)."""

    def __init__(self, app: "AppProtocol") -> None:
        super().__init__(app)
        # Pending operation state - owned by this handler
        self.pending_range: str = ""

    def range_format(self) -> None:
        """Set the format for the selected range."""
        self._app.push_screen(
            CommandInput(
                "Format: G, F0-F15, S0-S15, C0-C15, P0-P15, ,0-,15, D1-D9, T1-T4, H, +:"
            ),
            self._do_range_format,
        )

    def _do_range_format(self, result: str | None) -> None:
        if not result:
            return
        format_code = self._normalize_format_code(result)
        if format_code is None:
            self.notify(f"Invalid format: {result}", severity="error")
            return
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        changes = []
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                old_format = cell.format_code
                if old_format != format_code:
                    changes.append((r, c, format_code, old_format))
        if changes:
            cmd = RangeFormatCommand(spreadsheet=self.spreadsheet, changes=changes)
            self.undo_manager.execute(cmd)
        grid.refresh_grid()
        self.update_status()
        self.mark_dirty()
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
            self.update_status()
            self.mark_dirty()
        align_names = {"L": "Left", "R": "Right", "C": "Center"}
        self.notify(f"Label alignment set to {align_names.get(align_char, 'Left')}")

    def range_name(self) -> None:
        """Create a named range from the selection."""
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
        self.pending_range = range_str
        self._app.push_screen(
            CommandInput(f"Name for range {range_str}:"), self._do_range_name
        )

    def _do_range_name(self, result: str | None) -> None:
        if not result:
            return
        name = result.strip().upper()
        if not name:
            return
        try:
            self.spreadsheet.named_ranges.add_from_string(name, self.pending_range)
            self.mark_dirty()
            self.notify(f"Named range '{name}' created for {self.pending_range}")
        except ValueError as e:
            self.notify(str(e), severity="error")

    def _normalize_format_code(self, code: str) -> str | None:
        """Normalize and validate a format code.

        Args:
            code: User-entered format code

        Returns:
            Normalized format code, or None if invalid
        """
        code = code.strip().upper()
        if not code:
            return None

        # Single character formats
        if code in ("G", "H", "+"):
            return code

        # Formats with decimal places: F, S, C, P (0-15)
        if code[0] in ("F", "S", "C", "P"):
            if len(code) == 1:
                return code + "2"  # Default to 2 decimal places
            try:
                decimals = int(code[1:])
                if 0 <= decimals <= 15:
                    return f"{code[0]}{decimals}"
            except ValueError:
                pass
            return None

        # Comma format (,0-,15)
        if code.startswith(","):
            if len(code) == 1:
                return ",2"  # Default to 2 decimal places
            try:
                decimals = int(code[1:])
                if 0 <= decimals <= 15:
                    return f",{decimals}"
            except ValueError:
                pass
            return None

        # Date formats (D1-D9)
        if code[0] == "D":
            if len(code) == 1:
                return "D1"  # Default to D1
            try:
                variant = int(code[1:])
                if 1 <= variant <= 9:
                    return f"D{variant}"
            except ValueError:
                pass
            return None

        # Time formats (T1-T4)
        if code[0] == "T":
            if len(code) == 1:
                return "T1"  # Default to T1
            try:
                variant = int(code[1:])
                if 1 <= variant <= 4:
                    return f"T{variant}"
            except ValueError:
                pass
            return None

        return None

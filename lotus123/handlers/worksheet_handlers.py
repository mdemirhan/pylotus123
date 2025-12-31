"""Worksheet operation handler methods for LotusApp."""

from __future__ import annotations

from typing import TYPE_CHECKING

from ..core.reference import index_to_col
from ..ui import CommandInput
from ..utils.undo import (
    Command,
    CompositeCommand,
    DeleteColCommand,
    DeleteRowCommand,
    InsertColCommand,
    InsertRowCommand,
)
from .base import BaseHandler

if TYPE_CHECKING:
    from .base import AppProtocol


class WorksheetHandler(BaseHandler):
    """Handler for worksheet operations (insert/delete row, global settings)."""

    def __init__(self, app: "AppProtocol") -> None:
        super().__init__(app)

    def insert_rows(self) -> None:
        """Prompt for number of rows to insert."""
        self._app.push_screen(
            CommandInput("Number of rows to insert:", default="1"), self._do_insert_rows
        )

    def _do_insert_rows(self, result: str | None) -> None:
        count = 1
        if result and result.strip():
            try:
                count = int(result.strip())
                if count < 1:
                    self.notify("Count must be at least 1", severity="error")
                    return
            except ValueError:
                self.notify("Invalid number", severity="error")
                return

        grid = self.get_grid()
        if count == 1:
            cmd = InsertRowCommand(spreadsheet=self.spreadsheet, row=grid.cursor_row)
            self.undo_manager.execute(cmd)
        else:
            commands: list[Command] = [
                InsertRowCommand(spreadsheet=self.spreadsheet, row=grid.cursor_row)
                for _ in range(count)
            ]
            composite = CompositeCommand(commands, f"Insert {count} rows")
            self.undo_manager.execute(composite)
        grid.refresh_grid()
        self._app._mark_dirty()
        self.notify(f"{count} row(s) inserted at row {grid.cursor_row + 1}")

    def insert_columns(self) -> None:
        """Prompt for number of columns to insert."""
        self._app.push_screen(
            CommandInput("Number of columns to insert:", default="1"), self._do_insert_columns
        )

    def _do_insert_columns(self, result: str | None) -> None:
        count = 1
        if result and result.strip():
            try:
                count = int(result.strip())
                if count < 1:
                    self.notify("Count must be at least 1", severity="error")
                    return
            except ValueError:
                self.notify("Invalid number", severity="error")
                return

        grid = self.get_grid()
        if count == 1:
            cmd = InsertColCommand(spreadsheet=self.spreadsheet, col=grid.cursor_col)
            self.undo_manager.execute(cmd)
        else:
            commands: list[Command] = [
                InsertColCommand(spreadsheet=self.spreadsheet, col=grid.cursor_col)
                for _ in range(count)
            ]
            composite = CompositeCommand(commands, f"Insert {count} columns")
            self.undo_manager.execute(composite)
        grid.recalculate_visible_area()
        grid.refresh_grid()
        self._app._mark_dirty()
        self.notify(f"{count} column(s) inserted at column {index_to_col(grid.cursor_col)}")

    def delete_rows(self) -> None:
        """Prompt for number of rows to delete."""
        self._app.push_screen(
            CommandInput("Number of rows to delete:", default="1"), self._do_delete_rows
        )

    def _do_delete_rows(self, result: str | None) -> None:
        count = 1
        if result and result.strip():
            try:
                count = int(result.strip())
                if count < 1:
                    self.notify("Count must be at least 1", severity="error")
                    return
            except ValueError:
                self.notify("Invalid number", severity="error")
                return

        grid = self.get_grid()
        if count == 1:
            cmd = DeleteRowCommand(spreadsheet=self.spreadsheet, row=grid.cursor_row)
            self.undo_manager.execute(cmd)
        else:
            commands: list[Command] = [
                DeleteRowCommand(spreadsheet=self.spreadsheet, row=grid.cursor_row)
                for _ in range(count)
            ]
            composite = CompositeCommand(commands, f"Delete {count} rows")
            self.undo_manager.execute(composite)
        grid.refresh_grid()
        self._app._update_status()
        self._app._mark_dirty()
        self.notify(f"{count} row(s) deleted starting at row {grid.cursor_row + 1}")

    def delete_columns(self) -> None:
        """Prompt for number of columns to delete."""
        self._app.push_screen(
            CommandInput("Number of columns to delete:", default="1"), self._do_delete_columns
        )

    def _do_delete_columns(self, result: str | None) -> None:
        count = 1
        if result and result.strip():
            try:
                count = int(result.strip())
                if count < 1:
                    self.notify("Count must be at least 1", severity="error")
                    return
            except ValueError:
                self.notify("Invalid number", severity="error")
                return

        grid = self.get_grid()
        if count == 1:
            cmd = DeleteColCommand(spreadsheet=self.spreadsheet, col=grid.cursor_col)
            self.undo_manager.execute(cmd)
        else:
            commands: list[Command] = [
                DeleteColCommand(spreadsheet=self.spreadsheet, col=grid.cursor_col)
                for _ in range(count)
            ]
            composite = CompositeCommand(commands, f"Delete {count} columns")
            self.undo_manager.execute(composite)
        grid.recalculate_visible_area()
        grid.refresh_grid()
        self._app._update_status()
        self._app._mark_dirty()
        self.notify(f"{count} column(s) deleted starting at column {index_to_col(grid.cursor_col)}")

    def set_column_width(self) -> None:
        """Prompt for a new column width."""
        self._app.push_screen(
            CommandInput("Column width (3-50):"), self._do_set_width
        )

    def _do_set_width(self, result: str | None) -> None:
        if result:
            try:
                width = int(result)
                if width < 3 or width > 50:
                    self.notify("Width must be between 3 and 50", severity="error")
                    return
                grid = self.get_grid()
                self.spreadsheet.set_col_width(grid.cursor_col, width)
                self._app._mark_dirty()
                grid.recalculate_visible_area()
                grid.refresh_grid()
            except ValueError:
                self.notify("Invalid width value", severity="error")

    def global_format(self) -> None:
        """Set the default format for new cells."""
        self._app.push_screen(
            CommandInput(
                f"Default format (F=Fixed, S=Scientific, C=Currency, P=Percent, G=General) [{self._app._global_format_code}]:"
            ),
            self._do_global_format,
        )

    def _do_global_format(self, result: str | None) -> None:
        if not result:
            return
        format_char = result.upper()[0] if result else "G"
        format_map = {"F": "F2", "S": "S", "C": "C2", "P": "P2", "G": "G", ",": ",2"}
        self._app._global_format_code = format_map.get(format_char, "G")
        self._app._mark_dirty()
        self.notify(f"Default format set to {self._app._global_format_code}")

    def global_label_prefix(self) -> None:
        """Set the default label alignment."""
        self._app.push_screen(
            CommandInput("Default label alignment (L=Left, R=Right, C=Center):"),
            self._do_global_label_prefix,
        )

    def _do_global_label_prefix(self, result: str | None) -> None:
        if not result:
            return
        align_char = result.upper()[0] if result else "L"
        prefix_map = {"L": "'", "R": '"', "C": "^"}
        self._app._global_label_prefix = prefix_map.get(align_char, "'")
        self._app._mark_dirty()
        align_names = {"'": "Left", '"': "Right", "^": "Center"}
        self.notify(
            f"Default label alignment set to {align_names.get(self._app._global_label_prefix, 'Left')}"
        )

    def global_column_width(self) -> None:
        """Set the default column width."""
        self._app.push_screen(
            CommandInput(
                f"Default column width (3-50) [{self._app._global_col_width}]:"
            ),
            self._do_global_column_width,
        )

    def _do_global_column_width(self, result: str | None) -> None:
        if not result:
            return
        try:
            width = int(result)
            width = max(3, min(50, width))
            self._app._global_col_width = width
            self.spreadsheet.global_settings["default_col_width"] = width
            grid = self.get_grid()
            grid.default_col_width = width
            grid.recalculate_visible_area()
            grid.refresh_grid()
            self._app._mark_dirty()
            self.notify(f"Default column width set to {width}")
        except ValueError:
            self.notify("Invalid width", severity="error")

    def global_recalculation(self) -> None:
        """Toggle between automatic and manual recalculation."""
        if self._app._recalc_mode == "auto":
            self._app._recalc_mode = "manual"
            self.notify("Recalculation: Manual (press F9 to recalculate)")
        else:
            self._app._recalc_mode = "auto"
            self.spreadsheet._invalidate_cache()
            grid = self.get_grid()
            grid.refresh_grid()
            self.notify("Recalculation: Automatic")

    def global_zero(self) -> None:
        """Toggle display of zero values."""
        self._app._global_zero_display = not self._app._global_zero_display
        grid = self.get_grid()
        grid.show_zero = self._app._global_zero_display
        grid.refresh_grid()
        self._app._mark_dirty()
        if self._app._global_zero_display:
            self.notify("Zero values: Displayed")
        else:
            self.notify("Zero values: Hidden (blank)")

    def worksheet_erase(self) -> None:
        """Prompt to erase the entire worksheet."""
        self._app.push_screen(
            CommandInput("Erase entire worksheet? (Y/N):"), self._do_worksheet_erase
        )

    def _do_worksheet_erase(self, result: str | None) -> None:
        if result and result.upper().startswith("Y"):
            self.spreadsheet.clear()
            self.undo_manager.clear()
            grid = self.get_grid()
            grid.refresh_grid()
            self._app._update_status()
            self._app._mark_dirty()
            self.notify("Worksheet erased")

"""Clipboard-related handler methods for LotusApp."""

from __future__ import annotations

from typing import TYPE_CHECKING

from ..core import adjust_formula_references, make_cell_ref, parse_cell_ref
from ..ui import CommandInput
from ..utils.undo import CellChangeCommand, RangeChangeCommand
from .base import BaseHandler

if TYPE_CHECKING:
    from .base import AppProtocol


class ClipboardHandler(BaseHandler):
    """Handler for clipboard operations (copy, cut, paste, move)."""

    def __init__(self, app: "AppProtocol") -> None:
        super().__init__(app)

    def menu_copy(self) -> None:
        """Initiate a menu-based copy operation."""
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        source_range = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
        self._app._pending_source_range = (r1, c1, r2, c2)
        self._app.push_screen(
            CommandInput(f"Copy {source_range} TO (e.g., D1):"), self._do_menu_copy
        )

    def _do_menu_copy(self, result: str | None) -> None:
        if not result:
            return
        try:
            dest_row, dest_col = parse_cell_ref(result.upper())
            r1, c1, r2, c2 = self._app._pending_source_range
            changes = []
            for r_offset in range(r2 - r1 + 1):
                for c_offset in range(c2 - c1 + 1):
                    src_row, src_col = r1 + r_offset, c1 + c_offset
                    target_row, target_col = dest_row + r_offset, dest_col + c_offset
                    if (
                        target_row >= self.spreadsheet.rows
                        or target_col >= self.spreadsheet.cols
                    ):
                        continue
                    src_cell = self.spreadsheet.get_cell(src_row, src_col)
                    target_cell = self.spreadsheet.get_cell(target_row, target_col)
                    old_value = target_cell.raw_value
                    new_value = src_cell.raw_value
                    if new_value and (
                        new_value.startswith("=") or new_value.startswith("@")
                    ):
                        row_delta = target_row - src_row
                        col_delta = target_col - src_col
                        new_value = new_value[0] + adjust_formula_references(
                            new_value[1:], row_delta, col_delta
                        )
                    if new_value != old_value:
                        changes.append((target_row, target_col, new_value, old_value))
            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid = self.get_grid()
                grid.refresh_grid()
                self._app._update_status()
                self._app._mark_dirty()
                self.notify(f"Copied {len(changes)} cell(s)")
        except ValueError as e:
            self.notify(f"Invalid destination: {e}", severity="error")

    def menu_move(self) -> None:
        """Initiate a menu-based move operation."""
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        source_range = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
        self._app._pending_source_range = (r1, c1, r2, c2)
        self._app.push_screen(
            CommandInput(f"Move {source_range} TO (e.g., D1):"), self._do_menu_move
        )

    def _do_menu_move(self, result: str | None) -> None:
        if not result:
            return
        try:
            dest_row, dest_col = parse_cell_ref(result.upper())
            r1, c1, r2, c2 = self._app._pending_source_range
            changes = []
            for r_offset in range(r2 - r1 + 1):
                for c_offset in range(c2 - c1 + 1):
                    src_row, src_col = r1 + r_offset, c1 + c_offset
                    target_row, target_col = dest_row + r_offset, dest_col + c_offset
                    if (
                        target_row >= self.spreadsheet.rows
                        or target_col >= self.spreadsheet.cols
                    ):
                        continue
                    src_cell = self.spreadsheet.get_cell(src_row, src_col)
                    target_cell = self.spreadsheet.get_cell(target_row, target_col)
                    old_value = target_cell.raw_value
                    new_value = src_cell.raw_value
                    if new_value and (
                        new_value.startswith("=") or new_value.startswith("@")
                    ):
                        row_delta = target_row - src_row
                        col_delta = target_col - src_col
                        new_value = new_value[0] + adjust_formula_references(
                            new_value[1:], row_delta, col_delta
                        )
                    if new_value != old_value:
                        changes.append((target_row, target_col, new_value, old_value))
            for r_offset in range(r2 - r1 + 1):
                for c_offset in range(c2 - c1 + 1):
                    src_row, src_col = r1 + r_offset, c1 + c_offset
                    target_row, target_col = dest_row + r_offset, dest_col + c_offset
                    if src_row == target_row and src_col == target_col:
                        continue
                    src_cell = self.spreadsheet.get_cell(src_row, src_col)
                    if src_cell.raw_value:
                        changes.append((src_row, src_col, "", src_cell.raw_value))
            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid = self.get_grid()
                grid.clear_selection()
                grid.cursor_row = dest_row
                grid.cursor_col = dest_col
                grid.refresh_grid()
                self._app._update_status()
                self._app._mark_dirty()
                self.notify(f"Moved cells to {make_cell_ref(dest_row, dest_col)}")
        except ValueError as e:
            self.notify(f"Invalid destination: {e}", severity="error")

    def copy_cells(self) -> None:
        """Copy the current selection to clipboard."""
        grid = self.get_grid()
        r1, c1, r2, c2 = grid.selection_range
        self._app._range_clipboard = []
        self._app._clipboard_origin = (r1, c1)
        for r in range(r1, r2 + 1):
            row_data = []
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                row_data.append(cell.raw_value)
            self._app._range_clipboard.append(row_data)
        self._app._clipboard_is_cut = False
        cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)
        self._app._cell_clipboard = (grid.cursor_row, grid.cursor_col, cell.raw_value)
        cells_count = (r2 - r1 + 1) * (c2 - c1 + 1)
        self.notify(f"Copied {cells_count} cell(s)")

    def cut_cells(self) -> None:
        """Cut the current selection to clipboard."""
        self.copy_cells()
        self._app._clipboard_is_cut = True
        self.notify("Cut to clipboard")

    def paste_cells(self) -> None:
        """Paste from clipboard to current position."""
        if not self._app._range_clipboard:
            if self._app._cell_clipboard:
                grid = self.get_grid()
                cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)
                old_value = cell.raw_value
                new_value = self._app._cell_clipboard[2]
                if new_value.startswith("=") or new_value.startswith("@"):
                    row_delta = grid.cursor_row - self._app._cell_clipboard[0]
                    col_delta = grid.cursor_col - self._app._cell_clipboard[1]
                    new_value = new_value[0] + adjust_formula_references(
                        new_value[1:], row_delta, col_delta
                    )
                cmd = CellChangeCommand(
                    spreadsheet=self.spreadsheet,
                    row=grid.cursor_row,
                    col=grid.cursor_col,
                    new_value=new_value,
                    old_value=old_value,
                )
                self.undo_manager.execute(cmd)
                grid.refresh_grid()
                self._app._update_status()
                self._app._mark_dirty()
                self.notify("Pasted")
            return

        grid = self.get_grid()
        dest_row, dest_col = grid.cursor_row, grid.cursor_col
        src_row, src_col = getattr(self._app, "_clipboard_origin", (0, 0))
        changes = []
        for r_offset, row_data in enumerate(self._app._range_clipboard):
            for c_offset, value in enumerate(row_data):
                target_row = dest_row + r_offset
                target_col = dest_col + c_offset
                if (
                    target_row >= self.spreadsheet.rows
                    or target_col >= self.spreadsheet.cols
                ):
                    continue
                cell = self.spreadsheet.get_cell(target_row, target_col)
                old_value = cell.raw_value
                new_value = value
                if new_value and (
                    new_value.startswith("=") or new_value.startswith("@")
                ):
                    row_delta = target_row - (src_row + r_offset)
                    col_delta = target_col - (src_col + c_offset)
                    new_value = new_value[0] + adjust_formula_references(
                        new_value[1:], row_delta, col_delta
                    )
                if new_value != old_value:
                    changes.append((target_row, target_col, new_value, old_value))
        if changes:
            range_cmd = RangeChangeCommand(
                spreadsheet=self.spreadsheet, changes=changes
            )
            self.undo_manager.execute(range_cmd)
        was_cut = self._app._clipboard_is_cut
        if self._app._clipboard_is_cut:
            clear_changes = []
            for r_offset, row_data in enumerate(self._app._range_clipboard):
                for c_offset, value in enumerate(row_data):
                    if value:
                        clear_changes.append(
                            (src_row + r_offset, src_col + c_offset, "", value)
                        )
            if clear_changes:
                clear_cmd = RangeChangeCommand(
                    spreadsheet=self.spreadsheet, changes=clear_changes
                )
                self.undo_manager.execute(clear_cmd)
            self._app._clipboard_is_cut = False
        grid.refresh_grid()
        self._app._update_status()
        if changes or was_cut:
            self._app._mark_dirty()
        cells_count = (
            len(self._app._range_clipboard) * len(self._app._range_clipboard[0])
            if self._app._range_clipboard
            else 0
        )
        self.notify(f"Pasted {cells_count} cell(s)")

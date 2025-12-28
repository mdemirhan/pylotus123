"""Clipboard management for copy/cut/paste operations."""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


class ClipboardMode(Enum):
    """Mode of clipboard content."""
    EMPTY = auto()
    COPY = auto()
    CUT = auto()


@dataclass
class ClipboardCell:
    """A cell stored in the clipboard."""
    raw_value: str
    format_code: str = "G"
    is_formula: bool = False


@dataclass
class ClipboardContent:
    """Content stored in the clipboard."""
    mode: ClipboardMode = ClipboardMode.EMPTY
    # Source location
    source_row: int = 0
    source_col: int = 0
    # Dimensions
    num_rows: int = 0
    num_cols: int = 0
    # Cell data: dict of (relative_row, relative_col) -> ClipboardCell
    cells: dict[tuple[int, int], ClipboardCell] = field(default_factory=dict)


class Clipboard:
    """Manages clipboard for spreadsheet copy/cut/paste.

    Supports:
    - Single cell operations
    - Range operations
    - Copy with formula adjustment
    - Cut (move) operations
    """

    def __init__(self, spreadsheet: Spreadsheet) -> None:
        self.spreadsheet = spreadsheet
        self._content = ClipboardContent()

    @property
    def is_empty(self) -> bool:
        """Check if clipboard is empty."""
        return self._content.mode == ClipboardMode.EMPTY

    @property
    def mode(self) -> ClipboardMode:
        """Get current clipboard mode."""
        return self._content.mode

    @property
    def has_content(self) -> bool:
        """Check if clipboard has content to paste."""
        return self._content.mode != ClipboardMode.EMPTY and len(self._content.cells) > 0

    def copy_cell(self, row: int, col: int) -> None:
        """Copy a single cell to clipboard."""
        self.copy_range(row, col, row, col)

    def copy_range(self, start_row: int, start_col: int,
                   end_row: int, end_col: int) -> None:
        """Copy a range of cells to clipboard."""
        # Normalize range
        if start_row > end_row:
            start_row, end_row = end_row, start_row
        if start_col > end_col:
            start_col, end_col = end_col, start_col

        self._content = ClipboardContent(
            mode=ClipboardMode.COPY,
            source_row=start_row,
            source_col=start_col,
            num_rows=end_row - start_row + 1,
            num_cols=end_col - start_col + 1,
        )

        # Store cells
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                cell = self.spreadsheet.get_cell_if_exists(r, c)
                if cell and not cell.is_empty:
                    rel_r = r - start_row
                    rel_c = c - start_col
                    self._content.cells[(rel_r, rel_c)] = ClipboardCell(
                        raw_value=cell.raw_value,
                        format_code=cell.format_code,
                        is_formula=cell.is_formula,
                    )

    def cut_cell(self, row: int, col: int) -> None:
        """Cut a single cell to clipboard."""
        self.cut_range(row, col, row, col)

    def cut_range(self, start_row: int, start_col: int,
                  end_row: int, end_col: int) -> None:
        """Cut a range of cells to clipboard."""
        # First copy
        self.copy_range(start_row, start_col, end_row, end_col)
        # Mark as cut
        self._content.mode = ClipboardMode.CUT

    def paste(self, dest_row: int, dest_col: int,
              adjust_references: bool = True) -> list[tuple[int, int]]:
        """Paste clipboard content to destination.

        Args:
            dest_row: Destination row
            dest_col: Destination column
            adjust_references: Whether to adjust formula references

        Returns:
            List of (row, col) cells that were modified
        """
        if not self.has_content:
            return []

        modified = []
        row_delta = dest_row - self._content.source_row
        col_delta = dest_col - self._content.source_col

        for (rel_r, rel_c), clip_cell in self._content.cells.items():
            target_row = dest_row + rel_r
            target_col = dest_col + rel_c

            # Check bounds
            if target_row >= self.spreadsheet.rows or target_col >= self.spreadsheet.cols:
                continue

            # Check protection
            if self.spreadsheet.protection.is_cell_protected(target_row, target_col):
                continue

            # Prepare value
            value = clip_cell.raw_value
            if adjust_references and clip_cell.is_formula:
                from ..core.reference import adjust_formula_references
                value = adjust_formula_references(
                    value, row_delta, col_delta,
                    self.spreadsheet.rows - 1,
                    self.spreadsheet.cols - 1
                )

            # Set cell
            cell = self.spreadsheet.get_cell(target_row, target_col)
            cell.set_value(value)
            cell.format_code = clip_cell.format_code
            modified.append((target_row, target_col))

        # If cut, clear source cells
        if self._content.mode == ClipboardMode.CUT:
            for (rel_r, rel_c) in self._content.cells.keys():
                src_row = self._content.source_row + rel_r
                src_col = self._content.source_col + rel_c

                # Don't clear if pasting over source
                if (src_row, src_col) in modified:
                    continue

                cell = self.spreadsheet.get_cell_if_exists(src_row, src_col)
                if cell:
                    cell.set_value("")

            # Clear clipboard after cut-paste
            self.clear()

        self.spreadsheet._invalidate_cache()
        return modified

    def paste_special(self, dest_row: int, dest_col: int,
                      values_only: bool = False,
                      formats_only: bool = False,
                      transpose: bool = False) -> list[tuple[int, int]]:
        """Paste with special options.

        Args:
            dest_row: Destination row
            dest_col: Destination column
            values_only: Only paste computed values, not formulas
            formats_only: Only paste formats, not values
            transpose: Swap rows and columns

        Returns:
            List of modified cells
        """
        if not self.has_content:
            return []

        modified = []

        for (rel_r, rel_c), clip_cell in self._content.cells.items():
            # Apply transpose
            if transpose:
                target_row = dest_row + rel_c
                target_col = dest_col + rel_r
            else:
                target_row = dest_row + rel_r
                target_col = dest_col + rel_c

            # Check bounds
            if target_row >= self.spreadsheet.rows or target_col >= self.spreadsheet.cols:
                continue

            # Check protection
            if self.spreadsheet.protection.is_cell_protected(target_row, target_col):
                continue

            cell = self.spreadsheet.get_cell(target_row, target_col)

            if formats_only:
                cell.format_code = clip_cell.format_code
            elif values_only:
                # Evaluate formula and paste result
                if clip_cell.is_formula:
                    src_row = self._content.source_row + rel_r
                    src_col = self._content.source_col + rel_c
                    value = self.spreadsheet.get_value(src_row, src_col)
                    cell.set_value(str(value))
                else:
                    cell.set_value(clip_cell.raw_value)
            else:
                # Normal paste
                cell.set_value(clip_cell.raw_value)
                cell.format_code = clip_cell.format_code

            modified.append((target_row, target_col))

        self.spreadsheet._invalidate_cache()
        return modified

    def clear(self) -> None:
        """Clear the clipboard."""
        self._content = ClipboardContent()

    @property
    def size(self) -> tuple[int, int]:
        """Get dimensions of clipboard content (rows, cols)."""
        return (self._content.num_rows, self._content.num_cols)

    @property
    def source_range(self) -> tuple[int, int, int, int] | None:
        """Get source range (start_row, start_col, end_row, end_col)."""
        if self.is_empty:
            return None
        return (
            self._content.source_row,
            self._content.source_col,
            self._content.source_row + self._content.num_rows - 1,
            self._content.source_col + self._content.num_cols - 1,
        )

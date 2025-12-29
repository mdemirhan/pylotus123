"""Spreadsheet grid widget for displaying and editing cells.

Provides the main grid display with mouse support, keyboard navigation,
and range selection capabilities.
"""

from __future__ import annotations

from typing import Any

from rich.style import Style
from rich.text import Text
from textual import events
from textual.message import Message
from textual.reactive import reactive
from textual.widgets import Static

from ..core import Spreadsheet, index_to_col, parse_cell_ref
from .themes import Theme


class SpreadsheetGrid(Static, can_focus=True):
    """The main spreadsheet grid display with mouse support and range selection."""

    cursor_row = reactive(0)
    cursor_col = reactive(0)
    scroll_row = reactive(0)
    scroll_col = reactive(0)
    # Range selection (anchor point for shift+arrow selection)
    select_anchor_row = reactive(-1)
    select_anchor_col = reactive(-1)
    # Freeze titles (number of rows/cols to keep visible)
    freeze_rows = reactive(0)
    freeze_cols = reactive(0)

    class CellSelected(Message):
        """Sent when a cell is selected via keyboard."""

        def __init__(self, row: int, col: int) -> None:
            self.row = row
            self.col = col
            super().__init__()

    class CellClicked(Message):
        """Sent when a cell is clicked with mouse."""

        def __init__(self, row: int, col: int) -> None:
            self.row = row
            self.col = col
            super().__init__()

    class RangeSelected(Message):
        """Sent when a range of cells is selected."""

        def __init__(self, start_row: int, start_col: int, end_row: int, end_col: int) -> None:
            self.start_row = start_row
            self.start_col = start_col
            self.end_row = end_row
            self.end_col = end_col
            super().__init__()

    def __init__(self, spreadsheet: Spreadsheet, theme: Theme, **kwargs: Any) -> None:
        super().__init__(**kwargs)
        self.spreadsheet = spreadsheet
        self.theme = theme
        self._visible_rows = 20
        self._visible_cols = 8
        self._col_positions: list[tuple[int, int, int]] = []
        self._row_positions: list[tuple[int, int]] = []
        self.default_col_width = 10
        self.show_zero = True  # Whether to display zero values

    @property
    def has_selection(self) -> bool:
        """Check if there's an active range selection."""
        return self.select_anchor_row >= 0 and self.select_anchor_col >= 0

    @property
    def selection_range(self) -> tuple[int, int, int, int]:
        """Get the normalized selection range (top-left to bottom-right)."""
        if not self.has_selection:
            return (self.cursor_row, self.cursor_col, self.cursor_row, self.cursor_col)
        r1, c1 = (
            min(self.select_anchor_row, self.cursor_row),
            min(self.select_anchor_col, self.cursor_col),
        )
        r2, c2 = (
            max(self.select_anchor_row, self.cursor_row),
            max(self.select_anchor_col, self.cursor_col),
        )
        return (r1, c1, r2, c2)

    def start_selection(self) -> None:
        """Start a range selection from current cursor position."""
        self.select_anchor_row = self.cursor_row
        self.select_anchor_col = self.cursor_col

    def clear_selection(self) -> None:
        """Clear the range selection."""
        self.select_anchor_row = -1
        self.select_anchor_col = -1

    def is_in_selection(self, row: int, col: int) -> bool:
        """Check if a cell is within the current selection."""
        if not self.has_selection:
            return row == self.cursor_row and col == self.cursor_col
        r1, c1, r2, c2 = self.selection_range
        return r1 <= row <= r2 and c1 <= col <= c2

    @property
    def visible_rows(self) -> int:
        return self._visible_rows

    @property
    def visible_cols(self) -> int:
        return self._visible_cols

    def on_mount(self) -> None:
        self._calculate_visible_area()
        self.refresh_grid()

    def on_resize(self, event: events.Resize) -> None:
        self._calculate_visible_area()
        self.refresh_grid()

    def _calculate_visible_area(self) -> None:
        """Calculate how many rows/cols fit in the current size."""
        if self.size.height > 2:
            self._visible_rows = self.size.height - 2
        if self.size.width > 6:
            # Calculate visible columns based on actual column widths
            used_width = 5  # Row number area (4 chars + 1 border)
            visible_cols = 0
            for c in range(self.scroll_col, self.spreadsheet.cols):
                col_width = self.spreadsheet.get_col_width(c) + 1  # +1 for border
                if used_width + col_width > self.size.width:
                    break
                used_width += col_width
                visible_cols += 1
            self._visible_cols = max(1, visible_cols)

    def recalculate_visible_area(self) -> None:
        """Public method to recalculate visible area after column width changes."""
        self._calculate_visible_area()

    def watch_cursor_row(self, value: int) -> None:
        self._ensure_visible()
        self.refresh_grid()
        self.post_message(self.CellSelected(value, self.cursor_col))

    def watch_cursor_col(self, value: int) -> None:
        self._ensure_visible()
        self.refresh_grid()
        self.post_message(self.CellSelected(self.cursor_row, value))

    def watch_scroll_row(self, value: int) -> None:
        self.refresh_grid()

    def watch_scroll_col(self, value: int) -> None:
        self._calculate_visible_area()
        self.refresh_grid()

    def _ensure_visible(self) -> None:
        """Ensure cursor is within visible area, scrolling if needed."""
        if self.cursor_row < self.scroll_row:
            self.scroll_row = self.cursor_row
        elif self.cursor_row >= self.scroll_row + self._visible_rows:
            self.scroll_row = self.cursor_row - self._visible_rows + 1

        if self.cursor_col < self.scroll_col:
            self.scroll_col = self.cursor_col
        elif self.cursor_col >= self.scroll_col + self._visible_cols:
            self.scroll_col = self.cursor_col - self._visible_cols + 1

    def set_theme(self, theme: Theme) -> None:
        """Update the grid's theme."""
        self.theme = theme
        self.refresh_grid()

    def refresh_grid(self) -> None:
        """Redraw the grid."""
        lines = []
        t = self.theme

        header_style = Style(color=t.header_fg, bgcolor=t.header_bg, bold=True)
        cell_style = Style(color=t.cell_fg, bgcolor=t.cell_bg)
        selected_style = Style(color=t.selected_fg, bgcolor=t.selected_bg, bold=True)
        border_style = Style(color=t.border, bgcolor=t.cell_bg)

        self._col_positions = []
        self._row_positions = []

        # Header row
        header = Text()
        header.append("    ", header_style)
        header.append("\u2502", border_style)

        x_pos = 5
        for c in range(self.scroll_col, self.scroll_col + self._visible_cols):
            if c >= self.spreadsheet.cols:
                break
            col_width = self.spreadsheet.get_col_width(c)
            col_name = index_to_col(c)
            header.append(col_name.center(col_width), header_style)
            header.append("\u2502", border_style)
            self._col_positions.append((c, x_pos, x_pos + col_width))
            x_pos += col_width + 1
        lines.append(header)

        # Separator
        sep = Text()
        sep.append("\u2500" * 4, border_style)
        sep.append("\u253c", border_style)
        for c in range(self.scroll_col, self.scroll_col + self._visible_cols):
            if c >= self.spreadsheet.cols:
                break
            col_width = self.spreadsheet.get_col_width(c)
            sep.append("\u2500" * col_width, border_style)
            sep.append("\u253c", border_style)
        lines.append(sep)

        # Data rows
        for row_idx, r in enumerate(range(self.scroll_row, self.scroll_row + self._visible_rows)):
            if r >= self.spreadsheet.rows:
                break
            row_text = Text()
            row_num = str(r + 1).rjust(4)
            row_text.append(row_num, header_style)
            row_text.append("\u2502", border_style)

            self._row_positions.append((r, row_idx + 2))

            for c in range(self.scroll_col, self.scroll_col + self._visible_cols):
                if c >= self.spreadsheet.cols:
                    break
                col_width = self.spreadsheet.get_col_width(c)
                value = self.spreadsheet.get_display_value(r, c)
                # Hide zero values if show_zero is False
                if not self.show_zero and value in ("0", "0.0", "0.00"):
                    value = ""
                display = value[:col_width].ljust(col_width)

                if self.is_in_selection(r, c):
                    row_text.append(display, selected_style)
                else:
                    row_text.append(display, cell_style)
                row_text.append("\u2502", border_style)

            lines.append(row_text)

        # Build content
        content = Text()
        for i, line in enumerate(lines):
            content.append_text(line)
            if i < len(lines) - 1:
                content.append("\n")
        self.update(content)
        self.refresh()

    def on_click(self, event: events.Click) -> None:
        """Handle mouse clicks to select cells."""
        click_x = event.x
        click_y = event.y

        clicked_row = None
        for row_idx, y in self._row_positions:
            if y == click_y:
                clicked_row = row_idx
                break

        clicked_col = None
        for col_idx, start_x, end_x in self._col_positions:
            if start_x <= click_x < end_x:
                clicked_col = col_idx
                break

        if clicked_row is not None and clicked_col is not None:
            self.cursor_row = clicked_row
            self.cursor_col = clicked_col
            self.post_message(self.CellClicked(clicked_row, clicked_col))

    def move_cursor(self, dr: int, dc: int) -> None:
        """Move cursor by delta rows/columns."""
        new_row = max(0, min(self.spreadsheet.rows - 1, self.cursor_row + dr))
        new_col = max(0, min(self.spreadsheet.cols - 1, self.cursor_col + dc))
        self.cursor_row = new_row
        self.cursor_col = new_col

    def goto_cell(self, ref: str) -> None:
        """Go to a cell by reference (e.g., 'A1')."""
        try:
            row, col = parse_cell_ref(ref)
            self.cursor_row = max(0, min(self.spreadsheet.rows - 1, row))
            self.cursor_col = max(0, min(self.spreadsheet.cols - 1, col))
        except ValueError:
            pass

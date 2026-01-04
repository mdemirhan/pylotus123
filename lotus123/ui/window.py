"""Window splitting and frozen titles support.

Implements Lotus 1-2-3 style window management:
- Horizontal and vertical window splitting
- Frozen row/column titles
- Synchronized and unsynchronized scrolling
"""

from dataclasses import dataclass
from enum import Enum, auto

from ..core.spreadsheet_protocol import SpreadsheetProtocol


class SplitType(Enum):
    """Type of window split."""

    NONE = auto()
    HORIZONTAL = auto()
    VERTICAL = auto()


class TitleFreezeType(Enum):
    """Type of frozen titles."""

    NONE = auto()
    HORIZONTAL = auto()  # Freeze rows (horizontal titles)
    VERTICAL = auto()  # Freeze columns (vertical titles)
    BOTH = auto()  # Freeze both rows and columns


@dataclass
class ViewPort:
    """A viewport into the spreadsheet.

    Represents a visible portion of the spreadsheet with its own
    scroll position.
    """

    top_row: int = 0
    left_col: int = 0
    visible_rows: int = 20
    visible_cols: int = 10
    cursor_row: int = 0
    cursor_col: int = 0

    def scroll_to(self, row: int, col: int) -> None:
        """Scroll to make cell visible."""
        # Adjust vertical scroll
        if row < self.top_row:
            self.top_row = row
        elif row >= self.top_row + self.visible_rows:
            self.top_row = row - self.visible_rows + 1

        # Adjust horizontal scroll
        if col < self.left_col:
            self.left_col = col
        elif col >= self.left_col + self.visible_cols:
            self.left_col = col - self.visible_cols + 1

    def move_cursor(self, row: int, col: int) -> None:
        """Move cursor and scroll if needed."""
        self.cursor_row = row
        self.cursor_col = col
        self.scroll_to(row, col)

    def is_visible(self, row: int, col: int) -> bool:
        """Check if cell is visible in this viewport."""
        return (
            self.top_row <= row < self.top_row + self.visible_rows
            and self.left_col <= col < self.left_col + self.visible_cols
        )

    def get_visible_range(self) -> tuple[int, int, int, int]:
        """Get visible range as (start_row, end_row, start_col, end_col)."""
        return (
            self.top_row,
            self.top_row + self.visible_rows - 1,
            self.left_col,
            self.left_col + self.visible_cols - 1,
        )


@dataclass
class FrozenTitles:
    """Frozen row and column titles configuration."""

    freeze_type: TitleFreezeType = TitleFreezeType.NONE
    frozen_rows: int = 0  # Number of rows frozen at top
    frozen_cols: int = 0  # Number of columns frozen at left
    freeze_row: int = 0  # Row where freeze was applied (for reference)
    freeze_col: int = 0  # Column where freeze was applied

    def freeze_horizontal(self, at_row: int) -> None:
        """Freeze rows above current row."""
        self.frozen_rows = at_row
        self.freeze_row = at_row
        if self.frozen_cols > 0:
            self.freeze_type = TitleFreezeType.BOTH
        else:
            self.freeze_type = TitleFreezeType.HORIZONTAL

    def freeze_vertical(self, at_col: int) -> None:
        """Freeze columns to the left of current column."""
        self.frozen_cols = at_col
        self.freeze_col = at_col
        if self.frozen_rows > 0:
            self.freeze_type = TitleFreezeType.BOTH
        else:
            self.freeze_type = TitleFreezeType.VERTICAL

    def freeze_both(self, at_row: int, at_col: int) -> None:
        """Freeze both rows and columns."""
        self.frozen_rows = at_row
        self.frozen_cols = at_col
        self.freeze_row = at_row
        self.freeze_col = at_col
        self.freeze_type = TitleFreezeType.BOTH

    def clear(self) -> None:
        """Clear all frozen titles."""
        self.freeze_type = TitleFreezeType.NONE
        self.frozen_rows = 0
        self.frozen_cols = 0
        self.freeze_row = 0
        self.freeze_col = 0

    @property
    def has_frozen_rows(self) -> bool:
        return self.frozen_rows > 0

    @property
    def has_frozen_cols(self) -> bool:
        return self.frozen_cols > 0


@dataclass
class WindowSplit:
    """Window split configuration."""

    split_type: SplitType = SplitType.NONE
    split_position: int = 0  # Row for horizontal, column for vertical
    synchronized: bool = True  # Synchronize scrolling between panes

    def split_horizontal(self, at_row: int) -> None:
        """Split window horizontally at row."""
        self.split_type = SplitType.HORIZONTAL
        self.split_position = at_row

    def split_vertical(self, at_col: int) -> None:
        """Split window vertically at column."""
        self.split_type = SplitType.VERTICAL
        self.split_position = at_col

    def clear(self) -> None:
        """Remove window split."""
        self.split_type = SplitType.NONE
        self.split_position = 0

    @property
    def is_split(self) -> bool:
        return self.split_type != SplitType.NONE


class WindowManager:
    """Manages window splitting, frozen titles, and viewports.

    Provides Lotus 1-2-3 style window management:
    - /Worksheet Titles - Freeze rows/columns
    - /Worksheet Window - Split view
    """

    def __init__(self, spreadsheet: SpreadsheetProtocol | None = None) -> None:
        self.spreadsheet = spreadsheet
        self.titles = FrozenTitles()
        self.split = WindowSplit()

        # Main viewport (or top/left pane when split)
        self.primary = ViewPort()
        # Secondary viewport (bottom/right pane when split)
        self.secondary = ViewPort()

        # Which pane is active (0 = primary, 1 = secondary)
        self.active_pane: int = 0

    def freeze_titles_horizontal(self, at_row: int) -> None:
        """Freeze rows at the top (horizontal titles).

        Args:
            at_row: Rows 0 to at_row-1 will be frozen
        """
        self.titles.freeze_horizontal(at_row)

    def freeze_titles_vertical(self, at_col: int) -> None:
        """Freeze columns at the left (vertical titles).

        Args:
            at_col: Columns 0 to at_col-1 will be frozen
        """
        self.titles.freeze_vertical(at_col)

    def freeze_titles_both(self, at_row: int, at_col: int) -> None:
        """Freeze both rows and columns.

        Args:
            at_row: Rows 0 to at_row-1 will be frozen
            at_col: Columns 0 to at_col-1 will be frozen
        """
        self.titles.freeze_both(at_row, at_col)

    def clear_titles(self) -> None:
        """Unfreeze all titles."""
        self.titles.clear()

    def split_horizontal(self, at_row: int) -> None:
        """Split window horizontally.

        Creates two panes stacked vertically. The top pane shows
        rows before at_row, bottom pane shows rows from at_row.

        Args:
            at_row: Row where to split
        """
        self.split.split_horizontal(at_row)
        # Configure secondary viewport for bottom pane
        self.secondary.top_row = at_row
        self.secondary.cursor_row = at_row
        self.secondary.left_col = self.primary.left_col
        self.secondary.cursor_col = self.primary.cursor_col

    def split_vertical(self, at_col: int) -> None:
        """Split window vertically.

        Creates two panes side by side. Left pane shows columns
        before at_col, right pane shows columns from at_col.

        Args:
            at_col: Column where to split
        """
        self.split.split_vertical(at_col)
        # Configure secondary viewport for right pane
        self.secondary.left_col = at_col
        self.secondary.cursor_col = at_col
        self.secondary.top_row = self.primary.top_row
        self.secondary.cursor_row = self.primary.cursor_row

    def clear_split(self) -> None:
        """Remove window split."""
        self.split.clear()
        self.active_pane = 0

    def sync_scrolling(self) -> None:
        """Enable synchronized scrolling between panes."""
        self.split.synchronized = True

    def unsync_scrolling(self) -> None:
        """Disable synchronized scrolling between panes."""
        self.split.synchronized = False

    def switch_pane(self) -> None:
        """Switch to the other pane."""
        if self.split.is_split:
            self.active_pane = 1 - self.active_pane

    @property
    def active_viewport(self) -> ViewPort:
        """Get the currently active viewport."""
        if self.active_pane == 0:
            return self.primary
        return self.secondary

    def scroll(self, delta_row: int = 0, delta_col: int = 0) -> None:
        """Scroll the active viewport.

        If synchronized, also scrolls the other pane in the
        appropriate direction.
        """
        viewport = self.active_viewport
        viewport.top_row = max(0, viewport.top_row + delta_row)
        viewport.left_col = max(0, viewport.left_col + delta_col)

        # Handle synchronized scrolling
        if self.split.is_split and self.split.synchronized:
            other = self.secondary if self.active_pane == 0 else self.primary
            if self.split.split_type == SplitType.HORIZONTAL:
                # Horizontal split: sync horizontal scrolling
                other.left_col = viewport.left_col
            else:
                # Vertical split: sync vertical scrolling
                other.top_row = viewport.top_row

    def move_cursor(self, row: int, col: int) -> None:
        """Move cursor in active viewport."""
        # Check bounds with spreadsheet if available
        if self.spreadsheet:
            row = max(0, min(row, self.spreadsheet.rows - 1))
            col = max(0, min(col, self.spreadsheet.cols - 1))
        else:
            row = max(0, row)
            col = max(0, col)

        # Handle frozen titles - cursor can't enter frozen area
        if row < self.titles.frozen_rows:
            row = self.titles.frozen_rows
        if col < self.titles.frozen_cols:
            col = self.titles.frozen_cols

        self.active_viewport.move_cursor(row, col)

    def get_cursor_position(self) -> tuple[int, int]:
        """Get current cursor position."""
        vp = self.active_viewport
        return (vp.cursor_row, vp.cursor_col)

    def get_pane_count(self) -> int:
        """Get number of visible panes."""
        if self.split.is_split:
            return 2
        return 1

    def get_visible_regions(self) -> list[dict]:
        """Get description of all visible regions.

        Returns a list of region dictionaries for rendering.
        Each region has:
        - type: 'frozen_corner', 'frozen_row', 'frozen_col', 'main'
        - viewport: ViewPort for scrolling region
        - fixed_rows: True if rows don't scroll
        - fixed_cols: True if columns don't scroll
        - row_range: (start, end) row indices
        - col_range: (start, end) column indices
        """
        regions: list = []
        vp = self.active_viewport

        # Frozen corner (if both frozen)
        if self.titles.has_frozen_rows and self.titles.has_frozen_cols:
            regions.append(
                {
                    "type": "frozen_corner",
                    "viewport": None,
                    "fixed_rows": True,
                    "fixed_cols": True,
                    "row_range": (0, self.titles.frozen_rows - 1),
                    "col_range": (0, self.titles.frozen_cols - 1),
                }
            )

        # Frozen rows (top)
        if self.titles.has_frozen_rows:
            regions.append(
                {
                    "type": "frozen_row",
                    "viewport": vp,
                    "fixed_rows": True,
                    "fixed_cols": False,
                    "row_range": (0, self.titles.frozen_rows - 1),
                    "col_range": (vp.left_col, vp.left_col + vp.visible_cols - 1),
                }
            )

        # Frozen columns (left)
        if self.titles.has_frozen_cols:
            regions.append(
                {
                    "type": "frozen_col",
                    "viewport": vp,
                    "fixed_rows": False,
                    "fixed_cols": True,
                    "row_range": (vp.top_row, vp.top_row + vp.visible_rows - 1),
                    "col_range": (0, self.titles.frozen_cols - 1),
                }
            )

        # Main scrolling area
        main_start_row = vp.top_row
        main_start_col = vp.left_col
        if self.titles.has_frozen_rows:
            # Main area accounts for frozen rows visually
            pass  # Row range is based on viewport scroll position
        if self.titles.has_frozen_cols:
            # Main area accounts for frozen cols visually
            pass

        regions.append(
            {
                "type": "main",
                "viewport": vp,
                "fixed_rows": False,
                "fixed_cols": False,
                "row_range": (main_start_row, main_start_row + vp.visible_rows - 1),
                "col_range": (main_start_col, main_start_col + vp.visible_cols - 1),
                "is_active": self.active_pane == 0,
            }
        )

        # Secondary pane if split
        if self.split.is_split:
            sec = self.secondary
            regions.append(
                {
                    "type": "secondary",
                    "viewport": sec,
                    "fixed_rows": False,
                    "fixed_cols": False,
                    "row_range": (sec.top_row, sec.top_row + sec.visible_rows - 1),
                    "col_range": (sec.left_col, sec.left_col + sec.visible_cols - 1),
                    "is_active": self.active_pane == 1,
                }
            )

        return regions

    def resize(self, total_rows: int, total_cols: int) -> None:
        """Handle window resize.

        Args:
            total_rows: Total visible rows for spreadsheet area
            total_cols: Total visible columns for spreadsheet area
        """
        if self.split.is_split:
            if self.split.split_type == SplitType.HORIZONTAL:
                # Split rows between panes
                split_at = min(self.split.split_position, total_rows - 1)
                self.primary.visible_rows = split_at
                self.secondary.visible_rows = total_rows - split_at
                self.primary.visible_cols = total_cols
                self.secondary.visible_cols = total_cols
            else:
                # Split columns between panes
                split_at = min(self.split.split_position, total_cols - 1)
                self.primary.visible_cols = split_at
                self.secondary.visible_cols = total_cols - split_at
                self.primary.visible_rows = total_rows
                self.secondary.visible_rows = total_rows
        else:
            self.primary.visible_rows = total_rows
            self.primary.visible_cols = total_cols

    def get_status(self) -> str:
        """Get status string for display."""
        parts = []

        if self.titles.freeze_type != TitleFreezeType.NONE:
            if self.titles.freeze_type == TitleFreezeType.BOTH:
                parts.append(f"Titles:{self.titles.frozen_rows}R,{self.titles.frozen_cols}C")
            elif self.titles.freeze_type == TitleFreezeType.HORIZONTAL:
                parts.append(f"Titles:{self.titles.frozen_rows}R")
            else:
                parts.append(f"Titles:{self.titles.frozen_cols}C")

        if self.split.is_split:
            if self.split.split_type == SplitType.HORIZONTAL:
                parts.append(f"Split:H@{self.split.split_position}")
            else:
                parts.append(f"Split:V@{self.split.split_position}")
            if not self.split.synchronized:
                parts.append("Unsync")

        return " ".join(parts) if parts else ""

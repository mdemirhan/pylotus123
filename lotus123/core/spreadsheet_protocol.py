"""Protocol for Spreadsheet to avoid circular imports.

This protocol defines the interface that modules can depend on
without importing the full Spreadsheet class.
"""

from typing import Any, Iterator, Protocol

from .cell import Cell, TextAlignment
from .named_ranges import NamedRangeManager
from ..formula.context import EvaluationContext
from ..formula.recalc_types import RecalcMode, RecalcOrder


class SpreadsheetProtocol(Protocol):
    """Protocol defining the Spreadsheet interface.

    This allows other modules to type-hint Spreadsheet without
    creating circular import dependencies.
    """

    # Properties
    rows: int
    cols: int
    modified: bool
    filename: str
    named_ranges: NamedRangeManager
    global_settings: dict[str, Any]
    frozen_rows: int
    frozen_cols: int

    # Read-only properties (sparse storage access)
    @property
    def cells(self) -> dict[tuple[int, int], Cell]:
        """Read-only access to the sparse cell storage."""
        ...

    @property
    def col_widths(self) -> dict[int, int]:
        """Read-only access to column widths (sparse, non-default only)."""
        ...

    @property
    def row_heights(self) -> dict[int, int]:
        """Read-only access to row heights (sparse, non-default only)."""
        ...

    @property
    def needs_recalc(self) -> bool:
        """Check if spreadsheet needs recalculation."""
        ...

    @property
    def has_circular_refs(self) -> bool:
        """Check if circular references were detected."""
        ...

    # Cell operations
    def get_cell(self, row: int, col: int) -> Cell:
        """Get or create a cell at the given position."""
        ...

    def get_cell_if_exists(self, row: int, col: int) -> Cell | None:
        """Get cell if it exists, without creating."""
        ...

    def set_cell(self, row: int, col: int, value: str) -> None:
        """Set cell value and invalidate cache."""
        ...

    def get_value(self, row: int, col: int, context: Any = None) -> Any:
        """Get computed value of cell."""
        ...

    def get_value_by_ref(self, ref: str, context: EvaluationContext | None = None) -> Any:
        """Get value by cell reference string."""
        ...

    def set_cell_data(self, row: int, col: int, data: dict[str, Any]) -> None:
        """Set cell data from a dictionary."""
        ...

    def iter_cells(self) -> Iterator[tuple[int, int, Cell]]:
        """Iterate over all cells as (row, col, cell) tuples."""
        ...

    # Display operations
    def get_display_value(self, row: int, col: int) -> str:
        """Get formatted string for display."""
        ...

    def get_cell_alignment(self, row: int, col: int) -> TextAlignment:
        """Get the alignment for a cell."""
        ...

    # Column operations
    def get_col_width(self, col: int) -> int:
        """Get width of a column."""
        ...

    def set_col_width(self, col: int, width: int) -> None:
        """Set width of a column."""
        ...

    def get_cells_in_col(self, col: int) -> list[tuple[int, Cell]]:
        """Get all cells in a column."""
        ...

    def get_cells_in_row(self, row: int) -> list[tuple[int, Cell]]:
        """Get all cells in a row."""
        ...

    # Row operations
    def get_row_height(self, row: int) -> int:
        """Get height of a row."""
        ...

    def set_row_height(self, row: int, height: int) -> None:
        """Set height of a row."""
        ...

    def get_all_col_widths(self) -> dict[int, int]:
        """Get all non-default column widths."""
        ...

    def set_all_col_widths(self, widths: dict[int, int]) -> None:
        """Set column widths from a dictionary."""
        ...

    def get_all_row_heights(self) -> dict[int, int]:
        """Get all non-default row heights."""
        ...

    def set_all_row_heights(self, heights: dict[int, int]) -> None:
        """Set row heights from a dictionary."""
        ...

    # Row/column insertion/deletion
    def insert_row(self, row: int) -> None:
        """Insert a row at the given position."""
        ...

    def delete_row(self, row: int) -> None:
        """Delete a row at the given position."""
        ...

    def insert_col(self, col: int) -> None:
        """Insert a column at the given position."""
        ...

    def delete_col(self, col: int) -> None:
        """Delete a column at the given position."""
        ...

    # Range operations
    def get_range(
        self, start_ref: str, end_ref: str, context: EvaluationContext | None = None
    ) -> list[list[Any]]:
        """Get values from a range as 2D list."""
        ...

    def get_range_flat(
        self, start_ref: str, end_ref: str, context: EvaluationContext | None = None
    ) -> list[Any]:
        """Get values from a range as flat list."""
        ...

    def get_used_range(self) -> tuple[tuple[int, int], tuple[int, int]] | None:
        """Get the bounding box of used cells."""
        ...

    # Cache and dependency operations
    def invalidate_cache(self) -> None:
        """Invalidate the entire computation cache."""
        ...

    def invalidate_cell_cache(self, row: int, col: int) -> None:
        """Invalidate cache for a specific cell."""
        ...

    def clear_cache(self) -> None:
        """Clear the entire computation cache."""
        ...

    def update_cell_dependency(self, row: int, col: int, formula: str | None) -> None:
        """Update cell dependency tracking."""
        ...

    def rebuild_dependency_graph(self) -> None:
        """Rebuild the dependency graph."""
        ...

    # Recalculation
    def recalculate(self) -> None:
        """Recalculate all cells."""
        ...

    def set_recalc_mode(self, mode: RecalcMode) -> None:
        """Set the recalculation mode."""
        ...

    def get_recalc_mode(self) -> RecalcMode:
        """Get the current recalculation mode."""
        ...

    def set_recalc_order(self, order: RecalcOrder) -> None:
        """Set the recalculation order."""
        ...

    def get_recalc_order(self) -> RecalcOrder:
        """Get the current recalculation order."""
        ...

    # Save/load
    def save(self, filename: str) -> None:
        """Save spreadsheet to a file."""
        ...

    def load(self, filename: str) -> None:
        """Load spreadsheet from a file."""
        ...

    def clear(self) -> None:
        """Clear all cells."""
        ...

    # Copy operations
    def copy_cell(
        self, from_row: int, from_col: int, to_row: int, to_col: int, adjust_refs: bool = True
    ) -> None:
        """Copy a cell to another location."""
        ...

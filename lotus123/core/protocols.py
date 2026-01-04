"""Protocols for decoupling core components.

These protocols define interfaces that allow components to interact
without tight coupling to concrete implementations.
"""

from typing import Any, Iterator, Protocol

from .cell import Cell


class CellStore(Protocol):
    """Protocol defining cell storage operations needed by RecalcEngine.

    This allows RecalcEngine to work with any cell storage implementation
    without tight coupling to the Spreadsheet class.
    """

    def get_cell_if_exists(self, row: int, col: int) -> Cell | None:
        """Get cell if it exists, without creating."""
        ...

    def get_value(self, row: int, col: int, context: Any = None) -> Any:
        """Get computed value of cell."""
        ...

    def iter_cells(self) -> Iterator[tuple[int, int, "Cell"]]:
        """Iterate over all cells as (row, col, cell) tuples."""
        ...

    def invalidate_cell_cache(self, row: int, col: int) -> None:
        """Invalidate cache for a specific cell."""
        ...

    def clear_cache(self) -> None:
        """Clear the entire computation cache."""
        ...

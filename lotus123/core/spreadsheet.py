"""Main spreadsheet class with expanded dimensions and features."""

from __future__ import annotations

import math
from typing import TYPE_CHECKING, Any, Iterator

from .cell import Cell, TextAlignment
from .errors import FormulaError
from .formatting import format_value, parse_format_code
from .named_ranges import NamedRangeManager
from .reference import adjust_for_structural_change, adjust_formula_references, parse_cell_ref

if TYPE_CHECKING:
    from ..formula.recalc import RecalcEngine, RecalcMode, RecalcOrder
    from ..formula.evaluator import EvaluationContext


# Lotus 1-2-3 dimensions
MAX_ROWS = 65536  # 65,536 rows (early versions had 8,192)
MAX_COLS = 256  # 256 columns (A through IV)

# Default column width
DEFAULT_COL_WIDTH = 10
MIN_COL_WIDTH = 3
MAX_COL_WIDTH = 50

# Default row height
DEFAULT_ROW_HEIGHT = 1
MIN_ROW_HEIGHT = 1
MAX_ROW_HEIGHT = 72


class Spreadsheet:
    """Main spreadsheet class managing a grid of cells.

    Supports Lotus 1-2-3 compatible features:
    - 256 columns (A through IV) x 65,536 rows
    - Cell formulas with circular reference detection
    - Named ranges
    - Multiple recalculation modes
    - Row heights and column widths

    Cells are stored sparsely - only non-empty cells use memory.
    """

    def __init__(self, rows: int = MAX_ROWS, cols: int = MAX_COLS) -> None:
        """Initialize spreadsheet.

        Args:
            rows: Number of rows (default 65,536)
            cols: Number of columns (default 256)
        """
        self.rows = min(rows, MAX_ROWS)
        self.cols = min(cols, MAX_COLS)

        # Sparse cell storage
        self._cells: dict[tuple[int, int], Cell] = {}

        # Sparse row/column indices for O(k) row/column access
        self._row_index: dict[int, set[int]] = {}  # row -> set of col indices
        self._col_index: dict[int, set[int]] = {}  # col -> set of row indices

        # Cache for computed values
        self._cache: dict[tuple[int, int], Any] = {}

        # Track cells currently being computed (for circular reference detection)
        self._computing: set[tuple[int, int]] = set()

        # Column widths (sparse - only store non-default)
        self._col_widths: dict[int, int] = {}

        # Row heights (sparse - only store non-default)
        self._row_heights: dict[int, int] = {}

        # Named ranges
        self.named_ranges = NamedRangeManager()

        # File info
        self.filename: str = ""
        self.modified: bool = False

        # Frozen rows/cols (titles)
        self.frozen_rows: int = 0
        self.frozen_cols: int = 0

        # Recalculation engine (created lazily to avoid circular imports)
        self._recalc_engine: RecalcEngine | None = None
        self._init_recalc_engine()

        # Circular references detected
        self._circular_refs: set[tuple[int, int]] = set()

        # Global settings (synced from app for persistence)
        self.global_settings: dict[str, Any] = {
            "format_code": "G",
            "label_prefix": "'",
            "default_col_width": DEFAULT_COL_WIDTH,
            "zero_display": True,
        }

    # -------------------------------------------------------------------------
    # Public Accessors for Internal Data
    # -------------------------------------------------------------------------

    @property
    def cells(self) -> dict[tuple[int, int], Cell]:
        """Read-only access to the sparse cell storage."""
        return self._cells

    @property
    def col_widths(self) -> dict[int, int]:
        """Read-only access to column widths (sparse, non-default only)."""
        return self._col_widths

    @property
    def row_heights(self) -> dict[int, int]:
        """Read-only access to row heights (sparse, non-default only)."""
        return self._row_heights

    def _init_recalc_engine(self) -> None:
        """Initialize the recalculation engine."""
        from ..formula.recalc import RecalcEngine

        self._recalc_engine = RecalcEngine(self)

    def set_recalc_mode(self, mode: "RecalcMode") -> None:
        """Set the recalculation mode.

        Args:
            mode: RecalcMode.AUTOMATIC or RecalcMode.MANUAL
        """
        if self._recalc_engine:
            self._recalc_engine.set_mode(mode)

    def get_recalc_mode(self) -> "RecalcMode":
        """Get the current recalculation mode.

        Returns:
            Current RecalcMode
        """
        from ..formula.recalc import RecalcMode

        if self._recalc_engine:
            return self._recalc_engine.mode
        return RecalcMode.AUTOMATIC

    def set_recalc_order(self, order: "RecalcOrder") -> None:
        """Set the recalculation order.

        Args:
            order: RecalcOrder.NATURAL, COLUMN_WISE, or ROW_WISE
        """
        if self._recalc_engine:
            self._recalc_engine.set_order(order)

    def get_recalc_order(self) -> "RecalcOrder":
        """Get the current recalculation order.

        Returns:
            Current RecalcOrder
        """
        from ..formula.recalc import RecalcOrder

        if self._recalc_engine:
            return self._recalc_engine.order
        return RecalcOrder.NATURAL

    # -------------------------------------------------------------------------
    # Index Management (for O(k) row/column access)
    # -------------------------------------------------------------------------

    def _add_to_indices(self, row: int, col: int) -> None:
        """Add a cell to the row/column indices."""
        if row not in self._row_index:
            self._row_index[row] = set()
        self._row_index[row].add(col)

        if col not in self._col_index:
            self._col_index[col] = set()
        self._col_index[col].add(row)

    def _remove_from_indices(self, row: int, col: int) -> None:
        """Remove a cell from the row/column indices."""
        if row in self._row_index:
            self._row_index[row].discard(col)
            if not self._row_index[row]:
                del self._row_index[row]

        if col in self._col_index:
            self._col_index[col].discard(row)
            if not self._col_index[col]:
                del self._col_index[col]

    def _rebuild_indices(self) -> None:
        """Rebuild row/column indices from cells dict."""
        self._row_index.clear()
        self._col_index.clear()
        for row, col in self._cells:
            self._add_to_indices(row, col)

    # -------------------------------------------------------------------------
    # Cell Access
    # -------------------------------------------------------------------------

    def get_cell(self, row: int, col: int) -> Cell:
        """Get cell at position, creating if needed.

        Args:
            row: 0-based row index
            col: 0-based column index

        Returns:
            Cell at position (creates empty cell if none exists)
        """
        if (row, col) not in self._cells:
            self._cells[(row, col)] = Cell()
            self._add_to_indices(row, col)
        return self._cells[(row, col)]

    def get_cell_if_exists(self, row: int, col: int) -> Cell | None:
        """Get cell if it exists, without creating.

        Args:
            row: 0-based row index
            col: 0-based column index

        Returns:
            Cell if exists, None otherwise
        """
        return self._cells.get((row, col))

    def set_cell(self, row: int, col: int, value: str) -> None:
        """Set cell value and invalidate cache.

        Args:
            row: 0-based row index
            col: 0-based column index
            value: Raw value string
        """
        cell = self.get_cell(row, col)
        cell.set_value(value)

        # Update dependency graph incrementally
        self.mark_cell_dirty(row, col)
        new_formula = cell.formula if cell.is_formula else None
        self.update_cell_dependency(row, col, new_formula)

        self.modified = True

        # Only full invalidate if no engine (engine handles incremental invalidation)
        if not self._recalc_engine:
            self._invalidate_cache()

    def set_cell_by_ref(self, ref: str, value: str) -> None:
        """Set cell by reference string like 'A1'."""
        row, col = parse_cell_ref(ref)
        self.set_cell(row, col, value)

    def get_cell_by_ref(self, ref: str) -> Cell:
        """Get cell by reference string like 'A1'."""
        row, col = parse_cell_ref(ref)
        return self.get_cell(row, col)

    def delete_cell(self, row: int, col: int) -> None:
        """Delete a cell (remove from sparse storage)."""
        if (row, col) in self._cells:
            del self._cells[(row, col)]
            self._remove_from_indices(row, col)

            # Update dependency graph incrementally (remove dependencies)
            self.mark_cell_dirty(row, col)
            self.update_cell_dependency(row, col, None)

            self.modified = True

            if not self._recalc_engine:
                self._invalidate_cache()

    def cell_exists(self, row: int, col: int) -> bool:
        """Check if a cell has content."""
        cell = self._cells.get((row, col))
        return cell is not None and not cell.is_empty

    # -------------------------------------------------------------------------
    # Value Computation
    # -------------------------------------------------------------------------

    def get_value(self, row: int, col: int, context: "EvaluationContext | None" = None) -> Any:
        """Get computed value of cell.

        Args:
            row: 0-based row index
            col: 0-based column index
            context: Optional evaluation context for cycle detection

        Returns:
            Computed value (number, string, or error)
        """
        if (row, col) in self._cache:
            return self._cache[(row, col)]

        cell = self._cells.get((row, col))
        if not cell or cell.is_empty:
            return ""

        if cell.is_formula:
            # Pass context if provided
            value = self._evaluate_formula(row, col, context)
        else:
            value = self._parse_literal(cell.display_value)

        self._cache[(row, col)] = value
        return value

    def get_value_by_ref(self, ref: str, context: "EvaluationContext | None" = None) -> Any:
        """Get computed value by reference string."""
        row, col = parse_cell_ref(ref)
        return self.get_value(row, col, context)

    def _parse_literal(self, value: str) -> Any:
        """Parse a literal value (number or string)."""
        if not value:
            return ""
        # Try integer first (only for standard numeric formats)
        if "." not in value and "e" not in value.lower():
            try:
                return int(value.replace(",", ""))
            except ValueError:
                pass  # Fall through to try float
        # Try float (handles decimals, scientific notation, and special values like inf/nan)
        try:
            return float(value.replace(",", ""))
        except ValueError:
            return value

    def _evaluate_formula(
        self, row: int, col: int, context: "EvaluationContext | None" = None
    ) -> Any:
        """Evaluate a formula at the given cell position."""
        # Import here to avoid circular dependency
        from ..formula import FormulaParser
        from ..formula.evaluator import FormulaEvaluator

        # Use shared context if provided, else rely on simple self._computing set
        # (though ideally we should switch to full context threading)
        if context:
            evaluator = FormulaEvaluator(self, context=context)
            return evaluator.evaluate_cell(row, col)

        # Fallback for legacy/simple calls (like from RecalcEngine using topological sort order)
        if (row, col) in self._computing:
            self._circular_refs.add((row, col))
            return FormulaError.CIRC

        self._computing.add((row, col))
        try:
            cell = self.get_cell(row, col)
            parser = FormulaParser(self)
            result = parser.evaluate(cell.formula)
            return result
        except (ValueError, TypeError, ZeroDivisionError, OverflowError, KeyError, IndexError):
            return FormulaError.ERR
        finally:
            self._computing.discard((row, col))

    def _invalidate_cache(self) -> None:
        """Clear the computation cache."""
        self._cache.clear()
        self._circular_refs.clear()

    def invalidate_cache(self) -> None:
        """Public method to clear the computation cache.

        Call this after making changes that affect computed values.
        """
        self._invalidate_cache()

    def invalidate_cell_cache(self, row: int, col: int) -> None:
        """Invalidate cache for a specific cell.

        Part of CellStore protocol for RecalcEngine.
        """
        self._cache.pop((row, col), None)

    def clear_cache(self) -> None:
        """Clear the entire computation cache.

        Part of CellStore protocol for RecalcEngine.
        """
        self._cache.clear()
        self._circular_refs.clear()

    # -------------------------------------------------------------------------
    # Cell Store Interface (for undo system and I/O)
    # -------------------------------------------------------------------------

    def get_cell_data(self, row: int, col: int) -> dict | None:
        """Get cell data as a dictionary for serialization.

        Args:
            row: 0-based row index
            col: 0-based column index

        Returns:
            Dictionary representation of cell, or None if cell doesn't exist
        """
        cell = self._cells.get((row, col))
        if cell is None:
            return None
        return cell.to_dict()

    def set_cell_data(self, row: int, col: int, data: dict) -> None:
        """Set cell from dictionary representation.

        Args:
            row: 0-based row index
            col: 0-based column index
            data: Dictionary with cell data (from Cell.to_dict())
        """
        self._cells[(row, col)] = Cell.from_dict(data)

    def remove_cell(self, row: int, col: int) -> None:
        """Remove a cell from the spreadsheet.

        Args:
            row: 0-based row index
            col: 0-based column index
        """
        if (row, col) in self._cells:
            self._cells.pop((row, col), None)
            self._remove_from_indices(row, col)

    def iter_cells(self) -> "Iterator[tuple[int, int, Cell]]":
        """Iterate over all cells.

        Yields:
            Tuples of (row, col, cell) for each cell
        """
        for (r, c), cell in self._cells.items():
            yield r, c, cell

    def get_cells_in_row(self, row: int) -> list[tuple[int, Cell]]:
        """Get all cells in a specific row - O(k) where k = cells in row.

        Args:
            row: 0-based row index

        Returns:
            List of (col, cell) tuples for cells in the row
        """
        cols = self._row_index.get(row, set())
        return [(c, self._cells[(row, c)]) for c in cols if (row, c) in self._cells]

    def get_cells_in_col(self, col: int) -> list[tuple[int, Cell]]:
        """Get all cells in a specific column - O(k) where k = cells in column.

        Args:
            col: 0-based column index

        Returns:
            List of (row, cell) tuples for cells in the column
        """
        rows = self._col_index.get(col, set())
        return [(r, self._cells[(r, col)]) for r in rows if (r, col) in self._cells]

    # -------------------------------------------------------------------------
    # Bulk Accessors (for I/O)
    # -------------------------------------------------------------------------

    def get_all_col_widths(self) -> dict[int, int]:
        """Get all custom column widths.

        Returns:
            Dictionary mapping column index to width
        """
        return dict(self._col_widths)

    def set_all_col_widths(self, widths: dict[int, int]) -> None:
        """Set all column widths.

        Args:
            widths: Dictionary mapping column index to width
        """
        self._col_widths = dict(widths)

    def get_all_row_heights(self) -> dict[int, int]:
        """Get all custom row heights.

        Returns:
            Dictionary mapping row index to height
        """
        return dict(self._row_heights)

    def set_all_row_heights(self, heights: dict[int, int]) -> None:
        """Set all row heights.

        Args:
            heights: Dictionary mapping row index to height
        """
        self._row_heights = dict(heights)

    def recalculate(self) -> None:
        """Force recalculation of all cells."""
        self._invalidate_cache()
        if self._recalc_engine:
            self._recalc_engine.recalculate(full=True)

    def update_cell_dependency(self, row: int, col: int, formula: str | None) -> None:
        """Update the dependency graph for a cell.

        Args:
            row: 0-based row index
            col: 0-based column index
            formula: The formula string, or None if the cell has no formula
        """
        if self._recalc_engine:
            self._recalc_engine.update_cell_dependency(row, col, formula)

    def rebuild_dependency_graph(self) -> None:
        """Rebuild the entire dependency graph.

        Call this after structural changes like inserting/deleting rows/columns.
        """
        if self._recalc_engine:
            self._recalc_engine.rebuild_dependency_graph()

    def mark_cell_dirty(self, row: int, col: int) -> None:
        """Mark a cell as needing recalculation.

        Args:
            row: 0-based row index
            col: 0-based column index
        """
        if self._recalc_engine:
            self._recalc_engine.mark_dirty(row, col)

    @property
    def needs_recalc(self) -> bool:
        """Check if spreadsheet needs recalculation."""
        if self._recalc_engine:
            return self._recalc_engine.needs_recalc
        # In manual mode, check if cache is invalid
        return len(self._cache) == 0 and any(
            c.is_formula for c in self._cells.values() if not c.is_empty
        )

    @property
    def has_circular_refs(self) -> bool:
        """Check if circular references were detected."""
        if self._recalc_engine:
            return len(self._recalc_engine.get_circular_references()) > 0
        return len(self._circular_refs) > 0

    # -------------------------------------------------------------------------
    # Display
    # -------------------------------------------------------------------------

    def get_display_value(self, row: int, col: int) -> str:
        """Get formatted string for display.

        Args:
            row: 0-based row index
            col: 0-based column index

        Returns:
            Formatted value string
        """
        value = self.get_value(row, col)
        cell = self._cells.get((row, col))

        if cell and cell.format_code != "G":
            spec = parse_format_code(cell.format_code)
            return format_value(value, spec, self.get_col_width(col))

        # Default formatting
        if isinstance(value, float):
            # Handle IEEE 754 special values (infinity, NaN)
            if math.isfinite(value) and value == int(value):
                return str(int(value))
            return f"{value:.2f}" if math.isfinite(value) else str(value)
        return str(value) if value != "" else ""

    def get_cell_alignment(self, row: int, col: int) -> TextAlignment:
        """Get the alignment for a cell.

        Lotus 1-2-3 alignment rules:
        - Explicit prefix overrides default: ' (left), " (right), ^ (center), \\ (repeat)
        - Numbers, currency, dates, times: right-aligned by default
        - Text/labels: left-aligned by default
        - Formulas: based on computed result type
        - Errors: left-aligned

        Args:
            row: 0-based row index
            col: 0-based column index

        Returns:
            TextAlignment enum value
        """
        cell = self._cells.get((row, col))
        if not cell or cell.is_empty:
            return TextAlignment.DEFAULT

        # Check for explicit alignment prefix
        explicit_alignment = cell.alignment
        if explicit_alignment != TextAlignment.DEFAULT:
            return explicit_alignment

        # For formulas, determine alignment based on computed value type
        if cell.is_formula:
            value = self.get_value(row, col)
            # Errors are left-aligned
            if isinstance(value, str) and value.startswith("#"):
                return TextAlignment.LEFT
            # Numbers are right-aligned
            if isinstance(value, (int, float)):
                return TextAlignment.RIGHT
            # Text results are left-aligned
            return TextAlignment.LEFT

        # For non-formula cells, check the cell type
        # Numbers are right-aligned, text is left-aligned
        display = cell.display_value
        if display:
            try:
                float(display.replace(",", ""))
                return TextAlignment.RIGHT
            except ValueError:
                pass

        return TextAlignment.LEFT

    # -------------------------------------------------------------------------
    # Column/Row Dimensions
    # -------------------------------------------------------------------------

    def get_col_width(self, col: int) -> int:
        """Get width of column."""
        default: int = self.global_settings.get("default_col_width", DEFAULT_COL_WIDTH)
        return self._col_widths.get(col, default)

    def set_col_width(self, col: int, width: int) -> None:
        """Set width of column."""
        width = max(MIN_COL_WIDTH, min(MAX_COL_WIDTH, width))
        default = self.global_settings.get("default_col_width", DEFAULT_COL_WIDTH)
        if width == default:
            self._col_widths.pop(col, None)
        else:
            self._col_widths[col] = width
        self.modified = True

    def get_row_height(self, row: int) -> int:
        """Get height of row."""
        return self._row_heights.get(row, DEFAULT_ROW_HEIGHT)

    def set_row_height(self, row: int, height: int) -> None:
        """Set height of row."""
        height = max(MIN_ROW_HEIGHT, min(MAX_ROW_HEIGHT, height))
        if height == DEFAULT_ROW_HEIGHT:
            self._row_heights.pop(row, None)
        else:
            self._row_heights[row] = height
        self.modified = True

    # -------------------------------------------------------------------------
    # Range Operations
    # -------------------------------------------------------------------------

    def get_range(
        self, start_ref: str, end_ref: str, context: "EvaluationContext | None" = None
    ) -> list[list[Any]]:
        """Get values in a range as 2D list.

        Args:
            start_ref: Start cell reference (e.g., 'A1')
            end_ref: End cell reference (e.g., 'B10')
            context: Optional evaluation context for cycle detection

        Returns:
            2D list of values
        """
        start_row, start_col = parse_cell_ref(start_ref)
        end_row, end_col = parse_cell_ref(end_ref)

        # Normalize direction
        if start_row > end_row:
            start_row, end_row = end_row, start_row
        if start_col > end_col:
            start_col, end_col = end_col, start_col

        result = []
        for r in range(start_row, end_row + 1):
            row_vals = []
            for c in range(start_col, end_col + 1):
                row_vals.append(self.get_value(r, c, context))
            result.append(row_vals)
        return result

    def get_range_flat(
        self, start_ref: str, end_ref: str, context: "EvaluationContext | None" = None
    ) -> list[Any]:
        """Get values in a range as flat list."""
        rows = self.get_range(start_ref, end_ref, context)
        return [val for row in rows for val in row]

    def set_range_format(
        self, start_row: int, start_col: int, end_row: int, end_col: int, format_code: str
    ) -> None:
        """Set format for a range of cells."""
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                cell = self.get_cell(r, c)
                cell.format_code = format_code
        self.modified = True

    # -------------------------------------------------------------------------
    # Row/Column Operations
    # -------------------------------------------------------------------------

    def delete_row(self, row: int) -> None:
        """Delete a row and shift cells up."""
        new_cells = {}
        for (r, c), cell in self._cells.items():
            if r == row:
                continue

            # Adjust formula references for ALL remaining cells
            if cell.is_formula:
                adjusted = adjust_for_structural_change(
                    cell.raw_value, "row", row, -1, self.rows - 1, self.cols - 1
                )
                cell.set_value(adjusted)

            if r < row:
                new_cells[(r, c)] = cell
            elif r > row:
                new_cells[(r - 1, c)] = cell
        self._cells = new_cells
        self._rebuild_indices()

        # Adjust row heights
        new_heights = {}
        for r, h in self._row_heights.items():
            if r < row:
                new_heights[r] = h
            elif r > row:
                new_heights[r - 1] = h
        self._row_heights = new_heights

        # Adjust named ranges
        self.named_ranges.adjust_for_delete_row(row)

        self.modified = True
        self._invalidate_cache()
        self.rebuild_dependency_graph()

    def insert_row(self, row: int) -> None:
        """Insert a row and shift cells down."""
        new_cells = {}
        for (r, c), cell in self._cells.items():
            # Adjust formula references for ALL cells
            if cell.is_formula:
                adjusted = adjust_for_structural_change(
                    cell.raw_value, "row", row, 1, self.rows - 1, self.cols - 1
                )
                cell.set_value(adjusted)

            if r < row:
                new_cells[(r, c)] = cell
            else:
                new_cells[(r + 1, c)] = cell
        self._cells = new_cells
        self._rebuild_indices()

        # Adjust row heights
        new_heights = {}
        for r, h in self._row_heights.items():
            if r < row:
                new_heights[r] = h
            else:
                new_heights[r + 1] = h
        self._row_heights = new_heights

        # Adjust named ranges
        self.named_ranges.adjust_for_insert_row(row)

        self.modified = True
        self._invalidate_cache()
        self.rebuild_dependency_graph()

    def delete_col(self, col: int) -> None:
        """Delete a column and shift cells left."""
        new_cells = {}
        for (r, c), cell in self._cells.items():
            if c == col:
                continue

            # Adjust formula references for ALL remaining cells
            if cell.is_formula:
                adjusted = adjust_for_structural_change(
                    cell.raw_value, "col", col, -1, self.rows - 1, self.cols - 1
                )
                cell.set_value(adjusted)

            if c < col:
                new_cells[(r, c)] = cell
            elif c > col:
                new_cells[(r, c - 1)] = cell
        self._cells = new_cells
        self._rebuild_indices()

        # Adjust column widths
        new_widths = {}
        for c, w in self._col_widths.items():
            if c < col:
                new_widths[c] = w
            elif c > col:
                new_widths[c - 1] = w
        self._col_widths = new_widths

        # Adjust named ranges
        self.named_ranges.adjust_for_delete_col(col)

        self.modified = True
        self._invalidate_cache()
        self.rebuild_dependency_graph()

    def insert_col(self, col: int) -> None:
        """Insert a column and shift cells right."""
        new_cells = {}
        for (r, c), cell in self._cells.items():
            # Adjust formula references for ALL cells
            if cell.is_formula:
                adjusted = adjust_for_structural_change(
                    cell.raw_value, "col", col, 1, self.rows - 1, self.cols - 1
                )
                cell.set_value(adjusted)

            if c < col:
                new_cells[(r, c)] = cell
            else:
                new_cells[(r, c + 1)] = cell
        self._cells = new_cells
        self._rebuild_indices()

        # Adjust column widths
        new_widths = {}
        for c, w in self._col_widths.items():
            if c < col:
                new_widths[c] = w
            else:
                new_widths[c + 1] = w
        self._col_widths = new_widths

        # Adjust named ranges
        self.named_ranges.adjust_for_insert_col(col)

        self.modified = True
        self._invalidate_cache()
        self.rebuild_dependency_graph()

    # -------------------------------------------------------------------------
    # Copy Operations
    # -------------------------------------------------------------------------

    def copy_cell(
        self, from_row: int, from_col: int, to_row: int, to_col: int, adjust_refs: bool = True
    ) -> None:
        """Copy a cell to another location.

        Args:
            from_row, from_col: Source cell
            to_row, to_col: Destination cell
            adjust_refs: Whether to adjust relative references
        """
        src = self._cells.get((from_row, from_col))
        if not src:
            return

        value = src.raw_value
        if adjust_refs and src.is_formula:
            row_delta = to_row - from_row
            col_delta = to_col - from_col
            value = adjust_formula_references(
                value, row_delta, col_delta, self.rows - 1, self.cols - 1
            )

        dest = self.get_cell(to_row, to_col)
        dest.set_value(value)
        dest.format_code = src.format_code

        # Update dependency graph incrementally
        self.mark_cell_dirty(to_row, to_col)
        new_formula = dest.formula if dest.is_formula else None
        self.update_cell_dependency(to_row, to_col, new_formula)

        self.modified = True

        if not self._recalc_engine:
            self._invalidate_cache()

    def copy_range(
        self,
        src_start: tuple[int, int],
        src_end: tuple[int, int],
        dest_start: tuple[int, int],
        adjust_refs: bool = True,
    ) -> None:
        """Copy a range of cells.

        Args:
            src_start: (row, col) of source top-left
            src_end: (row, col) of source bottom-right
            dest_start: (row, col) of destination top-left
            adjust_refs: Whether to adjust relative references
        """
        src_r1, src_c1 = src_start
        src_r2, src_c2 = src_end
        dest_r, dest_c = dest_start

        for r in range(src_r1, src_r2 + 1):
            for c in range(src_c1, src_c2 + 1):
                dr = dest_r + (r - src_r1)
                dc = dest_c + (c - src_c1)
                self.copy_cell(r, c, dr, dc, adjust_refs)

    # -------------------------------------------------------------------------
    # File Operations
    # -------------------------------------------------------------------------

    def save(self, filename: str) -> None:
        """Save spreadsheet to JSON file."""
        from ..io.lotus_json import LotusJsonSerializer

        LotusJsonSerializer.save(self, filename)
        self.filename = filename
        self.modified = False

    def load(self, filename: str) -> None:
        """Load spreadsheet from JSON file."""
        from ..io.lotus_json import LotusJsonSerializer

        LotusJsonSerializer.load(self, filename)

        self.filename = filename
        self.modified = False
        self._invalidate_cache()

    def clear(self) -> None:
        """Clear all cells and reset to empty state."""
        self._cells.clear()
        self._row_index.clear()
        self._col_index.clear()
        self._cache.clear()
        self._col_widths.clear()
        self._row_heights.clear()
        self.named_ranges.clear()
        self.frozen_rows = 0
        self.frozen_cols = 0
        self.modified = False
        self._circular_refs.clear()
        self.global_settings = {
            "format_code": "G",
            "label_prefix": "'",
            "default_col_width": DEFAULT_COL_WIDTH,
            "zero_display": True,
        }

    # -------------------------------------------------------------------------
    # Iteration
    # -------------------------------------------------------------------------

    def get_used_range(self) -> tuple[tuple[int, int], tuple[int, int]] | None:
        """Get the bounding box of used cells.

        Returns:
            ((min_row, min_col), (max_row, max_col)) or None if empty
        """
        if not self._cells:
            return None

        min_row = min_col = float("inf")
        max_row = max_col = 0

        for (r, c), cell in self._cells.items():
            if not cell.is_empty:
                min_row = min(min_row, r)
                min_col = min(min_col, c)
                max_row = max(max_row, r)
                max_col = max(max_col, c)

        if min_row == float("inf"):
            return None

        return ((int(min_row), int(min_col)), (int(max_row), int(max_col)))

    # -------------------------------------------------------------------------
    # Utility
    # -------------------------------------------------------------------------

    def __repr__(self) -> str:
        cell_count = sum(1 for c in self._cells.values() if not c.is_empty)
        return f"Spreadsheet({self.rows}x{self.cols}, {cell_count} cells)"

"""Recalculation engine with multiple calculation modes."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum, auto
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


class RecalcMode(Enum):
    """Recalculation mode."""

    AUTOMATIC = auto()  # Recalculate immediately when cells change
    MANUAL = auto()  # Only recalculate when requested (F9)


class RecalcOrder(Enum):
    """Order of recalculation for cells."""

    NATURAL = auto()  # Dependency-based (smart recalc)
    COLUMN_WISE = auto()  # Left to right, top to bottom (A1, A2, ..., B1, B2, ...)
    ROW_WISE = auto()  # Top to bottom, left to right (A1, B1, ..., A2, B2, ...)


@dataclass
class RecalcStats:
    """Statistics from a recalculation run."""

    cells_evaluated: int = 0
    circular_refs_found: int = 0
    errors_found: int = 0
    elapsed_ms: float = 0.0


class RecalcEngine:
    """Engine for managing spreadsheet recalculation.

    Supports:
    - Automatic and manual recalculation modes
    - Natural (dependency-based), column-wise, and row-wise ordering
    - Circular reference detection
    - Incremental recalculation (only affected cells)
    """

    def __init__(self, spreadsheet: Spreadsheet) -> None:
        self.spreadsheet = spreadsheet
        self.mode = RecalcMode.AUTOMATIC
        self.order = RecalcOrder.NATURAL
        self._dirty_cells: set[tuple[int, int]] = set()
        self._dependency_graph: dict[tuple[int, int], set[tuple[int, int]]] = {}
        self._dependents: dict[tuple[int, int], set[tuple[int, int]]] = {}
        self._circular_refs: set[tuple[int, int]] = set()

    def set_mode(self, mode: RecalcMode) -> None:
        """Set recalculation mode."""
        self.mode = mode

    def set_order(self, order: RecalcOrder) -> None:
        """Set recalculation order."""
        self.order = order
        if order != RecalcOrder.NATURAL:
            # Clear dependency graph when not using natural order
            self._dependency_graph.clear()
            self._dependents.clear()

    def mark_dirty(self, row: int, col: int) -> None:
        """Mark a cell as needing recalculation.

        Also marks all cells that depend on this cell.
        """
        self._dirty_cells.add((row, col))

        # Mark dependents as dirty too
        for dependent in self._dependents.get((row, col), set()):
            if dependent not in self._dirty_cells:
                self._dirty_cells.add(dependent)
                # Recursively mark dependents
                self.mark_dirty(*dependent)

        if self.mode == RecalcMode.AUTOMATIC:
            self.recalculate()

    def update_cell_dependency(self, row: int, col: int, new_formula: str | None) -> None:
        """Update dependency graph for a single cell.

        Args:
            row: Cell row
            col: Cell column
            new_formula: New formula string (None/empty if not a formula)
        """
        # If it was a formula, remove old dependencies
        old_deps = self._dependency_graph.pop((row, col), set())
        for old_dep in old_deps:
            if old_dep in self._dependents:
                self._dependents[old_dep].discard((row, col))
                if not self._dependents[old_dep]:
                    del self._dependents[old_dep]

        if not new_formula:
            return

        # Calculate new dependencies
        # Import list here to avoid circular dependency
        from .evaluator import FormulaEvaluator

        # We can create a temporary evaluator just for parsing dependencies
        # This is lightweight as it doesn't need to actually evaluate
        evaluator = FormulaEvaluator(self.spreadsheet)
        new_deps = evaluator.get_dependencies(new_formula)

        self._dependency_graph[(row, col)] = new_deps

        # Update dependents (reverse graph)
        for dep in new_deps:
            if dep not in self._dependents:
                self._dependents[dep] = set()
            self._dependents[dep].add((row, col))

    def recalculate(self, full: bool = False) -> RecalcStats:
        """Perform recalculation.

        Args:
            full: If True, recalculate all formula cells.
                  If False, only recalculate dirty cells.

        Returns:
            Statistics from the recalculation
        """
        import time

        start = time.time()

        stats = RecalcStats()

        if full:
            # Only rebuild graph if explicitly requested (e.g. on load)
            # Otherwise assume graph is maintained incrementally
            if not self._dependency_graph:
                self._rebuild_dependency_graph()
            cells_to_calc = self._get_all_formula_cells()
        else:
            cells_to_calc = self._dirty_cells.copy()

        # Get calculation order
        ordered_cells = self._get_calculation_order(cells_to_calc)

        # Clear cache and calculate
        self.spreadsheet._cache.clear()
        self._circular_refs.clear()

        for row, col in ordered_cells:
            cell = self.spreadsheet.get_cell_if_exists(row, col)
            if cell and cell.is_formula:
                value = self.spreadsheet.get_value(row, col)
                stats.cells_evaluated += 1

                if isinstance(value, str):
                    if value == "#CIRC!":
                        self._circular_refs.add((row, col))
                        stats.circular_refs_found += 1
                    elif value.startswith("#"):
                        stats.errors_found += 1

        self._dirty_cells.clear()
        stats.elapsed_ms = (time.time() - start) * 1000

        return stats

    def _get_all_formula_cells(self) -> set[tuple[int, int]]:
        """Get all cells containing formulas."""
        cells = set()
        for row, col, cell in self.spreadsheet.iter_cells():
            if cell.is_formula:
                cells.add((row, col))
        return cells

    def _get_calculation_order(self, cells: set[tuple[int, int]]) -> list[tuple[int, int]]:
        """Get cells in calculation order based on current order setting."""
        if self.order == RecalcOrder.NATURAL:
            return self._topological_order(cells)
        elif self.order == RecalcOrder.COLUMN_WISE:
            return sorted(cells, key=lambda c: (c[1], c[0]))
        else:  # ROW_WISE
            return sorted(cells, key=lambda c: (c[0], c[1]))

    def _topological_order(self, cells: set[tuple[int, int]]) -> list[tuple[int, int]]:
        """Sort cells by dependencies (dependencies first)."""
        if not self._dependency_graph:
            self._rebuild_dependency_graph()

        # Filter to only requested cells
        relevant = {c: self._dependency_graph.get(c, set()) & cells for c in cells}

        # Kahn's algorithm using reverse edges for efficiency
        from collections import deque

        in_degree = {c: len(deps) for c, deps in relevant.items()}
        queue = deque([c for c, d in in_degree.items() if d == 0])
        result = []

        while queue:
            cell = queue.popleft()
            result.append(cell)

            for dependent in self._dependents.get(cell, set()):
                if dependent in in_degree:
                    in_degree[dependent] -= 1
                    if in_degree[dependent] == 0:
                        queue.append(dependent)

        # Add any remaining cells (possibly in cycles)
        for cell in cells:
            if cell not in result:
                result.append(cell)

        return result

    def _rebuild_dependency_graph(self) -> None:
        """Rebuild the dependency graph from scratch."""
        from .evaluator import FormulaEvaluator

        self._dependency_graph.clear()
        self._dependents.clear()

        evaluator = FormulaEvaluator(self.spreadsheet)

        for row, col, cell in self.spreadsheet.iter_cells():
            if cell.is_formula:
                deps = evaluator.get_dependencies(cell.formula)
                self._dependency_graph[(row, col)] = deps

                # Build reverse graph (dependents)
                for dep in deps:
                    if dep not in self._dependents:
                        self._dependents[dep] = set()
                    self._dependents[dep].add((row, col))

    def get_circular_references(self) -> set[tuple[int, int]]:
        """Get cells involved in circular references."""
        return self._circular_refs.copy()

    @property
    def needs_recalc(self) -> bool:
        """Check if there are dirty cells needing recalculation."""
        return len(self._dirty_cells) > 0

    def get_dependents(self, row: int, col: int) -> set[tuple[int, int]]:
        """Get cells that depend on the given cell."""
        return self._dependents.get((row, col), set()).copy()

    def get_dependencies(self, row: int, col: int) -> set[tuple[int, int]]:
        """Get cells that the given cell depends on."""
        return self._dependency_graph.get((row, col), set()).copy()


def create_recalc_engine(spreadsheet: Spreadsheet) -> RecalcEngine:
    """Factory function to create and attach a recalc engine to a spreadsheet."""
    engine = RecalcEngine(spreadsheet)
    spreadsheet._recalc_engine = engine
    return engine

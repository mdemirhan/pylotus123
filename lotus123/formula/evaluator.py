"""Formula evaluation engine with dependency tracking."""
from __future__ import annotations

from typing import Any, TYPE_CHECKING
from dataclasses import dataclass, field
import re

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


@dataclass
class EvaluationContext:
    """Context for formula evaluation.

    Tracks the current cell being evaluated and dependencies.
    """
    current_row: int = 0
    current_col: int = 0
    dependencies: set[tuple[int, int]] = field(default_factory=set)
    computing: set[tuple[int, int]] = field(default_factory=set)


class FormulaEvaluator:
    """Evaluates formulas with dependency tracking.

    This class wraps the FormulaParser and adds:
    - Dependency tracking for recalculation
    - Context awareness (current cell position)
    - Circular reference detection
    """

    # Pattern to find cell references in formulas
    CELL_REF_PATTERN = re.compile(r'\$?([A-Za-z]+)\$?(\d+)')
    RANGE_PATTERN = re.compile(r'(\$?[A-Za-z]+\$?\d+):(\$?[A-Za-z]+\$?\d+)')

    def __init__(self, spreadsheet: Spreadsheet) -> None:
        self.spreadsheet = spreadsheet
        self._context = EvaluationContext()

    def evaluate_cell(self, row: int, col: int) -> Any:
        """Evaluate the formula in a cell.

        Args:
            row: Cell row (0-based)
            col: Cell column (0-based)

        Returns:
            Computed value or error string
        """
        cell = self.spreadsheet.get_cell_if_exists(row, col)
        if not cell or cell.is_empty:
            return ""

        if not cell.is_formula:
            return self._parse_literal(cell.display_value)

        # Check for circular reference
        if (row, col) in self._context.computing:
            return "#CIRC!"

        # Set up context
        self._context.current_row = row
        self._context.current_col = col
        self._context.computing.add((row, col))

        try:
            from .parser import FormulaParser
            parser = FormulaParser(self.spreadsheet)
            result = parser.evaluate(cell.formula)
            return result
        finally:
            self._context.computing.discard((row, col))

    def get_dependencies(self, formula: str) -> set[tuple[int, int]]:
        """Extract cell dependencies from a formula.

        Args:
            formula: Formula string

        Returns:
            Set of (row, col) tuples that the formula depends on
        """
        deps = set()

        # Find range references first
        for match in self.RANGE_PATTERN.finditer(formula):
            start_ref, end_ref = match.groups()
            deps.update(self._expand_range(start_ref, end_ref))

        # Find single cell references (not part of ranges)
        # Remove range patterns first to avoid double-counting
        formula_no_ranges = self.RANGE_PATTERN.sub('', formula)
        for match in self.CELL_REF_PATTERN.finditer(formula_no_ranges):
            col_str, row_str = match.groups()
            try:
                from ..core.reference import parse_cell_ref
                row, col = parse_cell_ref(f"{col_str}{row_str}")
                deps.add((row, col))
            except ValueError:
                pass

        return deps

    def _expand_range(self, start_ref: str, end_ref: str) -> set[tuple[int, int]]:
        """Expand a range reference to individual cell coordinates."""
        deps = set()
        try:
            from ..core.reference import parse_cell_ref
            start_row, start_col = parse_cell_ref(start_ref.replace('$', ''))
            end_row, end_col = parse_cell_ref(end_ref.replace('$', ''))

            # Normalize
            if start_row > end_row:
                start_row, end_row = end_row, start_row
            if start_col > end_col:
                start_col, end_col = end_col, start_col

            for r in range(start_row, end_row + 1):
                for c in range(start_col, end_col + 1):
                    deps.add((r, c))
        except ValueError:
            pass
        return deps

    def _parse_literal(self, value: str) -> Any:
        """Parse a literal value."""
        if not value:
            return ""
        try:
            if '.' not in value and 'e' not in value.lower():
                return int(value.replace(",", ""))
            return float(value.replace(",", ""))
        except ValueError:
            return value

    def reset_context(self) -> None:
        """Reset the evaluation context."""
        self._context = EvaluationContext()


def build_dependency_graph(spreadsheet: Spreadsheet) -> dict[tuple[int, int], set[tuple[int, int]]]:
    """Build a graph of cell dependencies.

    Args:
        spreadsheet: The spreadsheet to analyze

    Returns:
        Dictionary mapping each cell to its dependencies
    """
    evaluator = FormulaEvaluator(spreadsheet)
    graph = {}

    for row, col, cell in spreadsheet.iter_cells():
        if cell.is_formula:
            deps = evaluator.get_dependencies(cell.formula)
            graph[(row, col)] = deps
        else:
            graph[(row, col)] = set()

    return graph


def find_circular_references(dependency_graph: dict[tuple[int, int], set[tuple[int, int]]]) -> list[tuple[int, int]]:
    """Find cells involved in circular references.

    Uses Tarjan's algorithm to find strongly connected components.

    Args:
        dependency_graph: Graph from build_dependency_graph

    Returns:
        List of cells involved in cycles
    """
    index_counter = [0]
    stack = []
    lowlinks = {}
    index = {}
    on_stack = {}
    circular = []

    def strongconnect(node: tuple[int, int]) -> None:
        index[node] = index_counter[0]
        lowlinks[node] = index_counter[0]
        index_counter[0] += 1
        stack.append(node)
        on_stack[node] = True

        for dep in dependency_graph.get(node, set()):
            if dep not in index:
                strongconnect(dep)
                lowlinks[node] = min(lowlinks[node], lowlinks.get(dep, index_counter[0]))
            elif on_stack.get(dep, False):
                lowlinks[node] = min(lowlinks[node], index.get(dep, index_counter[0]))

        if lowlinks[node] == index[node]:
            scc = []
            while True:
                w = stack.pop()
                on_stack[w] = False
                scc.append(w)
                if w == node:
                    break

            # If SCC has more than one node, or node references itself
            if len(scc) > 1 or (len(scc) == 1 and node in dependency_graph.get(node, set())):
                circular.extend(scc)

    for node in dependency_graph:
        if node not in index:
            strongconnect(node)

    return circular


def topological_sort(dependency_graph: dict[tuple[int, int], set[tuple[int, int]]]) -> list[tuple[int, int]]:
    """Sort cells in dependency order for recalculation.

    Args:
        dependency_graph: Graph from build_dependency_graph

    Returns:
        List of cells in order (dependencies first)
    """
    # Kahn's algorithm
    in_degree = {node: 0 for node in dependency_graph}

    # Build reverse graph and count in-degrees
    for node, deps in dependency_graph.items():
        for dep in deps:
            if dep in in_degree:
                # This node depends on dep, so dep must come first
                pass

    # Actually we need dependents, not dependencies
    dependents: dict[tuple[int, int], set[tuple[int, int]]] = {node: set() for node in dependency_graph}
    for node, deps in dependency_graph.items():
        for dep in deps:
            if dep in dependents:
                dependents[dep].add(node)
            in_degree[node] = in_degree.get(node, 0)

    # Recalculate in-degrees based on dependencies
    in_degree = {node: len(deps) for node, deps in dependency_graph.items()}

    # Start with nodes that have no dependencies
    queue = [node for node, degree in in_degree.items() if degree == 0]
    result = []

    while queue:
        node = queue.pop(0)
        result.append(node)

        for dependent in dependents.get(node, set()):
            in_degree[dependent] -= 1
            if in_degree[dependent] == 0:
                queue.append(dependent)

    return result

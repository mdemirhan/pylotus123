"""Formula evaluation engine with dependency tracking."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

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

    def __init__(self, spreadsheet: Spreadsheet, context: EvaluationContext | None = None) -> None:
        self.spreadsheet = spreadsheet
        self._context = context or EvaluationContext()

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

            parser = FormulaParser(self.spreadsheet, context=self._context)
            result = parser.evaluate(cell.formula)
            return result
        finally:
            self._context.computing.discard((row, col))

    def get_dependencies(self, formula: str) -> set[tuple[int, int]]:
        """Extract cell dependencies from a formula.

        Uses the Tokenizer to properly handle string literals and named ranges.

        Args:
            formula: Formula string

        Returns:
            Set of (row, col) tuples that the formula depends on
        """
        from .tokenizer import Tokenizer, TokenType

        deps: set[tuple[int, int]] = set()

        # Use tokenizer to properly parse the formula (skips string literals)
        tokenizer = Tokenizer(self.spreadsheet)
        tokens = tokenizer.tokenize(formula)

        i = 0
        while i < len(tokens):
            token = tokens[i]

            if token.type == TokenType.CELL:
                # Check if this is part of a range (CELL:CELL or CELL..CELL)
                if i + 2 < len(tokens) and tokens[i + 1].type == TokenType.COLON:
                    next_token = tokens[i + 2]
                    if next_token.type == TokenType.CELL:
                        # It's a range
                        deps.update(self._expand_range(token.value, next_token.value))
                        i += 3
                        continue
                # Single cell reference
                try:
                    from ..core.reference import parse_cell_ref

                    row, col = parse_cell_ref(token.value.replace("$", ""))
                    deps.add((row, col))
                except ValueError:
                    pass

            elif token.type == TokenType.RANGE:
                # Named range resolved to range string like "A1:B10"
                if ":" in token.value:
                    parts = token.value.split(":")
                    if len(parts) == 2:
                        deps.update(self._expand_range(parts[0], parts[1]))

            i += 1

        return deps

    def _expand_range(self, start_ref: str, end_ref: str) -> set[tuple[int, int]]:
        """Expand a range reference to individual cell coordinates."""
        deps = set()
        try:
            from ..core.reference import parse_cell_ref

            start_row, start_col = parse_cell_ref(start_ref.replace("$", ""))
            end_row, end_col = parse_cell_ref(end_ref.replace("$", ""))

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
            if "." not in value and "e" not in value.lower():
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


def find_circular_references(
    dependency_graph: dict[tuple[int, int], set[tuple[int, int]]],
) -> list[tuple[int, int]]:
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




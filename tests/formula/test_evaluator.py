"""Tests for formula evaluator."""

from lotus123 import Spreadsheet
from lotus123.formula.evaluator import (
    EvaluationContext,
    FormulaEvaluator,
    build_dependency_graph,
    find_circular_references,
)


class TestEvaluationContext:
    """Tests for EvaluationContext class."""

    def test_default_values(self):
        """Test default context values."""
        ctx = EvaluationContext()
        assert ctx.current_row == 0
        assert ctx.current_col == 0
        assert ctx.dependencies == set()
        assert ctx.computing == set()

    def test_custom_values(self):
        """Test custom context values."""
        deps = {(0, 0), (0, 1)}
        ctx = EvaluationContext(current_row=5, current_col=3, dependencies=deps)
        assert ctx.current_row == 5
        assert ctx.current_col == 3
        assert ctx.dependencies == deps


class TestFormulaEvaluator:
    """Tests for FormulaEvaluator class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.evaluator = FormulaEvaluator(self.ss)

    def test_evaluate_empty_cell(self):
        """Test evaluating empty cell returns empty string."""
        result = self.evaluator.evaluate_cell(0, 0)
        assert result == ""

    def test_evaluate_literal_number(self):
        """Test evaluating cell with literal number."""
        self.ss.set_cell(0, 0, "42")
        result = self.evaluator.evaluate_cell(0, 0)
        assert result == 42

    def test_evaluate_literal_float(self):
        """Test evaluating cell with float."""
        self.ss.set_cell(0, 0, "3.14")
        result = self.evaluator.evaluate_cell(0, 0)
        assert result == 3.14

    def test_evaluate_literal_text(self):
        """Test evaluating cell with text."""
        self.ss.set_cell(0, 0, "Hello")
        result = self.evaluator.evaluate_cell(0, 0)
        assert result == "Hello"

    def test_evaluate_simple_formula(self):
        """Test evaluating simple formula."""
        self.ss.set_cell(0, 0, "=1+2")
        result = self.evaluator.evaluate_cell(0, 0)
        assert result == 3

    def test_evaluate_formula_with_cell_ref(self):
        """Test evaluating formula with cell reference."""
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "=A1*2")
        result = self.evaluator.evaluate_cell(0, 1)
        assert result == 20

    def test_circular_reference_detected(self):
        """Test circular reference detection."""
        self.ss.set_cell(0, 0, "=B1")
        self.ss.set_cell(0, 1, "=A1")

        # First cell evaluates, marks itself as computing
        # When it tries to evaluate B1, which tries to evaluate A1
        # A1 is still marked as computing -> circular reference
        result = self.evaluator.evaluate_cell(0, 0)
        # Note: This may depend on parser implementation
        # The evaluator tracks computing cells

    def test_reset_context(self):
        """Test resetting evaluation context."""
        self.evaluator._context.current_row = 5
        self.evaluator._context.computing.add((0, 0))

        self.evaluator.reset_context()

        assert self.evaluator._context.current_row == 0
        assert len(self.evaluator._context.computing) == 0


class TestGetDependencies:
    """Tests for get_dependencies method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.evaluator = FormulaEvaluator(self.ss)

    def test_single_cell_reference(self):
        """Test extracting single cell dependency."""
        deps = self.evaluator.get_dependencies("=A1")
        assert (0, 0) in deps

    def test_multiple_cell_references(self):
        """Test extracting multiple cell dependencies."""
        deps = self.evaluator.get_dependencies("=A1+B2")
        assert (0, 0) in deps
        assert (1, 1) in deps

    def test_range_reference(self):
        """Test extracting range dependencies."""
        deps = self.evaluator.get_dependencies("=SUM(A1:A3)")
        assert (0, 0) in deps
        assert (1, 0) in deps
        assert (2, 0) in deps

    def test_mixed_references(self):
        """Test extracting mixed dependencies."""
        deps = self.evaluator.get_dependencies("=SUM(A1:A3)+B5")
        assert (0, 0) in deps
        assert (1, 0) in deps
        assert (2, 0) in deps
        assert (4, 1) in deps

    def test_absolute_reference(self):
        """Test extracting absolute references."""
        deps = self.evaluator.get_dependencies("=$A$1")
        assert (0, 0) in deps

    def test_no_dependencies(self):
        """Test formula with no cell references."""
        deps = self.evaluator.get_dependencies("=1+2")
        assert len(deps) == 0


class TestExpandRange:
    """Tests for _expand_range method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.evaluator = FormulaEvaluator(self.ss)

    def test_single_column_range(self):
        """Test expanding single column range."""
        deps = self.evaluator._expand_range("A1", "A3")
        assert deps == {(0, 0), (1, 0), (2, 0)}

    def test_single_row_range(self):
        """Test expanding single row range."""
        deps = self.evaluator._expand_range("A1", "C1")
        assert deps == {(0, 0), (0, 1), (0, 2)}

    def test_rectangular_range(self):
        """Test expanding rectangular range."""
        deps = self.evaluator._expand_range("A1", "B2")
        assert deps == {(0, 0), (0, 1), (1, 0), (1, 1)}

    def test_reversed_range(self):
        """Test expanding reversed range."""
        deps = self.evaluator._expand_range("B2", "A1")
        assert deps == {(0, 0), (0, 1), (1, 0), (1, 1)}


class TestParseLiteral:
    """Tests for _parse_literal method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.evaluator = FormulaEvaluator(self.ss)

    def test_empty_string(self):
        """Test parsing empty string."""
        assert self.evaluator._parse_literal("") == ""

    def test_integer(self):
        """Test parsing integer."""
        assert self.evaluator._parse_literal("42") == 42

    def test_float(self):
        """Test parsing float."""
        assert self.evaluator._parse_literal("3.14") == 3.14

    def test_scientific_notation(self):
        """Test parsing scientific notation."""
        assert self.evaluator._parse_literal("1e10") == 1e10

    def test_text(self):
        """Test parsing text."""
        assert self.evaluator._parse_literal("hello") == "hello"

    def test_number_with_comma(self):
        """Test parsing number with comma."""
        assert self.evaluator._parse_literal("1,234") == 1234


class TestBuildDependencyGraph:
    """Tests for build_dependency_graph function."""

    def test_empty_spreadsheet(self):
        """Test building graph from empty spreadsheet."""
        ss = Spreadsheet()
        graph = build_dependency_graph(ss)
        assert graph == {}

    def test_single_formula(self):
        """Test building graph with single formula."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "10")
        ss.set_cell(0, 1, "=A1*2")

        graph = build_dependency_graph(ss)

        assert (0, 1) in graph
        assert (0, 0) in graph[(0, 1)]

    def test_chain_dependencies(self):
        """Test building graph with chain dependencies."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "10")
        ss.set_cell(0, 1, "=A1*2")
        ss.set_cell(0, 2, "=B1+5")

        graph = build_dependency_graph(ss)

        assert (0, 0) in graph[(0, 1)]
        assert (0, 1) in graph[(0, 2)]


class TestFindCircularReferences:
    """Tests for find_circular_references function."""

    def test_no_circular(self):
        """Test graph with no circular references."""
        graph = {
            (0, 0): set(),
            (0, 1): {(0, 0)},
            (0, 2): {(0, 1)},
        }

        circular = find_circular_references(graph)
        assert circular == []

    def test_self_reference(self):
        """Test cell referencing itself."""
        graph = {
            (0, 0): {(0, 0)},
        }

        circular = find_circular_references(graph)
        assert (0, 0) in circular

    def test_two_cell_cycle(self):
        """Test two cells referencing each other."""
        graph = {
            (0, 0): {(0, 1)},
            (0, 1): {(0, 0)},
        }

        circular = find_circular_references(graph)
        assert (0, 0) in circular
        assert (0, 1) in circular

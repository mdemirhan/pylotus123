"""Tests for recalculation engine."""


from lotus123 import Spreadsheet
from lotus123.formula.recalc import (
    RecalcEngine,
    RecalcMode,
    RecalcOrder,
    RecalcStats,
    create_recalc_engine,
)


class TestRecalcMode:
    """Tests for RecalcMode enum."""

    def test_modes_exist(self):
        """Test all modes exist."""
        assert RecalcMode.AUTOMATIC
        assert RecalcMode.MANUAL


class TestRecalcOrder:
    """Tests for RecalcOrder enum."""

    def test_orders_exist(self):
        """Test all orders exist."""
        assert RecalcOrder.NATURAL
        assert RecalcOrder.COLUMN_WISE
        assert RecalcOrder.ROW_WISE


class TestRecalcStats:
    """Tests for RecalcStats dataclass."""

    def test_default_values(self):
        """Test default values."""
        stats = RecalcStats()
        assert stats.cells_evaluated == 0
        assert stats.circular_refs_found == 0
        assert stats.errors_found == 0
        assert stats.elapsed_ms == 0.0

    def test_custom_values(self):
        """Test custom values."""
        stats = RecalcStats(cells_evaluated=5, errors_found=2, elapsed_ms=10.5)
        assert stats.cells_evaluated == 5
        assert stats.errors_found == 2
        assert stats.elapsed_ms == 10.5


class TestRecalcEngine:
    """Tests for RecalcEngine class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.engine = RecalcEngine(self.ss)

    def test_default_mode(self):
        """Test default mode is automatic."""
        assert self.engine.mode == RecalcMode.AUTOMATIC

    def test_default_order(self):
        """Test default order is natural."""
        assert self.engine.order == RecalcOrder.NATURAL

    def test_set_mode(self):
        """Test setting recalc mode."""
        self.engine.set_mode(RecalcMode.MANUAL)
        assert self.engine.mode == RecalcMode.MANUAL

    def test_set_order(self):
        """Test setting recalc order."""
        self.engine.set_order(RecalcOrder.COLUMN_WISE)
        assert self.engine.order == RecalcOrder.COLUMN_WISE

    def test_set_order_clears_graph(self):
        """Test setting non-natural order clears dependency graph."""
        # Build some dependencies
        self.engine._dependency_graph[(0, 0)] = {(0, 1)}
        self.engine._dependents[(0, 1)] = {(0, 0)}

        self.engine.set_order(RecalcOrder.ROW_WISE)

        assert len(self.engine._dependency_graph) == 0
        assert len(self.engine._dependents) == 0

    def test_mark_dirty(self):
        """Test marking cell as dirty."""
        self.engine.set_mode(RecalcMode.MANUAL)  # Disable auto-recalc
        self.engine.mark_dirty(0, 0)
        assert (0, 0) in self.engine._dirty_cells

    def test_mark_dirty_marks_dependents(self):
        """Test marking dirty also marks dependents."""
        self.engine.set_mode(RecalcMode.MANUAL)

        # Set up dependency: B1 depends on A1
        self.engine._dependents[(0, 0)] = {(0, 1)}

        self.engine.mark_dirty(0, 0)

        assert (0, 0) in self.engine._dirty_cells
        assert (0, 1) in self.engine._dirty_cells

    def test_needs_recalc_empty(self):
        """Test needs_recalc when no dirty cells."""
        assert self.engine.needs_recalc is False

    def test_needs_recalc_with_dirty(self):
        """Test needs_recalc with dirty cells."""
        self.engine.set_mode(RecalcMode.MANUAL)
        self.engine._dirty_cells.add((0, 0))
        assert self.engine.needs_recalc is True

    def test_update_cell_dependency(self):
        """Test incremental dependency updates."""
        # Setup: B1 depends on A1 (0,0 -> 0,1)
        self.engine._dependency_graph[(0, 1)] = {(0, 0)}
        self.engine._dependents[(0, 0)] = {(0, 1)}
        
        # Change B1 to depend on C1 (0,2)
        # We need to simulate the formula change passed to update_cell_dependency
        self.engine.update_cell_dependency(0, 1, "=C1")
        
        # Verify old dependency removed
        assert (0, 0) not in self.engine.get_dependencies(0, 1)
        assert (0, 1) not in self.engine.get_dependents(0, 0)
        
        # Verify new dependency added
        assert (0, 2) in self.engine.get_dependencies(0, 1)
        assert (0, 1) in self.engine.get_dependents(0, 2)

    def test_update_cell_dependency_to_constant(self):
        """Test updating cell from formula to constant removes dependencies."""
        self.engine._dependency_graph[(0, 1)] = {(0, 0)}
        self.engine._dependents[(0, 0)] = {(0, 1)}
        
        self.engine.update_cell_dependency(0, 1, None)
        
        assert len(self.engine.get_dependencies(0, 1)) == 0
        assert len(self.engine.get_dependents(0, 0)) == 0


class TestRecalcCalculation:
    """Tests for recalculation functionality."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.engine = RecalcEngine(self.ss)
        self.engine.set_mode(RecalcMode.MANUAL)

    def test_recalculate_empty(self):
        """Test recalculating empty spreadsheet."""
        stats = self.engine.recalculate()
        assert stats.cells_evaluated == 0

    def test_recalculate_with_formula(self):
        """Test recalculating spreadsheet with formulas."""
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "=A1*2")

        self.engine._dirty_cells.add((0, 1))
        stats = self.engine.recalculate()

        assert stats.cells_evaluated == 1

    def test_recalculate_full(self):
        """Test full recalculation."""
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "=A1*2")
        self.ss.set_cell(0, 2, "=B1+5")

        stats = self.engine.recalculate(full=True)

        assert stats.cells_evaluated == 2

    def test_recalculate_clears_dirty(self):
        """Test recalculation clears dirty cells."""
        self.engine._dirty_cells.add((0, 0))
        self.engine.recalculate()
        assert len(self.engine._dirty_cells) == 0

    def test_recalculate_tracks_elapsed_time(self):
        """Test recalculation tracks elapsed time."""
        stats = self.engine.recalculate()
        assert stats.elapsed_ms >= 0


class TestRecalcOrder:
    """Tests for calculation order."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.engine = RecalcEngine(self.ss)
        self.engine.set_mode(RecalcMode.MANUAL)

    def test_column_wise_order(self):
        """Test column-wise ordering."""
        self.engine.set_order(RecalcOrder.COLUMN_WISE)
        cells = {(0, 0), (1, 0), (0, 1), (1, 1)}

        order = self.engine._get_calculation_order(cells)

        # Column-wise: (col, row) -> A1, A2, B1, B2
        assert order == [(0, 0), (1, 0), (0, 1), (1, 1)]

    def test_row_wise_order(self):
        """Test row-wise ordering."""
        self.engine.set_order(RecalcOrder.ROW_WISE)
        cells = {(0, 0), (1, 0), (0, 1), (1, 1)}

        order = self.engine._get_calculation_order(cells)

        # Row-wise: (row, col) -> A1, B1, A2, B2
        assert order == [(0, 0), (0, 1), (1, 0), (1, 1)]

    def test_natural_order_simple(self):
        """Test natural ordering with simple dependency."""
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "=A1*2")

        self.engine._rebuild_dependency_graph()
        cells = {(0, 0), (0, 1)}

        order = self.engine._get_calculation_order(cells)

        # A1 should come before B1
        assert order.index((0, 0)) < order.index((0, 1))


class TestDependencyGraph:
    """Tests for dependency graph functionality."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.engine = RecalcEngine(self.ss)

    def test_rebuild_dependency_graph(self):
        """Test rebuilding dependency graph."""
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "=A1*2")

        self.engine._rebuild_dependency_graph()

        # B1 depends on A1
        assert (0, 0) in self.engine._dependency_graph.get((0, 1), set())

    def test_dependents_built(self):
        """Test dependents map is built."""
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "=A1*2")

        self.engine._rebuild_dependency_graph()

        # A1 has B1 as dependent
        assert (0, 1) in self.engine._dependents.get((0, 0), set())

    def test_get_dependents(self):
        """Test get_dependents method."""
        self.engine._dependents[(0, 0)] = {(0, 1), (0, 2)}

        deps = self.engine.get_dependents(0, 0)

        assert deps == {(0, 1), (0, 2)}

    def test_get_dependents_empty(self):
        """Test get_dependents with no dependents."""
        deps = self.engine.get_dependents(0, 0)
        assert deps == set()

    def test_get_dependencies(self):
        """Test get_dependencies method."""
        self.engine._dependency_graph[(0, 1)] = {(0, 0)}

        deps = self.engine.get_dependencies(0, 1)

        assert deps == {(0, 0)}

    def test_get_dependencies_empty(self):
        """Test get_dependencies with no dependencies."""
        deps = self.engine.get_dependencies(0, 0)
        assert deps == set()


class TestGetAllFormulaCells:
    """Tests for _get_all_formula_cells method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.engine = RecalcEngine(self.ss)

    def test_no_formulas(self):
        """Test with no formula cells."""
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "20")

        cells = self.engine._get_all_formula_cells()

        assert len(cells) == 0

    def test_with_formulas(self):
        """Test with formula cells."""
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "=A1*2")
        self.ss.set_cell(0, 2, "=A1+B1")

        cells = self.engine._get_all_formula_cells()

        assert (0, 1) in cells
        assert (0, 2) in cells
        assert len(cells) == 2


class TestCircularReferences:
    """Tests for circular reference handling."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.engine = RecalcEngine(self.ss)
        self.engine.set_mode(RecalcMode.MANUAL)

    def test_get_circular_references_empty(self):
        """Test get_circular_references with no circulars."""
        refs = self.engine.get_circular_references()
        assert refs == set()

    def test_get_circular_references_returns_copy(self):
        """Test get_circular_references returns a copy."""
        self.engine._circular_refs.add((0, 0))

        refs = self.engine.get_circular_references()
        refs.add((1, 1))

        # Original shouldn't be modified
        assert (1, 1) not in self.engine._circular_refs


class TestCreateRecalcEngine:
    """Tests for create_recalc_engine factory function."""

    def test_creates_engine(self):
        """Test factory creates engine."""
        ss = Spreadsheet()
        engine = create_recalc_engine(ss)

        assert isinstance(engine, RecalcEngine)
        assert engine.spreadsheet is ss

    def test_attaches_to_spreadsheet(self):
        """Test factory attaches engine to spreadsheet."""
        ss = Spreadsheet()
        engine = create_recalc_engine(ss)

        assert ss._recalc_engine is engine

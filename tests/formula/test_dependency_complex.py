
import pytest
from lotus123.core.spreadsheet import Spreadsheet
from lotus123.formula.recalc import create_recalc_engine, RecalcMode
from lotus123.utils.undo import DeleteRowCommand, UndoManager

class TestDependencyComplex:
    def setup_method(self):
        self.sheet = Spreadsheet()
        self.engine = create_recalc_engine(self.sheet)
        self.engine.set_mode(RecalcMode.MANUAL)
        self.undo_manager = UndoManager()

    def test_chain_dependency_updates(self):
        """Test A->B->C chain updates correctly."""
        # C1 = 10
        # B1 = C1 * 2
        # A1 = B1 + 5
        self.sheet.set_cell(2, 0, "10")
        self.sheet.set_cell(1, 0, "=A3*2")
        self.sheet.set_cell(0, 0, "=A2+5")

        # Initial check
        self.sheet.recalculate()
        assert self.sheet.get_value(0, 0) == 25.0  # 10*2 + 5

        # Modify C1 (A3)
        self.sheet.set_cell(2, 0, "20")
        
        # Verify dirty propagation (Manual mode)
        # Note: Recalc logic marks dependents dirty
        # A3 changed -> marks A2 dirty -> marks A1 dirty
        
        self.sheet.recalculate()
        assert self.sheet.get_value(0, 0) == 45.0  # 20*2 + 5

    def test_structural_insert_row_maintains_dependency(self):
        """Test that inserting a row shifts dependencies correctly."""
        # A1 = A2 * 2. A2 = 10.
        # Dep: A1(0,0) -> A2(1,0)
        self.sheet.set_cell(1, 0, "10")
        self.sheet.set_cell(0, 0, "=A2*2")
        
        self.sheet.recalculate()
        assert self.sheet.get_value(0, 0) == 20.0
        assert self.engine.get_dependencies(0, 0) == {(1, 0)}

        # Insert Row 1 (between A1 and A2)
        # A1 stays at (0,0).
        # A2 moves to (2,0) and becomes A3.
        # Formula at A1 should update to =A3.
        self.sheet.insert_row(1)

        # Check Formula
        assert self.sheet.get_cell(0, 0).raw_value == "=A3*2"
        
        # Check Value (should be preserved or recomputable)
        self.sheet.recalculate()
        assert self.sheet.get_value(0, 0) == 20.0
        
        # Check Dependency Graph
        # A1(0,0) should depend on A3(2,0)
        deps = self.engine.get_dependencies(0, 0)
        assert deps == {(2, 0)}, f"Expected deps {{(2, 0)}}, got {deps}"

    def test_structural_delete_row_maintains_dependency(self):
        """Test deleting a row updates pointers."""
        # A1 = A3. A3 = 10.
        # Insert extra row at 1 to delete.
        self.sheet.set_cell(2, 0, "10")
        self.sheet.set_cell(0, 0, "=A3")
        
        # Delete Row 1 (between A1 and A3)
        # A3 moves to A2.
        # A1 formula should become =A2.
        self.sheet.delete_row(1)
        
        assert self.sheet.get_cell(0, 0).raw_value == "=A2"
        self.sheet.recalculate()
        assert self.sheet.get_value(0, 0) == 10.0
        
        # Check Dependency
        # A1(0,0) -> A2(1,0)
        assert self.engine.get_dependencies(0, 0) == {(1, 0)}

    def test_undo_row_delete_restores_dependency(self):
        """Test that Undo of DeleteRow restores the graph."""
        # A1 = 10, B1 = A1
        self.sheet.set_cell(0, 0, "10")
        self.sheet.set_cell(0, 1, "=A1")
        
        assert self.engine.get_dependencies(0, 1) == {(0, 0)}
        
        # Delete Row 0
        cmd = DeleteRowCommand(self.sheet, 0)
        self.undo_manager.execute(cmd)
        
        # Verify B1 is gone
        assert not self.sheet.cell_exists(0, 1)
        
        # Undo
        self.undo_manager.undo()
        
        # Verify B1 is back
        assert self.sheet.get_cell(0, 1).raw_value == "=A1"
        
        # Verify dependency restored
        assert self.engine.get_dependencies(0, 1) == {(0, 0)}

    def test_complex_graph_restructure(self):
        """More complex structural change involving multiple updates."""
        # A1=1
        # B1=A1+1
        # C1=B1+1
        # D1=C1+1
        self.sheet.set_cell(0, 0, "1")
        self.sheet.set_cell(0, 1, "=A1+1")
        self.sheet.set_cell(0, 2, "=B1+1")
        self.sheet.set_cell(0, 3, "=C1+1")
        
        self.sheet.recalculate()
        assert self.sheet.get_value(0, 3) == 4.0
        
        # Insert Column at B (Col 1).
        # Old B1 moves to C1.
        # Old C1 moves to D1.
        # Old D1 moves to E1.
        # A1 stays.
        # Relationships should preserve: E1->D1->C1->A1.
        self.sheet.insert_col(1)
        
        # Check formulas
        # Old B1 (now at 0,2 C1) should refer to A1 (0,0).
        assert self.sheet.get_cell(0, 2).raw_value == "=A1+1" # C1=A1+1
        # Old C1 (now at 0,3 D1) should refer to C1 (0,2).
        assert self.sheet.get_cell(0, 3).raw_value == "=C1+1" # D1=C1+1
        
        self.sheet.recalculate()
        assert self.sheet.get_value(0, 4) == 4.0 # E1
        
        # Verify Graph
        # E1(0,4) -> D1(0,3)
        assert self.engine.get_dependencies(0, 4) == {(0, 3)}
        # D1(0,3) -> C1(0,2)
        assert self.engine.get_dependencies(0, 3) == {(0, 2)}
        # C1(0,2) -> A1(0,0)
        assert self.engine.get_dependencies(0, 2) == {(0, 0)}



import pytest
from lotus123.core.spreadsheet import Spreadsheet
from lotus123.formula.recalc import create_recalc_engine, RecalcMode
from lotus123.utils.undo import (
    CellChangeCommand,
    ClearRangeCommand,
    DeleteColCommand,
    DeleteRowCommand,
    InsertColCommand,
    InsertRowCommand,
    RangeChangeCommand,
    UndoManager,
)

class TestUndoDependencyIntegrity:
    """Tests that Undo/Redo operations correctly maintain the dependency graph."""

    def setup_method(self):
        self.sheet = Spreadsheet()
        self.engine = create_recalc_engine(self.sheet)
        self.engine.set_mode(RecalcMode.MANUAL)
        self.undo_manager = UndoManager()

    def test_cell_change_undo_redo_dependencies(self):
        """Test CellChangeCommand updates dependencies on undo/redo."""
        # Setup: A1=10, B1=A1*2
        self.sheet.set_cell(0, 0, "10")
        self.sheet.set_cell(0, 1, "=A1*2")
        
        # Initial dependency: B1->A1
        assert self.engine.get_dependencies(0, 1) == {(0, 0)}

        # Execute: Change B1 to constant "20"
        cmd = CellChangeCommand(self.sheet, 0, 1, "20")
        self.undo_manager.execute(cmd)
        
        # Verify dependency removed
        assert not self.engine.get_dependencies(0, 1)

        # Undo: Restore B1 to "=A1*2"
        self.undo_manager.undo()
        
        # Verify dependency restored
        assert self.engine.get_dependencies(0, 1) == {(0, 0)}
        
        # Redo: Change B1 back to "20"
        self.undo_manager.redo()
        
        # Verify dependency removed again
        assert not self.engine.get_dependencies(0, 1)

    def test_range_change_undo_dependencies(self):
        """Test RangeChangeCommand updates dependencies."""
        # A1=1, A2=2
        # B1=A1, B2=A2
        self.sheet.set_cell(0, 0, "1")
        self.sheet.set_cell(1, 0, "2")
        self.sheet.set_cell(0, 1, "=A1")
        self.sheet.set_cell(1, 1, "=A2")
        
        assert self.engine.get_dependencies(0, 1) == {(0, 0)}
        assert self.engine.get_dependencies(1, 1) == {(1, 0)}
        
        # Execute: Change B1:B2 to constants
        cmd = RangeChangeCommand(self.sheet, [
            (0, 1, "10", "=A1"),
            (1, 1, "20", "=A2")
        ])
        self.undo_manager.execute(cmd)
        
        assert not self.engine.get_dependencies(0, 1)
        assert not self.engine.get_dependencies(1, 1)
        
        # Undo
        self.undo_manager.undo()
        
        assert self.engine.get_dependencies(0, 1) == {(0, 0)}
        assert self.engine.get_dependencies(1, 1) == {(1, 0)}

    def test_delete_row_undo_dependencies(self):
        """Test DeleteRowCommand logic with dependencies."""
        # A1=10
        # A2=A1+5
        self.sheet.set_cell(0, 0, "10")
        self.sheet.set_cell(1, 0, "=A1+5")
        
        assert self.engine.get_dependencies(1, 0) == {(0, 0)}
        
        # Delete Row 0
        # A2 moves to A1. A1 (new) formula becomes #ERR or invalid if it referred to deleted row, 
        # BUT here A2 referred to A1 (which is being deleted).
        # Actually, let's use a safer case: Insert row first, then delete.
        
        # A1=10. B1=A1.
        self.sheet.set_cell(0, 0, "10")
        self.sheet.set_cell(0, 1, "=A1")
        
        # Delete Row 0.
        cmd = DeleteRowCommand(self.sheet, 0)
        self.undo_manager.execute(cmd)
        
        # Dependencies should be gone
        assert not self.engine.get_dependencies(0, 1)
        
        # Undo
        self.undo_manager.undo()
        
        # Verify B1 exists and depends on A1
        assert self.sheet.get_cell(0, 1).raw_value == "=A1"
        assert self.engine.get_dependencies(0, 1) == {(0, 0)}

    def test_insert_row_undo_dependencies(self):
        """Test InsertRowCommand shifts dependencies correctly on undo."""
        # A1=10, A2=A1
        self.sheet.set_cell(0, 0, "10")
        self.sheet.set_cell(1, 0, "=A1")
        
        assert self.engine.get_dependencies(1, 0) == {(0, 0)}

        # Insert Row 1
        # A1 stays.
        # A2 moves to A3. Formula A3 refers to A1.
        cmd = InsertRowCommand(self.sheet, 1)
        self.undo_manager.execute(cmd)
        
        # Verify A3 -> A1
        assert self.engine.get_dependencies(2, 0) == {(0, 0)}
        assert not self.engine.get_dependencies(1, 0) # Row 1 is empty

        # Undo (Delete inserted row)
        self.undo_manager.undo()
        
        # Verify A2 -> A1 restored
        assert self.engine.get_dependencies(1, 0) == {(0, 0)}
        # A3 should be empty/gone
        assert self.sheet.get_cell_if_exists(2, 0) is None

    def test_clear_range_undo_dependencies(self):
        """Test ClearRangeCommand removes and restores dependencies."""
        # A1=10, B1=A1
        self.sheet.set_cell(0, 0, "10")
        self.sheet.set_cell(0, 1, "=A1")
        
        assert self.engine.get_dependencies(0, 1) == {(0, 0)}
        
        # Clear B1
        cmd = ClearRangeCommand(self.sheet, 0, 1, 0, 1)
        self.undo_manager.execute(cmd)
        
        # Verify dependency gone
        assert not self.engine.get_dependencies(0, 1)
        
        # Undo
        self.undo_manager.undo()
        
        # Verify dependency restored
        assert self.engine.get_dependencies(0, 1) == {(0, 0)}

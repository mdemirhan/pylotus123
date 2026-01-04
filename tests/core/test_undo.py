"""Tests for undo/redo operations."""

from lotus123 import Spreadsheet
from lotus123.utils.undo import (
    CellChangeCommand,
    ClearRangeCommand,
    CompositeCommand,
    DeleteColCommand,
    DeleteRowCommand,
    InsertColCommand,
    InsertRowCommand,
    RangeChangeCommand,
    UndoManager,
)


class TestUndoManager:
    """Tests for UndoManager class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.manager = UndoManager()

    def test_initial_state(self):
        """Test initial manager state."""
        assert self.manager.can_undo is False
        assert self.manager.can_redo is False
        assert self.manager.undo_count == 0
        assert self.manager.redo_count == 0
        assert self.manager.undo_description == ""
        assert self.manager.redo_description == ""

    def test_execute_adds_to_undo_stack(self):
        """Test execute adds command to undo stack."""
        cmd = CellChangeCommand(self.ss, 0, 0, "test")
        self.manager.execute(cmd)

        assert self.manager.can_undo is True
        assert self.manager.undo_count == 1

    def test_undo_basic(self):
        """Test basic undo operation."""
        cmd = CellChangeCommand(self.ss, 0, 0, "test")
        self.manager.execute(cmd)

        result = self.manager.undo()

        assert result == cmd
        assert self.manager.can_undo is False
        assert self.manager.can_redo is True

    def test_redo_basic(self):
        """Test basic redo operation."""
        cmd = CellChangeCommand(self.ss, 0, 0, "test")
        self.manager.execute(cmd)
        self.manager.undo()

        result = self.manager.redo()

        assert result == cmd
        assert self.manager.can_undo is True
        assert self.manager.can_redo is False

    def test_undo_empty_returns_none(self):
        """Test undo on empty stack returns None."""
        assert self.manager.undo() is None

    def test_redo_empty_returns_none(self):
        """Test redo on empty stack returns None."""
        assert self.manager.redo() is None

    def test_new_execute_clears_redo_stack(self):
        """Test new execute clears redo stack."""
        cmd1 = CellChangeCommand(self.ss, 0, 0, "test1")
        cmd2 = CellChangeCommand(self.ss, 0, 0, "test2")

        self.manager.execute(cmd1)
        self.manager.undo()
        assert self.manager.can_redo is True

        self.manager.execute(cmd2)
        assert self.manager.can_redo is False

    def test_clear(self):
        """Test clear removes all history."""
        cmd = CellChangeCommand(self.ss, 0, 0, "test")
        self.manager.execute(cmd)
        self.manager.undo()

        self.manager.clear()

        assert self.manager.can_undo is False
        assert self.manager.can_redo is False

    def test_max_history(self):
        """Test max history is respected."""
        manager = UndoManager(max_history=3)

        for i in range(5):
            cmd = CellChangeCommand(self.ss, 0, i, f"test{i}")
            manager.execute(cmd)

        assert manager.undo_count == 3

    def test_undo_description(self):
        """Test undo_description returns correct description."""
        cmd = CellChangeCommand(self.ss, 0, 0, "test")
        self.manager.execute(cmd)

        assert "Edit A1" in self.manager.undo_description

    def test_redo_description(self):
        """Test redo_description returns correct description."""
        cmd = CellChangeCommand(self.ss, 0, 0, "test")
        self.manager.execute(cmd)
        self.manager.undo()

        assert "Edit A1" in self.manager.redo_description


class TestCellChangeCommand:
    """Tests for CellChangeCommand class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()

    def test_execute(self):
        """Test execute sets cell value."""
        self.ss.set_cell(0, 0, "old")
        cmd = CellChangeCommand(self.ss, 0, 0, "new")
        cmd.execute()

        assert self.ss.get_value(0, 0) == "new"
        assert cmd.old_value == "old"

    def test_undo(self):
        """Test undo restores old value."""
        self.ss.set_cell(0, 0, "old")
        cmd = CellChangeCommand(self.ss, 0, 0, "new")
        cmd.execute()
        cmd.undo()

        assert self.ss.get_value(0, 0) == "old"

    def test_redo(self):
        """Test redo restores new value."""
        self.ss.set_cell(0, 0, "old")
        cmd = CellChangeCommand(self.ss, 0, 0, "new")
        cmd.execute()
        cmd.undo()
        cmd.redo()

        assert self.ss.get_value(0, 0) == "new"

    def test_with_format(self):
        """Test execute with format change."""
        cmd = CellChangeCommand(self.ss, 0, 0, "100", new_format="C2")
        cmd.execute()

        cell = self.ss.get_cell(0, 0)
        assert cell.format_code == "C2"

    def test_description(self):
        """Test description property."""
        cmd = CellChangeCommand(self.ss, 0, 0, "test")
        assert "A1" in cmd.description


class TestRangeChangeCommand:
    """Tests for RangeChangeCommand class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()

    def test_execute(self):
        """Test execute applies all changes."""
        self.ss.set_cell(0, 0, "old1")
        self.ss.set_cell(0, 1, "old2")

        cmd = RangeChangeCommand(
            self.ss,
            changes=[
                (0, 0, "new1", "old1"),
                (0, 1, "new2", "old2"),
            ],
        )
        cmd.execute()

        assert self.ss.get_value(0, 0) == "new1"
        assert self.ss.get_value(0, 1) == "new2"

    def test_undo(self):
        """Test undo restores all old values."""
        cmd = RangeChangeCommand(
            self.ss,
            changes=[
                (0, 0, "new1", "old1"),
                (0, 1, "new2", "old2"),
            ],
        )
        cmd.execute()
        cmd.undo()

        assert self.ss.get_value(0, 0) == "old1"
        assert self.ss.get_value(0, 1) == "old2"

    def test_description(self):
        """Test description property."""
        cmd = RangeChangeCommand(
            self.ss,
            changes=[
                (0, 0, "new1", "old1"),
                (0, 1, "new2", "old2"),
            ],
        )
        assert "2 cells" in cmd.description


class TestInsertRowCommand:
    """Tests for InsertRowCommand class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()

    def test_execute(self):
        """Test execute inserts row."""
        self.ss.set_cell(0, 0, "row0")
        self.ss.set_cell(1, 0, "row1")

        cmd = InsertRowCommand(self.ss, 1)
        cmd.execute()

        assert self.ss.get_value(0, 0) == "row0"
        assert self.ss.get_value(2, 0) == "row1"  # Shifted down

    def test_undo(self):
        """Test undo removes inserted row."""
        self.ss.set_cell(0, 0, "row0")
        self.ss.set_cell(1, 0, "row1")

        cmd = InsertRowCommand(self.ss, 1)
        cmd.execute()
        cmd.undo()

        assert self.ss.get_value(0, 0) == "row0"
        assert self.ss.get_value(1, 0) == "row1"

    def test_description(self):
        """Test description property."""
        cmd = InsertRowCommand(self.ss, 0)
        assert "row 1" in cmd.description


class TestDeleteRowCommand:
    """Tests for DeleteRowCommand class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()

    def test_execute(self):
        """Test execute deletes row."""
        self.ss.set_cell(0, 0, "row0")
        self.ss.set_cell(1, 0, "row1")
        self.ss.set_cell(2, 0, "row2")

        cmd = DeleteRowCommand(self.ss, 1)
        cmd.execute()

        assert self.ss.get_value(0, 0) == "row0"
        assert self.ss.get_value(1, 0) == "row2"  # Shifted up

    def test_undo(self):
        """Test undo restores deleted row."""
        self.ss.set_cell(0, 0, "row0")
        self.ss.set_cell(1, 0, "row1")
        self.ss.set_cell(2, 0, "row2")

        cmd = DeleteRowCommand(self.ss, 1)
        cmd.execute()
        cmd.undo()

        assert self.ss.get_value(0, 0) == "row0"
        assert self.ss.get_value(1, 0) == "row1"
        assert self.ss.get_value(2, 0) == "row2"

    def test_description(self):
        """Test description property."""
        cmd = DeleteRowCommand(self.ss, 0)
        assert "row 1" in cmd.description


class TestInsertColCommand:
    """Tests for InsertColCommand class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()

    def test_execute(self):
        """Test execute inserts column."""
        self.ss.set_cell(0, 0, "col0")
        self.ss.set_cell(0, 1, "col1")

        cmd = InsertColCommand(self.ss, 1)
        cmd.execute()

        assert self.ss.get_value(0, 0) == "col0"
        assert self.ss.get_value(0, 2) == "col1"  # Shifted right

    def test_undo(self):
        """Test undo removes inserted column."""
        self.ss.set_cell(0, 0, "col0")
        self.ss.set_cell(0, 1, "col1")

        cmd = InsertColCommand(self.ss, 1)
        cmd.execute()
        cmd.undo()

        assert self.ss.get_value(0, 0) == "col0"
        assert self.ss.get_value(0, 1) == "col1"

    def test_description(self):
        """Test description property."""
        cmd = InsertColCommand(self.ss, 0)
        assert "column A" in cmd.description


class TestDeleteColCommand:
    """Tests for DeleteColCommand class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()

    def test_execute(self):
        """Test execute deletes column."""
        self.ss.set_cell(0, 0, "col0")
        self.ss.set_cell(0, 1, "col1")
        self.ss.set_cell(0, 2, "col2")

        cmd = DeleteColCommand(self.ss, 1)
        cmd.execute()

        assert self.ss.get_value(0, 0) == "col0"
        assert self.ss.get_value(0, 1) == "col2"  # Shifted left

    def test_undo(self):
        """Test undo restores deleted column."""
        self.ss.set_cell(0, 0, "col0")
        self.ss.set_cell(0, 1, "col1")
        self.ss.set_cell(0, 2, "col2")

        cmd = DeleteColCommand(self.ss, 1)
        cmd.execute()
        cmd.undo()

        assert self.ss.get_value(0, 0) == "col0"
        assert self.ss.get_value(0, 1) == "col1"
        assert self.ss.get_value(0, 2) == "col2"

    def test_description(self):
        """Test description property."""
        cmd = DeleteColCommand(self.ss, 0)
        assert "column A" in cmd.description


class TestClearRangeCommand:
    """Tests for ClearRangeCommand class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()

    def test_execute(self):
        """Test execute clears range."""
        self.ss.set_cell(0, 0, "A")
        self.ss.set_cell(0, 1, "B")
        self.ss.set_cell(1, 0, "C")
        self.ss.set_cell(1, 1, "D")

        cmd = ClearRangeCommand(self.ss, 0, 0, 1, 1)
        cmd.execute()

        assert self.ss.get_value(0, 0) == ""
        assert self.ss.get_value(0, 1) == ""
        assert self.ss.get_value(1, 0) == ""
        assert self.ss.get_value(1, 1) == ""

    def test_undo(self):
        """Test undo restores cleared data."""
        self.ss.set_cell(0, 0, "A")
        self.ss.set_cell(0, 1, "B")

        cmd = ClearRangeCommand(self.ss, 0, 0, 0, 1)
        cmd.execute()
        cmd.undo()

        assert self.ss.get_value(0, 0) == "A"
        assert self.ss.get_value(0, 1) == "B"

    def test_redo(self):
        """Test redo clears data again."""
        self.ss.set_cell(0, 0, "A")

        cmd = ClearRangeCommand(self.ss, 0, 0, 0, 0)
        cmd.execute()
        cmd.undo()
        cmd.redo()

        assert self.ss.get_value(0, 0) == ""

    def test_description(self):
        """Test description property."""
        cmd = ClearRangeCommand(self.ss, 0, 0, 1, 1)
        assert "Clear" in cmd.description


class TestCompositeCommand:
    """Tests for CompositeCommand class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()

    def test_execute(self):
        """Test execute runs all commands."""
        cmd1 = CellChangeCommand(self.ss, 0, 0, "A")
        cmd2 = CellChangeCommand(self.ss, 0, 1, "B")

        composite = CompositeCommand([cmd1, cmd2], "Multiple edits")
        composite.execute()

        assert self.ss.get_value(0, 0) == "A"
        assert self.ss.get_value(0, 1) == "B"

    def test_undo(self):
        """Test undo reverses all commands."""
        self.ss.set_cell(0, 0, "old1")
        self.ss.set_cell(0, 1, "old2")

        cmd1 = CellChangeCommand(self.ss, 0, 0, "new1")
        cmd2 = CellChangeCommand(self.ss, 0, 1, "new2")

        composite = CompositeCommand([cmd1, cmd2], "Multiple edits")
        composite.execute()
        composite.undo()

        assert self.ss.get_value(0, 0) == "old1"
        assert self.ss.get_value(0, 1) == "old2"

    def test_redo(self):
        """Test redo re-runs all commands."""
        cmd1 = CellChangeCommand(self.ss, 0, 0, "A")
        cmd2 = CellChangeCommand(self.ss, 0, 1, "B")

        composite = CompositeCommand([cmd1, cmd2], "Multiple edits")
        composite.execute()
        composite.undo()
        composite.redo()

        assert self.ss.get_value(0, 0) == "A"
        assert self.ss.get_value(0, 1) == "B"

    def test_description(self):
        """Test description property."""
        composite = CompositeCommand([], "My description")
        assert composite.description == "My description"

    def test_composite_insert_rows(self):
        """Test CompositeCommand with multiple row inserts undoes all at once."""
        self.ss.set_cell(0, 0, "row0")
        self.ss.set_cell(1, 0, "row1")

        commands = [
            InsertRowCommand(self.ss, 1),
            InsertRowCommand(self.ss, 1),
            InsertRowCommand(self.ss, 1),
        ]
        composite = CompositeCommand(commands, "Insert 3 rows")
        composite.execute()

        # After 3 inserts at row 1, row1 data should be at row 4
        assert self.ss.get_value(0, 0) == "row0"
        assert self.ss.get_value(4, 0) == "row1"

        # Single undo should revert all 3 inserts
        composite.undo()
        assert self.ss.get_value(0, 0) == "row0"
        assert self.ss.get_value(1, 0) == "row1"

    def test_composite_insert_cols(self):
        """Test CompositeCommand with multiple column inserts undoes all at once."""
        self.ss.set_cell(0, 0, "col0")
        self.ss.set_cell(0, 1, "col1")

        commands = [
            InsertColCommand(self.ss, 1),
            InsertColCommand(self.ss, 1),
        ]
        composite = CompositeCommand(commands, "Insert 2 columns")
        composite.execute()

        # After 2 inserts at col 1, col1 data should be at col 3
        assert self.ss.get_value(0, 0) == "col0"
        assert self.ss.get_value(0, 3) == "col1"

        # Single undo should revert all 2 inserts
        composite.undo()
        assert self.ss.get_value(0, 0) == "col0"
        assert self.ss.get_value(0, 1) == "col1"

    def test_composite_delete_rows(self):
        """Test CompositeCommand with multiple row deletes undoes all at once."""
        self.ss.set_cell(0, 0, "row0")
        self.ss.set_cell(1, 0, "row1")
        self.ss.set_cell(2, 0, "row2")
        self.ss.set_cell(3, 0, "row3")

        commands = [
            DeleteRowCommand(self.ss, 1),
            DeleteRowCommand(self.ss, 1),
        ]
        composite = CompositeCommand(commands, "Delete 2 rows")
        composite.execute()

        # After 2 deletes at row 1, should have row0 and row3
        assert self.ss.get_value(0, 0) == "row0"
        assert self.ss.get_value(1, 0) == "row3"

        # Single undo should restore both rows
        composite.undo()
        assert self.ss.get_value(0, 0) == "row0"
        assert self.ss.get_value(1, 0) == "row1"
        assert self.ss.get_value(2, 0) == "row2"
        assert self.ss.get_value(3, 0) == "row3"


class TestDeleteRowColumnFormulaRestoration:
    """Tests for formula restoration when undoing row/column deletes.

    These tests verify that formulas referencing deleted rows/columns
    are properly restored to their original state after undo.
    """

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()

    def test_delete_column_restores_formula_referencing_deleted_column(self):
        """Test that deleting a column and undoing restores formulas that referenced it."""
        # Set up data in column A
        self.ss.set_cell(0, 0, "10")  # A1 = 10
        self.ss.set_cell(1, 0, "20")  # A2 = 20

        # Set up formula in column B that references column A
        self.ss.set_cell(0, 1, "=A1*2")  # B1 = A1*2 = 20
        self.ss.set_cell(1, 1, "=A2+5")  # B2 = A2+5 = 25

        # Verify formulas work
        assert self.ss.get_value(0, 1) == 20
        assert self.ss.get_value(1, 1) == 25

        # Delete column A
        cmd = DeleteColCommand(self.ss, 0)
        cmd.execute()

        # Now B moved to A, formulas should be broken (#REF!)
        # But we're testing undo, so let's undo
        cmd.undo()

        # After undo, formulas should be restored
        assert self.ss.get_value(0, 0) == 10  # A1 restored
        assert self.ss.get_value(1, 0) == 20  # A2 restored

        # Formulas should work again
        cell_b1 = self.ss.get_cell(0, 1)
        cell_b2 = self.ss.get_cell(1, 1)
        assert cell_b1.raw_value == "=A1*2"
        assert cell_b2.raw_value == "=A2+5"
        assert self.ss.get_value(0, 1) == 20
        assert self.ss.get_value(1, 1) == 25

    def test_delete_row_restores_formula_referencing_deleted_row(self):
        """Test that deleting a row and undoing restores formulas that referenced it."""
        # Set up data in row 1
        self.ss.set_cell(0, 0, "10")  # A1 = 10
        self.ss.set_cell(0, 1, "20")  # B1 = 20

        # Set up formula in row 2 that references row 1
        self.ss.set_cell(1, 0, "=A1*3")  # A2 = A1*3 = 30
        self.ss.set_cell(1, 1, "=B1-5")  # B2 = B1-5 = 15

        # Verify formulas work
        assert self.ss.get_value(1, 0) == 30
        assert self.ss.get_value(1, 1) == 15

        # Delete row 1 (index 0)
        cmd = DeleteRowCommand(self.ss, 0)
        cmd.execute()

        # Undo the delete
        cmd.undo()

        # After undo, formulas should be restored
        assert self.ss.get_value(0, 0) == 10  # A1 restored
        assert self.ss.get_value(0, 1) == 20  # B1 restored

        # Formulas should work again
        cell_a2 = self.ss.get_cell(1, 0)
        cell_b2 = self.ss.get_cell(1, 1)
        assert cell_a2.raw_value == "=A1*3"
        assert cell_b2.raw_value == "=B1-5"
        assert self.ss.get_value(1, 0) == 30
        assert self.ss.get_value(1, 1) == 15

    def test_delete_column_restores_sum_formula(self):
        """Test that SUM formulas referencing deleted columns are restored."""
        # Set up data
        self.ss.set_cell(0, 0, "1")  # A1
        self.ss.set_cell(0, 1, "2")  # B1
        self.ss.set_cell(0, 2, "3")  # C1
        self.ss.set_cell(0, 3, "=SUM(A1:C1)")  # D1 = SUM(A1:C1) = 6

        # Verify formula works
        assert self.ss.get_value(0, 3) == 6

        # Delete column B (index 1)
        cmd = DeleteColCommand(self.ss, 1)
        cmd.execute()

        # Undo the delete
        cmd.undo()

        # After undo, SUM formula should be restored
        cell_d1 = self.ss.get_cell(0, 3)
        assert cell_d1.raw_value == "=SUM(A1:C1)"
        assert self.ss.get_value(0, 3) == 6

    def test_delete_row_restores_sum_formula(self):
        """Test that SUM formulas referencing deleted rows are restored."""
        # Set up data
        self.ss.set_cell(0, 0, "1")  # A1
        self.ss.set_cell(1, 0, "2")  # A2
        self.ss.set_cell(2, 0, "3")  # A3
        self.ss.set_cell(3, 0, "=SUM(A1:A3)")  # A4 = SUM(A1:A3) = 6

        # Verify formula works
        assert self.ss.get_value(3, 0) == 6

        # Delete row 2 (index 1)
        cmd = DeleteRowCommand(self.ss, 1)
        cmd.execute()

        # Undo the delete
        cmd.undo()

        # After undo, SUM formula should be restored
        cell_a4 = self.ss.get_cell(3, 0)
        assert cell_a4.raw_value == "=SUM(A1:A3)"
        assert self.ss.get_value(3, 0) == 6

    def test_delete_column_multiple_formulas_restored(self):
        """Test that multiple formulas are all restored after column delete undo."""
        # Set up data
        self.ss.set_cell(0, 0, "10")  # A1
        self.ss.set_cell(1, 0, "20")  # A2
        self.ss.set_cell(2, 0, "30")  # A3

        # Set up multiple formulas in column B referencing column A
        self.ss.set_cell(0, 1, "=A1+1")  # B1
        self.ss.set_cell(1, 1, "=A2+2")  # B2
        self.ss.set_cell(2, 1, "=A3+3")  # B3
        self.ss.set_cell(3, 1, "=SUM(A1:A3)")  # B4

        # Delete column A
        cmd = DeleteColCommand(self.ss, 0)
        cmd.execute()

        # Undo
        cmd.undo()

        # All formulas should be restored
        assert self.ss.get_cell(0, 1).raw_value == "=A1+1"
        assert self.ss.get_cell(1, 1).raw_value == "=A2+2"
        assert self.ss.get_cell(2, 1).raw_value == "=A3+3"
        assert self.ss.get_cell(3, 1).raw_value == "=SUM(A1:A3)"

        # Values should be correct
        assert self.ss.get_value(0, 1) == 11
        assert self.ss.get_value(1, 1) == 22
        assert self.ss.get_value(2, 1) == 33
        assert self.ss.get_value(3, 1) == 60

    def test_redo_after_undo_delete_column(self):
        """Test that redo after undo of column delete works correctly."""
        self.ss.set_cell(0, 0, "10")  # A1
        self.ss.set_cell(0, 1, "=A1*2")  # B1

        cmd = DeleteColCommand(self.ss, 0)
        cmd.execute()

        # Undo
        cmd.undo()
        assert self.ss.get_value(0, 0) == 10
        assert self.ss.get_value(0, 1) == 20

        # Redo
        cmd.redo()

        # After redo, column A should be deleted again
        # B1 (now A1) should have the formula but with #REF!
        # The important thing is redo doesn't crash

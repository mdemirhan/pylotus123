"""Tests for undo/redo operations."""

import pytest

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

        cmd = RangeChangeCommand(self.ss, changes=[
            (0, 0, "new1", "old1"),
            (0, 1, "new2", "old2"),
        ])
        cmd.execute()

        assert self.ss.get_value(0, 0) == "new1"
        assert self.ss.get_value(0, 1) == "new2"

    def test_undo(self):
        """Test undo restores all old values."""
        cmd = RangeChangeCommand(self.ss, changes=[
            (0, 0, "new1", "old1"),
            (0, 1, "new2", "old2"),
        ])
        cmd.execute()
        cmd.undo()

        assert self.ss.get_value(0, 0) == "old1"
        assert self.ss.get_value(0, 1) == "old2"

    def test_description(self):
        """Test description property."""
        cmd = RangeChangeCommand(self.ss, changes=[
            (0, 0, "new1", "old1"),
            (0, 1, "new2", "old2"),
        ])
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

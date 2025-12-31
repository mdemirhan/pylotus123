"""Undo/Redo system using the Command pattern."""

from __future__ import annotations

from abc import ABC, abstractmethod
from collections import deque
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


class Command(ABC):
    """Abstract base class for undoable commands."""

    @abstractmethod
    def execute(self) -> None:
        """Execute the command."""
        pass

    @abstractmethod
    def undo(self) -> None:
        """Undo the command."""
        pass

    @abstractmethod
    def redo(self) -> None:
        """Redo the command (usually same as execute)."""
        pass

    @property
    def description(self) -> str:
        """Human-readable description of the command."""
        return self.__class__.__name__


@dataclass
class CellChangeCommand(Command):
    """Command for changing a single cell's value."""

    spreadsheet: Spreadsheet
    row: int
    col: int
    new_value: str
    old_value: str = ""
    old_format: str = "G"
    new_format: str | None = None

    def execute(self) -> None:
        """Set the new value."""
        cell = self.spreadsheet.get_cell(self.row, self.col)
        self.old_value = cell.raw_value
        self.old_format = cell.format_code
        cell.set_value(self.new_value)
        if self.new_format:
            cell.format_code = self.new_format

        # Update dependencies
        formula = cell.formula if cell.is_formula else None
        self.spreadsheet.update_cell_dependency(self.row, self.col, formula)

        self.spreadsheet._invalidate_cache()

    def undo(self) -> None:
        """Restore the old value."""
        cell = self.spreadsheet.get_cell(self.row, self.col)
        cell.set_value(self.old_value)
        cell.format_code = self.old_format

        # Update dependencies
        formula = cell.formula if cell.is_formula else None
        self.spreadsheet.update_cell_dependency(self.row, self.col, formula)

        self.spreadsheet._invalidate_cache()

    def redo(self) -> None:
        """Set the new value again."""
        cell = self.spreadsheet.get_cell(self.row, self.col)
        cell.set_value(self.new_value)
        if self.new_format:
            cell.format_code = self.new_format

        # Update dependencies
        formula = cell.formula if cell.is_formula else None
        self.spreadsheet.update_cell_dependency(self.row, self.col, formula)

        self.spreadsheet._invalidate_cache()

    @property
    def description(self) -> str:
        from ..core.reference import make_cell_ref

        return f"Edit {make_cell_ref(self.row, self.col)}"


@dataclass
class RangeChangeCommand(Command):
    """Command for changing a range of cells."""

    spreadsheet: Spreadsheet
    changes: list[tuple[int, int, str, str]] = field(default_factory=list)
    # Each tuple: (row, col, new_value, old_value)

    def execute(self) -> None:
        """Apply all changes."""
        for row, col, new_val, _ in self.changes:
            cell = self.spreadsheet.get_cell(row, col)
            cell.set_value(new_val)
            # Update dependencies
            formula = cell.formula if cell.is_formula else None
            self.spreadsheet.update_cell_dependency(row, col, formula)
        self.spreadsheet._invalidate_cache()

    def undo(self) -> None:
        """Restore all old values."""
        for row, col, _, old_val in self.changes:
            cell = self.spreadsheet.get_cell(row, col)
            cell.set_value(old_val)
            # Update dependencies
            formula = cell.formula if cell.is_formula else None
            self.spreadsheet.update_cell_dependency(row, col, formula)
        self.spreadsheet._invalidate_cache()

    def redo(self) -> None:
        """Apply all changes again."""
        self.execute()

    @property
    def description(self) -> str:
        return f"Edit {len(self.changes)} cells"


@dataclass
class InsertRowCommand(Command):
    """Command for inserting a row."""

    spreadsheet: Spreadsheet
    row: int
    deleted_data: dict = field(default_factory=dict)

    def execute(self) -> None:
        """Insert the row."""
        self.spreadsheet.insert_row(self.row)

    def undo(self) -> None:
        """Delete the inserted row."""
        self.spreadsheet.delete_row(self.row)

    def redo(self) -> None:
        """Insert the row again."""
        self.execute()

    @property
    def description(self) -> str:
        return f"Insert row {self.row + 1}"


@dataclass
class DeleteRowCommand(Command):
    """Command for deleting a row."""

    spreadsheet: Spreadsheet
    row: int
    saved_data: dict = field(default_factory=dict)
    saved_formulas: dict = field(default_factory=dict)

    def execute(self) -> None:
        """Delete the row, saving its data."""
        # Save all cells in this row
        self.saved_data = {}
        for (r, c), cell in list(self.spreadsheet._cells.items()):
            if r == self.row:
                self.saved_data[c] = cell.to_dict()

        # Save ALL formula cells before deletion (they may be modified by adjust_for_structural_change)
        self.saved_formulas = {}
        for (r, c), cell in self.spreadsheet._cells.items():
            if cell.is_formula:
                self.saved_formulas[(r, c)] = cell.raw_value

        self.spreadsheet.delete_row(self.row)

    def undo(self) -> None:
        """Restore the deleted row."""
        # First insert the row back
        self.spreadsheet.insert_row(self.row)

        # Restore saved cells - use the Cell class from spreadsheet module
        from ..core.cell import Cell

        for col, cell_data in self.saved_data.items():
            self.spreadsheet._cells[(self.row, col)] = Cell.from_dict(cell_data)

        # Restore all formulas to their original state
        # After insert_row, cells are back to their original positions
        for (r, c), formula in self.saved_formulas.items():
            cell = self.spreadsheet._cells.get((r, c))
            if cell and cell.is_formula:
                cell.set_value(formula)

        self.spreadsheet._invalidate_cache()
        self.spreadsheet.rebuild_dependency_graph()

    def redo(self) -> None:
        """Delete the row again."""
        self.spreadsheet.delete_row(self.row)

    @property
    def description(self) -> str:
        return f"Delete row {self.row + 1}"


@dataclass
class InsertColCommand(Command):
    """Command for inserting a column."""

    spreadsheet: Spreadsheet
    col: int

    def execute(self) -> None:
        """Insert the column."""
        self.spreadsheet.insert_col(self.col)

    def undo(self) -> None:
        """Delete the inserted column."""
        self.spreadsheet.delete_col(self.col)

    def redo(self) -> None:
        """Insert the column again."""
        self.execute()

    @property
    def description(self) -> str:
        from ..core.reference import index_to_col

        return f"Insert column {index_to_col(self.col)}"


@dataclass
class DeleteColCommand(Command):
    """Command for deleting a column."""

    spreadsheet: Spreadsheet
    col: int
    saved_data: dict = field(default_factory=dict)
    saved_width: int | None = None
    saved_formulas: dict = field(default_factory=dict)

    def execute(self) -> None:
        """Delete the column, saving its data."""
        # Save all cells in this column
        self.saved_data = {}
        for (r, c), cell in list(self.spreadsheet._cells.items()):
            if c == self.col:
                self.saved_data[r] = cell.to_dict()

        # Save column width
        self.saved_width = self.spreadsheet.get_col_width(self.col)

        # Save ALL formula cells before deletion (they may be modified by adjust_for_structural_change)
        self.saved_formulas = {}
        for (r, c), cell in self.spreadsheet._cells.items():
            if cell.is_formula:
                self.saved_formulas[(r, c)] = cell.raw_value

        self.spreadsheet.delete_col(self.col)

    def undo(self) -> None:
        """Restore the deleted column."""
        # First insert the column back
        self.spreadsheet.insert_col(self.col)

        # Restore saved cells - use the Cell class from spreadsheet module
        from ..core.cell import Cell

        for row, cell_data in self.saved_data.items():
            self.spreadsheet._cells[(row, self.col)] = Cell.from_dict(cell_data)

        # Restore width
        if self.saved_width is not None:
            self.spreadsheet.set_col_width(self.col, self.saved_width)

        # Restore all formulas to their original state
        # After insert_col, cells are back to their original positions
        for (r, c), formula in self.saved_formulas.items():
            cell = self.spreadsheet._cells.get((r, c))
            if cell and cell.is_formula:
                cell.set_value(formula)

        self.spreadsheet._invalidate_cache()
        self.spreadsheet.rebuild_dependency_graph()

    def redo(self) -> None:
        """Delete the column again."""
        self.spreadsheet.delete_col(self.col)

    @property
    def description(self) -> str:
        from ..core.reference import index_to_col

        return f"Delete column {index_to_col(self.col)}"


@dataclass
class ClearRangeCommand(Command):
    """Command for clearing a range of cells."""

    spreadsheet: Spreadsheet
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    saved_data: dict = field(default_factory=dict)

    def execute(self) -> None:
        """Clear the range, saving its data."""
        self.saved_data = {}
        for r in range(self.start_row, self.end_row + 1):
            for c in range(self.start_col, self.end_col + 1):
                cell = self.spreadsheet.get_cell_if_exists(r, c)
                if cell and not cell.is_empty:
                    self.saved_data[(r, c)] = cell.to_dict()
                    cell.set_value("")
                    # Update dependencies (remove)
                    self.spreadsheet.update_cell_dependency(r, c, None)

        self.spreadsheet._invalidate_cache()

    def undo(self) -> None:
        """Restore the cleared data."""
        from ..core.cell import Cell

        for (r, c), cell_data in self.saved_data.items():
            self.spreadsheet._cells[(r, c)] = cell = Cell.from_dict(cell_data)
            # Update dependencies
            formula = cell.formula if cell.is_formula else None
            self.spreadsheet.update_cell_dependency(r, c, formula)

        self.spreadsheet._invalidate_cache()

    def redo(self) -> None:
        """Clear the range again."""
        for r in range(self.start_row, self.end_row + 1):
            for c in range(self.start_col, self.end_col + 1):
                cell = self.spreadsheet.get_cell_if_exists(r, c)
                if cell:
                    cell.set_value("")
                    # Update dependencies (remove)
                    self.spreadsheet.update_cell_dependency(r, c, None)

        self.spreadsheet._invalidate_cache()

    @property
    def description(self) -> str:
        return "Clear range"


@dataclass
class RangeFormatCommand(Command):
    """Command for changing the format of a range of cells."""

    spreadsheet: Spreadsheet
    changes: list[tuple[int, int, str, str]] = field(default_factory=list)
    # Each tuple: (row, col, new_format, old_format)

    def execute(self) -> None:
        """Apply format to all cells."""
        for row, col, new_fmt, _ in self.changes:
            cell = self.spreadsheet.get_cell(row, col)
            cell.format_code = new_fmt

    def undo(self) -> None:
        """Restore all old formats."""
        for row, col, _, old_fmt in self.changes:
            cell = self.spreadsheet.get_cell(row, col)
            cell.format_code = old_fmt

    def redo(self) -> None:
        """Apply format again."""
        self.execute()

    @property
    def description(self) -> str:
        return f"Format {len(self.changes)} cells"


class CompositeCommand(Command):
    """Command that groups multiple commands."""

    def __init__(self, commands: list[Command], description: str = "Multiple changes") -> None:
        self._commands = commands
        self._description = description

    def execute(self) -> None:
        """Execute all commands in order."""
        for cmd in self._commands:
            cmd.execute()

    def undo(self) -> None:
        """Undo all commands in reverse order."""
        for cmd in reversed(self._commands):
            cmd.undo()

    def redo(self) -> None:
        """Redo all commands in order."""
        for cmd in self._commands:
            cmd.redo()

    @property
    def description(self) -> str:
        return self._description


class UndoManager:
    """Manages undo and redo operations.

    Uses a stack-based approach with configurable history size.
    """

    def __init__(self, max_history: int = 100) -> None:
        self.max_history = max_history
        self._undo_stack: deque[Command] = deque(maxlen=max_history)
        self._redo_stack: deque[Command] = deque(maxlen=max_history)

    def execute(self, command: Command) -> None:
        """Execute a command and add it to the undo stack."""
        command.execute()
        self._undo_stack.append(command)
        self._redo_stack.clear()  # Clear redo stack on new action

    def undo(self) -> Command | None:
        """Undo the last command.

        Returns:
            The command that was undone, or None if nothing to undo
        """
        if not self._undo_stack:
            return None

        command = self._undo_stack.pop()
        command.undo()
        self._redo_stack.append(command)
        return command

    def redo(self) -> Command | None:
        """Redo the last undone command.

        Returns:
            The command that was redone, or None if nothing to redo
        """
        if not self._redo_stack:
            return None

        command = self._redo_stack.pop()
        command.redo()
        self._undo_stack.append(command)
        return command

    @property
    def can_undo(self) -> bool:
        """Check if there's something to undo."""
        return len(self._undo_stack) > 0

    @property
    def can_redo(self) -> bool:
        """Check if there's something to redo."""
        return len(self._redo_stack) > 0

    @property
    def undo_description(self) -> str:
        """Get description of next undo operation."""
        if self._undo_stack:
            return self._undo_stack[-1].description
        return ""

    @property
    def redo_description(self) -> str:
        """Get description of next redo operation."""
        if self._redo_stack:
            return self._redo_stack[-1].description
        return ""

    def clear(self) -> None:
        """Clear all undo/redo history."""
        self._undo_stack.clear()
        self._redo_stack.clear()

    @property
    def undo_count(self) -> int:
        """Number of operations that can be undone."""
        return len(self._undo_stack)

    @property
    def redo_count(self) -> int:
        """Number of operations that can be redone."""
        return len(self._redo_stack)

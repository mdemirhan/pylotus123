"""Spreadsheet data model with cells and formula support.

This module provides backwards compatibility by re-exporting from the new
modular core package. New code should import from lotus123.core directly.
"""
from __future__ import annotations

# Re-export from new modular structure for backwards compatibility
from .core.reference import (
    col_to_index,
    index_to_col,
    parse_cell_ref,
    make_cell_ref,
)
from .core.cell import Cell as _NewCell
from .core.spreadsheet import Spreadsheet as _NewSpreadsheet

import json
from dataclasses import dataclass
from typing import Any


# Lotus 1-2-3 alignment prefix characters
ALIGNMENT_PREFIXES = {"'", '"', "^", "\\"}


# Backwards-compatible Cell class that mirrors the old interface
@dataclass
class Cell:
    """Represents a single cell in the spreadsheet.

    This is a backwards-compatible wrapper. New code should use lotus123.core.Cell.
    """
    raw_value: str = ""
    format_str: str = ""
    width: int = 10

    @property
    def format_code(self) -> str:
        """Alias for format_str for compatibility with core.Cell."""
        return self.format_str or "G"

    @format_code.setter
    def format_code(self, value: str) -> None:
        """Set format code."""
        self.format_str = value

    @property
    def is_empty(self) -> bool:
        """Check if cell has no value."""
        return not self.raw_value

    @property
    def is_formula(self) -> bool:
        return (self.raw_value.startswith('=') or
                self.raw_value.startswith('+') or
                self.raw_value.startswith('@'))

    @property
    def formula(self) -> str:
        """Return the formula without the leading =, + or @."""
        if self.is_formula:
            return self.raw_value[1:]
        return ""

    @property
    def display_value(self) -> str:
        """Get the value for display, stripping alignment prefixes.

        Lotus 1-2-3 uses prefix characters for alignment:
        - ' (apostrophe) = left aligned
        - " (quote) = right aligned
        - ^ (caret) = centered
        - \\ (backslash) = repeating
        """
        if not self.raw_value:
            return ""
        if self.raw_value[0] in ALIGNMENT_PREFIXES:
            return self.raw_value[1:]
        return self.raw_value

    def set_value(self, value: str) -> None:
        """Set the cell's raw value."""
        self.raw_value = value

    def to_dict(self) -> dict:
        return {"raw_value": self.raw_value, "format_str": self.format_str, "width": self.width}

    @classmethod
    def from_dict(cls, data: dict) -> Cell:
        return cls(
            raw_value=data.get("raw_value", ""),
            format_str=data.get("format_str", ""),
            width=data.get("width", 10),
        )


class Spreadsheet:
    """Main spreadsheet class managing a grid of cells.

    This is a backwards-compatible wrapper around the new core.Spreadsheet.
    New code should use lotus123.core.Spreadsheet directly for full features.

    Grid size matches original Lotus 1-2-3 Release 2: 256 columns × 8192 rows.
    """

    def __init__(self, rows: int = 8192, cols: int = 256):
        self.rows = rows
        self.cols = cols
        self._cells: dict[tuple[int, int], Cell] = {}
        self._cache: dict[tuple[int, int], Any] = {}
        self._computing: set[tuple[int, int]] = set()
        self._circular_refs: set[tuple[int, int]] = set()
        self.col_widths: dict[int, int] = {}
        self.filename: str = ""
        self.modified: bool = False
        # For new formula parser compatibility
        self.named_ranges = _DummyNamedRanges()

    def get_cell(self, row: int, col: int) -> Cell:
        """Get cell at position, creating if needed."""
        if (row, col) not in self._cells:
            self._cells[(row, col)] = Cell()
        return self._cells[(row, col)]

    def get_cell_if_exists(self, row: int, col: int) -> Cell | None:
        """Get cell if it exists, without creating."""
        return self._cells.get((row, col))

    def set_cell(self, row: int, col: int, value: str) -> None:
        """Set cell value and invalidate cache."""
        cell = self.get_cell(row, col)
        cell.raw_value = value
        self._invalidate_cache()

    def set_cell_by_ref(self, ref: str, value: str) -> None:
        """Set cell by reference like 'A1'."""
        row, col = parse_cell_ref(ref)
        self.set_cell(row, col, value)

    def get_value(self, row: int, col: int) -> Any:
        """Get computed value of cell."""
        if (row, col) in self._cache:
            return self._cache[(row, col)]

        cell = self._cells.get((row, col))
        if not cell or not cell.raw_value:
            return ""

        if cell.is_formula:
            value = self._evaluate_formula(row, col)
        else:
            # Use display_value to strip alignment prefixes
            value = self._parse_literal(cell.display_value)

        self._cache[(row, col)] = value
        return value

    def get_value_by_ref(self, ref: str) -> Any:
        """Get computed value by reference like 'A1'."""
        row, col = parse_cell_ref(ref)
        return self.get_value(row, col)

    def _parse_literal(self, value: str) -> Any:
        """Parse a literal value (number or string)."""
        try:
            if '.' in value:
                return float(value)
            return int(value)
        except ValueError:
            return value

    def _evaluate_formula(self, row: int, col: int) -> Any:
        """Evaluate a formula at the given cell position."""
        from .formula.parser import FormulaParser

        if (row, col) in self._computing:
            return "#CIRC!"

        self._computing.add((row, col))
        try:
            cell = self.get_cell(row, col)
            parser = FormulaParser(self)
            result = parser.evaluate(cell.formula)
            return result
        except Exception as e:
            return f"#ERR!"
        finally:
            self._computing.discard((row, col))

    def _invalidate_cache(self) -> None:
        """Clear the computation cache."""
        self._cache.clear()

    @property
    def needs_recalc(self) -> bool:
        """Check if spreadsheet needs recalculation."""
        return len(self._cache) == 0 and any(
            c.is_formula for c in self._cells.values() if c.raw_value
        )

    @property
    def has_circular_refs(self) -> bool:
        """Check if circular references were detected."""
        return len(self._circular_refs) > 0

    def get_display_value(self, row: int, col: int) -> str:
        """Get string representation for display, using cell's format code."""
        from .core.formatting import parse_format_code, format_value

        value = self.get_value(row, col)
        cell = self._cells.get((row, col))

        # Use cell's format code if set
        if cell and cell.format_str:
            spec = parse_format_code(cell.format_str)
            return format_value(value, spec)

        # Default formatting for unformatted cells
        if isinstance(value, float):
            if value == int(value):
                return str(int(value))
            return f"{value:.2f}"
        return str(value)

    def get_col_width(self, col: int) -> int:
        """Get width of column."""
        return self.col_widths.get(col, 10)

    def set_col_width(self, col: int, width: int) -> None:
        """Set width of column."""
        self.col_widths[col] = max(3, min(50, width))

    def get_range(self, start_ref: str, end_ref: str) -> list[list[Any]]:
        """Get values in a range like A1:B5."""
        start_row, start_col = parse_cell_ref(start_ref)
        end_row, end_col = parse_cell_ref(end_ref)

        if start_row > end_row:
            start_row, end_row = end_row, start_row
        if start_col > end_col:
            start_col, end_col = end_col, start_col

        result = []
        for r in range(start_row, end_row + 1):
            row_vals = []
            for c in range(start_col, end_col + 1):
                row_vals.append(self.get_value(r, c))
            result.append(row_vals)
        return result

    def get_range_flat(self, start_ref: str, end_ref: str) -> list[Any]:
        """Get values in a range as flat list."""
        rows = self.get_range(start_ref, end_ref)
        return [val for row in rows for val in row]

    def save(self, filename: str) -> None:
        """Save spreadsheet to JSON file."""
        data = {
            "rows": self.rows,
            "cols": self.cols,
            "col_widths": self.col_widths,
            "cells": {f"{r},{c}": cell.to_dict() for (r, c), cell in self._cells.items() if cell.raw_value}
        }
        with open(filename, 'w') as f:
            json.dump(data, f, indent=2)
        self.filename = filename

    def load(self, filename: str) -> None:
        """Load spreadsheet from JSON file."""
        with open(filename, 'r') as f:
            data = json.load(f)

        self.rows = data.get("rows", 100)
        self.cols = data.get("cols", 26)
        self.col_widths = {int(k): v for k, v in data.get("col_widths", {}).items()}
        self._cells.clear()

        for key, cell_data in data.get("cells", {}).items():
            r, c = map(int, key.split(','))
            self._cells[(r, c)] = Cell.from_dict(cell_data)

        self._invalidate_cache()
        self.filename = filename

    def clear(self) -> None:
        """Clear all cells."""
        self._cells.clear()
        self._cache.clear()
        self.col_widths.clear()

    def delete_row(self, row: int) -> None:
        """Delete a row and shift cells up."""
        new_cells = {}
        for (r, c), cell in self._cells.items():
            if r < row:
                new_cells[(r, c)] = cell
            elif r > row:
                new_cells[(r - 1, c)] = cell
        self._cells = new_cells
        self._invalidate_cache()

    def insert_row(self, row: int) -> None:
        """Insert a row and shift cells down."""
        new_cells = {}
        for (r, c), cell in self._cells.items():
            if r < row:
                new_cells[(r, c)] = cell
            else:
                new_cells[(r + 1, c)] = cell
        self._cells = new_cells
        self._invalidate_cache()

    def delete_col(self, col: int) -> None:
        """Delete a column and shift cells left."""
        new_cells = {}
        for (r, c), cell in self._cells.items():
            if c < col:
                new_cells[(r, c)] = cell
            elif c > col:
                new_cells[(r, c - 1)] = cell
        self._cells = new_cells
        self._invalidate_cache()

    def insert_col(self, col: int) -> None:
        """Insert a column and shift cells right."""
        new_cells = {}
        for (r, c), cell in self._cells.items():
            if c < col:
                new_cells[(r, c)] = cell
            else:
                new_cells[(r, c + 1)] = cell
        self._cells = new_cells
        self._invalidate_cache()

    def copy_cell(self, from_row: int, from_col: int, to_row: int, to_col: int) -> None:
        """Copy a cell to another location."""
        src = self.get_cell(from_row, from_col)
        self.set_cell(to_row, to_col, src.raw_value)


class _DummyNamedRanges:
    """Dummy named ranges for backwards compatibility."""

    def exists(self, name: str) -> bool:
        return False

    def get(self, name: str) -> None:
        return None

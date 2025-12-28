"""Tests for spreadsheet operations."""
import pytest
import json
import tempfile
import os
from lotus123.spreadsheet import Spreadsheet, Cell
from lotus123.spreadsheet import col_to_index, index_to_col, parse_cell_ref, make_cell_ref


class TestCellReferenceUtils:
    """Tests for cell reference utility functions."""

    def test_col_to_index_single_letter(self):
        """Test single letter column conversion."""
        assert col_to_index("A") == 0
        assert col_to_index("B") == 1
        assert col_to_index("Z") == 25

    def test_col_to_index_double_letter(self):
        """Test double letter column conversion."""
        assert col_to_index("AA") == 26
        assert col_to_index("AB") == 27
        assert col_to_index("AZ") == 51
        assert col_to_index("BA") == 52

    def test_col_to_index_triple_letter(self):
        """Test triple letter column conversion."""
        assert col_to_index("AAA") == 702

    def test_index_to_col_single_letter(self):
        """Test index to single letter column."""
        assert index_to_col(0) == "A"
        assert index_to_col(1) == "B"
        assert index_to_col(25) == "Z"

    def test_index_to_col_double_letter(self):
        """Test index to double letter column."""
        assert index_to_col(26) == "AA"
        assert index_to_col(27) == "AB"
        assert index_to_col(51) == "AZ"
        assert index_to_col(52) == "BA"

    def test_parse_cell_ref(self):
        """Test parsing cell references."""
        assert parse_cell_ref("A1") == (0, 0)
        assert parse_cell_ref("B2") == (1, 1)
        assert parse_cell_ref("Z10") == (9, 25)
        assert parse_cell_ref("AA100") == (99, 26)

    def test_parse_cell_ref_lowercase(self):
        """Test parsing lowercase cell references."""
        assert parse_cell_ref("a1") == (0, 0)
        assert parse_cell_ref("b2") == (1, 1)

    def test_make_cell_ref(self):
        """Test making cell references."""
        assert make_cell_ref(0, 0) == "A1"
        assert make_cell_ref(1, 1) == "B2"
        assert make_cell_ref(9, 25) == "Z10"


class TestCell:
    """Tests for Cell class."""

    def test_empty_cell(self):
        """Test empty cell defaults."""
        cell = Cell()
        assert cell.raw_value == ""
        assert cell.format_str == ""
        assert cell.width == 10
        assert cell.is_formula is False
        assert cell.formula == ""

    def test_value_cell(self):
        """Test cell with value."""
        cell = Cell(raw_value="Hello")
        assert cell.raw_value == "Hello"
        assert cell.is_formula is False

    def test_formula_cell_equals(self):
        """Test cell with = formula."""
        cell = Cell(raw_value="=A1+B1")
        assert cell.is_formula is True
        assert cell.formula == "A1+B1"

    def test_formula_cell_plus(self):
        """Test cell with + formula."""
        cell = Cell(raw_value="+A1+B1")
        assert cell.is_formula is True
        assert cell.formula == "A1+B1"

    def test_formula_cell_at(self):
        """Test cell with @ formula."""
        cell = Cell(raw_value="@SUM(A1:A10)")
        assert cell.is_formula is True
        assert cell.formula == "SUM(A1:A10)"

    def test_cell_to_dict(self):
        """Test cell serialization."""
        cell = Cell(raw_value="test", format_str="0.00", width=15)
        d = cell.to_dict()
        assert d["raw_value"] == "test"
        assert d["format_str"] == "0.00"
        assert d["width"] == 15

    def test_cell_from_dict(self):
        """Test cell deserialization."""
        d = {"raw_value": "test", "format_str": "0.00", "width": 15}
        cell = Cell.from_dict(d)
        assert cell.raw_value == "test"
        assert cell.format_str == "0.00"
        assert cell.width == 15


class TestSpreadsheetBasics:
    """Basic spreadsheet operation tests."""

    def test_initialization(self):
        """Test spreadsheet initialization.

        Default grid size matches Lotus 1-2-3 Release 2: 256 columns × 8192 rows.
        """
        ss = Spreadsheet()
        assert ss.rows == 8192
        assert ss.cols == 256

    def test_custom_size(self):
        """Test custom spreadsheet size."""
        ss = Spreadsheet(rows=50, cols=10)
        assert ss.rows == 50
        assert ss.cols == 10

    def test_set_and_get_cell(self):
        """Test setting and getting cell values."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "Hello")
        cell = ss.get_cell(0, 0)
        assert cell.raw_value == "Hello"

    def test_get_value_number(self):
        """Test getting numeric value."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "123")
        assert ss.get_value(0, 0) == 123

    def test_get_value_float(self):
        """Test getting float value."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "123.45")
        assert ss.get_value(0, 0) == 123.45

    def test_get_value_string(self):
        """Test getting string value."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "Hello")
        assert ss.get_value(0, 0) == "Hello"

    def test_get_empty_cell(self):
        """Test getting empty cell value."""
        ss = Spreadsheet()
        assert ss.get_value(0, 0) == ""

    def test_set_cell_by_ref(self):
        """Test setting cell by reference."""
        ss = Spreadsheet()
        ss.set_cell_by_ref("A1", "Test")
        assert ss.get_value(0, 0) == "Test"

    def test_get_value_by_ref(self):
        """Test getting value by reference."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "Test")
        assert ss.get_value_by_ref("A1") == "Test"

    def test_get_cell_if_exists(self):
        """Test get_cell_if_exists method."""
        ss = Spreadsheet()
        assert ss.get_cell_if_exists(0, 0) is None

        ss.set_cell(0, 0, "Test")
        cell = ss.get_cell_if_exists(0, 0)
        assert cell is not None
        assert cell.raw_value == "Test"


class TestSpreadsheetRange:
    """Tests for range operations."""

    def test_get_range(self):
        """Test getting range of values."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "1")
        ss.set_cell(0, 1, "2")
        ss.set_cell(1, 0, "3")
        ss.set_cell(1, 1, "4")

        result = ss.get_range("A1", "B2")
        assert result == [[1, 2], [3, 4]]

    def test_get_range_flat(self):
        """Test getting flattened range."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "1")
        ss.set_cell(0, 1, "2")
        ss.set_cell(1, 0, "3")
        ss.set_cell(1, 1, "4")

        result = ss.get_range_flat("A1", "B2")
        assert result == [1, 2, 3, 4]

    def test_get_range_reversed(self):
        """Test range with reversed coordinates."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "1")
        ss.set_cell(1, 1, "4")

        result = ss.get_range("B2", "A1")
        assert result == [[1, ""], ["", 4]]


class TestSpreadsheetColumnWidth:
    """Tests for column width management."""

    def test_default_col_width(self):
        """Test default column width."""
        ss = Spreadsheet()
        assert ss.get_col_width(0) == 10

    def test_set_col_width(self):
        """Test setting column width."""
        ss = Spreadsheet()
        ss.set_col_width(0, 20)
        assert ss.get_col_width(0) == 20

    def test_col_width_min_bound(self):
        """Test column width minimum bound."""
        ss = Spreadsheet()
        ss.set_col_width(0, 1)
        assert ss.get_col_width(0) == 3

    def test_col_width_max_bound(self):
        """Test column width maximum bound."""
        ss = Spreadsheet()
        ss.set_col_width(0, 100)
        assert ss.get_col_width(0) == 50


class TestSpreadsheetDisplayValue:
    """Tests for display value formatting."""

    def test_display_int(self):
        """Test displaying integer."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "123")
        assert ss.get_display_value(0, 0) == "123"

    def test_display_float(self):
        """Test displaying float."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "123.456")
        assert ss.get_display_value(0, 0) == "123.46"

    def test_display_whole_float(self):
        """Test displaying float that equals integer."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "=10/2")
        assert ss.get_display_value(0, 0) == "5"

    def test_display_string(self):
        """Test displaying string."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "Hello")
        assert ss.get_display_value(0, 0) == "Hello"


class TestSpreadsheetRowColumnOps:
    """Tests for row and column operations."""

    def test_delete_row(self):
        """Test deleting a row."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "A")
        ss.set_cell(1, 0, "B")
        ss.set_cell(2, 0, "C")

        ss.delete_row(1)

        assert ss.get_value(0, 0) == "A"
        assert ss.get_value(1, 0) == "C"

    def test_insert_row(self):
        """Test inserting a row."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "A")
        ss.set_cell(1, 0, "B")

        ss.insert_row(1)

        assert ss.get_value(0, 0) == "A"
        assert ss.get_value(1, 0) == ""
        assert ss.get_value(2, 0) == "B"

    def test_delete_col(self):
        """Test deleting a column."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "A")
        ss.set_cell(0, 1, "B")
        ss.set_cell(0, 2, "C")

        ss.delete_col(1)

        assert ss.get_value(0, 0) == "A"
        assert ss.get_value(0, 1) == "C"

    def test_insert_col(self):
        """Test inserting a column."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "A")
        ss.set_cell(0, 1, "B")

        ss.insert_col(1)

        assert ss.get_value(0, 0) == "A"
        assert ss.get_value(0, 1) == ""
        assert ss.get_value(0, 2) == "B"

    def test_copy_cell(self):
        """Test copying a cell."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "Test")
        ss.copy_cell(0, 0, 1, 1)
        assert ss.get_value(1, 1) == "Test"

    def test_clear(self):
        """Test clearing spreadsheet."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "A")
        ss.set_cell(1, 1, "B")
        ss.set_col_width(0, 20)

        ss.clear()

        assert ss.get_value(0, 0) == ""
        assert ss.get_value(1, 1) == ""
        assert ss.get_col_width(0) == 10


class TestSpreadsheetSaveLoad:
    """Tests for save and load operations."""

    def test_save_and_load(self):
        """Test saving and loading spreadsheet."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "Hello")
        ss.set_cell(1, 0, "123")
        ss.set_cell(2, 0, "=A1")
        ss.set_col_width(0, 15)

        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            filename = f.name

        try:
            ss.save(filename)

            ss2 = Spreadsheet()
            ss2.load(filename)

            assert ss2.get_value(0, 0) == "Hello"
            assert ss2.get_value(1, 0) == 123
            assert ss2.get_col_width(0) == 15
            assert ss2.filename == filename
        finally:
            os.unlink(filename)

    def test_save_sets_filename(self):
        """Test that save sets filename."""
        ss = Spreadsheet()

        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            filename = f.name

        try:
            ss.save(filename)
            assert ss.filename == filename
        finally:
            os.unlink(filename)

    def test_load_clears_existing(self):
        """Test that load clears existing data."""
        ss = Spreadsheet()
        ss.set_cell(5, 5, "Existing")

        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump({"rows": 100, "cols": 26, "col_widths": {}, "cells": {}}, f)
            filename = f.name

        try:
            ss.load(filename)
            assert ss.get_value(5, 5) == ""
        finally:
            os.unlink(filename)


class TestSpreadsheetFormulas:
    """Tests for formula evaluation."""

    def test_simple_formula(self):
        """Test simple formula evaluation."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "10")
        ss.set_cell(0, 1, "20")
        ss.set_cell(0, 2, "=A1+B1")
        assert ss.get_value(0, 2) == 30

    def test_formula_with_reference(self):
        """Test formula with cell reference."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "100")
        ss.set_cell(1, 0, "=A1*2")
        assert ss.get_value(1, 0) == 200

    def test_nested_formula(self):
        """Test nested formula references."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "10")
        ss.set_cell(1, 0, "=A1+5")
        ss.set_cell(2, 0, "=A2*2")
        assert ss.get_value(2, 0) == 30

    def test_circular_reference_detection(self):
        """Test circular reference detection."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "=B1")
        ss.set_cell(0, 1, "=A1")
        assert "#CIRC!" in str(ss.get_value(0, 0))

    def test_formula_error(self):
        """Test formula error handling."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "=INVALID()")
        result = str(ss.get_value(0, 0))
        # Unknown function returns #NAME? error
        assert "#NAME?" in result or "ERR" in result or "error" in result.lower()

    def test_cache_invalidation(self):
        """Test cache is invalidated on cell change."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "10")
        ss.set_cell(1, 0, "=A1*2")

        # First computation
        assert ss.get_value(1, 0) == 20

        # Change source cell
        ss.set_cell(0, 0, "20")

        # Should recompute
        assert ss.get_value(1, 0) == 40

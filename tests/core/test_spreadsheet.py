"""Tests for spreadsheet data model."""

import os
import tempfile

import pytest

from lotus123 import Cell, Spreadsheet, col_to_index, index_to_col, make_cell_ref, parse_cell_ref


class TestCellReferenceConversions:
    def test_col_to_index(self):
        assert col_to_index("A") == 0
        assert col_to_index("B") == 1
        assert col_to_index("Z") == 25
        assert col_to_index("AA") == 26
        assert col_to_index("AB") == 27
        assert col_to_index("AZ") == 51
        assert col_to_index("BA") == 52

    def test_index_to_col(self):
        assert index_to_col(0) == "A"
        assert index_to_col(1) == "B"
        assert index_to_col(25) == "Z"
        assert index_to_col(26) == "AA"
        assert index_to_col(27) == "AB"
        assert index_to_col(51) == "AZ"
        assert index_to_col(52) == "BA"

    def test_parse_cell_ref(self):
        assert parse_cell_ref("A1") == (0, 0)
        assert parse_cell_ref("B2") == (1, 1)
        assert parse_cell_ref("Z10") == (9, 25)
        assert parse_cell_ref("AA100") == (99, 26)

    def test_parse_cell_ref_invalid(self):
        with pytest.raises(ValueError):
            parse_cell_ref("invalid")
        with pytest.raises(ValueError):
            parse_cell_ref("123")

    def test_make_cell_ref(self):
        assert make_cell_ref(0, 0) == "A1"
        assert make_cell_ref(1, 1) == "B2"
        assert make_cell_ref(9, 25) == "Z10"
        assert make_cell_ref(99, 26) == "AA100"


class TestCell:
    def test_empty_cell(self):
        cell = Cell()
        assert cell.raw_value == ""
        assert not cell.is_formula
        assert cell.formula == ""

    def test_value_cell(self):
        cell = Cell(raw_value="123")
        assert cell.raw_value == "123"
        assert not cell.is_formula

    def test_formula_cell_equals(self):
        cell = Cell(raw_value="=A1+B1")
        assert cell.is_formula
        assert cell.formula == "A1+B1"

    def test_formula_cell_plus(self):
        cell = Cell(raw_value="+A1+B1")
        assert cell.is_formula
        assert cell.formula == "+A1+B1"  # + prefix is kept for formula evaluation

    def test_to_dict_from_dict(self):
        cell = Cell(raw_value="=SUM(A1:A10)", format_code="F2")
        d = cell.to_dict()
        restored = Cell.from_dict(d)
        assert restored.raw_value == cell.raw_value
        assert restored.format_code == cell.format_code


class TestSpreadsheet:
    def test_empty_spreadsheet(self):
        ss = Spreadsheet()
        assert ss.get_value(0, 0) == ""
        assert ss.get_display_value(0, 0) == ""

    def test_set_and_get_number(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "42")
        assert ss.get_value(0, 0) == 42

    def test_set_and_get_float(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "3.14")
        assert ss.get_value(0, 0) == 3.14

    def test_set_and_get_string(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "Hello")
        assert ss.get_value(0, 0) == "Hello"

    def test_set_cell_by_ref(self):
        ss = Spreadsheet()
        ss.set_cell_by_ref("B5", "100")
        assert ss.get_value(4, 1) == 100

    def test_get_value_by_ref(self):
        ss = Spreadsheet()
        ss.set_cell(4, 1, "200")
        assert ss.get_value_by_ref("B5") == 200

    def test_simple_formula(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "10")
        ss.set_cell(0, 1, "20")
        ss.set_cell(0, 2, "=A1+B1")
        assert ss.get_value(0, 2) == 30

    def test_circular_reference(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "=A1")
        assert "#CIRC!" in str(ss.get_value(0, 0))

    def test_get_range(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "1")
        ss.set_cell(0, 1, "2")
        ss.set_cell(1, 0, "3")
        ss.set_cell(1, 1, "4")
        result = ss.get_range("A1", "B2")
        assert result == [[1, 2], [3, 4]]

    def test_get_range_flat(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "1")
        ss.set_cell(0, 1, "2")
        ss.set_cell(1, 0, "3")
        ss.set_cell(1, 1, "4")
        result = ss.get_range_flat("A1", "B2")
        assert result == [1, 2, 3, 4]

    def test_col_width(self):
        ss = Spreadsheet()
        assert ss.get_col_width(0) == 10
        ss.set_col_width(0, 20)
        assert ss.get_col_width(0) == 20

    def test_col_width_bounds(self):
        ss = Spreadsheet()
        ss.set_col_width(0, 1)
        assert ss.get_col_width(0) == 3
        ss.set_col_width(0, 100)
        assert ss.get_col_width(0) == 50

    def test_display_value_int(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "42")
        assert ss.get_display_value(0, 0) == "42"

    def test_display_value_float(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "3.14159")
        assert ss.get_display_value(0, 0) == "3.14"

    def test_display_value_whole_float(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "5.0")
        assert ss.get_display_value(0, 0) == "5"

    def test_clear(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "test")
        ss.set_col_width(0, 15)
        ss.clear()
        assert ss.get_value(0, 0) == ""
        assert ss.get_col_width(0) == 10

    def test_delete_row(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "A")
        ss.set_cell(1, 0, "B")
        ss.set_cell(2, 0, "C")
        ss.delete_row(1)
        assert ss.get_value(0, 0) == "A"
        assert ss.get_value(1, 0) == "C"

    def test_insert_row(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "A")
        ss.set_cell(1, 0, "B")
        ss.insert_row(1)
        assert ss.get_value(0, 0) == "A"
        assert ss.get_value(1, 0) == ""
        assert ss.get_value(2, 0) == "B"

    def test_delete_col(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "A")
        ss.set_cell(0, 1, "B")
        ss.set_cell(0, 2, "C")
        ss.delete_col(1)
        assert ss.get_value(0, 0) == "A"
        assert ss.get_value(0, 1) == "C"

    def test_insert_col(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "A")
        ss.set_cell(0, 1, "B")
        ss.insert_col(1)
        assert ss.get_value(0, 0) == "A"
        assert ss.get_value(0, 1) == ""
        assert ss.get_value(0, 2) == "B"

    def test_copy_cell(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "Test")
        ss.copy_cell(0, 0, 1, 1)
        assert ss.get_value(1, 1) == "Test"


class TestSpreadsheetSaveLoad:
    def test_save_and_load(self):
        ss = Spreadsheet()
        ss.set_cell(0, 0, "100")
        ss.set_cell(0, 1, "=A1*2")
        ss.set_col_width(0, 15)

        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
            temp_path = f.name

        try:
            ss.save(temp_path)

            ss2 = Spreadsheet()
            ss2.load(temp_path)

            assert ss2.get_value(0, 0) == 100
            assert ss2.get_value(0, 1) == 200
            assert ss2.get_col_width(0) == 15
        finally:
            os.unlink(temp_path)

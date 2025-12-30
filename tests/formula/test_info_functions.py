"""Tests for information functions."""

import pytest

from lotus123.formula.functions.info import (
    INFO_FUNCTIONS,
    fn_areas,
    fn_cell,
    fn_cellpointer,
    fn_error_type,
    fn_info,
    fn_isformula,
    fn_n,
    fn_sheet,
    fn_sheets,
    fn_type,
    fn_version,
)


class TestType:
    """Tests for TYPE function."""

    def test_type_number(self):
        """Test TYPE returns 1 for numbers."""
        assert fn_type(42) == 1
        assert fn_type(3.14) == 1

    def test_type_text(self):
        """Test TYPE returns 2 for text."""
        assert fn_type("hello") == 2
        assert fn_type("") == 2

    def test_type_boolean(self):
        """Test TYPE returns 4 for boolean."""
        assert fn_type(True) == 4
        assert fn_type(False) == 4

    def test_type_error(self):
        """Test TYPE returns 16 for error."""
        assert fn_type("#DIV/0!") == 16
        assert fn_type("#VALUE!") == 16

    def test_type_array(self):
        """Test TYPE returns 64 for array."""
        assert fn_type([1, 2, 3]) == 64


class TestCell:
    """Tests for CELL function."""

    def test_cell_address(self):
        """Test CELL with address info type."""
        result = fn_cell("address")
        assert result == "$A$1"

    def test_cell_col(self):
        """Test CELL with col info type."""
        result = fn_cell("col")
        assert result == 1

    def test_cell_row(self):
        """Test CELL with row info type."""
        result = fn_cell("row")
        assert result == 1

    def test_cell_contents(self):
        """Test CELL with contents info type."""
        result = fn_cell("contents", "test")
        assert result == "test"

    def test_cell_contents_empty(self):
        """Test CELL contents with no reference."""
        result = fn_cell("contents")
        assert result == ""

    def test_cell_type_blank(self):
        """Test CELL type for blank."""
        result = fn_cell("type", None)
        assert result == "b"

    def test_cell_type_value(self):
        """Test CELL type for value."""
        result = fn_cell("type", 42)
        assert result == "v"

    def test_cell_type_label(self):
        """Test CELL type for label."""
        result = fn_cell("type", "text")
        assert result == "l"

    def test_cell_width(self):
        """Test CELL with width info type."""
        result = fn_cell("width")
        assert result == 9

    def test_cell_format(self):
        """Test CELL with format info type."""
        result = fn_cell("format")
        assert result == "G"

    def test_cell_protect(self):
        """Test CELL with protect info type."""
        result = fn_cell("protect")
        assert result == 0

    def test_cell_prefix(self):
        """Test CELL with prefix info type."""
        result = fn_cell("prefix")
        assert result == "'"

    def test_cell_unknown(self):
        """Test CELL with unknown info type."""
        result = fn_cell("unknown")
        assert result == ""

    def test_cell_quoted_type(self):
        """Test CELL strips quotes from type."""
        result = fn_cell('"address"')
        assert result == "$A$1"


class TestCellPointer:
    """Tests for CELLPOINTER function."""

    def test_cellpointer_default(self):
        """Test CELLPOINTER default attribute."""
        result = fn_cellpointer()
        assert result == ""  # Default is contents

    def test_cellpointer_address(self):
        """Test CELLPOINTER with address."""
        result = fn_cellpointer("address")
        assert result == "$A$1"


class TestInfo:
    """Tests for INFO function."""

    def test_info_directory(self):
        """Test INFO with directory type."""
        result = fn_info("directory")
        assert isinstance(result, str)
        assert len(result) > 0

    def test_info_numfile(self):
        """Test INFO with numfile type."""
        result = fn_info("numfile")
        assert result == 1

    def test_info_origin(self):
        """Test INFO with origin type."""
        result = fn_info("origin")
        assert "$A" in result

    def test_info_osversion(self):
        """Test INFO with osversion type."""
        result = fn_info("osversion")
        assert isinstance(result, str)

    def test_info_recalc(self):
        """Test INFO with recalc type."""
        result = fn_info("recalc")
        assert result == "Automatic"

    def test_info_release(self):
        """Test INFO with release type."""
        result = fn_info("release")
        assert result == "1.0"

    def test_info_system(self):
        """Test INFO with system type."""
        result = fn_info("system")
        assert isinstance(result, str)

    def test_info_totmem(self):
        """Test INFO with totmem type."""
        result = fn_info("totmem")
        assert result > 0

    def test_info_usedmem(self):
        """Test INFO with usedmem type."""
        result = fn_info("usedmem")
        assert result > 0

    def test_info_unknown(self):
        """Test INFO with unknown type."""
        result = fn_info("unknown")
        assert result == ""


class TestErrorType:
    """Tests for ERROR.TYPE function."""

    def test_error_type_null(self):
        """Test ERROR.TYPE for #NULL!"""
        assert fn_error_type("#NULL!") == 1

    def test_error_type_div0(self):
        """Test ERROR.TYPE for #DIV/0!"""
        assert fn_error_type("#DIV/0!") == 2

    def test_error_type_value(self):
        """Test ERROR.TYPE for #VALUE!"""
        assert fn_error_type("#VALUE!") == 3

    def test_error_type_ref(self):
        """Test ERROR.TYPE for #REF!"""
        assert fn_error_type("#REF!") == 4

    def test_error_type_name(self):
        """Test ERROR.TYPE for #NAME?"""
        assert fn_error_type("#NAME?") == 5

    def test_error_type_num(self):
        """Test ERROR.TYPE for #NUM!"""
        assert fn_error_type("#NUM!") == 6

    def test_error_type_na(self):
        """Test ERROR.TYPE for #N/A"""
        assert fn_error_type("#N/A") == 7

    def test_error_type_circ(self):
        """Test ERROR.TYPE for #CIRC!"""
        assert fn_error_type("#CIRC!") == 8

    def test_error_type_err(self):
        """Test ERROR.TYPE for #ERR!"""
        assert fn_error_type("#ERR!") == 3

    def test_error_type_not_error(self):
        """Test ERROR.TYPE for non-error returns 0."""
        assert fn_error_type("hello") == 0
        assert fn_error_type(42) == 0


class TestSheet:
    """Tests for SHEET and SHEETS functions."""

    def test_sheet(self):
        """Test SHEET returns 1."""
        assert fn_sheet() == 1
        assert fn_sheet("anything") == 1

    def test_sheets(self):
        """Test SHEETS returns 1."""
        assert fn_sheets() == 1
        assert fn_sheets("anything") == 1


class TestAreas:
    """Tests for AREAS function."""

    def test_areas(self):
        """Test AREAS returns 1."""
        assert fn_areas("A1:B10") == 1


class TestIsFormula:
    """Tests for ISFORMULA function."""

    def test_isformula_equals(self):
        """Test ISFORMULA with = prefix."""
        assert fn_isformula("=A1+B1") is True

    def test_isformula_plus(self):
        """Test ISFORMULA with + prefix."""
        assert fn_isformula("+A1") is True

    def test_isformula_not_formula(self):
        """Test ISFORMULA with non-formula."""
        assert fn_isformula("hello") is False
        assert fn_isformula(42) is False


class TestN:
    """Tests for N function."""

    def test_n_number(self):
        """Test N with number."""
        assert fn_n(42) == 42.0
        assert fn_n(3.14) == 3.14

    def test_n_boolean(self):
        """Test N with boolean."""
        assert fn_n(True) == 1.0
        assert fn_n(False) == 0.0

    def test_n_text(self):
        """Test N with text returns 0."""
        assert fn_n("hello") == 0.0

    def test_n_none(self):
        """Test N with None returns 0."""
        assert fn_n(None) == 0.0


class TestVersion:
    """Tests for VERSION function."""

    def test_version(self):
        """Test VERSION returns version string."""
        result = fn_version()
        assert "Lotus" in result or "lotus" in result.lower()


class TestFunctionRegistry:
    """Test the function registry."""

    def test_all_functions_registered(self):
        """Test that all functions are in the registry."""
        assert "TYPE" in INFO_FUNCTIONS
        assert "CELL" in INFO_FUNCTIONS
        assert "CELLPOINTER" in INFO_FUNCTIONS
        assert "INFO" in INFO_FUNCTIONS
        assert "VERSION" in INFO_FUNCTIONS
        assert "ERROR.TYPE" in INFO_FUNCTIONS
        assert "SHEET" in INFO_FUNCTIONS
        assert "SHEETS" in INFO_FUNCTIONS
        assert "AREAS" in INFO_FUNCTIONS
        assert "ISFORMULA" in INFO_FUNCTIONS
        assert "N" in INFO_FUNCTIONS

    def test_functions_callable(self):
        """Test that all registered functions are callable."""
        for name, func in INFO_FUNCTIONS.items():
            assert callable(func), f"{name} is not callable"

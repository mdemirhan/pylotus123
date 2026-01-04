"""Tests for XLSX import/export functionality."""

import tempfile
from pathlib import Path

import pytest

from lotus123.core.spreadsheet import Spreadsheet
from lotus123.io.xlsx import XlsxReader, XlsxWriter, XlsxImportWarnings, get_xlsx_sheet_names


class TestXlsxImportWarnings:
    """Tests for XlsxImportWarnings dataclass."""

    def test_no_warnings(self):
        """Test has_warnings returns False when empty."""
        warnings = XlsxImportWarnings()
        assert not warnings.has_warnings()
        assert warnings.to_message() == ""

    def test_merged_cells_warning(self):
        """Test merged cells warning."""
        warnings = XlsxImportWarnings(merged_cells=["A1:B2", "C3:D4", "E5:F6", "G7:H8", "I9:J10"])
        assert warnings.has_warnings()
        assert "5 merged" in warnings.to_message()

    def test_unsupported_functions_warning(self):
        """Test unsupported functions warning."""
        warnings = XlsxImportWarnings(
            unsupported_formulas=[("A1", "=XLOOKUP(...)"), ("B2", "=LAMBDA(...)")]
        )
        assert warnings.has_warnings()
        assert "formulas have unsupported" in warnings.to_message()

    def test_multiple_warnings(self):
        """Test multiple warnings combined."""
        warnings = XlsxImportWarnings(
            merged_cells=["A1:B2", "C3:D4", "E5:F6"],
            conditional_formatting=True,
            data_validations=True,
        )
        assert warnings.has_warnings()
        msg = warnings.to_message()
        assert "3 merged" in msg
        assert "conditional" in msg.lower()
        assert "validation" in msg.lower()


class TestXlsxWriter:
    """Tests for XlsxWriter class."""

    def test_export_basic_values(self):
        """Test exporting basic cell values."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "Hello")
        ss.set_cell(0, 1, "123")
        ss.set_cell(1, 0, "45.67")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            writer = XlsxWriter(ss)
            writer.save(filepath)

            # Verify file exists and can be read
            assert Path(filepath).exists()
            assert Path(filepath).stat().st_size > 0
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_export_formulas(self):
        """Test exporting formulas."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "10")
        ss.set_cell(0, 1, "20")
        ss.set_cell(0, 2, "=SUM(A1:B1)")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            writer = XlsxWriter(ss)
            writer.save(filepath)
            assert Path(filepath).exists()
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_export_column_widths(self):
        """Test exporting column widths."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "Test")
        ss.set_col_width(0, 20)
        ss.set_col_width(2, 15)

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            writer = XlsxWriter(ss)
            writer.save(filepath)
            assert Path(filepath).exists()
        finally:
            Path(filepath).unlink(missing_ok=True)


class TestXlsxReader:
    """Tests for XlsxReader class."""

    def test_import_basic_values(self):
        """Test importing basic cell values."""
        # Create a test file first
        ss1 = Spreadsheet()
        ss1.set_cell(0, 0, "Hello")
        ss1.set_cell(0, 1, "123")
        ss1.set_cell(1, 0, "45.67")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            # Now import it
            ss2 = Spreadsheet()
            reader = XlsxReader(ss2)
            reader.load(filepath)

            assert ss2.get_cell(0, 0).raw_value == "Hello"
            assert ss2.get_cell(0, 1).raw_value == "123"
            assert ss2.get_cell(1, 0).raw_value == "45.67"
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_get_sheet_names(self):
        """Test getting sheet names from workbook."""
        ss = Spreadsheet()
        ss.set_cell(0, 0, "Test")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss).save(filepath)
            names = get_xlsx_sheet_names(filepath)
            assert len(names) >= 1
            assert names[0] == "Sheet1"  # Default sheet name from openpyxl
        finally:
            Path(filepath).unlink(missing_ok=True)


class TestXlsxRoundTrip:
    """Round-trip tests for XLSX export/import."""

    def test_roundtrip_basic_values(self):
        """Test basic values survive round-trip."""
        ss1 = Spreadsheet()
        ss1.set_cell(0, 0, "Hello World")
        ss1.set_cell(0, 1, "123.456")
        ss1.set_cell(1, 0, "-999")
        ss1.set_cell(5, 5, "Sparse cell")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            assert ss2.get_cell(0, 0).raw_value == "Hello World"
            assert ss2.get_cell(0, 1).raw_value == "123.456"
            assert ss2.get_cell(1, 0).raw_value == "-999"
            assert ss2.get_cell(5, 5).raw_value == "Sparse cell"
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_roundtrip_formulas(self):
        """Test formulas survive round-trip."""
        ss1 = Spreadsheet()
        ss1.set_cell(0, 0, "10")
        ss1.set_cell(0, 1, "20")
        ss1.set_cell(0, 2, "=SUM(A1:B1)")
        ss1.set_cell(1, 0, "=A1+B1")
        ss1.set_cell(1, 1, "=$A$1*2")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            assert ss2.get_cell(0, 2).raw_value == "=SUM(A1:B1)"
            assert ss2.get_cell(1, 0).raw_value == "=A1+B1"
            assert ss2.get_cell(1, 1).raw_value == "=$A$1*2"
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_roundtrip_function_translation(self):
        """Test function names are correctly translated in round-trip."""
        ss1 = Spreadsheet()
        ss1.set_cell(0, 0, "10")
        ss1.set_cell(0, 1, "20")
        ss1.set_cell(0, 2, "30")
        ss1.set_cell(1, 0, "=AVG(A1:C1)")
        ss1.set_cell(1, 1, "=STD(A1:C1)")
        ss1.set_cell(1, 2, "=COLS(A1:C1)")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            # Verify formulas are correctly translated back
            assert ss2.get_cell(1, 0).raw_value == "=AVG(A1:C1)"
            assert ss2.get_cell(1, 1).raw_value == "=STD(A1:C1)"
            assert ss2.get_cell(1, 2).raw_value == "=COLS(A1:C1)"
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_roundtrip_alignment_prefixes(self):
        """Test alignment prefixes survive round-trip."""
        ss1 = Spreadsheet()
        ss1.set_cell(0, 0, "'Left aligned")
        ss1.set_cell(0, 1, '"Right aligned')
        ss1.set_cell(0, 2, "^Center aligned")
        ss1.set_cell(0, 3, "No alignment")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            assert ss2.get_cell(0, 0).raw_value == "'Left aligned"
            assert ss2.get_cell(0, 1).raw_value == '"Right aligned'
            assert ss2.get_cell(0, 2).raw_value == "^Center aligned"
            assert ss2.get_cell(0, 3).raw_value == "No alignment"
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_roundtrip_format_codes(self):
        """Test format codes survive round-trip."""
        ss1 = Spreadsheet()
        # Set values with different formats
        ss1.set_cell(0, 0, "123.456")
        ss1.get_cell(0, 0).format_code = "F2"
        ss1.set_cell(0, 1, "1000")
        ss1.get_cell(0, 1).format_code = "C2"
        ss1.set_cell(0, 2, "0.75")
        ss1.get_cell(0, 2).format_code = "P0"
        ss1.set_cell(0, 3, "45000")  # Date serial
        ss1.get_cell(0, 3).format_code = "D1"

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            assert ss2.get_cell(0, 0).format_code == "F2"
            assert ss2.get_cell(0, 1).format_code == "C2"
            assert ss2.get_cell(0, 2).format_code == "P0"
            assert ss2.get_cell(0, 3).format_code == "D1"
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_roundtrip_column_widths(self):
        """Test column widths survive round-trip."""
        ss1 = Spreadsheet()
        ss1.set_cell(0, 0, "Test")
        ss1.set_col_width(0, 20)
        ss1.set_col_width(2, 15)
        ss1.set_col_width(5, 25)

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            assert ss2.get_col_width(0) == 20
            assert ss2.get_col_width(2) == 15
            assert ss2.get_col_width(5) == 25
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_roundtrip_row_heights(self):
        """Test row heights survive round-trip."""
        ss1 = Spreadsheet()
        ss1.set_cell(0, 0, "Test")
        ss1.set_row_height(0, 2)
        ss1.set_row_height(2, 3)

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            assert ss2.get_row_height(0) == 2
            assert ss2.get_row_height(2) == 3
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_roundtrip_named_ranges(self):
        """Test named ranges survive round-trip."""
        ss1 = Spreadsheet()
        ss1.set_cell(0, 0, "100")
        ss1.set_cell(0, 1, "200")
        ss1.named_ranges.add_from_string("RATE", "A1")
        ss1.named_ranges.add_from_string("DATA", "A1:B1")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            rate = ss2.named_ranges.get("RATE")
            data = ss2.named_ranges.get("DATA")
            assert rate is not None
            assert data is not None
            # Check the references are correctly restored
            assert str(rate.reference) == "A1"
            assert str(data.reference) == "A1:B1"
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_roundtrip_frozen_panes(self):
        """Test frozen panes survive round-trip."""
        ss1 = Spreadsheet()
        ss1.set_cell(0, 0, "Test")
        ss1.frozen_rows = 2
        ss1.frozen_cols = 1

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            assert ss2.frozen_rows == 2
            assert ss2.frozen_cols == 1
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_roundtrip_empty_cells_sparse(self):
        """Test sparse storage is preserved (empty cells not stored)."""
        ss1 = Spreadsheet()
        ss1.set_cell(0, 0, "A")
        ss1.set_cell(100, 100, "Z")  # Far away sparse cell

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            assert ss2.get_cell(0, 0).raw_value == "A"
            assert ss2.get_cell(100, 100).raw_value == "Z"
            # Intermediate cells should not exist
            assert ss2.get_cell_if_exists(50, 50) is None
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_roundtrip_special_characters(self):
        """Test special characters in text survive round-trip."""
        ss1 = Spreadsheet()
        ss1.set_cell(0, 0, "Hello, World!")
        ss1.set_cell(0, 1, '"Quoted"')
        ss1.set_cell(0, 2, "Line1\nLine2")
        ss1.set_cell(0, 3, "Tab\there")
        ss1.set_cell(0, 4, "Unicode: café ñ 日本語")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            assert ss2.get_cell(0, 0).raw_value == "Hello, World!"
            assert ss2.get_cell(0, 1).raw_value == '"Quoted"'
            assert ss2.get_cell(0, 2).raw_value == "Line1\nLine2"
            assert ss2.get_cell(0, 3).raw_value == "Tab\there"
            assert ss2.get_cell(0, 4).raw_value == "Unicode: café ñ 日本語"
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_roundtrip_numeric_precision(self):
        """Test numeric precision is preserved in round-trip."""
        ss1 = Spreadsheet()
        ss1.set_cell(0, 0, "3.141592653589793")
        ss1.set_cell(0, 1, "0.0000001")
        ss1.set_cell(0, 2, "99999999999999")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            # Values should be preserved
            val1 = float(ss2.get_cell(0, 0).raw_value)
            assert abs(val1 - 3.141592653589793) < 1e-10
        finally:
            Path(filepath).unlink(missing_ok=True)

    def test_roundtrip_text_that_looks_like_formula(self):
        """Test that text labels starting with = or @ are preserved as text."""
        ss1 = Spreadsheet()
        # These are TEXT labels (with ' prefix), not formulas
        ss1.set_cell(0, 0, "'=SUM(A1:A10)")  # Text that looks like formula
        ss1.set_cell(0, 1, "'@Function")  # Text starting with @
        ss1.set_cell(0, 2, "'+Positive")  # Text starting with +
        ss1.set_cell(0, 3, "'-Negative")  # Text starting with -

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            filepath = f.name

        try:
            XlsxWriter(ss1).save(filepath)

            ss2 = Spreadsheet()
            XlsxReader(ss2).load(filepath)

            # Should be imported as text labels (with ' prefix)
            assert ss2.get_cell(0, 0).raw_value == "'=SUM(A1:A10)"
            assert ss2.get_cell(0, 1).raw_value == "'@Function"
            assert ss2.get_cell(0, 2).raw_value == "'+Positive"
            assert ss2.get_cell(0, 3).raw_value == "'-Negative"
        finally:
            Path(filepath).unlink(missing_ok=True)

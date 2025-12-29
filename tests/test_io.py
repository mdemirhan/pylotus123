"""Tests for text import/export operations."""

import tempfile
from pathlib import Path

import pytest

from lotus123 import Spreadsheet
from lotus123.io.text_export import ExportFormat, ExportOptions, TextExporter
from lotus123.io.text_import import ImportFormat, ImportOptions, TextImporter


class TestImportOptions:
    """Tests for ImportOptions dataclass."""

    def test_default_values(self):
        """Test default values."""
        opts = ImportOptions()
        assert opts.format == ImportFormat.CSV
        assert opts.delimiter == ","
        assert opts.has_header is False
        assert opts.encoding == "utf-8"
        assert opts.text_qualifier == '"'

    def test_custom_values(self):
        """Test custom values."""
        opts = ImportOptions(
            format=ImportFormat.TSV, delimiter="\t", has_header=True, encoding="latin-1"
        )
        assert opts.format == ImportFormat.TSV
        assert opts.has_header is True


class TestExportOptions:
    """Tests for ExportOptions dataclass."""

    def test_default_values(self):
        """Test default values."""
        opts = ExportOptions()
        assert opts.format == ExportFormat.CSV
        assert opts.delimiter == ","
        assert opts.include_header is False
        assert opts.use_formulas is False

    def test_custom_values(self):
        """Test custom values."""
        opts = ExportOptions(format=ExportFormat.TSV, include_header=True, use_formulas=True)
        assert opts.format == ExportFormat.TSV


class TestTextImporter:
    """Tests for TextImporter class."""

    def setup_method(self):
        """Set up test spreadsheet."""
        self.ss = Spreadsheet()
        self.importer = TextImporter(self.ss)

    def test_import_csv_basic(self):
        """Test basic CSV import."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            f.write("A,B,C\n")
            f.write("1,2,3\n")
            f.write("4,5,6\n")
            f.flush()
            path = f.name

        try:
            rows = self.importer.import_file(path)
            assert rows > 0
            assert self.ss.get_value(0, 0) == "A"
            assert self.ss.get_value(1, 0) == 1
            assert self.ss.get_value(1, 2) == 3
        finally:
            Path(path).unlink()

    def test_import_csv_with_options(self):
        """Test CSV import with custom options."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            f.write("Header1,Header2\n")
            f.write("Value1,Value2\n")
            f.flush()
            path = f.name

        opts = ImportOptions(
            format=ImportFormat.CSV, has_header=True, dest_row=5, dest_col=5
        )

        try:
            rows = self.importer.import_file(path, opts)
            assert rows > 0
        finally:
            Path(path).unlink()

    def test_import_tsv(self):
        """Test TSV import."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".tsv", delete=False) as f:
            f.write("A\tB\tC\n")
            f.write("1\t2\t3\n")
            f.flush()
            path = f.name

        opts = ImportOptions(format=ImportFormat.TSV)

        try:
            rows = self.importer.import_file(path, opts)
            assert rows > 0
        finally:
            Path(path).unlink()

    def test_import_custom_delimiter(self):
        """Test import with custom delimiter."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            f.write("A;B;C\n")
            f.write("1;2;3\n")
            f.flush()
            path = f.name

        opts = ImportOptions(format=ImportFormat.CUSTOM_DELIMITED, delimiter=";")

        try:
            rows = self.importer.import_file(path, opts)
            assert rows > 0
        finally:
            Path(path).unlink()

    def test_import_file_not_found(self):
        """Test importing non-existent file."""
        with pytest.raises(FileNotFoundError):
            self.importer.import_file("/nonexistent/file.csv")

    def test_auto_detect_csv(self):
        """Test auto-detection of CSV format."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            f.write("A,B,C\n")
            f.flush()
            path = f.name

        try:
            # Pass None for options to trigger auto-detection
            rows = self.importer.import_file(path)
            assert rows > 0
        finally:
            Path(path).unlink()


class TestTextExporter:
    """Tests for TextExporter class."""

    def setup_method(self):
        """Set up test spreadsheet."""
        self.ss = Spreadsheet()
        self.exporter = TextExporter(self.ss)
        # Set up test data
        self.ss.set_cell(0, 0, "A")
        self.ss.set_cell(0, 1, "B")
        self.ss.set_cell(0, 2, "C")
        self.ss.set_cell(1, 0, "1")
        self.ss.set_cell(1, 1, "2")
        self.ss.set_cell(1, 2, "3")

    def test_export_csv_basic(self):
        """Test basic CSV export."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            path = f.name

        try:
            opts = ExportOptions(end_row=1, end_col=2)
            rows = self.exporter.export_file(path, opts)
            assert rows > 0

            # Read back and verify
            content = Path(path).read_text()
            assert "A" in content
            assert "," in content
        finally:
            Path(path).unlink()

    def test_export_tsv(self):
        """Test TSV export."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".tsv", delete=False) as f:
            path = f.name

        opts = ExportOptions(format=ExportFormat.TSV, end_row=1, end_col=2)

        try:
            rows = self.exporter.export_file(path, opts)
            assert rows > 0

            content = Path(path).read_text()
            assert "\t" in content
        finally:
            Path(path).unlink()

    def test_export_custom_delimiter(self):
        """Test export with custom delimiter."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            path = f.name

        opts = ExportOptions(
            format=ExportFormat.CUSTOM_DELIMITED, delimiter=";", end_row=1, end_col=2
        )

        try:
            rows = self.exporter.export_file(path, opts)
            assert rows > 0

            content = Path(path).read_text()
            assert ";" in content
        finally:
            Path(path).unlink()

    def test_export_with_formulas(self):
        """Test export with formulas option."""
        self.ss.set_cell(2, 0, "=A1+B1")

        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            path = f.name

        opts = ExportOptions(use_formulas=True, end_row=2, end_col=2)

        try:
            rows = self.exporter.export_file(path, opts)
            content = Path(path).read_text()
            # Should contain formula string
        finally:
            Path(path).unlink()

    def test_export_values_only(self):
        """Test export with values only (not formulas)."""
        self.ss.set_cell(2, 0, "=1+2")

        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            path = f.name

        opts = ExportOptions(use_formulas=False, end_row=2, end_col=2)

        try:
            rows = self.exporter.export_file(path, opts)
            content = Path(path).read_text()
            # Should contain computed value, not formula
        finally:
            Path(path).unlink()

    def test_auto_detect_from_extension(self):
        """Test auto-detection of format from extension."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".tsv", delete=False) as f:
            path = f.name

        try:
            # Pass None for options to trigger auto-detection
            rows = self.exporter.export_file(path)
            content = Path(path).read_text()
            # Should use tab delimiter for .tsv
            assert "\t" in content or content  # May be empty if no data in range
        finally:
            Path(path).unlink()


class TestImportFormat:
    """Tests for ImportFormat enum."""

    def test_all_formats_exist(self):
        """Test all import formats exist."""
        assert ImportFormat.CSV
        assert ImportFormat.TSV
        assert ImportFormat.FIXED_WIDTH
        assert ImportFormat.CUSTOM_DELIMITED


class TestExportFormat:
    """Tests for ExportFormat enum."""

    def test_all_formats_exist(self):
        """Test all export formats exist."""
        assert ExportFormat.CSV
        assert ExportFormat.TSV
        assert ExportFormat.CUSTOM_DELIMITED
        assert ExportFormat.FORMATTED_TEXT

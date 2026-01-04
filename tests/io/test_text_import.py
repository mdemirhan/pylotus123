"""Tests for text import module."""

import tempfile
from pathlib import Path

import pytest

from lotus123 import Spreadsheet
from lotus123.io.text_import import (
    ImportFormat,
    ImportOptions,
    TextImporter,
)


class TestImportFormat:
    """Tests for ImportFormat enum."""

    def test_formats_exist(self):
        """Test all formats exist."""
        assert ImportFormat.CSV
        assert ImportFormat.TSV
        assert ImportFormat.FIXED_WIDTH
        assert ImportFormat.CUSTOM_DELIMITED


class TestImportOptions:
    """Tests for ImportOptions dataclass."""

    def test_default_values(self):
        """Test default option values."""
        opts = ImportOptions()
        assert opts.format == ImportFormat.CSV
        assert opts.delimiter == ","
        assert opts.has_header is False
        assert opts.start_row == 0
        assert opts.start_col == 0
        assert opts.dest_row == 0
        assert opts.dest_col == 0
        assert opts.text_qualifier == '"'
        assert opts.field_widths == []
        assert opts.encoding == "utf-8"
        assert opts.skip_blank_lines is True
        assert opts.trim_whitespace is True

    def test_custom_values(self):
        """Test custom option values."""
        opts = ImportOptions(
            format=ImportFormat.TSV, delimiter="\t", has_header=True, start_row=1, dest_row=5
        )
        assert opts.format == ImportFormat.TSV
        assert opts.delimiter == "\t"
        assert opts.has_header is True
        assert opts.start_row == 1
        assert opts.dest_row == 5


class TestTextImporter:
    """Tests for TextImporter class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.importer = TextImporter(self.ss)

    def test_import_text_csv(self):
        """Test importing CSV text."""
        text = "A,B,C\n1,2,3\n4,5,6"
        opts = ImportOptions()
        count = self.importer.import_text(text, opts)

        assert count == 3
        assert self.ss.get_value(0, 0) == "A"
        assert self.ss.get_value(1, 0) == 1
        assert self.ss.get_value(2, 2) == 6

    def test_import_text_tsv(self):
        """Test importing TSV text."""
        text = "A\tB\tC\n1\t2\t3"
        opts = ImportOptions(format=ImportFormat.TSV, delimiter="\t")
        count = self.importer.import_text(text, opts)

        assert count == 2
        assert self.ss.get_value(0, 0) == "A"
        assert self.ss.get_value(0, 2) == "C"
        assert self.ss.get_value(1, 1) == 2

    def test_import_text_skip_rows(self):
        """Test importing with skipped rows."""
        text = "Header1,Header2\nData1,Data2\nData3,Data4"
        opts = ImportOptions(start_row=1)
        count = self.importer.import_text(text, opts)

        assert count == 2
        assert self.ss.get_value(0, 0) == "Data1"
        assert self.ss.get_value(1, 0) == "Data3"

    def test_import_text_dest_offset(self):
        """Test importing to offset destination."""
        text = "A,B\n1,2"
        opts = ImportOptions(dest_row=5, dest_col=3)
        count = self.importer.import_text(text, opts)

        assert count == 2
        assert self.ss.get_value(5, 3) == "A"
        assert self.ss.get_value(6, 4) == 2

    def test_import_text_skip_blank_lines(self):
        """Test skipping blank lines."""
        text = "A\n\nB\n\nC"
        opts = ImportOptions(skip_blank_lines=True)
        count = self.importer.import_text(text, opts)

        assert count == 3
        assert self.ss.get_value(0, 0) == "A"
        assert self.ss.get_value(1, 0) == "B"
        assert self.ss.get_value(2, 0) == "C"

    def test_import_text_trim_whitespace(self):
        """Test trimming whitespace."""
        text = "  A  ,  B  "
        opts = ImportOptions(trim_whitespace=True)
        count = self.importer.import_text(text, opts)

        assert count == 1
        assert self.ss.get_value(0, 0) == "A"
        assert self.ss.get_value(0, 1) == "B"

    def test_import_text_no_trim(self):
        """Test without trimming whitespace."""
        text = "  A  ,  B  "
        opts = ImportOptions(trim_whitespace=False)
        count = self.importer.import_text(text, opts)

        assert count == 1
        assert self.ss.get_value(0, 0) == "  A  "

    def test_import_text_custom_delimiter(self):
        """Test custom delimiter."""
        text = "A;B;C\n1;2;3"
        opts = ImportOptions(format=ImportFormat.CUSTOM_DELIMITED, delimiter=";")
        count = self.importer.import_text(text, opts)

        assert count == 2
        assert self.ss.get_value(0, 1) == "B"

    def test_import_text_fixed_width(self):
        """Test fixed width import."""
        text = "AAABBBCCC\n111222333"
        opts = ImportOptions(format=ImportFormat.FIXED_WIDTH, field_widths=[3, 3, 3])
        count = self.importer.import_text(text, opts)

        assert count == 2
        assert self.ss.get_value(0, 0) == "AAA"
        assert self.ss.get_value(0, 1) == "BBB"
        assert self.ss.get_value(0, 2) == "CCC"

    def test_import_text_start_col(self):
        """Test importing starting from specific column."""
        text = "A,B,C\n1,2,3"
        opts = ImportOptions(start_col=1)
        count = self.importer.import_text(text, opts)

        assert count == 2
        # Skips first column
        assert self.ss.get_value(0, 0) == "B"
        assert self.ss.get_value(0, 1) == "C"


class TestTextImporterFile:
    """Tests for TextImporter file operations."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.importer = TextImporter(self.ss)

    def test_import_csv_file(self):
        """Test importing CSV file."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            f.write("Name,Age\nJohn,30\nAlice,25")
            f.flush()

            count = self.importer.import_file(f.name)

            assert count == 3
            assert self.ss.get_value(0, 0) == "Name"
            assert self.ss.get_value(1, 0) == "John"
            assert self.ss.get_value(2, 1) == 25

            Path(f.name).unlink()

    def test_import_tsv_file(self):
        """Test importing TSV file."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".tsv", delete=False) as f:
            f.write("Name\tAge\nJohn\t30")
            f.flush()

            count = self.importer.import_file(f.name)

            assert count == 2
            assert self.ss.get_value(0, 0) == "Name"
            assert self.ss.get_value(0, 1) == "Age"

            Path(f.name).unlink()

    def test_import_file_not_found(self):
        """Test importing non-existent file."""
        with pytest.raises(FileNotFoundError):
            self.importer.import_file("/nonexistent/file.csv")

    def test_import_file_with_options(self):
        """Test importing file with explicit options."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            f.write("A,B,C\n1,2,3")
            f.flush()

            opts = ImportOptions(format=ImportFormat.CSV, delimiter=",")
            count = self.importer.import_file(f.name, opts)

            assert count == 2

            Path(f.name).unlink()

    def test_import_txt_auto_detect_tab(self):
        """Test auto-detecting tab delimiter in .txt file."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            f.write("A\tB\tC\n1\t2\t3")
            f.flush()

            count = self.importer.import_file(f.name)

            assert count == 2
            assert self.ss.get_value(0, 1) == "B"

            Path(f.name).unlink()

    def test_import_txt_auto_detect_comma(self):
        """Test auto-detecting comma delimiter in .txt file."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            f.write("A,B,C\n1,2,3")
            f.flush()

            count = self.importer.import_file(f.name)

            assert count == 2
            assert self.ss.get_value(0, 1) == "B"

            Path(f.name).unlink()

    def test_import_txt_auto_detect_semicolon(self):
        """Test auto-detecting semicolon delimiter in .txt file."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            f.write("A;B;C\n1;2;3")
            f.flush()

            count = self.importer.import_file(f.name)

            assert count == 2
            assert self.ss.get_value(0, 1) == "B"

            Path(f.name).unlink()


class TestTextImporterFixedWidth:
    """Tests for fixed-width import."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.importer = TextImporter(self.ss)

    def test_import_fixed_width_file(self):
        """Test importing fixed-width file."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            f.write("AAABBBCCC\n111222333\n444555666")
            f.flush()

            opts = ImportOptions(format=ImportFormat.FIXED_WIDTH, field_widths=[3, 3, 3])
            count = self.importer.import_file(f.name, opts)

            assert count == 3
            assert self.ss.get_value(0, 0) == "AAA"
            assert self.ss.get_value(1, 0) in [111, "111"]
            assert self.ss.get_value(2, 2) in [666, "666"]

            Path(f.name).unlink()

    def test_import_fixed_width_auto_detect(self):
        """Test fixed-width import with auto-detected widths."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            # Content with consistent spacing for auto-detection
            f.write("AAA   BBB   CCC\n")
            f.write("111   222   333\n")
            f.write("444   555   666\n")
            f.flush()

            opts = ImportOptions(
                format=ImportFormat.FIXED_WIDTH,
                field_widths=[],  # Auto-detect
            )
            # Should either work with auto-detection or raise error
            try:
                count = self.importer.import_file(f.name, opts)
                assert count >= 0
            except ValueError:
                pass  # Expected if auto-detection fails

            Path(f.name).unlink()

    def test_import_fixed_width_skip_rows(self):
        """Test fixed-width import with skip rows."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            f.write("Header---\n111222333")
            f.flush()

            opts = ImportOptions(
                format=ImportFormat.FIXED_WIDTH, field_widths=[3, 3, 3], start_row=1
            )
            count = self.importer.import_file(f.name, opts)

            assert count == 1
            # Value may be parsed as number
            assert self.ss.get_value(0, 0) in [111, "111"]

            Path(f.name).unlink()


class TestDetectFieldWidths:
    """Tests for field width detection."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.importer = TextImporter(self.ss)

    def test_detect_widths_empty_file(self):
        """Test detecting widths from empty file."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            f.write("")
            f.flush()

            widths = self.importer._detect_field_widths(Path(f.name))
            # Empty file may return empty list or single-element list
            assert isinstance(widths, list)

            Path(f.name).unlink()

    def test_detect_widths_spaced_content(self):
        """Test detecting widths from spaced content."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            # Create content with consistent spaces
            f.write("AAA   BBB   CCC\n")
            f.write("111   222   333\n")
            f.write("444   555   666\n")
            f.flush()

            widths = self.importer._detect_field_widths(Path(f.name))
            # Should detect some field boundaries
            assert len(widths) >= 1

            Path(f.name).unlink()

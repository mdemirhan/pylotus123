"""Tests for text export module."""

import tempfile
from pathlib import Path


from lotus123 import Spreadsheet
from lotus123.io.text_export import (
    ExportFormat,
    ExportOptions,
    TextExporter,
)


class TestExportFormat:
    """Tests for ExportFormat enum."""

    def test_formats_exist(self):
        """Test all formats exist."""
        assert ExportFormat.CSV
        assert ExportFormat.TSV
        assert ExportFormat.CUSTOM_DELIMITED
        assert ExportFormat.FORMATTED_TEXT


class TestExportOptions:
    """Tests for ExportOptions dataclass."""

    def test_default_values(self):
        """Test default option values."""
        opts = ExportOptions()
        assert opts.format == ExportFormat.CSV
        assert opts.delimiter == ","
        assert opts.include_header is False
        assert opts.start_row == 0
        assert opts.start_col == 0
        assert opts.end_row is None
        assert opts.end_col is None
        assert opts.text_qualifier == '"'
        assert opts.encoding == "utf-8"
        assert opts.use_formulas is False
        assert opts.line_ending == "\n"
        assert opts.quote_all is False

    def test_custom_values(self):
        """Test custom option values."""
        opts = ExportOptions(
            format=ExportFormat.TSV,
            delimiter="\t",
            include_header=True,
            start_row=1,
            end_row=10
        )
        assert opts.format == ExportFormat.TSV
        assert opts.include_header is True
        assert opts.start_row == 1
        assert opts.end_row == 10


class TestTextExporter:
    """Tests for TextExporter class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.exporter = TextExporter(self.ss)
        # Set up sample data
        self.ss.set_cell(0, 0, "Name")
        self.ss.set_cell(0, 1, "Age")
        self.ss.set_cell(1, 0, "John")
        self.ss.set_cell(1, 1, "30")
        self.ss.set_cell(2, 0, "Alice")
        self.ss.set_cell(2, 1, "25")


class TestExportToString:
    """Tests for export_to_string method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.exporter = TextExporter(self.ss)
        self.ss.set_cell(0, 0, "A")
        self.ss.set_cell(0, 1, "B")
        self.ss.set_cell(1, 0, "1")
        self.ss.set_cell(1, 1, "2")

    def test_export_csv_string(self):
        """Test exporting to CSV string."""
        result = self.exporter.export_to_string()
        assert "A" in result
        assert "B" in result
        assert "," in result

    def test_export_tsv_string(self):
        """Test exporting to TSV string."""
        opts = ExportOptions(format=ExportFormat.TSV, delimiter="\t")
        result = self.exporter.export_to_string(opts)
        assert "\t" in result

    def test_export_custom_delimiter(self):
        """Test exporting with custom delimiter."""
        opts = ExportOptions(delimiter=";")
        result = self.exporter.export_to_string(opts)
        assert ";" in result

    def test_export_quote_all(self):
        """Test exporting with all values quoted."""
        opts = ExportOptions(quote_all=True)
        result = self.exporter.export_to_string(opts)
        assert '"' in result

    def test_export_custom_line_ending(self):
        """Test exporting with custom line ending."""
        opts = ExportOptions(line_ending="\r\n")
        result = self.exporter.export_to_string(opts)
        assert "\r\n" in result

    def test_export_values_with_delimiter(self):
        """Test exporting values containing delimiter."""
        self.ss.set_cell(0, 0, "Hello,World")
        opts = ExportOptions()
        result = self.exporter.export_to_string(opts)
        # Should quote the value
        assert '"Hello,World"' in result

    def test_export_values_with_quotes(self):
        """Test exporting values containing quotes."""
        self.ss.set_cell(0, 0, 'Say "Hello"')
        opts = ExportOptions()
        result = self.exporter.export_to_string(opts)
        # Should quote the value
        assert '"' in result

    def test_export_empty_spreadsheet(self):
        """Test exporting empty spreadsheet."""
        ss = Spreadsheet()
        exporter = TextExporter(ss)
        result = exporter.export_to_string()
        assert result == ""


class TestExportFile:
    """Tests for export_file method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.exporter = TextExporter(self.ss)
        self.ss.set_cell(0, 0, "Name")
        self.ss.set_cell(0, 1, "Age")
        self.ss.set_cell(1, 0, "John")
        self.ss.set_cell(1, 1, "30")

    def test_export_csv_file(self):
        """Test exporting to CSV file."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            count = self.exporter.export_file(f.name)

            assert count == 2

            content = Path(f.name).read_text()
            assert "Name" in content
            assert "John" in content

            Path(f.name).unlink()

    def test_export_tsv_file(self):
        """Test exporting to TSV file."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".tsv", delete=False) as f:
            count = self.exporter.export_file(f.name)

            assert count == 2

            content = Path(f.name).read_text()
            assert "\t" in content

            Path(f.name).unlink()

    def test_export_txt_file(self):
        """Test exporting to TXT file (formatted)."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            count = self.exporter.export_file(f.name)

            assert count == 2

            content = Path(f.name).read_text()
            assert "Name" in content

            Path(f.name).unlink()

    def test_export_with_header(self):
        """Test exporting with column headers."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            opts = ExportOptions(include_header=True)
            count = self.exporter.export_file(f.name, opts)

            assert count == 3  # Header + 2 data rows

            content = Path(f.name).read_text()
            assert "A" in content  # Column A header
            assert "B" in content  # Column B header

            Path(f.name).unlink()

    def test_export_with_options(self):
        """Test exporting with explicit options."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            opts = ExportOptions(
                format=ExportFormat.CSV,
                delimiter=",",
                start_row=1,
                end_row=1
            )
            count = self.exporter.export_file(f.name, opts)

            assert count == 1  # Only row 1

            content = Path(f.name).read_text()
            assert "John" in content
            assert "Name" not in content  # Header row skipped

            Path(f.name).unlink()


class TestExportFormatted:
    """Tests for formatted text export."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.exporter = TextExporter(self.ss)
        self.ss.set_cell(0, 0, "Name")
        self.ss.set_cell(0, 1, "Age")
        self.ss.set_cell(1, 0, "John")
        self.ss.set_cell(1, 1, "30")

    def test_export_formatted(self):
        """Test formatted text export."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            opts = ExportOptions(format=ExportFormat.FORMATTED_TEXT)
            count = self.exporter.export_file(f.name, opts)

            assert count == 2

            content = Path(f.name).read_text()
            assert "Name" in content

            Path(f.name).unlink()

    def test_export_formatted_with_header(self):
        """Test formatted export with column headers."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            opts = ExportOptions(
                format=ExportFormat.FORMATTED_TEXT,
                include_header=True
            )
            count = self.exporter.export_file(f.name, opts)

            # Header row + separator + data rows
            assert count == 4

            content = Path(f.name).read_text()
            # Should have separator line
            assert "-" in content

            Path(f.name).unlink()


class TestExportRange:
    """Tests for export_range method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.exporter = TextExporter(self.ss)
        # Set up sample data
        for r in range(5):
            for c in range(5):
                self.ss.set_cell(r, c, f"R{r}C{c}")

    def test_export_specific_range(self):
        """Test exporting specific range."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            count = self.exporter.export_range(1, 1, 3, 3, f.name)

            assert count == 3  # Rows 1, 2, 3

            content = Path(f.name).read_text()
            assert "R1C1" in content
            assert "R3C3" in content
            assert "R0C0" not in content

            Path(f.name).unlink()

    def test_export_range_with_options(self):
        """Test exporting range with options."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False) as f:
            opts = ExportOptions(quote_all=True)
            count = self.exporter.export_range(0, 0, 1, 1, f.name, opts)

            assert count == 2

            content = Path(f.name).read_text()
            assert '"' in content

            Path(f.name).unlink()


class TestExportFormulas:
    """Tests for exporting with formulas."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.exporter = TextExporter(self.ss)
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "=A1*2")

    def test_export_values_default(self):
        """Test exporting values by default."""
        result = self.exporter.export_to_string()
        # Should export evaluated value, not formula
        lines = result.split("\n")
        values = lines[0].split(",")
        # The second value should be the calculated result, not the formula
        assert values[0] == "10"

    def test_export_formulas(self):
        """Test exporting formulas."""
        opts = ExportOptions(use_formulas=True)
        result = self.exporter.export_to_string(opts)
        # Should export formula
        assert "=A1*2" in result


class TestGetExportRange:
    """Tests for _get_export_range method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.exporter = TextExporter(self.ss)

    def test_empty_spreadsheet(self):
        """Test getting range from empty spreadsheet."""
        opts = ExportOptions()
        result = self.exporter._get_export_range(opts)
        assert result == (0, 0, 0, 0)

    def test_range_with_data(self):
        """Test getting range with data."""
        self.ss.set_cell(0, 0, "A")
        self.ss.set_cell(2, 3, "B")

        opts = ExportOptions()
        result = self.exporter._get_export_range(opts)
        # Should span from (0,0) to (2,3)
        start_row, start_col, end_row, end_col = result
        assert start_row == 0
        assert start_col == 0
        assert end_row == 2
        assert end_col == 3

    def test_range_with_options(self):
        """Test getting range with explicit options."""
        self.ss.set_cell(0, 0, "A")
        self.ss.set_cell(5, 5, "B")

        opts = ExportOptions(start_row=1, start_col=1, end_row=3, end_col=3)
        result = self.exporter._get_export_range(opts)
        start_row, start_col, end_row, end_col = result
        assert start_row == 1
        assert start_col == 1
        assert end_row == 3
        assert end_col == 3


class TestDetectOptions:
    """Tests for _detect_options method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.exporter = TextExporter(self.ss)

    def test_detect_csv(self):
        """Test detecting CSV format."""
        opts = self.exporter._detect_options("file.csv")
        assert opts.format == ExportFormat.CSV
        assert opts.delimiter == ","

    def test_detect_tsv(self):
        """Test detecting TSV format."""
        opts = self.exporter._detect_options("file.tsv")
        assert opts.format == ExportFormat.TSV
        assert opts.delimiter == "\t"

    def test_detect_txt(self):
        """Test detecting TXT format."""
        opts = self.exporter._detect_options("file.txt")
        assert opts.format == ExportFormat.FORMATTED_TEXT

    def test_detect_unknown(self):
        """Test detecting unknown format defaults to CSV."""
        opts = self.exporter._detect_options("file.xyz")
        assert opts.format == ExportFormat.CSV

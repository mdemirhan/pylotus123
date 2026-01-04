"""Text file export functionality.

Supports exporting to:
- CSV files (comma-separated)
- TSV files (tab-separated)
- Other delimited files
- Plain text with formatting
"""

import csv
from dataclasses import dataclass
from enum import Enum, auto
from pathlib import Path

from ..core.spreadsheet_protocol import SpreadsheetProtocol


class ExportFormat(Enum):
    """Export file format."""

    CSV = auto()
    TSV = auto()
    CUSTOM_DELIMITED = auto()
    FORMATTED_TEXT = auto()


@dataclass
class ExportOptions:
    """Options for text export.

    Attributes:
        format: File format
        delimiter: Custom delimiter
        include_header: Include column headers
        start_row: First row to export
        start_col: First column to export
        end_row: Last row to export (None for all)
        end_col: Last column to export (None for all)
        text_qualifier: Character for quoting text
        encoding: Text encoding
        use_formulas: Export formulas instead of values
        line_ending: Line ending character(s)
    """

    format: ExportFormat = ExportFormat.CSV
    delimiter: str = ","
    include_header: bool = False
    start_row: int = 0
    start_col: int = 0
    end_row: int | None = None
    end_col: int | None = None
    text_qualifier: str = '"'
    encoding: str = "utf-8"
    use_formulas: bool = False
    line_ending: str = "\n"
    quote_all: bool = False


class TextExporter:
    """Export spreadsheet data to text files."""

    def __init__(self, spreadsheet: SpreadsheetProtocol) -> None:
        self.spreadsheet = spreadsheet

    def export_file(self, filename: str | Path, options: ExportOptions | None = None) -> int:
        """Export to a text file.

        Args:
            filename: Path to file
            options: Export options

        Returns:
            Number of rows exported
        """
        if options is None:
            options = self._detect_options(filename)

        path = Path(filename)

        if options.format == ExportFormat.CSV:
            return self._export_csv(path, options)
        elif options.format == ExportFormat.TSV:
            options.delimiter = "\t"
            return self._export_csv(path, options)
        elif options.format == ExportFormat.FORMATTED_TEXT:
            return self._export_formatted(path, options)
        else:
            return self._export_csv(path, options)

    def _detect_options(self, filename: str | Path) -> ExportOptions:
        """Detect export options from filename."""
        path = Path(filename)
        options = ExportOptions()

        ext = path.suffix.lower()
        if ext == ".tsv":
            options.format = ExportFormat.TSV
            options.delimiter = "\t"
        elif ext == ".csv":
            options.format = ExportFormat.CSV
            options.delimiter = ","
        elif ext == ".txt":
            options.format = ExportFormat.FORMATTED_TEXT

        return options

    def _get_export_range(self, options: ExportOptions) -> tuple[int, int, int, int]:
        """Determine the range to export."""
        used = self.spreadsheet.get_used_range()
        if not used:
            return (0, 0, 0, 0)

        (min_row, min_col), (max_row, max_col) = used

        start_row = max(options.start_row, min_row)
        start_col = max(options.start_col, min_col)
        end_row = options.end_row if options.end_row is not None else max_row
        end_col = options.end_col if options.end_col is not None else max_col

        return (start_row, start_col, end_row, end_col)

    def _export_csv(self, path: Path, options: ExportOptions) -> int:
        """Export to CSV/TSV format."""
        start_row, start_col, end_row, end_col = self._get_export_range(options)
        rows_exported = 0

        with open(path, "w", encoding=options.encoding, newline="") as f:
            writer = csv.writer(
                f,
                delimiter=options.delimiter,
                quotechar=options.text_qualifier,
                quoting=csv.QUOTE_ALL if options.quote_all else csv.QUOTE_MINIMAL,
            )

            # Header row with column letters
            if options.include_header:
                from ..core.reference import index_to_col

                headers = [index_to_col(c) for c in range(start_col, end_col + 1)]
                writer.writerow(headers)
                rows_exported += 1

            # Data rows
            for row in range(start_row, end_row + 1):
                row_data = []
                for col in range(start_col, end_col + 1):
                    if options.use_formulas:
                        cell = self.spreadsheet.get_cell_if_exists(row, col)
                        value = cell.raw_value if cell else ""
                    else:
                        value = self.spreadsheet.get_display_value(row, col)
                        # Preserve text that looks like formulas by adding ' prefix
                        # This prevents re-import from treating them as formulas
                        cell = self.spreadsheet.get_cell_if_exists(row, col)
                        if value and not (cell and cell.is_formula):
                            if value[0] in "=@+-":
                                value = "'" + value
                    row_data.append(value)

                writer.writerow(row_data)
                rows_exported += 1

        return rows_exported

    def _export_formatted(self, path: Path, options: ExportOptions) -> int:
        """Export to formatted text with aligned columns."""
        start_row, start_col, end_row, end_col = self._get_export_range(options)
        rows_exported = 0

        # Get column widths
        col_widths = []
        for col in range(start_col, end_col + 1):
            width = self.spreadsheet.get_col_width(col)
            col_widths.append(width)

        with open(path, "w", encoding=options.encoding) as f:
            # Header row
            if options.include_header:
                from ..core.reference import index_to_col

                header_line = ""
                for i, col in enumerate(range(start_col, end_col + 1)):
                    col_name = index_to_col(col)
                    header_line += col_name.ljust(col_widths[i]) + " "
                f.write(header_line.rstrip() + options.line_ending)

                # Separator
                sep_line = ""
                for width in col_widths:
                    sep_line += "-" * width + " "
                f.write(sep_line.rstrip() + options.line_ending)
                rows_exported += 2

            # Data rows
            for row in range(start_row, end_row + 1):
                row_line = ""
                for i, col in enumerate(range(start_col, end_col + 1)):
                    if options.use_formulas:
                        cell = self.spreadsheet.get_cell_if_exists(row, col)
                        value = cell.raw_value if cell else ""
                    else:
                        value = self.spreadsheet.get_display_value(row, col)

                    # Align based on content type
                    cell = self.spreadsheet.get_cell_if_exists(row, col)
                    if cell:
                        aligned = cell.get_aligned_display(col_widths[i])
                    else:
                        aligned = value.ljust(col_widths[i])[: col_widths[i]]

                    row_line += aligned + " "

                f.write(row_line.rstrip() + options.line_ending)
                rows_exported += 1

        return rows_exported

    def export_to_string(self, options: ExportOptions | None = None) -> str:
        """Export to a string.

        Args:
            options: Export options

        Returns:
            Exported data as string
        """
        if options is None:
            options = ExportOptions()

        start_row, start_col, end_row, end_col = self._get_export_range(options)
        lines = []

        for row in range(start_row, end_row + 1):
            row_data = []
            for col in range(start_col, end_col + 1):
                if options.use_formulas:
                    cell = self.spreadsheet.get_cell_if_exists(row, col)
                    value = cell.raw_value if cell else ""
                else:
                    value = self.spreadsheet.get_display_value(row, col)
                    # Preserve text that looks like formulas by adding ' prefix
                    cell = self.spreadsheet.get_cell_if_exists(row, col)
                    if value and not (cell and cell.is_formula):
                        if value[0] in "=@+-":
                            value = "'" + value

                # Quote if needed
                if options.delimiter in str(value) or options.text_qualifier in str(value):
                    value = f"{options.text_qualifier}{value}{options.text_qualifier}"
                elif options.quote_all:
                    value = f"{options.text_qualifier}{value}{options.text_qualifier}"

                row_data.append(str(value))

            lines.append(options.delimiter.join(row_data))

        return options.line_ending.join(lines)

    def export_range(
        self,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int,
        filename: str | Path,
        options: ExportOptions | None = None,
    ) -> int:
        """Export a specific range.

        Args:
            start_row, start_col: Top-left of range
            end_row, end_col: Bottom-right of range
            filename: Output file path
            options: Export options

        Returns:
            Number of rows exported
        """
        if options is None:
            options = self._detect_options(filename)

        options.start_row = start_row
        options.start_col = start_col
        options.end_row = end_row
        options.end_col = end_col

        return self.export_file(filename, options)

"""Text file import functionality.

Supports importing:
- CSV files (comma-separated)
- TSV files (tab-separated)
- Fixed-width files
- Other delimited files
"""

import csv
from dataclasses import dataclass, field
from enum import Enum, auto
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


class ImportFormat(Enum):
    """Import file format."""

    CSV = auto()
    TSV = auto()
    FIXED_WIDTH = auto()
    CUSTOM_DELIMITED = auto()


@dataclass
class ImportOptions:
    """Options for text import.

    Attributes:
        format: File format
        delimiter: Custom delimiter for CUSTOM_DELIMITED
        has_header: First row is header
        start_row: Row to start import (0-based in file)
        start_col: Column to start import in spreadsheet
        dest_row: Destination row in spreadsheet
        dest_col: Destination column in spreadsheet
        text_qualifier: Character for quoting text (usually ")
        field_widths: List of widths for FIXED_WIDTH format
        encoding: Text encoding
    """

    format: ImportFormat = ImportFormat.CSV
    delimiter: str = ","
    has_header: bool = False
    start_row: int = 0
    start_col: int = 0
    dest_row: int = 0
    dest_col: int = 0
    text_qualifier: str = '"'
    field_widths: list[int] = field(default_factory=list)
    encoding: str = "utf-8"
    skip_blank_lines: bool = True
    trim_whitespace: bool = True


class TextImporter:
    """Import text files into the spreadsheet."""

    def __init__(self, spreadsheet: Spreadsheet) -> None:
        self.spreadsheet = spreadsheet

    def import_file(self, filename: str | Path, options: ImportOptions | None = None) -> int:
        """Import a text file into the spreadsheet.

        Args:
            filename: Path to file
            options: Import options (auto-detected if None)

        Returns:
            Number of rows imported
        """
        if options is None:
            options = self._detect_options(filename)

        path = Path(filename)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {filename}")

        if options.format == ImportFormat.CSV:
            return self._import_csv(path, options)
        elif options.format == ImportFormat.TSV:
            options.delimiter = "\t"
            return self._import_csv(path, options)
        elif options.format == ImportFormat.FIXED_WIDTH:
            return self._import_fixed_width(path, options)
        else:
            return self._import_csv(path, options)

    def _detect_options(self, filename: str | Path) -> ImportOptions:
        """Auto-detect import options from file."""
        path = Path(filename)
        options = ImportOptions()

        # Detect format from extension
        ext = path.suffix.lower()
        if ext == ".tsv":
            options.format = ImportFormat.TSV
            options.delimiter = "\t"
        elif ext == ".csv":
            options.format = ImportFormat.CSV
            options.delimiter = ","
        elif ext == ".txt":
            # Try to detect delimiter from content
            try:
                with open(path, "r", encoding="utf-8") as f:
                    first_line = f.readline()
                    if "\t" in first_line:
                        options.format = ImportFormat.TSV
                        options.delimiter = "\t"
                    elif "," in first_line:
                        options.format = ImportFormat.CSV
                        options.delimiter = ","
                    elif ";" in first_line:
                        options.format = ImportFormat.CUSTOM_DELIMITED
                        options.delimiter = ";"
            except Exception:
                pass

        return options

    def _import_csv(self, path: Path, options: ImportOptions) -> int:
        """Import CSV/TSV file."""
        rows_imported = 0

        with open(path, "r", encoding=options.encoding, newline="") as f:
            reader = csv.reader(f, delimiter=options.delimiter, quotechar=options.text_qualifier)

            # Skip rows if needed
            for _ in range(options.start_row):
                try:
                    next(reader)
                except StopIteration:
                    return 0

            # Import rows
            for row_idx, row in enumerate(reader):
                if options.skip_blank_lines and not any(row):
                    continue

                dest_row = options.dest_row + rows_imported

                for col_idx, value in enumerate(row[options.start_col :]):
                    if options.trim_whitespace:
                        value = value.strip()

                    dest_col = options.dest_col + col_idx
                    self.spreadsheet.set_cell(dest_row, dest_col, value)

                rows_imported += 1

        self.spreadsheet.invalidate_cache()
        return rows_imported

    def _import_fixed_width(self, path: Path, options: ImportOptions) -> int:
        """Import fixed-width file."""
        if not options.field_widths:
            # Try to auto-detect widths
            options.field_widths = self._detect_field_widths(path)

        if not options.field_widths:
            raise ValueError("Field widths must be specified for fixed-width import")

        rows_imported = 0

        with open(path, "r", encoding=options.encoding) as f:
            # Skip rows
            for _ in range(options.start_row):
                if not f.readline():
                    return 0

            for line in f:
                if options.skip_blank_lines and not line.strip():
                    continue

                dest_row = options.dest_row + rows_imported

                # Parse fixed-width fields
                pos = 0
                for col_idx, width in enumerate(options.field_widths):
                    if col_idx < options.start_col:
                        pos += width
                        continue

                    value = line[pos : pos + width]
                    if options.trim_whitespace:
                        value = value.strip()

                    dest_col = options.dest_col + (col_idx - options.start_col)
                    self.spreadsheet.set_cell(dest_row, dest_col, value)
                    pos += width

                rows_imported += 1

        self.spreadsheet.invalidate_cache()
        return rows_imported

    def _detect_field_widths(self, path: Path) -> list[int]:
        """Try to detect field widths from file content."""
        # Simple heuristic: look for consistent spacing
        with open(path, "r", encoding="utf-8") as f:
            lines = [f.readline() for _ in range(min(10, 1000))]

        if not lines:
            return []

        # Find positions where spaces appear consistently
        max_len = max(len(line) for line in lines)
        space_counts = [0] * max_len

        for line in lines:
            for idx, char in enumerate(line):
                if char == " ":
                    space_counts[idx] += 1

        # Find field boundaries (positions with high space counts)
        threshold = len(lines) * 0.8
        boundaries = [0]
        in_space = False

        for i, count in enumerate(space_counts):
            if count >= threshold:
                if not in_space and i > 0:
                    boundaries.append(i)
                in_space = True
            else:
                in_space = False

        # Calculate widths from boundaries
        widths = []
        for i in range(len(boundaries) - 1):
            widths.append(boundaries[i + 1] - boundaries[i])

        if boundaries:
            widths.append(max_len - boundaries[-1])

        return widths if widths else []

    def import_text(self, text: str, options: ImportOptions | None = None) -> int:
        """Import from a text string.

        Args:
            text: Text content to import
            options: Import options

        Returns:
            Number of rows imported
        """
        if options is None:
            options = ImportOptions()

        rows_imported = 0
        lines = text.split("\n")

        for line in lines[options.start_row :]:
            if options.skip_blank_lines and not line.strip():
                continue

            dest_row = options.dest_row + rows_imported

            if options.format == ImportFormat.FIXED_WIDTH:
                # Fixed-width parsing
                pos = 0
                for col_idx, width in enumerate(options.field_widths):
                    if col_idx >= options.start_col:
                        value = line[pos : pos + width]
                        if options.trim_whitespace:
                            value = value.strip()
                        dest_col = options.dest_col + (col_idx - options.start_col)
                        self.spreadsheet.set_cell(dest_row, dest_col, value)
                    pos += width
            else:
                # Delimited parsing
                values = line.split(options.delimiter)
                for col_idx, value in enumerate(values[options.start_col :]):
                    if options.trim_whitespace:
                        value = value.strip()
                        # Remove quotes if present
                        if value.startswith(options.text_qualifier) and value.endswith(
                            options.text_qualifier
                        ):
                            value = value[1:-1]
                    dest_col = options.dest_col + col_idx
                    self.spreadsheet.set_cell(dest_row, dest_col, value)

            rows_imported += 1

        self.spreadsheet.invalidate_cache()
        return rows_imported

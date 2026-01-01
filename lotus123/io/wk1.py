"""Lotus WK1 format handler.

WK1 is a binary format used by Lotus 1-2-3 Release 2.
File structure: sequence of variable-length records.
Each record: 2-byte opcode (little-endian) + 2-byte length + data.

References:
- http://fileformats.archiveteam.org/wiki/Lotus_1-2-3
- https://docs.fileformat.com/spreadsheet/123/
"""

from __future__ import annotations

import struct
from typing import TYPE_CHECKING, BinaryIO

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


# WK1 Record Opcodes (per Lotus 1-2-3 Release 2 specification)
BOF = 0x0000      # Beginning of file
EOF = 0x0001      # End of file
CALCMODE = 0x0002  # Calculation mode
CALCORDER = 0x0003  # Calculation order
SPLIT = 0x0004    # Split window type
SYNC = 0x0005     # Split window sync
RANGE = 0x0006    # Active worksheet range
WINDOW1 = 0x0007  # Window 1 record
COLW1 = 0x0008    # Column width, window 1
BLANK = 0x000C    # Blank cell
LABEL = 0x000F    # Label (string) cell
INTEGER = 0x000D  # Integer number cell (16-bit signed)
NUMBER = 0x000E   # Floating point number cell (IEEE 754 double)
FORMULA = 0x0010  # Formula cell

# Version codes
VERSION_WK1 = 0x0406  # Lotus 1-2-3 Release 2 (WK1)
VERSION_WKS = 0x0404  # Lotus 1-2-3 Release 1A (WKS)


class Wk1Reader:
    """Read Lotus WK1 files."""

    def __init__(self, spreadsheet: Spreadsheet) -> None:
        """Initialize reader with target spreadsheet.

        Args:
            spreadsheet: Spreadsheet to load data into
        """
        self.spreadsheet = spreadsheet

    def load(self, filepath: str) -> None:
        """Load WK1 file into spreadsheet.

        Args:
            filepath: Path to WK1 file

        Raises:
            ValueError: If file format is invalid
            FileNotFoundError: If file doesn't exist
        """
        self.spreadsheet.clear()

        with open(filepath, "rb") as f:
            self._read_file(f)

        self.spreadsheet.filename = filepath
        self.spreadsheet.modified = False
        self.spreadsheet.rebuild_dependency_graph()

    def _read_file(self, f: BinaryIO) -> None:
        """Read records from WK1 file.

        Args:
            f: Binary file handle
        """
        # Read BOF record
        opcode, length = self._read_record_header(f)
        if opcode != BOF:
            raise ValueError("Invalid WK1 file: missing BOF record")

        data = f.read(length)
        if len(data) < 2:
            raise ValueError("Invalid WK1 file: truncated BOF record")

        version = struct.unpack("<H", data[:2])[0]
        if version not in (VERSION_WK1, VERSION_WKS):
            raise ValueError(f"Unsupported WK1 version: {version:#06x}")

        # Read records until EOF
        while True:
            opcode, length = self._read_record_header(f)
            if opcode == EOF:
                break
            if opcode is None:
                break  # End of file

            data = f.read(length)
            if len(data) < length:
                break  # Truncated file

            self._process_record(opcode, data)

    def _read_record_header(self, f: BinaryIO) -> tuple[int | None, int]:
        """Read 4-byte record header.

        Args:
            f: Binary file handle

        Returns:
            Tuple of (opcode, length), or (None, 0) if EOF
        """
        header = f.read(4)
        if len(header) < 4:
            return None, 0
        opcode, length = struct.unpack("<HH", header)
        return opcode, length

    def _process_record(self, opcode: int, data: bytes) -> None:
        """Process a single record.

        Args:
            opcode: Record type
            data: Record data
        """
        if opcode == LABEL:
            self._read_label(data)
        elif opcode == INTEGER:
            self._read_integer(data)
        elif opcode == NUMBER:
            self._read_number(data)
        elif opcode == FORMULA:
            self._read_formula(data)
        elif opcode == COLW1:
            self._read_column_width(data)
        elif opcode == BLANK:
            pass  # Blank cells are ignored
        # Ignore other records (RANGE, WINDOW1, etc.)

    def _read_label(self, data: bytes) -> None:
        """Read string cell.

        Format: format(1) + col(2) + row(2) + string(null-terminated)
        """
        if len(data) < 6:
            return

        # Format byte contains alignment info
        _fmt = data[0]
        col = struct.unpack("<H", data[1:3])[0]
        row = struct.unpack("<H", data[3:5])[0]

        # String is null-terminated, may have prefix character
        string_data = data[5:]
        null_pos = string_data.find(0)
        if null_pos != -1:
            string_data = string_data[:null_pos]

        # Decode using CP437 (DOS codepage)
        try:
            text = string_data.decode("cp437", errors="replace")
        except Exception:
            text = string_data.decode("latin-1", errors="replace")

        # Keep Lotus label prefix if present for alignment
        # ' = left, " = right, ^ = center, \ = repeat
        self.spreadsheet.set_cell(row, col, text)

    def _read_integer(self, data: bytes) -> None:
        """Read integer cell.

        Format: format(1) + col(2) + row(2) + value(2)
        """
        if len(data) < 7:
            return

        col = struct.unpack("<H", data[1:3])[0]
        row = struct.unpack("<H", data[3:5])[0]
        value = struct.unpack("<h", data[5:7])[0]  # Signed 16-bit

        self.spreadsheet.set_cell(row, col, str(value))

    def _read_number(self, data: bytes) -> None:
        """Read floating-point cell.

        Format: format(1) + col(2) + row(2) + value(8)
        """
        if len(data) < 13:
            return

        col = struct.unpack("<H", data[1:3])[0]
        row = struct.unpack("<H", data[3:5])[0]
        value = struct.unpack("<d", data[5:13])[0]  # IEEE 754 double

        # Format nicely - avoid unnecessary decimals
        if value == int(value):
            self.spreadsheet.set_cell(row, col, str(int(value)))
        else:
            self.spreadsheet.set_cell(row, col, str(value))

    def _read_formula(self, data: bytes) -> None:
        """Read formula cell - store calculated value only.

        Format: format(1) + col(2) + row(2) + value(8) + length(2) + formula_bytes

        Note: We only store the cached value as formula bytecode
        parsing is very complex and not implemented.
        """
        if len(data) < 13:
            return

        col = struct.unpack("<H", data[1:3])[0]
        row = struct.unpack("<H", data[3:5])[0]
        value = struct.unpack("<d", data[5:13])[0]  # Cached result

        # Store as value (formula parsing is too complex)
        if value == int(value):
            self.spreadsheet.set_cell(row, col, str(int(value)))
        else:
            self.spreadsheet.set_cell(row, col, str(value))

    def _read_column_width(self, data: bytes) -> None:
        """Read column width.

        Format: col(2) + width(1)
        """
        if len(data) < 3:
            return

        col = struct.unpack("<H", data[0:2])[0]
        width = data[2]

        # Default width in WK1 is 9
        if width != 9:
            self.spreadsheet.set_col_width(col, width)


class Wk1Writer:
    """Write Lotus WK1 files."""

    def __init__(self, spreadsheet: Spreadsheet) -> None:
        """Initialize writer with source spreadsheet.

        Args:
            spreadsheet: Spreadsheet to save
        """
        self.spreadsheet = spreadsheet

    def save(self, filepath: str) -> None:
        """Save spreadsheet to WK1 format.

        Args:
            filepath: Path to save file

        Note:
            - Formulas are saved as their calculated values
            - Only basic cell types are supported
            - Maximum dimensions: 256 cols x 8192 rows
        """
        with open(filepath, "wb") as f:
            self._write_file(f)

    def _write_file(self, f: BinaryIO) -> None:
        """Write WK1 file.

        Args:
            f: Binary file handle
        """
        # Write BOF (Beginning of File)
        self._write_record(f, BOF, struct.pack("<H", VERSION_WK1))

        # Write RANGE (used area)
        used = self.spreadsheet.get_used_range()
        if used:
            (min_row, min_col), (max_row, max_col) = used
            # WK1 RANGE format: start_col, start_row, end_col, end_row (all 2 bytes)
            range_data = struct.pack("<HHHH", min_col, min_row, max_col, max_row)
            self._write_record(f, RANGE, range_data)

        # Write column widths
        for col, width in sorted(self.spreadsheet._col_widths.items()):
            colw_data = struct.pack("<HB", col, width)
            self._write_record(f, COLW1, colw_data)

        # Write cells
        for (row, col), cell in sorted(self.spreadsheet._cells.items()):
            if cell.is_empty:
                continue

            # Get calculated value for formulas
            value = self.spreadsheet.get_value(row, col)

            if value is None or value == "":
                continue
            elif isinstance(value, str):
                # Check if it's an error
                if value.startswith("#"):
                    self._write_label(f, row, col, value)
                else:
                    self._write_label(f, row, col, value)
            elif isinstance(value, bool):
                # Booleans as integers
                self._write_integer(f, row, col, 1 if value else 0)
            elif isinstance(value, int) and -32768 <= value <= 32767:
                self._write_integer(f, row, col, value)
            elif isinstance(value, (int, float)):
                self._write_number(f, row, col, float(value))
            else:
                # Fallback: convert to string
                self._write_label(f, row, col, str(value))

        # Write EOF (End of File)
        self._write_record(f, EOF, b"")

    def _write_record(self, f: BinaryIO, opcode: int, data: bytes) -> None:
        """Write a record.

        Args:
            f: Binary file handle
            opcode: Record type
            data: Record data
        """
        header = struct.pack("<HH", opcode, len(data))
        f.write(header)
        f.write(data)

    def _write_label(self, f: BinaryIO, row: int, col: int, text: str) -> None:
        """Write string cell.

        Args:
            f: Binary file handle
            row: Row index (0-based)
            col: Column index (0-based)
            text: Cell text
        """
        # Check if text already has alignment prefix
        if text and text[0] in "'\"^\\":
            prefix = ""
        else:
            prefix = "'"  # Default left-aligned

        # Encode to CP437 with null terminator
        try:
            encoded = (prefix + text).encode("cp437", errors="replace") + b"\x00"
        except Exception:
            encoded = (prefix + text).encode("latin-1", errors="replace") + b"\x00"

        # Format: format_byte(1) + col(2) + row(2) + string
        data = struct.pack("<BHH", 0, col, row) + encoded
        self._write_record(f, LABEL, data)

    def _write_integer(self, f: BinaryIO, row: int, col: int, value: int) -> None:
        """Write integer cell.

        Args:
            f: Binary file handle
            row: Row index (0-based)
            col: Column index (0-based)
            value: Integer value (-32768 to 32767)
        """
        # Format: format_byte(1) + col(2) + row(2) + value(2)
        data = struct.pack("<BHHh", 0, col, row, value)
        self._write_record(f, INTEGER, data)

    def _write_number(self, f: BinaryIO, row: int, col: int, value: float) -> None:
        """Write floating-point cell.

        Args:
            f: Binary file handle
            row: Row index (0-based)
            col: Column index (0-based)
            value: Float value
        """
        # Format: format_byte(1) + col(2) + row(2) + value(8)
        data = struct.pack("<BHHd", 0, col, row, value)
        self._write_record(f, NUMBER, data)

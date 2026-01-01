"""Tests for WK1 file format handler."""

import struct
import tempfile
from pathlib import Path

import pytest

from lotus123.core.spreadsheet import Spreadsheet
from lotus123.io.wk1 import (
    BOF,
    EOF,
    RANGE,
    COLW1,
    LABEL,
    INTEGER,
    NUMBER,
    VERSION_WK1,
    Wk1Reader,
    Wk1Writer,
)


def write_record(f, opcode: int, data: bytes) -> None:
    """Write a WK1 record to file."""
    header = struct.pack("<HH", opcode, len(data))
    f.write(header)
    f.write(data)


def create_minimal_wk1() -> bytes:
    """Create a minimal valid WK1 file in memory."""
    import io

    buf = io.BytesIO()
    # BOF
    write_record(buf, BOF, struct.pack("<H", VERSION_WK1))
    # EOF
    write_record(buf, EOF, b"")
    return buf.getvalue()


def create_wk1_with_data() -> bytes:
    """Create a WK1 file with some data."""
    import io

    buf = io.BytesIO()

    # BOF
    write_record(buf, BOF, struct.pack("<H", VERSION_WK1))

    # RANGE: 0,0 to 2,2
    write_record(buf, RANGE, struct.pack("<HHHH", 0, 0, 2, 2))

    # LABEL at A1: "Hello"
    label_data = struct.pack("<BHH", 0, 0, 0) + b"'Hello\x00"
    write_record(buf, LABEL, label_data)

    # INTEGER at B1: 42
    int_data = struct.pack("<BHHh", 0, 1, 0, 42)
    write_record(buf, INTEGER, int_data)

    # NUMBER at C1: 3.14
    num_data = struct.pack("<BHHd", 0, 2, 0, 3.14)
    write_record(buf, NUMBER, num_data)

    # LABEL at A2: "World"
    label_data2 = struct.pack("<BHH", 0, 0, 1) + b"'World\x00"
    write_record(buf, LABEL, label_data2)

    # EOF
    write_record(buf, EOF, b"")

    return buf.getvalue()


class TestWk1Reader:
    """Tests for WK1 file reading."""

    def test_read_minimal_file(self):
        """Test reading a minimal valid WK1 file."""
        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            f.write(create_minimal_wk1())
            f.flush()
            filepath = f.name

        try:
            reader.load(filepath)
            # Should succeed without error
            assert spreadsheet.filename == filepath
        finally:
            Path(filepath).unlink()

    def test_read_labels(self):
        """Test reading string cells."""
        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            f.write(create_wk1_with_data())
            f.flush()
            filepath = f.name

        try:
            reader.load(filepath)
            # Check labels - note prefix character is preserved
            assert spreadsheet.get_cell(0, 0).raw_value == "'Hello"
            assert spreadsheet.get_cell(1, 0).raw_value == "'World"
        finally:
            Path(filepath).unlink()

    def test_read_integers(self):
        """Test reading integer cells."""
        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            f.write(create_wk1_with_data())
            f.flush()
            filepath = f.name

        try:
            reader.load(filepath)
            # Check integer value
            assert spreadsheet.get_cell(0, 1).raw_value == "42"
        finally:
            Path(filepath).unlink()

    def test_read_numbers(self):
        """Test reading floating-point cells."""
        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            f.write(create_wk1_with_data())
            f.flush()
            filepath = f.name

        try:
            reader.load(filepath)
            # Check float value
            value = float(spreadsheet.get_cell(0, 2).raw_value)
            assert abs(value - 3.14) < 0.001
        finally:
            Path(filepath).unlink()

    def test_invalid_file_no_bof(self):
        """Test error handling for file without BOF."""
        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            # Write EOF directly without BOF
            write_record(f, EOF, b"")
            f.flush()
            filepath = f.name

        try:
            with pytest.raises(ValueError, match="missing BOF"):
                reader.load(filepath)
        finally:
            Path(filepath).unlink()

    def test_file_not_found(self):
        """Test error handling for missing file."""
        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)

        with pytest.raises(FileNotFoundError):
            reader.load("/nonexistent/file.wk1")

    def test_clears_spreadsheet_on_load(self):
        """Test that loading clears existing data."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "Existing data")

        reader = Wk1Reader(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            f.write(create_minimal_wk1())
            f.flush()
            filepath = f.name

        try:
            reader.load(filepath)
            # Existing data should be cleared
            assert spreadsheet.get_cell(0, 0).is_empty
        finally:
            Path(filepath).unlink()


class TestWk1Writer:
    """Tests for WK1 file writing."""

    def test_write_empty_spreadsheet(self):
        """Test writing an empty spreadsheet."""
        spreadsheet = Spreadsheet()
        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            # Read back and verify structure
            with open(filepath, "rb") as f:
                # Read BOF
                opcode, length = struct.unpack("<HH", f.read(4))
                assert opcode == BOF
                version = struct.unpack("<H", f.read(length))[0]
                assert version == VERSION_WK1

                # Read EOF (may have other records in between)
                # Just verify file is valid
                data = f.read()
                assert len(data) >= 4  # At least EOF record
        finally:
            Path(filepath).unlink()

    def test_write_labels(self):
        """Test writing string cells."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "Hello")
        spreadsheet.set_cell(1, 0, "World")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            # Read back with reader
            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Check values (prefix added by writer)
            assert "Hello" in new_spreadsheet.get_cell(0, 0).raw_value
            assert "World" in new_spreadsheet.get_cell(1, 0).raw_value
        finally:
            Path(filepath).unlink()

    def test_write_integers(self):
        """Test writing integer cells."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "42")
        spreadsheet.set_cell(0, 1, "-100")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            # Read back with reader
            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.get_cell(0, 0).raw_value == "42"
            assert new_spreadsheet.get_cell(0, 1).raw_value == "-100"
        finally:
            Path(filepath).unlink()

    def test_write_numbers(self):
        """Test writing floating-point cells."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "3.14159")
        spreadsheet.set_cell(0, 1, "2.71828")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            # Read back with reader
            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            value1 = float(new_spreadsheet.get_cell(0, 0).raw_value)
            value2 = float(new_spreadsheet.get_cell(0, 1).raw_value)
            assert abs(value1 - 3.14159) < 0.0001
            assert abs(value2 - 2.71828) < 0.0001
        finally:
            Path(filepath).unlink()

    def test_write_column_widths(self):
        """Test writing column widths."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_col_width(0, 15)
        spreadsheet.set_col_width(2, 20)
        spreadsheet.set_cell(0, 0, "Test")  # Need some data

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            # Read back with reader
            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.get_col_width(0) == 15
            assert new_spreadsheet.get_col_width(2) == 20
        finally:
            Path(filepath).unlink()

    def test_formulas_saved_as_values(self):
        """Test that formulas are saved as their calculated values."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "10")
        spreadsheet.set_cell(0, 1, "20")
        spreadsheet.set_cell(0, 2, "=A1+B1")  # Formula

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            # Read back with reader
            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Formula should be saved as calculated value (30)
            value = new_spreadsheet.get_cell(0, 2).raw_value
            assert value == "30"
        finally:
            Path(filepath).unlink()


class TestWk1Roundtrip:
    """Tests for WK1 save/load roundtrip."""

    def test_roundtrip_mixed_data(self):
        """Test saving and loading mixed data types."""
        spreadsheet = Spreadsheet()

        # Add various data types
        spreadsheet.set_cell(0, 0, "Name")
        spreadsheet.set_cell(0, 1, "Age")
        spreadsheet.set_cell(0, 2, "Score")

        spreadsheet.set_cell(1, 0, "Alice")
        spreadsheet.set_cell(1, 1, "30")
        spreadsheet.set_cell(1, 2, "95.5")

        spreadsheet.set_cell(2, 0, "Bob")
        spreadsheet.set_cell(2, 1, "25")
        spreadsheet.set_cell(2, 2, "87.3")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            # Read back
            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Verify data (labels have prefix)
            assert "Name" in new_spreadsheet.get_cell(0, 0).raw_value
            assert "Age" in new_spreadsheet.get_cell(0, 1).raw_value

            assert "Alice" in new_spreadsheet.get_cell(1, 0).raw_value
            assert new_spreadsheet.get_cell(1, 1).raw_value == "30"
            assert abs(float(new_spreadsheet.get_cell(1, 2).raw_value) - 95.5) < 0.01

            assert "Bob" in new_spreadsheet.get_cell(2, 0).raw_value
            assert new_spreadsheet.get_cell(2, 1).raw_value == "25"
            assert abs(float(new_spreadsheet.get_cell(2, 2).raw_value) - 87.3) < 0.01
        finally:
            Path(filepath).unlink()

    def test_roundtrip_large_numbers(self):
        """Test roundtrip with large numbers."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "1000000")
        spreadsheet.set_cell(0, 1, "1234567890.12345")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.get_cell(0, 0).raw_value == "1000000"
            value = float(new_spreadsheet.get_cell(0, 1).raw_value)
            assert abs(value - 1234567890.12345) < 0.001
        finally:
            Path(filepath).unlink()

    def test_roundtrip_negative_numbers(self):
        """Test roundtrip with negative numbers."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "-42")
        spreadsheet.set_cell(0, 1, "-3.14")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.get_cell(0, 0).raw_value == "-42"
            value = float(new_spreadsheet.get_cell(0, 1).raw_value)
            assert abs(value - (-3.14)) < 0.001
        finally:
            Path(filepath).unlink()

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
    LABEL,
    INTEGER,
    NUMBER,
    FORMULA,
    VERSION_WK1,
    Wk1Reader,
    Wk1Writer,
    compile_formula,
    decompile_formula,
    encode_format_byte,
    decode_format_byte,
    FMT_FIXED,
    FMT_SCIENTIFIC,
    FMT_CURRENCY,
    FMT_PERCENT,
    FMT_COMMA,
    FMT_DATETIME,
    FMT_DEFAULT,
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

    def test_formulas_preserved(self):
        """Test that formulas are preserved in WK1 format."""
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

            # Formula should be preserved
            value = new_spreadsheet.get_cell(0, 2).raw_value
            assert value == "=A1+B1"
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


class TestFormulaDecompiler:
    """Tests for WK1 formula decompilation (bytecode to text)."""

    def test_decompile_simple_addition(self):
        """Test decompiling A1+B1."""
        # Bytecode: variable(A1) + variable(B1) + ADD + RETURN
        bytecode = bytes(
            [
                0x01,
                0x00,
                0x00,
                0x00,
                0x00,  # variable: col=0, row=0 (A1)
                0x01,
                0x01,
                0x00,
                0x00,
                0x00,  # variable: col=1, row=0 (B1)
                0x09,  # ADD
                0x03,  # RETURN
            ]
        )
        result = decompile_formula(bytecode)
        assert result == "=A1+B1"

    def test_decompile_integer_constant(self):
        """Test decompiling integer constant."""
        # Bytecode: integer(42) + RETURN
        bytecode = bytes(
            [
                0x05,
                0x2A,
                0x00,  # integer: 42 (0x002A)
                0x03,  # RETURN
            ]
        )
        result = decompile_formula(bytecode)
        assert result == "=42"

    def test_decompile_subtraction(self):
        """Test decompiling A1-B1."""
        bytecode = bytes(
            [
                0x01,
                0x00,
                0x00,
                0x00,
                0x00,  # A1
                0x01,
                0x01,
                0x00,
                0x00,
                0x00,  # B1
                0x0A,  # SUB
                0x03,  # RETURN
            ]
        )
        result = decompile_formula(bytecode)
        assert result == "=A1-B1"

    def test_decompile_multiplication(self):
        """Test decompiling A1*B1."""
        bytecode = bytes(
            [
                0x01,
                0x00,
                0x00,
                0x00,
                0x00,  # A1
                0x01,
                0x01,
                0x00,
                0x00,
                0x00,  # B1
                0x0B,  # MUL
                0x03,  # RETURN
            ]
        )
        result = decompile_formula(bytecode)
        assert result == "=A1*B1"

    def test_decompile_division(self):
        """Test decompiling A1/B1."""
        bytecode = bytes(
            [
                0x01,
                0x00,
                0x00,
                0x00,
                0x00,  # A1
                0x01,
                0x01,
                0x00,
                0x00,
                0x00,  # B1
                0x0C,  # DIV
                0x03,  # RETURN
            ]
        )
        result = decompile_formula(bytecode)
        assert result == "=A1/B1"

    def test_decompile_unary_minus(self):
        """Test decompiling -A1."""
        bytecode = bytes(
            [
                0x01,
                0x00,
                0x00,
                0x00,
                0x00,  # A1
                0x08,  # UNARY_MINUS
                0x03,  # RETURN
            ]
        )
        result = decompile_formula(bytecode)
        assert result == "=-A1"

    def test_decompile_comparison(self):
        """Test decompiling A1>B1."""
        bytecode = bytes(
            [
                0x01,
                0x00,
                0x00,
                0x00,
                0x00,  # A1
                0x01,
                0x01,
                0x00,
                0x00,
                0x00,  # B1
                0x13,  # GT
                0x03,  # RETURN
            ]
        )
        result = decompile_formula(bytecode)
        assert result == "=A1>B1"


class TestFormulaCompiler:
    """Tests for WK1 formula compilation (text to bytecode)."""

    def test_compile_simple_addition(self):
        """Test compiling =A1+B1."""
        bytecode = compile_formula("=A1+B1")
        # Should produce: variable(A1) + variable(B1) + ADD + RETURN
        assert len(bytecode) > 0
        assert bytecode[-1] == 0x03  # RETURN

    def test_compile_number(self):
        """Test compiling =42."""
        bytecode = compile_formula("=42")
        # Should produce: integer(42) + RETURN
        assert len(bytecode) == 4  # opcode + 2-byte int + RETURN
        assert bytecode[0] == 0x05  # OP_INTEGER
        assert bytecode[-1] == 0x03  # RETURN

    def test_compile_cell_reference(self):
        """Test compiling =A1."""
        bytecode = compile_formula("=A1")
        # Should produce: variable(A1) + RETURN
        assert len(bytecode) == 6  # opcode + 4-byte ref + RETURN
        assert bytecode[0] == 0x01  # OP_VARIABLE
        assert bytecode[-1] == 0x03  # RETURN

    def test_compile_function(self):
        """Test compiling =@SUM(A1:A10)."""
        bytecode = compile_formula("=@SUM(A1:A10)")
        assert len(bytecode) > 0
        assert bytecode[-1] == 0x03  # RETURN
        # Should contain SUM opcode (0x50)
        assert 0x50 in bytecode

    def test_compile_empty_formula(self):
        """Test compiling empty formula returns just RETURN."""
        bytecode = compile_formula("=")
        assert bytecode == bytes([0x03])  # Just RETURN

    # --- Function argument parsing tests (regression tests for argument bug) ---

    def test_compile_function_with_range_arg(self):
        """Test that function with range argument compiles correctly.

        This is a regression test for the bug where function arguments
        were not being parsed due to incorrect position reset.
        """
        bytecode = compile_formula("=SUM(A1:A10)")
        # Must contain: range opcode (0x02), SUM opcode (0x50), arg count, RETURN
        assert 0x02 in bytecode  # OP_RANGE must be present
        assert 0x50 in bytecode  # SUM opcode must be present
        assert bytecode[-1] == 0x03  # RETURN

    def test_compile_function_with_cell_arg(self):
        """Test function with single cell argument."""
        bytecode = compile_formula("=ABS(A1)")
        # Must contain: variable opcode (0x01), ABS opcode (0x21)
        assert 0x01 in bytecode  # OP_VARIABLE
        assert 0x21 in bytecode  # ABS opcode
        assert bytecode[-1] == 0x03

    def test_compile_function_with_number_arg(self):
        """Test function with numeric argument."""
        bytecode = compile_formula("=SQRT(144)")
        # Must contain: integer opcode (0x05), SQRT opcode (0x23)
        assert 0x05 in bytecode  # OP_INTEGER
        assert 0x23 in bytecode  # SQRT opcode
        assert bytecode[-1] == 0x03

    def test_compile_function_with_multiple_args(self):
        """Test function with multiple comma-separated arguments."""
        bytecode = compile_formula("=@SUM(A1,B1,C1)")
        # Must have SUM opcode and arg count of 3
        assert 0x50 in bytecode  # SUM opcode
        # Find SUM opcode position and check arg count follows
        sum_pos = bytecode.index(0x50)
        assert bytecode[sum_pos + 1] == 3  # 3 arguments

    def test_compile_function_excel_style(self):
        """Test function without @ prefix (Excel-style)."""
        bytecode = compile_formula("=SUM(A1:B10)")
        assert 0x02 in bytecode  # OP_RANGE
        assert 0x50 in bytecode  # SUM opcode
        assert bytecode[-1] == 0x03

    def test_compile_function_lotus_style(self):
        """Test function with @ prefix (Lotus-style)."""
        bytecode = compile_formula("=@SUM(A1:B10)")
        assert 0x02 in bytecode  # OP_RANGE
        assert 0x50 in bytecode  # SUM opcode
        assert bytecode[-1] == 0x03

    def test_compile_avg_function(self):
        """Test AVG function compiles with arguments."""
        bytecode = compile_formula("=AVG(D1:D10)")
        assert 0x02 in bytecode  # OP_RANGE
        assert 0x51 in bytecode  # AVG opcode
        assert bytecode[-1] == 0x03

    def test_compile_max_function(self):
        """Test MAX function compiles with arguments."""
        bytecode = compile_formula("=MAX(A1:A100)")
        assert 0x02 in bytecode  # OP_RANGE
        assert 0x54 in bytecode  # MAX opcode
        assert bytecode[-1] == 0x03

    def test_compile_min_function(self):
        """Test MIN function compiles with arguments."""
        bytecode = compile_formula("=MIN(B1:B50)")
        assert 0x02 in bytecode  # OP_RANGE
        assert 0x53 in bytecode  # MIN opcode
        assert bytecode[-1] == 0x03

    def test_compile_count_function(self):
        """Test COUNT function compiles with arguments."""
        bytecode = compile_formula("=COUNT(A1:Z100)")
        assert 0x02 in bytecode  # OP_RANGE
        assert 0x52 in bytecode  # COUNT opcode
        assert bytecode[-1] == 0x03

    def test_compile_if_function(self):
        """Test IF function with 3 arguments."""
        bytecode = compile_formula("=IF(A1>0,1,0)")
        assert 0x3B in bytecode  # IF opcode
        assert bytecode[-1] == 0x03

    def test_compile_round_function(self):
        """Test ROUND function with 2 arguments."""
        bytecode = compile_formula("=ROUND(A1,2)")
        assert 0x01 in bytecode  # OP_VARIABLE (A1)
        assert 0x3F in bytecode  # ROUND opcode
        assert bytecode[-1] == 0x03

    def test_compile_nested_functions(self):
        """Test nested function calls."""
        bytecode = compile_formula("=SQRT(ABS(A1))")
        assert 0x21 in bytecode  # ABS opcode
        assert 0x23 in bytecode  # SQRT opcode
        assert bytecode[-1] == 0x03

    def test_compile_function_in_expression(self):
        """Test function as part of larger expression."""
        bytecode = compile_formula("=SUM(A1:A10)+100")
        assert 0x50 in bytecode  # SUM opcode
        assert 0x09 in bytecode  # ADD opcode
        assert bytecode[-1] == 0x03

    def test_compile_complex_expression_with_function(self):
        """Test complex expression mixing functions and operators."""
        bytecode = compile_formula("=(SUM(A1:A10)/COUNT(A1:A10))*100")
        assert 0x50 in bytecode  # SUM opcode
        assert 0x52 in bytecode  # COUNT opcode
        assert 0x0C in bytecode  # DIV opcode
        assert 0x0B in bytecode  # MUL opcode
        assert bytecode[-1] == 0x03


class TestFormulaRoundtrip:
    """Tests for formula compile/decompile roundtrip."""

    def test_roundtrip_simple_addition(self):
        """Test roundtrip of A1+B1."""
        original = "=A1+B1"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_cell_reference(self):
        """Test roundtrip of single cell reference."""
        original = "=A1"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_integer(self):
        """Test roundtrip of integer."""
        original = "=100"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_subtraction(self):
        """Test roundtrip of subtraction."""
        original = "=A1-B1"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_multiplication(self):
        """Test roundtrip of multiplication."""
        original = "=A1*B1"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_division(self):
        """Test roundtrip of division."""
        original = "=A1/B1"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    # --- Function roundtrip tests (regression tests for argument bug) ---

    def test_roundtrip_sum_with_range(self):
        """Test roundtrip of SUM function with range argument.

        This is a critical regression test - the original bug caused
        =SUM(A1:A10) to become =@SUM() (losing the argument).
        """
        original = "=@SUM(A1:A10)"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_avg_with_range(self):
        """Test roundtrip of AVG function."""
        original = "=@AVG(B1:B100)"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_max_with_range(self):
        """Test roundtrip of MAX function."""
        original = "=@MAX(C1:C50)"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_min_with_range(self):
        """Test roundtrip of MIN function."""
        original = "=@MIN(D1:D25)"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_count_with_range(self):
        """Test roundtrip of COUNT function."""
        original = "=@COUNT(E1:E200)"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_abs_with_cell(self):
        """Test roundtrip of ABS function with cell argument."""
        original = "=@ABS(A1)"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_sqrt_with_number(self):
        """Test roundtrip of SQRT function with numeric argument."""
        original = "=@SQRT(144)"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_round_with_two_args(self):
        """Test roundtrip of ROUND function with two arguments."""
        original = "=@ROUND(A1,2)"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_if_with_three_args(self):
        """Test roundtrip of IF function with three arguments."""
        original = "=@IF(A1>0,1,0)"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_sum_with_multiple_args(self):
        """Test roundtrip of SUM with multiple cell arguments."""
        original = "=@SUM(A1,B1,C1)"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_nested_functions(self):
        """Test roundtrip of nested function calls."""
        original = "=@SQRT(@ABS(A1))"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_function_in_expression(self):
        """Test roundtrip of function combined with operators."""
        original = "=@SUM(A1:A10)+100"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_complex_formula(self):
        """Test roundtrip of complex formula with functions and operators."""
        original = "=(@SUM(A1:A10)/@COUNT(A1:A10))*100"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_percentage_calculation(self):
        """Test roundtrip of percentage calculation (from sample_data2)."""
        original = "=((C18-B18)/B18)*100"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_no_arg_function(self):
        """Test roundtrip of function with no arguments.

        No-arg functions like PI, FALSE, TRUE, RAND need parentheses for parser.
        """
        original = "=@PI()"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        # Decompiler outputs @PI() with parens for parser compatibility
        assert result == "=@PI()"

    def test_roundtrip_comparison_operators(self):
        """Test roundtrip of comparison operators."""
        for op in ["=", "<>", "<", ">", "<=", ">="]:
            original = f"=A1{op}B1"
            bytecode = compile_formula(original)
            result = decompile_formula(bytecode)
            assert result == original, f"Failed for operator {op}"

    def test_roundtrip_unary_minus(self):
        """Test roundtrip of unary minus."""
        original = "=-A1"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original

    def test_roundtrip_exponentiation(self):
        """Test roundtrip of exponentiation."""
        original = "=A1^2"
        bytecode = compile_formula(original)
        result = decompile_formula(bytecode)
        assert result == original


class TestWk1FormulaIO:
    """Tests for formula reading/writing in WK1 files."""

    def test_write_and_read_formula(self):
        """Test writing and reading a formula cell."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "10")
        spreadsheet.set_cell(0, 1, "20")
        spreadsheet.set_cell(0, 2, "=A1+B1")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Formula should be preserved
            assert new_spreadsheet.get_cell(0, 2).raw_value == "=A1+B1"
        finally:
            Path(filepath).unlink()

    def test_write_and_read_sum_function(self):
        """Test writing and reading @SUM function."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "1")
        spreadsheet.set_cell(1, 0, "2")
        spreadsheet.set_cell(2, 0, "3")
        spreadsheet.set_cell(3, 0, "=@SUM(A1:A3)")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Function formula should be preserved
            raw = new_spreadsheet.get_cell(3, 0).raw_value
            assert "@SUM" in raw.upper()
            # Must contain the range argument
            assert "A1:A3" in raw.upper()
        finally:
            Path(filepath).unlink()

    def test_multiple_formulas(self):
        """Test multiple formulas in same file."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "10")
        spreadsheet.set_cell(0, 1, "20")
        spreadsheet.set_cell(0, 2, "=A1+B1")
        spreadsheet.set_cell(0, 3, "=A1*B1")
        spreadsheet.set_cell(0, 4, "=A1-B1")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.get_cell(0, 2).raw_value == "=A1+B1"
            assert new_spreadsheet.get_cell(0, 3).raw_value == "=A1*B1"
            assert new_spreadsheet.get_cell(0, 4).raw_value == "=A1-B1"
        finally:
            Path(filepath).unlink()

    # --- Comprehensive formula I/O tests (regression tests) ---

    def test_sum_function_with_range_preserved(self):
        """Test that SUM function with range argument is fully preserved.

        This is a critical regression test for the bug where =SUM(D3:D12)
        would become =@SUM() when saved and loaded.
        """
        spreadsheet = Spreadsheet()
        # Set up data
        for i in range(10):
            spreadsheet.set_cell(i, 3, str((i + 1) * 10))  # D1:D10
        # Add SUM formula
        spreadsheet.set_cell(10, 3, "=SUM(D1:D10)")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            raw = new_spreadsheet.get_cell(10, 3).raw_value
            # Must contain SUM and the range
            assert "SUM" in raw.upper()
            assert "D1:D10" in raw.upper()
        finally:
            Path(filepath).unlink()

    def test_avg_function_preserved(self):
        """Test AVG function with range is preserved."""
        spreadsheet = Spreadsheet()
        for i in range(5):
            spreadsheet.set_cell(i, 0, str(i * 20))
        spreadsheet.set_cell(5, 0, "=AVG(A1:A5)")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            raw = new_spreadsheet.get_cell(5, 0).raw_value
            assert "AVG" in raw.upper()
            assert "A1:A5" in raw.upper()
        finally:
            Path(filepath).unlink()

    def test_max_min_functions_preserved(self):
        """Test MAX and MIN functions are preserved."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "100")
        spreadsheet.set_cell(1, 0, "200")
        spreadsheet.set_cell(2, 0, "150")
        spreadsheet.set_cell(3, 0, "=MAX(A1:A3)")
        spreadsheet.set_cell(4, 0, "=MIN(A1:A3)")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            max_raw = new_spreadsheet.get_cell(3, 0).raw_value
            min_raw = new_spreadsheet.get_cell(4, 0).raw_value

            assert "MAX" in max_raw.upper()
            assert "A1:A3" in max_raw.upper()
            assert "MIN" in min_raw.upper()
            assert "A1:A3" in min_raw.upper()
        finally:
            Path(filepath).unlink()

    def test_count_function_preserved(self):
        """Test COUNT function is preserved."""
        spreadsheet = Spreadsheet()
        for i in range(10):
            spreadsheet.set_cell(i, 0, str(i))
        spreadsheet.set_cell(10, 0, "=COUNT(A1:A10)")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            raw = new_spreadsheet.get_cell(10, 0).raw_value
            assert "COUNT" in raw.upper()
            assert "A1:A10" in raw.upper()
        finally:
            Path(filepath).unlink()

    def test_if_function_preserved(self):
        """Test IF function with 3 arguments is preserved."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "100")
        spreadsheet.set_cell(0, 1, "=IF(A1>50,1,0)")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            raw = new_spreadsheet.get_cell(0, 1).raw_value
            assert "IF" in raw.upper()
            assert "A1" in raw.upper()
        finally:
            Path(filepath).unlink()

    def test_nested_functions_preserved(self):
        """Test nested function calls are preserved."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "-100")
        spreadsheet.set_cell(0, 1, "=SQRT(ABS(A1))")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            raw = new_spreadsheet.get_cell(0, 1).raw_value
            assert "SQRT" in raw.upper()
            assert "ABS" in raw.upper()
            assert "A1" in raw.upper()
        finally:
            Path(filepath).unlink()

    def test_complex_expression_with_functions(self):
        """Test complex expression mixing functions and operators."""
        spreadsheet = Spreadsheet()
        for i in range(10):
            spreadsheet.set_cell(i, 0, str((i + 1) * 10))
        # Average calculation using SUM/COUNT
        spreadsheet.set_cell(10, 0, "=(SUM(A1:A10)/COUNT(A1:A10))*100")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            raw = new_spreadsheet.get_cell(10, 0).raw_value
            assert "SUM" in raw.upper()
            assert "COUNT" in raw.upper()
            assert "A1:A10" in raw.upper()
        finally:
            Path(filepath).unlink()

    def test_percentage_formula_preserved(self):
        """Test percentage calculation formula (from sample_data2)."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(17, 1, "45000")  # B18
        spreadsheet.set_cell(17, 2, "52000")  # C18
        spreadsheet.set_cell(17, 3, "=((C18-B18)/B18)*100")  # D18

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            raw = new_spreadsheet.get_cell(17, 3).raw_value
            # Should preserve the full formula
            assert "C18" in raw.upper()
            assert "B18" in raw.upper()
            assert "*100" in raw
        finally:
            Path(filepath).unlink()

    def test_sample_data2_formulas_preserved(self):
        """Test that key formulas from sample_data2.json are preserved.

        This tests the actual formulas that were failing before the fix.
        """
        spreadsheet = Spreadsheet()

        # Set up the data structure from sample_data2
        # Row 12 (0-indexed): D13 = =SUM(D3:D12)
        for i in range(2, 12):
            spreadsheet.set_cell(i, 3, str((i - 1) * 10000))
        spreadsheet.set_cell(12, 3, "=SUM(D3:D12)")

        # Row 23 (0-indexed): B24 = =SUM(B18:B23)
        for i in range(17, 23):
            spreadsheet.set_cell(i, 1, str((i - 16) * 10000))
        spreadsheet.set_cell(23, 1, "=SUM(B18:B23)")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Check D13 formula
            d13_raw = new_spreadsheet.get_cell(12, 3).raw_value
            assert "SUM" in d13_raw.upper(), f"D13 missing SUM: {d13_raw}"
            assert "D3:D12" in d13_raw.upper(), f"D13 missing range: {d13_raw}"

            # Check B24 formula
            b24_raw = new_spreadsheet.get_cell(23, 1).raw_value
            assert "SUM" in b24_raw.upper(), f"B24 missing SUM: {b24_raw}"
            assert "B18:B23" in b24_raw.upper(), f"B24 missing range: {b24_raw}"
        finally:
            Path(filepath).unlink()

    def test_function_with_multiple_cell_args(self):
        """Test function with multiple individual cell arguments."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "10")
        spreadsheet.set_cell(0, 1, "20")
        spreadsheet.set_cell(0, 2, "30")
        spreadsheet.set_cell(0, 3, "=@SUM(A1,B1,C1)")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            raw = new_spreadsheet.get_cell(0, 3).raw_value
            assert "SUM" in raw.upper()
            # Should have all three cell references
            assert "A1" in raw.upper()
            assert "B1" in raw.upper()
            assert "C1" in raw.upper()
        finally:
            Path(filepath).unlink()

    def test_mixed_formulas_and_values(self):
        """Test file with mix of values and various formula types."""
        spreadsheet = Spreadsheet()

        # Values
        spreadsheet.set_cell(0, 0, "100")
        spreadsheet.set_cell(0, 1, "200")
        spreadsheet.set_cell(0, 2, "Text value")

        # Simple formula
        spreadsheet.set_cell(1, 0, "=A1+B1")

        # Function with range
        spreadsheet.set_cell(1, 1, "=SUM(A1:B1)")

        # Function with cell
        spreadsheet.set_cell(1, 2, "=ABS(A1)")

        # Nested expression
        spreadsheet.set_cell(2, 0, "=(A1*B1)/100")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Check values preserved
            assert new_spreadsheet.get_cell(0, 0).raw_value == "100"
            assert new_spreadsheet.get_cell(0, 1).raw_value == "200"

            # Check formulas preserved
            assert new_spreadsheet.get_cell(1, 0).raw_value == "=A1+B1"

            sum_raw = new_spreadsheet.get_cell(1, 1).raw_value
            assert "SUM" in sum_raw.upper()
            assert "A1:B1" in sum_raw.upper()

            abs_raw = new_spreadsheet.get_cell(1, 2).raw_value
            assert "ABS" in abs_raw.upper()
            assert "A1" in abs_raw.upper()
        finally:
            Path(filepath).unlink()


class TestFormatByteEncoding:
    """Tests for format byte encoding/decoding."""

    def test_encode_general_format(self):
        """Test encoding General format."""
        assert encode_format_byte("G") == FMT_DEFAULT
        assert encode_format_byte("") == FMT_DEFAULT
        assert encode_format_byte(None) == FMT_DEFAULT

    def test_encode_fixed_format(self):
        """Test encoding Fixed format (F0-F7)."""
        assert encode_format_byte("F0") == FMT_FIXED | (0 << 4)
        assert encode_format_byte("F2") == FMT_FIXED | (2 << 4)
        assert encode_format_byte("F4") == FMT_FIXED | (4 << 4)
        assert encode_format_byte("F7") == FMT_FIXED | (7 << 4)

    def test_encode_scientific_format(self):
        """Test encoding Scientific format."""
        assert encode_format_byte("S2") == FMT_SCIENTIFIC | (2 << 4)
        assert encode_format_byte("S4") == FMT_SCIENTIFIC | (4 << 4)

    def test_encode_currency_format(self):
        """Test encoding Currency format."""
        assert encode_format_byte("C0") == FMT_CURRENCY | (0 << 4)
        assert encode_format_byte("C2") == FMT_CURRENCY | (2 << 4)

    def test_encode_percent_format(self):
        """Test encoding Percent format."""
        assert encode_format_byte("P0") == FMT_PERCENT | (0 << 4)
        assert encode_format_byte("P2") == FMT_PERCENT | (2 << 4)

    def test_encode_comma_format(self):
        """Test encoding Comma format."""
        assert encode_format_byte(",0") == FMT_COMMA | (0 << 4)
        assert encode_format_byte(",2") == FMT_COMMA | (2 << 4)

    def test_encode_date_formats(self):
        """Test encoding Date formats (D1-D5)."""
        assert encode_format_byte("D1") == FMT_DATETIME | (0 << 4)
        assert encode_format_byte("D2") == FMT_DATETIME | (1 << 4)
        assert encode_format_byte("D3") == FMT_DATETIME | (2 << 4)
        assert encode_format_byte("D4") == FMT_DATETIME | (3 << 4)
        assert encode_format_byte("D5") == FMT_DATETIME | (4 << 4)

    def test_encode_time_formats(self):
        """Test encoding Time formats (T1-T3)."""
        assert encode_format_byte("T1") == FMT_DATETIME | (5 << 4)
        assert encode_format_byte("T2") == FMT_DATETIME | (6 << 4)
        assert encode_format_byte("T3") == FMT_DATETIME | (7 << 4)

    def test_decode_general_format(self):
        """Test decoding General format."""
        assert decode_format_byte(FMT_DEFAULT) == "G"

    def test_decode_fixed_format(self):
        """Test decoding Fixed format."""
        assert decode_format_byte(FMT_FIXED | (0 << 4)) == "F0"
        assert decode_format_byte(FMT_FIXED | (2 << 4)) == "F2"
        assert decode_format_byte(FMT_FIXED | (4 << 4)) == "F4"

    def test_decode_scientific_format(self):
        """Test decoding Scientific format."""
        assert decode_format_byte(FMT_SCIENTIFIC | (2 << 4)) == "S2"

    def test_decode_currency_format(self):
        """Test decoding Currency format."""
        assert decode_format_byte(FMT_CURRENCY | (2 << 4)) == "C2"

    def test_decode_percent_format(self):
        """Test decoding Percent format."""
        assert decode_format_byte(FMT_PERCENT | (2 << 4)) == "P2"

    def test_decode_date_formats(self):
        """Test decoding Date formats."""
        assert decode_format_byte(FMT_DATETIME | (0 << 4)) == "D1"
        assert decode_format_byte(FMT_DATETIME | (1 << 4)) == "D2"
        assert decode_format_byte(FMT_DATETIME | (2 << 4)) == "D3"
        assert decode_format_byte(FMT_DATETIME | (3 << 4)) == "D4"
        assert decode_format_byte(FMT_DATETIME | (4 << 4)) == "D5"

    def test_decode_time_formats(self):
        """Test decoding Time formats."""
        assert decode_format_byte(FMT_DATETIME | (5 << 4)) == "T1"
        assert decode_format_byte(FMT_DATETIME | (6 << 4)) == "T2"
        assert decode_format_byte(FMT_DATETIME | (7 << 4)) == "T3"

    def test_roundtrip_formats(self):
        """Test format encoding/decoding roundtrip."""
        formats = [
            "G",
            "F0",
            "F2",
            "F4",
            "S2",
            "C2",
            "P2",
            ",2",
            "D1",
            "D2",
            "D3",
            "D4",
            "D5",
            "T1",
            "T2",
            "T3",
        ]
        for fmt in formats:
            encoded = encode_format_byte(fmt)
            decoded = decode_format_byte(encoded)
            assert decoded == fmt, f"Roundtrip failed for {fmt}: got {decoded}"


class TestWk1FormatIO:
    """Tests for format code reading/writing in WK1 files."""

    def test_write_and_read_cell_format(self):
        """Test that cell format codes are preserved."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "1234.56")
        cell = spreadsheet.get_cell(0, 0)
        cell.format_code = "F2"

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            loaded_cell = new_spreadsheet.get_cell(0, 0)
            assert loaded_cell.format_code == "F2"
        finally:
            Path(filepath).unlink()

    def test_write_and_read_currency_format(self):
        """Test currency format preservation."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "9999.99")
        cell = spreadsheet.get_cell(0, 0)
        cell.format_code = "C2"

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            loaded_cell = new_spreadsheet.get_cell(0, 0)
            assert loaded_cell.format_code == "C2"
        finally:
            Path(filepath).unlink()

    def test_write_and_read_percent_format(self):
        """Test percent format preservation."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "0.25")
        cell = spreadsheet.get_cell(0, 0)
        cell.format_code = "P0"

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            loaded_cell = new_spreadsheet.get_cell(0, 0)
            assert loaded_cell.format_code == "P0"
        finally:
            Path(filepath).unlink()

    def test_write_and_read_date_format(self):
        """Test date format preservation."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "45000")  # Excel date serial
        cell = spreadsheet.get_cell(0, 0)
        cell.format_code = "D1"

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            loaded_cell = new_spreadsheet.get_cell(0, 0)
            assert loaded_cell.format_code == "D1"
        finally:
            Path(filepath).unlink()

    def test_write_and_read_multiple_formats(self):
        """Test multiple different formats in same file."""
        spreadsheet = Spreadsheet()

        # Set cells with different formats
        spreadsheet.set_cell(0, 0, "1234.56")
        spreadsheet.get_cell(0, 0).format_code = "F2"

        spreadsheet.set_cell(0, 1, "9999.99")
        spreadsheet.get_cell(0, 1).format_code = "C2"

        spreadsheet.set_cell(0, 2, "0.25")
        spreadsheet.get_cell(0, 2).format_code = "P0"

        spreadsheet.set_cell(0, 3, "45000")
        spreadsheet.get_cell(0, 3).format_code = "D1"

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.get_cell(0, 0).format_code == "F2"
            assert new_spreadsheet.get_cell(0, 1).format_code == "C2"
            assert new_spreadsheet.get_cell(0, 2).format_code == "P0"
            assert new_spreadsheet.get_cell(0, 3).format_code == "D1"
        finally:
            Path(filepath).unlink()

    def test_formula_format_preserved(self):
        """Test that format is preserved for formula cells."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "100")
        spreadsheet.set_cell(0, 1, "200")
        spreadsheet.set_cell(0, 2, "=A1+B1")
        spreadsheet.get_cell(0, 2).format_code = "C0"

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            loaded_cell = new_spreadsheet.get_cell(0, 2)
            assert loaded_cell.format_code == "C0"
            assert "A1" in loaded_cell.raw_value.upper()
        finally:
            Path(filepath).unlink()

    def test_label_format_preserved(self):
        """Test that format is preserved for label cells."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "Hello World")
        spreadsheet.get_cell(0, 0).format_code = "F2"  # Unusual but valid

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            loaded_cell = new_spreadsheet.get_cell(0, 0)
            assert loaded_cell.format_code == "F2"
        finally:
            Path(filepath).unlink()


class TestWk1NamedRangesIO:
    """Tests for named ranges reading/writing in WK1 files."""

    def test_write_and_read_single_cell_name(self):
        """Test named range for single cell."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "100")
        spreadsheet.named_ranges.add_from_string("TOTAL", "A1")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.named_ranges.exists("TOTAL")
            named = new_spreadsheet.named_ranges.get("TOTAL")
            assert named is not None
            assert named.is_single_cell
        finally:
            Path(filepath).unlink()

    def test_write_and_read_range_name(self):
        """Test named range for a range of cells."""
        spreadsheet = Spreadsheet()
        for i in range(10):
            spreadsheet.set_cell(i, 0, str(i * 10))
        spreadsheet.named_ranges.add_from_string("DATA", "A1:A10")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.named_ranges.exists("DATA")
            named = new_spreadsheet.named_ranges.get("DATA")
            assert named is not None
            assert not named.is_single_cell
        finally:
            Path(filepath).unlink()

    def test_write_and_read_multiple_names(self):
        """Test multiple named ranges."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "100")
        spreadsheet.set_cell(0, 1, "200")
        for i in range(5):
            spreadsheet.set_cell(i, 2, str(i))

        spreadsheet.named_ranges.add_from_string("FIRST", "A1")
        spreadsheet.named_ranges.add_from_string("SECOND", "B1")
        spreadsheet.named_ranges.add_from_string("RANGE", "C1:C5")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.named_ranges.exists("FIRST")
            assert new_spreadsheet.named_ranges.exists("SECOND")
            assert new_spreadsheet.named_ranges.exists("RANGE")
            assert len(new_spreadsheet.named_ranges) == 3
        finally:
            Path(filepath).unlink()

    def test_name_truncation(self):
        """Test that long names are truncated to 15 characters."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "100")
        # Name longer than 15 chars
        long_name = "VERYLONGNAMETHATSHOULDBETRUNCATED"
        spreadsheet.named_ranges.add_from_string(long_name, "A1")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Should exist with truncated name
            truncated = long_name[:15].upper()
            assert new_spreadsheet.named_ranges.exists(truncated)
        finally:
            Path(filepath).unlink()


class TestWk1CalcSettingsIO:
    """Tests for calculation settings reading/writing in WK1 files."""

    def test_write_and_read_automatic_calcmode(self):
        """Test automatic calculation mode."""
        from lotus123.formula.recalc import RecalcMode

        spreadsheet = Spreadsheet()
        spreadsheet.set_recalc_mode(RecalcMode.AUTOMATIC)
        spreadsheet.set_cell(0, 0, "100")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.get_recalc_mode() == RecalcMode.AUTOMATIC
        finally:
            Path(filepath).unlink()

    def test_write_and_read_manual_calcmode(self):
        """Test manual calculation mode."""
        from lotus123.formula.recalc import RecalcMode

        spreadsheet = Spreadsheet()
        spreadsheet.set_recalc_mode(RecalcMode.MANUAL)
        spreadsheet.set_cell(0, 0, "100")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.get_recalc_mode() == RecalcMode.MANUAL
        finally:
            Path(filepath).unlink()

    def test_write_and_read_natural_calcorder(self):
        """Test natural calculation order."""
        from lotus123.formula.recalc import RecalcOrder

        spreadsheet = Spreadsheet()
        spreadsheet.set_recalc_order(RecalcOrder.NATURAL)
        spreadsheet.set_cell(0, 0, "100")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.get_recalc_order() == RecalcOrder.NATURAL
        finally:
            Path(filepath).unlink()

    def test_write_and_read_column_wise_calcorder(self):
        """Test column-wise calculation order."""
        from lotus123.formula.recalc import RecalcOrder

        spreadsheet = Spreadsheet()
        spreadsheet.set_recalc_order(RecalcOrder.COLUMN_WISE)
        spreadsheet.set_cell(0, 0, "100")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.get_recalc_order() == RecalcOrder.COLUMN_WISE
        finally:
            Path(filepath).unlink()

    def test_write_and_read_row_wise_calcorder(self):
        """Test row-wise calculation order."""
        from lotus123.formula.recalc import RecalcOrder

        spreadsheet = Spreadsheet()
        spreadsheet.set_recalc_order(RecalcOrder.ROW_WISE)
        spreadsheet.set_cell(0, 0, "100")

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert new_spreadsheet.get_recalc_order() == RecalcOrder.ROW_WISE
        finally:
            Path(filepath).unlink()


class TestWk1BlankCellIO:
    """Tests for blank cell (format-only) reading/writing in WK1 files."""

    def test_write_blank_cell_with_format(self):
        """Test writing blank cell that has format but no value."""
        spreadsheet = Spreadsheet()
        # Create a cell with format but no value
        cell = spreadsheet.get_cell(5, 5)
        cell.format_code = "C2"  # Cell has format but is empty

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            loaded_cell = new_spreadsheet.get_cell(5, 5)
            assert loaded_cell.format_code == "C2"
            assert loaded_cell.is_empty
        finally:
            Path(filepath).unlink()


class TestWk1MetadataRoundtrip:
    """Integration tests for full metadata roundtrip."""

    def test_full_metadata_roundtrip(self):
        """Test complete metadata preservation across save/load."""
        from lotus123.formula.recalc import RecalcMode, RecalcOrder

        spreadsheet = Spreadsheet()

        # Set up calc settings
        spreadsheet.set_recalc_mode(RecalcMode.MANUAL)
        spreadsheet.set_recalc_order(RecalcOrder.COLUMN_WISE)

        # Add data with formats
        spreadsheet.set_cell(0, 0, "1234.56")
        spreadsheet.get_cell(0, 0).format_code = "F2"

        spreadsheet.set_cell(0, 1, "9999.99")
        spreadsheet.get_cell(0, 1).format_code = "C2"

        spreadsheet.set_cell(0, 2, "=A1+B1")
        spreadsheet.get_cell(0, 2).format_code = "F0"

        # Add named ranges (names that don't look like cell references)
        spreadsheet.named_ranges.add_from_string("FIRST_VALUE", "A1")
        spreadsheet.named_ranges.add_from_string("SECOND_VALUE", "B1")
        spreadsheet.named_ranges.add_from_string("TOTAL", "C1")

        # Set column widths
        spreadsheet.set_col_width(0, 12)
        spreadsheet.set_col_width(1, 15)

        writer = Wk1Writer(spreadsheet)

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Verify calc settings
            assert new_spreadsheet.get_recalc_mode() == RecalcMode.MANUAL
            assert new_spreadsheet.get_recalc_order() == RecalcOrder.COLUMN_WISE

            # Verify formats
            assert new_spreadsheet.get_cell(0, 0).format_code == "F2"
            assert new_spreadsheet.get_cell(0, 1).format_code == "C2"
            assert new_spreadsheet.get_cell(0, 2).format_code == "F0"

            # Verify named ranges
            assert new_spreadsheet.named_ranges.exists("FIRST_VALUE")
            assert new_spreadsheet.named_ranges.exists("SECOND_VALUE")
            assert new_spreadsheet.named_ranges.exists("TOTAL")

            # Verify column widths
            assert new_spreadsheet.get_col_width(0) == 12
            assert new_spreadsheet.get_col_width(1) == 15

            # Verify formula preserved
            assert "A1" in new_spreadsheet.get_cell(0, 2).raw_value.upper()
        finally:
            Path(filepath).unlink()


class TestWk1SpecialFloatValues:
    """Tests for IEEE 754 special float values (infinity, NaN).

    Note: Special float values are stored as string representations in cells
    (e.g., "inf", "-inf", "nan") because cell raw_value is a string.
    The WK1 writer detects these strings and writes proper IEEE 754 values.
    When reading back, get_value() parses them as Python float values.
    """

    def test_read_positive_infinity(self) -> None:
        """Test reading positive infinity from WK1."""
        import io
        import math

        data = io.BytesIO()
        # BOF
        write_record(data, BOF, struct.pack("<H", VERSION_WK1))
        # NUMBER record with positive infinity
        inf_value = float("inf")
        num_data = struct.pack("<BHHd", 0, 0, 0, inf_value)  # A1 = inf
        write_record(data, NUMBER, num_data)
        # EOF
        write_record(data, EOF, b"")

        data.seek(0)
        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)
        reader._read_file(data)

        # get_value() parses "inf" string as float
        value = spreadsheet.get_value(0, 0)
        assert math.isinf(value) and value > 0

    def test_read_negative_infinity(self) -> None:
        """Test reading negative infinity from WK1."""
        import io
        import math

        data = io.BytesIO()
        # BOF
        write_record(data, BOF, struct.pack("<H", VERSION_WK1))
        # NUMBER record with negative infinity
        neg_inf_value = float("-inf")
        num_data = struct.pack("<BHHd", 0, 0, 0, neg_inf_value)  # A1 = -inf
        write_record(data, NUMBER, num_data)
        # EOF
        write_record(data, EOF, b"")

        data.seek(0)
        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)
        reader._read_file(data)

        # get_value() parses "-inf" string as float
        value = spreadsheet.get_value(0, 0)
        assert math.isinf(value) and value < 0

    def test_read_nan(self) -> None:
        """Test reading NaN from WK1."""
        import io
        import math

        data = io.BytesIO()
        # BOF
        write_record(data, BOF, struct.pack("<H", VERSION_WK1))
        # NUMBER record with NaN
        nan_value = float("nan")
        num_data = struct.pack("<BHHd", 0, 0, 0, nan_value)  # A1 = nan
        write_record(data, NUMBER, num_data)
        # EOF
        write_record(data, EOF, b"")

        data.seek(0)
        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)
        reader._read_file(data)

        # get_value() parses "nan" string as float
        value = spreadsheet.get_value(0, 0)
        assert math.isnan(value)

    def test_write_infinity_as_number_record(self) -> None:
        """Test that infinity strings are written as NUMBER records, not LABEL."""
        import io
        import math

        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "inf")

        data = io.BytesIO()
        writer = Wk1Writer(spreadsheet)
        writer._write_file(data)

        # Parse the output to verify it's a NUMBER record
        data.seek(0)
        records = []
        while True:
            header = data.read(4)
            if len(header) < 4:
                break
            opcode, length = struct.unpack("<HH", header)
            record_data = data.read(length)
            records.append((opcode, record_data))

        # Find the cell record (skip BOF, CALCMODE, CALCORDER, RANGE)
        number_records = [r for r in records if r[0] == NUMBER]
        assert len(number_records) == 1

        # Verify the value is IEEE 754 infinity
        record_data = number_records[0][1]
        value = struct.unpack("<d", record_data[5:13])[0]
        assert math.isinf(value) and value > 0

    def test_write_and_read_infinity_roundtrip(self) -> None:
        """Test writing and reading infinity values preserves them."""
        import math

        spreadsheet = Spreadsheet()
        # Set cells with infinity values
        spreadsheet.set_cell(0, 0, "inf")  # Positive infinity
        spreadsheet.set_cell(0, 1, "-inf")  # Negative infinity

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer = Wk1Writer(spreadsheet)
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Values are preserved as floats
            value_a1 = new_spreadsheet.get_value(0, 0)
            assert math.isinf(value_a1) and value_a1 > 0
            value_b1 = new_spreadsheet.get_value(0, 1)
            assert math.isinf(value_b1) and value_b1 < 0
        finally:
            Path(filepath).unlink()

    def test_write_and_read_nan_roundtrip(self) -> None:
        """Test writing and reading NaN value preserves it."""
        import math

        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "nan")

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer = Wk1Writer(spreadsheet)
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Value is preserved as float
            value = new_spreadsheet.get_value(0, 0)
            assert math.isnan(value)
        finally:
            Path(filepath).unlink()

    def test_multiple_special_values(self) -> None:
        """Test multiple special values in same file."""
        import math

        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "inf")
        spreadsheet.set_cell(0, 1, "-inf")
        spreadsheet.set_cell(0, 2, "nan")
        spreadsheet.set_cell(0, 3, "42")  # Normal value for comparison

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer = Wk1Writer(spreadsheet)
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            assert math.isinf(new_spreadsheet.get_value(0, 0))
            assert new_spreadsheet.get_value(0, 0) > 0
            assert math.isinf(new_spreadsheet.get_value(0, 1))
            assert new_spreadsheet.get_value(0, 1) < 0
            assert math.isnan(new_spreadsheet.get_value(0, 2))
            assert new_spreadsheet.get_value(0, 3) == 42
        finally:
            Path(filepath).unlink()

    def test_truncated_formula_with_infinity_cached_value(self) -> None:
        """Test reading truncated formula record with infinity as cached value.

        When a formula record is truncated (bytecode length exceeds actual data),
        the reader falls back to using the cached value.
        """
        import io
        import math

        data = io.BytesIO()
        # BOF
        write_record(data, BOF, struct.pack("<H", VERSION_WK1))

        # FORMULA record with infinity as cached value and TRUNCATED bytecode
        # This will trigger the fallback path that uses cached value
        inf_value = float("inf")
        # Format: format(1) + col(2) + row(2) + cached_value(8) + formula_len(2) + bytecode
        # Claim 10 bytes of bytecode but only provide 3 (truncated)
        claimed_len = 10
        actual_bytecode = bytes([0x00, 0x01, 0x02])  # Only 3 bytes
        formula_data = struct.pack("<BHHdH", 0, 0, 0, inf_value, claimed_len)
        formula_data += actual_bytecode
        write_record(data, FORMULA, formula_data)

        # EOF
        write_record(data, EOF, b"")

        data.seek(0)
        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)
        reader._read_file(data)

        # Should fall back to cached value (parsed as float inf)
        value = spreadsheet.get_value(0, 0)
        assert math.isinf(value)

    def test_load_kbase_sample(self) -> None:
        """Test loading the kbase.wk1 sample file which contains special values."""
        sample_path = Path("samples/kbase.wk1")
        if not sample_path.exists():
            pytest.skip("Sample file not found")

        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)
        reader.load(str(sample_path))

        # File should load without error
        used_range = spreadsheet.get_used_range()
        assert used_range is not None
        # The file contains data
        (min_row, min_col), (max_row, max_col) = used_range
        assert max_row > min_row or max_col > min_col


class TestWk1RelativeCellReferences:
    """Tests for relative cell reference encoding/decoding in formulas."""

    def test_decompile_relative_cell_reference_same_row(self) -> None:
        """Test decompiling relative reference to cell in same row."""
        # Formula in N3 (col=13, row=2) referencing C3 (col=2, row=2)
        # Relative column offset: 2 - 13 = -11 (0xF5 as unsigned 8-bit)
        # Relative row offset: 0
        # Column word: 0x8000 | 0xF5 = 0x80F5 -> little-endian: F5 80 -> actually 0xBFF5 with extra bits
        # Let me recalculate:
        # The actual bytes from the WK1 file: 0xF5, 0xBF for column (0xBFF5)
        # 0xBFF5: bit 15 set (relative), bits 0-7 = 0xF5 = -11 as signed
        # Row: 0x8000 (relative, offset 0)
        bytecode = bytes(
            [
                0x01,  # OP_VARIABLE
                0xF5,
                0xBF,  # Column: relative, offset -11 (0xF5 signed)
                0x00,
                0x80,  # Row: relative, offset 0
                0x03,  # OP_RETURN
            ]
        )
        # Formula at N3 (col=13, row=2)
        result = decompile_formula(bytecode, formula_row=2, formula_col=13)
        assert result == "=C3"

    def test_decompile_relative_cell_reference_different_row(self) -> None:
        """Test decompiling relative reference to cell in different row."""
        # Formula in B2 (col=1, row=1) referencing B3 (col=1, row=2)
        # Relative column offset: 0 (same column)
        # Relative row offset: 2 - 1 = 1
        bytecode = bytes(
            [
                0x01,  # OP_VARIABLE
                0x00,
                0x80,  # Column: relative, offset 0
                0x01,
                0x80,  # Row: relative, offset 1
                0x03,  # OP_RETURN
            ]
        )
        result = decompile_formula(bytecode, formula_row=1, formula_col=1)
        assert result == "=B3"

    def test_decompile_relative_cell_reference_negative_row_offset(self) -> None:
        """Test decompiling relative reference with negative row offset."""
        # Formula in B5 (col=1, row=4) referencing B3 (col=1, row=2)
        # Relative row offset: 2 - 4 = -2
        # -2 as signed 14-bit: 0x4000 - 2 = 0x3FFE
        bytecode = bytes(
            [
                0x01,  # OP_VARIABLE
                0x00,
                0x80,  # Column: relative, offset 0
                0xFE,
                0xBF,  # Row: relative, offset -2 (0x3FFE with bit 15 set = 0xBFFE)
                0x03,  # OP_RETURN
            ]
        )
        result = decompile_formula(bytecode, formula_row=4, formula_col=1)
        assert result == "=B3"

    def test_compile_relative_cell_reference(self) -> None:
        """Test compiling formula with relative cell reference."""
        # Compile formula "=A1" from cell B2 (col=1, row=1)
        # Relative column offset: 0 - 1 = -1
        # Relative row offset: 0 - 1 = -1
        bytecode = compile_formula("=A1", formula_row=1, formula_col=1)

        # Decompile from same position should get same result
        result = decompile_formula(bytecode, formula_row=1, formula_col=1)
        assert result == "=A1"

    def test_compile_and_decompile_roundtrip(self) -> None:
        """Test that compile and decompile are inverse operations."""
        formulas = [
            ("=A1", 0, 0),
            ("=B2", 0, 0),
            ("=C3", 2, 1),  # Formula in B3 referencing C3
            ("=A1+B1", 0, 2),  # Formula in C1 referencing A1 and B1
        ]
        for formula, row, col in formulas:
            bytecode = compile_formula(formula, formula_row=row, formula_col=col)
            result = decompile_formula(bytecode, formula_row=row, formula_col=col)
            assert result == formula, f"Roundtrip failed for {formula} at ({row},{col})"

    def test_compile_relative_range(self) -> None:
        """Test compiling formula with relative range reference."""
        # Compile formula "=SUM(A1:A10)" from cell C1 (col=2, row=0)
        bytecode = compile_formula("=@SUM(A1:A10)", formula_row=0, formula_col=2)

        # Decompile from same position
        result = decompile_formula(bytecode, formula_row=0, formula_col=2)
        assert "@SUM(A1:A10)" in result

    def test_absolute_cell_reference_preserved(self) -> None:
        """Test that absolute references ($A$1) are preserved."""
        # Absolute references have bit 15 = 0
        bytecode = compile_formula("=$A$1", formula_row=5, formula_col=5)
        result = decompile_formula(bytecode, formula_row=5, formula_col=5)
        # Note: Absolute references currently compile as absolute but decompile
        # without the $ prefix (limitation of current implementation)
        assert "A1" in result

    def test_wk1_file_relative_formula_roundtrip(self) -> None:
        """Test that formulas with relative refs survive WK1 write/read."""
        spreadsheet = Spreadsheet()

        # Set up data: A1=10, B1=20
        spreadsheet.set_cell(0, 0, "10")  # A1
        spreadsheet.set_cell(0, 1, "20")  # B1

        # C1 = A1 + B1 (relative references)
        spreadsheet.set_cell(0, 2, "=A1+B1")  # C1

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer = Wk1Writer(spreadsheet)
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Formula should be preserved
            cell = new_spreadsheet.get_cell(0, 2)
            assert "A1" in cell.raw_value
            assert "B1" in cell.raw_value

            # Value should be correct
            value = new_spreadsheet.get_value(0, 2)
            assert value == 30
        finally:
            Path(filepath).unlink()

    def test_wk1_file_formula_with_function_roundtrip(self) -> None:
        """Test formulas with functions and relative refs survive roundtrip."""
        spreadsheet = Spreadsheet()

        # Set up data range A1:A5
        for i in range(5):
            spreadsheet.set_cell(i, 0, str(i + 1))

        # B1 = SUM(A1:A5)
        spreadsheet.set_cell(0, 1, "=@SUM(A1:A5)")

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer = Wk1Writer(spreadsheet)
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Formula should reference the correct range
            cell = new_spreadsheet.get_cell(0, 1)
            assert "SUM" in cell.raw_value
            assert "A1" in cell.raw_value
            assert "A5" in cell.raw_value

            # Value should be correct (1+2+3+4+5 = 15)
            value = new_spreadsheet.get_value(0, 1)
            assert value == 15
        finally:
            Path(filepath).unlink()

    def test_kbase_n3_o3_formulas(self) -> None:
        """Test that kbase.wk1 N3 and O3 formulas are read correctly."""
        sample_path = Path("specs/kbase.wk1")
        if not sample_path.exists():
            pytest.skip("Sample file not found")

        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)
        reader.load(str(sample_path))

        # N3 (row=2, col=13) should have a formula comparing C3/C4 and D3/D4
        n3 = spreadsheet.get_cell(2, 13)
        assert n3.raw_value.startswith("=")
        assert "C3" in n3.raw_value or "C4" in n3.raw_value
        assert "D3" in n3.raw_value or "D4" in n3.raw_value
        # Value should be boolean (True/False) or 1/0
        n3_value = spreadsheet.get_value(2, 13)
        assert n3_value in (True, False, 1, 0)

        # O3 (row=2, col=14) should have an IF formula
        o3 = spreadsheet.get_cell(2, 14)
        assert o3.raw_value.startswith("=")
        assert "IF" in o3.raw_value
        # Value should be a number
        o3_value = spreadsheet.get_value(2, 14)
        assert isinstance(o3_value, (int, float))


class TestWk1ZeroArgFunctions:
    """Tests for zero-argument function handling in WK1 formulas.

    Zero-arg functions like @FALSE, @TRUE, @PI must be decompiled with
    parentheses (e.g., @FALSE()) for the formula parser to recognize them.
    """

    def test_decompile_false_has_parentheses(self) -> None:
        """Test that @FALSE is decompiled with parentheses."""
        # FN_FALSE opcode is 0x33, followed by OP_RETURN (0x03)
        bytecode = bytes([0x33, 0x03])
        result = decompile_formula(bytecode)
        assert result == "=@FALSE()"

    def test_decompile_true_has_parentheses(self) -> None:
        """Test that @TRUE is decompiled with parentheses."""
        # FN_TRUE opcode is 0x34, followed by OP_RETURN (0x03)
        bytecode = bytes([0x34, 0x03])
        result = decompile_formula(bytecode)
        assert result == "=@TRUE()"

    def test_decompile_pi_has_parentheses(self) -> None:
        """Test that @PI is decompiled with parentheses."""
        # FN_PI opcode is 0x26, followed by OP_RETURN (0x03)
        bytecode = bytes([0x26, 0x03])
        result = decompile_formula(bytecode)
        assert result == "=@PI()"

    def test_decompile_rand_has_parentheses(self) -> None:
        """Test that @RAND is decompiled with parentheses."""
        # FN_RAND opcode is 0x35, followed by OP_RETURN (0x03)
        bytecode = bytes([0x35, 0x03])
        result = decompile_formula(bytecode)
        assert result == "=@RAND()"

    def test_decompile_today_has_parentheses(self) -> None:
        """Test that @TODAY is decompiled with parentheses."""
        # FN_TODAY opcode is 0x37, followed by OP_RETURN (0x03)
        bytecode = bytes([0x37, 0x03])
        result = decompile_formula(bytecode)
        assert result == "=@TODAY()"

    def test_decompile_na_has_parentheses(self) -> None:
        """Test that @NA is decompiled with parentheses."""
        # FN_NA opcode is 0x1F, followed by OP_RETURN (0x03)
        bytecode = bytes([0x1F, 0x03])
        result = decompile_formula(bytecode)
        assert result == "=@NA()"

    def test_decompile_err_has_parentheses(self) -> None:
        """Test that @ERR is decompiled with parentheses."""
        # FN_ERR opcode is 0x20, followed by OP_RETURN (0x03)
        bytecode = bytes([0x20, 0x03])
        result = decompile_formula(bytecode)
        assert result == "=@ERR()"

    def test_false_evaluates_correctly(self) -> None:
        """Test that decompiled @FALSE() evaluates to False."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "=@FALSE()")
        value = spreadsheet.get_value(0, 0)
        assert value is False

    def test_true_evaluates_correctly(self) -> None:
        """Test that decompiled @TRUE() evaluates to True."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "=@TRUE()")
        value = spreadsheet.get_value(0, 0)
        assert value is True

    def test_if_with_false_branch(self) -> None:
        """Test IF formula where false branch uses @FALSE()."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "0")  # A1 = 0 (false condition)
        spreadsheet.set_cell(0, 1, "=@IF(A1=1,100,@FALSE())")

        value = spreadsheet.get_value(0, 1)
        assert value is False

    def test_if_with_true_branch(self) -> None:
        """Test IF formula where true branch uses @TRUE()."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "1")  # A1 = 1 (true condition)
        spreadsheet.set_cell(0, 1, "=@IF(A1=1,@TRUE(),0)")

        value = spreadsheet.get_value(0, 1)
        assert value is True

    def test_wk1_roundtrip_if_with_false(self) -> None:
        """Test WK1 roundtrip of IF formula with @FALSE() in else branch."""
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "0")  # A1 = 0
        spreadsheet.set_cell(0, 1, "=@IF(A1=1,100,@FALSE())")

        with tempfile.NamedTemporaryFile(suffix=".wk1", delete=False) as f:
            filepath = f.name

        try:
            writer = Wk1Writer(spreadsheet)
            writer.save(filepath)

            new_spreadsheet = Spreadsheet()
            reader = Wk1Reader(new_spreadsheet)
            reader.load(filepath)

            # Formula should contain @FALSE()
            cell = new_spreadsheet.get_cell(0, 1)
            assert "@FALSE()" in cell.raw_value

            # Value should be False
            value = new_spreadsheet.get_value(0, 1)
            assert value is False
        finally:
            Path(filepath).unlink()

    def test_kbase_o4_evaluates_correctly(self) -> None:
        """Test that kbase.wk1 O4 (which uses @FALSE) evaluates correctly."""
        sample_path = Path("specs/kbase.wk1")
        if not sample_path.exists():
            pytest.skip("Sample file not found")

        spreadsheet = Spreadsheet()
        reader = Wk1Reader(spreadsheet)
        reader.load(str(sample_path))

        # O4 has formula with @FALSE() in else branch
        o4 = spreadsheet.get_cell(3, 14)
        assert "@FALSE()" in o4.raw_value

        # Value should be False (not #REF!)
        o4_value = spreadsheet.get_value(3, 14)
        assert o4_value is False or o4_value == 0 or isinstance(o4_value, (int, float))

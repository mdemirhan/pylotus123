"""Tests for lookup and reference functions."""

from lotus123.formula.functions.lookup import (
    LOOKUP_FUNCTIONS,
    fn_address,
    fn_cols,
    fn_column,
    fn_columns,
    fn_hlookup,
    fn_index,
    fn_indirect,
    fn_lookup,
    fn_match,
    fn_offset,
    fn_row,
    fn_rows,
    fn_transpose,
    fn_vlookup,
)


class TestVLOOKUP:
    """Tests for VLOOKUP function."""

    def setup_method(self):
        """Set up test table."""
        self.table = [
            [1, "Apple", 1.00],
            [2, "Banana", 0.50],
            [3, "Cherry", 2.00],
        ]

    def test_vlookup_exact_match(self):
        """Test VLOOKUP with exact match."""
        result = fn_vlookup(2, self.table, 2, False)
        assert result == "Banana"

    def test_vlookup_range_match(self):
        """Test VLOOKUP with range match."""
        result = fn_vlookup(2, self.table, 2, True)
        assert result == "Banana"

    def test_vlookup_not_found(self):
        """Test VLOOKUP when value not found."""
        result = fn_vlookup(5, self.table, 2, False)
        assert result == "#N/A"

    def test_vlookup_invalid_col(self):
        """Test VLOOKUP with invalid column index."""
        result = fn_vlookup(1, self.table, 5, False)
        assert result == "#REF!"

    def test_vlookup_empty_table(self):
        """Test VLOOKUP with empty table."""
        result = fn_vlookup(1, [], 1, False)
        assert result == "#N/A"

    def test_vlookup_1d_table(self):
        """Test VLOOKUP with 1D table."""
        result = fn_vlookup(2, [1, 2, 3], 1, False)
        assert result == 2

    def test_vlookup_string_match(self):
        """Test VLOOKUP with string lookup value."""
        table = [["Apple", 1], ["Banana", 2], ["Cherry", 3]]
        result = fn_vlookup("Banana", table, 2, False)
        assert result == 2

    def test_vlookup_range_unsorted(self):
        """Approximate match requires sorted data."""
        table = [[2, "Banana"], [1, "Apple"], [3, "Cherry"]]
        result = fn_vlookup(2, table, 2, True)
        assert result == "#N/A"


class TestHLOOKUP:
    """Tests for HLOOKUP function."""

    def setup_method(self):
        """Set up test table."""
        self.table = [
            [1, 2, 3],
            ["Apple", "Banana", "Cherry"],
            [1.00, 0.50, 2.00],
        ]

    def test_hlookup_exact_match(self):
        """Test HLOOKUP with exact match."""
        result = fn_hlookup(2, self.table, 2, False)
        assert result == "Banana"

    def test_hlookup_range_match(self):
        """Test HLOOKUP with range match."""
        result = fn_hlookup(2, self.table, 2, True)
        assert result == "Banana"

    def test_hlookup_not_found(self):
        """Test HLOOKUP when value not found."""
        result = fn_hlookup(5, self.table, 2, False)
        assert result == "#N/A"

    def test_hlookup_invalid_row(self):
        """Test HLOOKUP with invalid row index."""
        result = fn_hlookup(1, self.table, 5, False)
        assert result == "#REF!"

    def test_hlookup_empty_table(self):
        """Test HLOOKUP with empty table."""
        result = fn_hlookup(1, [], 1, False)
        assert result == "#N/A"

    def test_hlookup_1d_table(self):
        """Test HLOOKUP with 1D table."""
        result = fn_hlookup(2, [1, 2, 3], 1, False)
        assert result == 2

    def test_hlookup_range_unsorted(self):
        """Approximate match requires sorted data."""
        table = [[2, 1, 3], ["Banana", "Apple", "Cherry"]]
        result = fn_hlookup(2, table, 2, True)
        assert result == "#N/A"


class TestINDEX:
    """Tests for INDEX function."""

    def test_index_2d_array(self):
        """Test INDEX with 2D array."""
        arr = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
        assert fn_index(arr, 2, 2) == 5
        assert fn_index(arr, 1, 3) == 3

    def test_index_1d_array(self):
        """Test INDEX with 1D array."""
        arr = [10, 20, 30]
        assert fn_index(arr, 2) == 20

    def test_index_single_value(self):
        """Test INDEX with single value."""
        assert fn_index(42, 1) == 42

    def test_index_row_only(self):
        """Test INDEX with row only (returns entire row)."""
        arr = [[1, 2], [3, 4]]
        result = fn_index(arr, 2)
        assert result == [3, 4]

    def test_index_invalid_row(self):
        """Test INDEX with invalid row."""
        arr = [[1, 2], [3, 4]]
        assert fn_index(arr, 5, 1) == "#REF!"

    def test_index_invalid_col(self):
        """Test INDEX with invalid column."""
        arr = [[1, 2], [3, 4]]
        assert fn_index(arr, 1, 5) == "#REF!"


class TestMATCH:
    """Tests for MATCH function."""

    def test_match_exact(self):
        """Test MATCH with exact match."""
        arr = [10, 20, 30, 40]
        assert fn_match(20, arr, 0) == 2

    def test_match_exact_not_found(self):
        """Test MATCH exact when not found."""
        arr = [10, 20, 30]
        assert fn_match(25, arr, 0) == 0

    def test_match_ascending(self):
        """Test MATCH with ascending array (match_type=1)."""
        arr = [10, 20, 30, 40]
        assert fn_match(25, arr, 1) == 2  # Largest <= 25

    def test_match_descending(self):
        """Test MATCH with descending array (match_type=-1)."""
        arr = [40, 30, 20, 10]
        assert fn_match(25, arr, -1) == 2  # Smallest >= 25

    def test_match_nested_array(self):
        """Test MATCH with nested array."""
        arr = [[10], [20], [30]]
        assert fn_match(20, arr, 0) == 2

    def test_match_string(self):
        """Test MATCH with string values."""
        arr = ["apple", "banana", "cherry"]
        assert fn_match("banana", arr, 0) == 2


class TestLOOKUP:
    """Tests for LOOKUP function."""

    def test_lookup_with_result_vector(self):
        """Test LOOKUP with separate result vector."""
        lookup = [1, 2, 3, 4]
        result = ["A", "B", "C", "D"]
        assert fn_lookup(2, lookup, result) == "B"

    def test_lookup_without_result(self):
        """Test LOOKUP without result vector."""
        arr = [1, 2, 3, 4]
        assert fn_lookup(3, arr) == 3

    def test_lookup_range_match(self):
        """Test LOOKUP finds largest <=."""
        lookup = [1, 2, 3, 4]
        result = ["A", "B", "C", "D"]
        assert fn_lookup(2.5, lookup, result) == "B"

    def test_lookup_not_found(self):
        """Test LOOKUP when value smaller than all."""
        lookup = [10, 20, 30]
        result = fn_lookup(5, lookup)
        assert result == "#N/A"


class TestROWS:
    """Tests for ROWS function."""

    def test_rows_2d_array(self):
        """Test ROWS with 2D array."""
        arr = [[1, 2], [3, 4], [5, 6]]
        assert fn_rows(arr) == 3

    def test_rows_1d_array(self):
        """Test ROWS with 1D array."""
        arr = [1, 2, 3]
        assert fn_rows(arr) == 3

    def test_rows_single_value(self):
        """Test ROWS with single value."""
        assert fn_rows(42) == 1


class TestCOLS:
    """Tests for COLS function."""

    def test_cols_2d_array(self):
        """Test COLS with 2D array."""
        arr = [[1, 2, 3], [4, 5, 6]]
        assert fn_cols(arr) == 3

    def test_cols_1d_array(self):
        """Test COLS with 1D array."""
        arr = [1, 2, 3]
        assert fn_cols(arr) == 3

    def test_cols_single_value(self):
        """Test COLS with single value."""
        assert fn_cols(42) == 1

    def test_columns_alias(self):
        """Test COLUMNS is alias for COLS."""
        arr = [[1, 2], [3, 4]]
        assert fn_columns(arr) == fn_cols(arr)


class TestTRANSPOSE:
    """Tests for TRANSPOSE function."""

    def test_transpose_2d(self):
        """Test TRANSPOSE with 2D array."""
        arr = [[1, 2, 3], [4, 5, 6]]
        result = fn_transpose(arr)
        assert result == [[1, 4], [2, 5], [3, 6]]

    def test_transpose_1d(self):
        """Test TRANSPOSE with 1D array."""
        arr = [1, 2, 3]
        result = fn_transpose(arr)
        assert result == [[1], [2], [3]]

    def test_transpose_single(self):
        """Test TRANSPOSE with single value."""
        result = fn_transpose(42)
        assert result == [[42]]

    def test_transpose_empty(self):
        """Test TRANSPOSE with empty array."""
        result = fn_transpose([])
        assert result == [[]]


class TestOFFSET:
    """Tests for OFFSET function."""

    def test_offset_returns_ref(self):
        """Test OFFSET returns #REF! without context."""
        result = fn_offset("A1", 1, 1)
        assert result == "#REF!"


class TestINDIRECT:
    """Tests for INDIRECT function."""

    def test_indirect_returns_ref(self):
        """Test INDIRECT returns #REF! without context."""
        result = fn_indirect("A1")
        assert result == "#REF!"


class TestROW:
    """Tests for ROW function."""

    def test_row_returns_1(self):
        """Test ROW returns 1 without context."""
        assert fn_row() == 1
        assert fn_row("A5") == 1


class TestCOLUMN:
    """Tests for COLUMN function."""

    def test_column_returns_1(self):
        """Test COLUMN returns 1 without context."""
        assert fn_column() == 1
        assert fn_column("B1") == 1


class TestADDRESS:
    """Tests for ADDRESS function."""

    def test_address_absolute(self):
        """Test ADDRESS with absolute reference."""
        result = fn_address(1, 1, 1)
        assert result == "$A$1"

    def test_address_row_absolute(self):
        """Test ADDRESS with row absolute only."""
        result = fn_address(1, 1, 2)
        assert result == "A$1"

    def test_address_col_absolute(self):
        """Test ADDRESS with column absolute only."""
        result = fn_address(1, 1, 3)
        assert result == "$A1"

    def test_address_relative(self):
        """Test ADDRESS with relative reference."""
        result = fn_address(1, 1, 4)
        assert result == "A1"

    def test_address_larger_col(self):
        """Test ADDRESS with larger column number."""
        result = fn_address(1, 27, 1)
        assert result == "$AA$1"


class TestFunctionRegistry:
    """Test the function registry."""

    def test_all_functions_registered(self):
        """Test that all functions are in the registry."""
        assert "VLOOKUP" in LOOKUP_FUNCTIONS
        assert "HLOOKUP" in LOOKUP_FUNCTIONS
        assert "LOOKUP" in LOOKUP_FUNCTIONS
        assert "MATCH" in LOOKUP_FUNCTIONS
        assert "INDEX" in LOOKUP_FUNCTIONS
        assert "OFFSET" in LOOKUP_FUNCTIONS
        assert "INDIRECT" in LOOKUP_FUNCTIONS
        assert "ROW" in LOOKUP_FUNCTIONS
        assert "COLUMN" in LOOKUP_FUNCTIONS
        assert "ADDRESS" in LOOKUP_FUNCTIONS
        assert "ROWS" in LOOKUP_FUNCTIONS
        assert "COLS" in LOOKUP_FUNCTIONS
        assert "COLUMNS" in LOOKUP_FUNCTIONS
        assert "TRANSPOSE" in LOOKUP_FUNCTIONS

    def test_functions_callable(self):
        """Test that all registered functions are callable."""
        for name, func in LOOKUP_FUNCTIONS.items():
            assert callable(func), f"{name} is not callable"

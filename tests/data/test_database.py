"""Tests for database operations module."""

import pytest

from lotus123 import Spreadsheet
from lotus123.data.database import (
    DatabaseOperations,
    SortKey,
    SortOrder,
)


class TestSortOrder:
    """Tests for SortOrder enum."""

    def test_orders_exist(self):
        """Test all sort orders exist."""
        assert SortOrder.ASCENDING
        assert SortOrder.DESCENDING


class TestSortKey:
    """Tests for SortKey dataclass."""

    def test_default_order(self):
        """Test default order is ascending."""
        key = SortKey(column=0)
        assert key.column == 0
        assert key.order == SortOrder.ASCENDING

    def test_custom_order(self):
        """Test custom order."""
        key = SortKey(column=1, order=SortOrder.DESCENDING)
        assert key.column == 1
        assert key.order == SortOrder.DESCENDING


class TestDatabaseOperations:
    """Tests for DatabaseOperations class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.db = DatabaseOperations(self.ss)

    def _setup_sample_data(self):
        """Set up sample data in spreadsheet."""
        # Headers
        self.ss.set_cell(0, 0, "Name")
        self.ss.set_cell(0, 1, "Age")
        self.ss.set_cell(0, 2, "City")
        # Data
        self.ss.set_cell(1, 0, "John")
        self.ss.set_cell(1, 1, "30")
        self.ss.set_cell(1, 2, "NYC")
        self.ss.set_cell(2, 0, "Alice")
        self.ss.set_cell(2, 1, "25")
        self.ss.set_cell(2, 2, "LA")
        self.ss.set_cell(3, 0, "Bob")
        self.ss.set_cell(3, 1, "35")
        self.ss.set_cell(3, 2, "Chicago")


class TestSortRange:
    """Tests for sort_range method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.db = DatabaseOperations(self.ss)

    def _setup_numeric_data(self):
        """Set up numeric test data."""
        self.ss.set_cell(0, 0, "Name")
        self.ss.set_cell(0, 1, "Score")
        self.ss.set_cell(1, 0, "Alice")
        self.ss.set_cell(1, 1, "90")
        self.ss.set_cell(2, 0, "Bob")
        self.ss.set_cell(2, 1, "75")
        self.ss.set_cell(3, 0, "Charlie")
        self.ss.set_cell(3, 1, "85")

    def test_sort_ascending_numeric(self):
        """Test sorting numeric column ascending."""
        self._setup_numeric_data()
        keys = [SortKey(column=1, order=SortOrder.ASCENDING)]
        self.db.sort_range(0, 0, 3, 1, keys, has_header=True)

        # Bob (75), Charlie (85), Alice (90)
        assert self.ss.get_value(1, 0) == "Bob"
        assert self.ss.get_value(2, 0) == "Charlie"
        assert self.ss.get_value(3, 0) == "Alice"

    def test_sort_descending_numeric(self):
        """Test sorting numeric column descending."""
        self._setup_numeric_data()
        keys = [SortKey(column=1, order=SortOrder.DESCENDING)]
        self.db.sort_range(0, 0, 3, 1, keys, has_header=True)

        # Alice (90), Charlie (85), Bob (75)
        assert self.ss.get_value(1, 0) == "Alice"
        assert self.ss.get_value(2, 0) == "Charlie"
        assert self.ss.get_value(3, 0) == "Bob"

    def test_sort_string_ascending(self):
        """Test sorting string column ascending."""
        self._setup_numeric_data()
        keys = [SortKey(column=0, order=SortOrder.ASCENDING)]
        self.db.sort_range(0, 0, 3, 1, keys, has_header=True)

        # Alice, Bob, Charlie
        assert self.ss.get_value(1, 0) == "Alice"
        assert self.ss.get_value(2, 0) == "Bob"
        assert self.ss.get_value(3, 0) == "Charlie"

    def test_sort_string_descending(self):
        """Test sorting string column descending."""
        self._setup_numeric_data()
        keys = [SortKey(column=0, order=SortOrder.DESCENDING)]
        self.db.sort_range(0, 0, 3, 1, keys, has_header=True)

        # Charlie, Bob, Alice
        assert self.ss.get_value(1, 0) == "Charlie"
        assert self.ss.get_value(2, 0) == "Bob"
        assert self.ss.get_value(3, 0) == "Alice"

    def test_sort_no_header(self):
        """Test sorting without header."""
        self.ss.set_cell(0, 0, "C")
        self.ss.set_cell(1, 0, "A")
        self.ss.set_cell(2, 0, "B")

        keys = [SortKey(column=0, order=SortOrder.ASCENDING)]
        self.db.sort_range(0, 0, 2, 0, keys, has_header=False)

        assert self.ss.get_value(0, 0) == "A"
        assert self.ss.get_value(1, 0) == "B"
        assert self.ss.get_value(2, 0) == "C"

    def test_sort_empty_range(self):
        """Test sorting empty/minimal range."""
        self.ss.set_cell(0, 0, "Header")
        keys = [SortKey(column=0)]
        # No data rows - should not crash
        self.db.sort_range(0, 0, 0, 0, keys, has_header=True)

    def test_sort_reversed_range(self):
        """Test sorting with reversed row/col indices."""
        self._setup_numeric_data()
        keys = [SortKey(column=0, order=SortOrder.ASCENDING)]
        # Pass end_row before start_row
        self.db.sort_range(3, 0, 0, 1, keys, has_header=True)
        # Should normalize and sort correctly
        assert self.ss.get_value(1, 0) == "Alice"


class TestQuery:
    """Tests for query method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.db = DatabaseOperations(self.ss)
        # Set up sample data
        self.ss.set_cell(0, 0, "Name")
        self.ss.set_cell(0, 1, "Age")
        self.ss.set_cell(1, 0, "John")
        self.ss.set_cell(1, 1, "30")
        self.ss.set_cell(2, 0, "Alice")
        self.ss.set_cell(2, 1, "25")
        self.ss.set_cell(3, 0, "Bob")
        self.ss.set_cell(3, 1, "35")

    def test_query_no_criteria(self):
        """Test query without criteria returns all rows."""
        rows = self.db.query((0, 0, 3, 1))
        assert rows == [1, 2, 3]

    def test_query_with_func(self):
        """Test query with criteria function."""
        rows = self.db.query(
            (0, 0, 3, 1),
            criteria_func=lambda r: r[0] == "John"
        )
        assert rows == [1]

    def test_query_age_criteria(self):
        """Test query with age criteria."""
        rows = self.db.query(
            (0, 0, 3, 1),
            criteria_func=lambda r: int(r[1]) > 28
        )
        assert rows == [1, 3]  # John (30) and Bob (35)

    def test_query_no_matches(self):
        """Test query with no matches."""
        rows = self.db.query(
            (0, 0, 3, 1),
            criteria_func=lambda r: r[0] == "Nobody"
        )
        assert rows == []


class TestExtract:
    """Tests for extract method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.db = DatabaseOperations(self.ss)
        # Set up sample data
        self.ss.set_cell(0, 0, "Name")
        self.ss.set_cell(0, 1, "Age")
        self.ss.set_cell(1, 0, "John")
        self.ss.set_cell(1, 1, "30")
        self.ss.set_cell(2, 0, "Alice")
        self.ss.set_cell(2, 1, "25")
        self.ss.set_cell(3, 0, "Bob")
        self.ss.set_cell(3, 1, "35")

    def test_extract_all_columns(self):
        """Test extract all columns."""
        matching = [1, 3]  # John and Bob
        count = self.db.extract(
            (0, 0, 3, 1),
            (10, 0),
            matching
        )

        assert count == 2
        # Header
        assert self.ss.get_value(10, 0) == "Name"
        assert self.ss.get_value(10, 1) == "Age"
        # Data (values may be parsed as numbers)
        assert self.ss.get_value(11, 0) == "John"
        assert self.ss.get_value(11, 1) in [30, "30"]
        assert self.ss.get_value(12, 0) == "Bob"
        assert self.ss.get_value(12, 1) in [35, "35"]

    def test_extract_specific_columns(self):
        """Test extract specific columns."""
        matching = [1, 2]
        count = self.db.extract(
            (0, 0, 3, 1),
            (10, 0),
            matching,
            columns=[0]  # Only Name column
        )

        assert count == 2
        assert self.ss.get_value(10, 0) == "Name"
        assert self.ss.get_value(11, 0) == "John"
        assert self.ss.get_value(12, 0) == "Alice"

    def test_extract_no_matches(self):
        """Test extract with no matches."""
        count = self.db.extract(
            (0, 0, 3, 1),
            (10, 0),
            []
        )
        assert count == 0


class TestDeleteMatching:
    """Tests for delete_matching method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.db = DatabaseOperations(self.ss)
        # Set up sample data
        self.ss.set_cell(0, 0, "Name")
        self.ss.set_cell(1, 0, "John")
        self.ss.set_cell(2, 0, "Alice")
        self.ss.set_cell(3, 0, "Bob")

    def test_delete_single_row(self):
        """Test deleting single row."""
        count = self.db.delete_matching((0, 0, 3, 0), [2])  # Delete Alice
        assert count == 1
        # Alice row should be gone, Bob shifted up
        assert self.ss.get_value(1, 0) == "John"
        assert self.ss.get_value(2, 0) == "Bob"

    def test_delete_multiple_rows(self):
        """Test deleting multiple rows."""
        count = self.db.delete_matching((0, 0, 3, 0), [1, 3])  # Delete John and Bob
        assert count == 2
        # Only Alice should remain at row 1
        assert self.ss.get_value(1, 0) == "Alice"

    def test_delete_no_rows(self):
        """Test deleting no rows."""
        count = self.db.delete_matching((0, 0, 3, 0), [])
        assert count == 0


class TestUnique:
    """Tests for unique method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.db = DatabaseOperations(self.ss)
        # Set up data with duplicates
        self.ss.set_cell(0, 0, "Name")
        self.ss.set_cell(0, 1, "Dept")
        self.ss.set_cell(1, 0, "John")
        self.ss.set_cell(1, 1, "Sales")
        self.ss.set_cell(2, 0, "Alice")
        self.ss.set_cell(2, 1, "HR")
        self.ss.set_cell(3, 0, "John")  # Duplicate name
        self.ss.set_cell(3, 1, "IT")
        self.ss.set_cell(4, 0, "Alice")  # Duplicate name
        self.ss.set_cell(4, 1, "HR")  # Same dept too

    def test_unique_single_column(self):
        """Test unique on single column."""
        unique_rows = self.db.unique((0, 0, 4, 1), key_columns=[0])
        # John, Alice (first occurrences)
        assert unique_rows == [1, 2]

    def test_unique_multiple_columns(self):
        """Test unique on multiple columns."""
        unique_rows = self.db.unique((0, 0, 4, 1), key_columns=[0, 1])
        # (John, Sales), (Alice, HR), (John, IT) - row 4 is duplicate of row 2
        assert unique_rows == [1, 2, 3]


class TestSubtotal:
    """Tests for subtotal method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.db = DatabaseOperations(self.ss)
        # Set up data
        self.ss.set_cell(0, 0, "Dept")
        self.ss.set_cell(0, 1, "Sales")
        self.ss.set_cell(0, 2, "Units")
        # Sales dept
        self.ss.set_cell(1, 0, "Sales")
        self.ss.set_cell(1, 1, "100")
        self.ss.set_cell(1, 2, "10")
        self.ss.set_cell(2, 0, "Sales")
        self.ss.set_cell(2, 1, "200")
        self.ss.set_cell(2, 2, "20")
        # HR dept
        self.ss.set_cell(3, 0, "HR")
        self.ss.set_cell(3, 1, "50")
        self.ss.set_cell(3, 2, "5")

    def test_subtotal_single_column(self):
        """Test subtotal on single column."""
        totals = self.db.subtotal((0, 0, 3, 2), group_col=0, sum_cols=[1])
        # Sales total: 100 + 200 = 300
        assert totals["Sales"][1] == 300
        # HR total: 50
        assert totals["HR"][1] == 50

    def test_subtotal_multiple_columns(self):
        """Test subtotal on multiple columns."""
        totals = self.db.subtotal((0, 0, 3, 2), group_col=0, sum_cols=[1, 2])
        assert totals["Sales"][1] == 300  # Sales sum
        assert totals["Sales"][2] == 30   # Units sum
        assert totals["HR"][1] == 50
        assert totals["HR"][2] == 5

    def test_subtotal_with_non_numeric(self):
        """Test subtotal ignores non-numeric values."""
        self.ss.set_cell(4, 0, "Sales")
        self.ss.set_cell(4, 1, "invalid")  # Non-numeric
        self.ss.set_cell(4, 2, "5")

        totals = self.db.subtotal((0, 0, 4, 2), group_col=0, sum_cols=[1, 2])
        # Sales total should still be 300 (ignores "invalid")
        assert totals["Sales"][1] == 300
        # Units includes the new row
        assert totals["Sales"][2] == 35

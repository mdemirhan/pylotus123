"""Tests for named range management."""

import pytest

from lotus123.core.named_ranges import NAME_PATTERN, NamedRange, NamedRangeManager
from lotus123.core.reference import CellReference, RangeReference


class TestNamedRange:
    """Tests for NamedRange dataclass."""

    def test_name_normalized_to_uppercase(self):
        """Test name is normalized to uppercase."""
        ref = CellReference(0, 0)
        nr = NamedRange("myname", ref)
        assert nr.name == "MYNAME"

    def test_is_single_cell_true(self):
        """Test is_single_cell for cell reference."""
        ref = CellReference(0, 0)
        nr = NamedRange("test", ref)
        assert nr.is_single_cell is True

    def test_is_single_cell_false(self):
        """Test is_single_cell for range reference."""
        start = CellReference(0, 0)
        end = CellReference(5, 5)
        ref = RangeReference(start, end)
        nr = NamedRange("test", ref)
        assert nr.is_single_cell is False

    def test_to_dict_cell(self):
        """Test to_dict for cell reference."""
        ref = CellReference(0, 0)
        nr = NamedRange("test", ref, "A description")
        data = nr.to_dict()
        assert data["name"] == "TEST"
        assert "reference" in data
        assert data["is_range"] is False
        assert data["description"] == "A description"

    def test_to_dict_range(self):
        """Test to_dict for range reference."""
        start = CellReference(0, 0)
        end = CellReference(5, 5)
        ref = RangeReference(start, end)
        nr = NamedRange("test", ref)
        data = nr.to_dict()
        assert data["is_range"] is True

    def test_from_dict_cell(self):
        """Test from_dict for cell reference."""
        data = {"name": "TEST", "reference": "A1", "is_range": False}
        nr = NamedRange.from_dict(data)
        assert nr.name == "TEST"
        assert isinstance(nr.reference, CellReference)

    def test_from_dict_range(self):
        """Test from_dict for range reference."""
        data = {"name": "TEST", "reference": "A1:B5", "is_range": True}
        nr = NamedRange.from_dict(data)
        assert nr.name == "TEST"
        assert isinstance(nr.reference, RangeReference)

    def test_from_dict_with_description(self):
        """Test from_dict with description."""
        data = {"name": "TEST", "reference": "A1", "description": "My range"}
        nr = NamedRange.from_dict(data)
        assert nr.description == "My range"


class TestNamedRangeManager:
    """Tests for NamedRangeManager class."""

    def test_add_cell_reference(self):
        """Test adding a cell reference."""
        mgr = NamedRangeManager()
        ref = CellReference(0, 0)
        nr = mgr.add("Sales", ref, "Sales total")
        assert nr.name == "SALES"
        assert nr.description == "Sales total"

    def test_add_invalid_name(self):
        """Test adding with invalid name raises error."""
        mgr = NamedRangeManager()
        ref = CellReference(0, 0)
        with pytest.raises(ValueError):
            mgr.add("123invalid", ref)

    def test_add_from_string_cell(self):
        """Test add_from_string for cell."""
        mgr = NamedRangeManager()
        nr = mgr.add_from_string("Total", "A1")
        assert isinstance(nr.reference, CellReference)

    def test_add_from_string_range(self):
        """Test add_from_string for range."""
        mgr = NamedRangeManager()
        nr = mgr.add_from_string("Data", "A1:B10")
        assert isinstance(nr.reference, RangeReference)

    def test_delete_existing(self):
        """Test deleting existing name."""
        mgr = NamedRangeManager()
        mgr.add("Test", CellReference(0, 0))
        assert mgr.delete("Test") is True
        assert mgr.get("Test") is None

    def test_delete_nonexistent(self):
        """Test deleting nonexistent name."""
        mgr = NamedRangeManager()
        assert mgr.delete("Nonexistent") is False

    def test_get_existing(self):
        """Test getting existing name."""
        mgr = NamedRangeManager()
        mgr.add("Test", CellReference(0, 0))
        nr = mgr.get("test")  # Case insensitive
        assert nr is not None
        assert nr.name == "TEST"

    def test_get_nonexistent(self):
        """Test getting nonexistent name."""
        mgr = NamedRangeManager()
        assert mgr.get("Nonexistent") is None

    def test_get_reference(self):
        """Test get_reference."""
        mgr = NamedRangeManager()
        ref = CellReference(0, 0)
        mgr.add("Test", ref)
        result = mgr.get_reference("Test")
        assert result == ref

    def test_get_reference_nonexistent(self):
        """Test get_reference for nonexistent."""
        mgr = NamedRangeManager()
        assert mgr.get_reference("Nonexistent") is None

    def test_exists(self):
        """Test exists method."""
        mgr = NamedRangeManager()
        mgr.add("Test", CellReference(0, 0))
        assert mgr.exists("test") is True
        assert mgr.exists("Other") is False

    def test_list_all_sorted(self):
        """Test list_all returns sorted names."""
        mgr = NamedRangeManager()
        mgr.add("Zebra", CellReference(0, 0))
        mgr.add("Alpha", CellReference(0, 1))
        mgr.add("Middle", CellReference(0, 2))
        names = mgr.list_all()
        assert [n.name for n in names] == ["ALPHA", "MIDDLE", "ZEBRA"]

    def test_find_by_cell_single(self):
        """Test find_by_cell for single cell reference."""
        mgr = NamedRangeManager()
        mgr.add("Test", CellReference(5, 5))
        result = mgr.find_by_cell(5, 5)
        assert len(result) == 1
        assert result[0].name == "TEST"

    def test_find_by_cell_in_range(self):
        """Test find_by_cell for cell in range."""
        mgr = NamedRangeManager()
        start = CellReference(0, 0)
        end = CellReference(10, 10)
        mgr.add("Range", RangeReference(start, end))
        result = mgr.find_by_cell(5, 5)
        assert len(result) == 1

    def test_find_by_cell_not_found(self):
        """Test find_by_cell when cell not in any range."""
        mgr = NamedRangeManager()
        mgr.add("Test", CellReference(0, 0))
        result = mgr.find_by_cell(5, 5)
        assert len(result) == 0

    def test_adjust_for_insert_row_cell(self):
        """Test adjust_for_insert_row for cell reference."""
        mgr = NamedRangeManager()
        mgr.add("Test", CellReference(5, 0))
        mgr.adjust_for_insert_row(3)
        ref = mgr.get_reference("Test")
        assert isinstance(ref, CellReference)
        assert ref.row == 6

    def test_adjust_for_insert_row_range(self):
        """Test adjust_for_insert_row for range reference."""
        mgr = NamedRangeManager()
        start = CellReference(5, 0)
        end = CellReference(10, 5)
        mgr.add("Range", RangeReference(start, end))
        mgr.adjust_for_insert_row(3)
        ref = mgr.get_reference("Range")
        assert isinstance(ref, RangeReference)
        assert ref.start.row == 6
        assert ref.end.row == 11

    def test_adjust_for_delete_row_cell(self):
        """Test adjust_for_delete_row for cell reference."""
        mgr = NamedRangeManager()
        mgr.add("Test", CellReference(5, 0))
        mgr.adjust_for_delete_row(3)
        ref = mgr.get_reference("Test")
        assert isinstance(ref, CellReference)
        assert ref.row == 4

    def test_adjust_for_delete_row_invalidates(self):
        """Test adjust_for_delete_row invalidates cell on deleted row."""
        mgr = NamedRangeManager()
        mgr.add("Test", CellReference(3, 0))
        invalidated = mgr.adjust_for_delete_row(3)
        assert "TEST" in invalidated
        assert mgr.get("Test") is None

    def test_adjust_for_insert_col_cell(self):
        """Test adjust_for_insert_col for cell reference."""
        mgr = NamedRangeManager()
        mgr.add("Test", CellReference(0, 5))
        mgr.adjust_for_insert_col(3)
        ref = mgr.get_reference("Test")
        assert isinstance(ref, CellReference)
        assert ref.col == 6

    def test_adjust_for_delete_col_cell(self):
        """Test adjust_for_delete_col for cell reference."""
        mgr = NamedRangeManager()
        mgr.add("Test", CellReference(0, 5))
        mgr.adjust_for_delete_col(3)
        ref = mgr.get_reference("Test")
        assert isinstance(ref, CellReference)
        assert ref.col == 4

    def test_adjust_for_delete_col_invalidates(self):
        """Test adjust_for_delete_col invalidates cell on deleted col."""
        mgr = NamedRangeManager()
        mgr.add("Test", CellReference(0, 3))
        invalidated = mgr.adjust_for_delete_col(3)
        assert "TEST" in invalidated
        assert mgr.get("Test") is None

    def test_clear(self):
        """Test clear removes all names."""
        mgr = NamedRangeManager()
        # Use names with underscores - names like "Test1" match cell reference pattern
        mgr.add("Test_One", CellReference(0, 0))
        mgr.add("Test_Two", CellReference(0, 1))
        mgr.clear()
        assert len(mgr) == 0

    def test_len(self):
        """Test __len__."""
        mgr = NamedRangeManager()
        assert len(mgr) == 0
        mgr.add("TestName", CellReference(0, 0))
        assert len(mgr) == 1

    def test_iter(self):
        """Test __iter__."""
        mgr = NamedRangeManager()
        mgr.add("Test_One", CellReference(0, 0))
        mgr.add("Test_Two", CellReference(0, 1))
        names = [nr.name for nr in mgr]
        assert len(names) == 2

    def test_contains(self):
        """Test __contains__."""
        mgr = NamedRangeManager()
        mgr.add("TestName", CellReference(0, 0))
        assert "testname" in mgr
        assert "Other" not in mgr

    def test_is_valid_name_valid(self):
        """Test is_valid_name for valid names."""
        assert NamedRangeManager.is_valid_name("Sales") is True
        # Names like "Total2024" match cell reference pattern (letters+digits)
        # Use underscores or mixed case to avoid matching
        assert NamedRangeManager.is_valid_name("Total_2024") is True
        assert NamedRangeManager.is_valid_name("My_Range") is True

    def test_is_valid_name_invalid(self):
        """Test is_valid_name for invalid names."""
        assert NamedRangeManager.is_valid_name("123abc") is False  # Starts with number
        assert NamedRangeManager.is_valid_name("A1") is False  # Cell reference
        assert NamedRangeManager.is_valid_name("AB123") is False  # Cell reference
        assert NamedRangeManager.is_valid_name("my-range") is False  # Has hyphen

    def test_to_dict_from_dict(self):
        """Test serialization roundtrip."""
        mgr = NamedRangeManager()
        mgr.add("Test_One", CellReference(0, 0), "First")
        mgr.add("Test_Two", CellReference(0, 1), "Second")

        data = mgr.to_dict()

        mgr2 = NamedRangeManager()
        mgr2.from_dict(data)

        assert len(mgr2) == 2
        assert mgr2.get("Test_One").description == "First"


class TestNamePattern:
    """Tests for NAME_PATTERN regex."""

    def test_valid_patterns(self):
        """Test valid name patterns."""
        assert NAME_PATTERN.match("Test") is not None
        assert NAME_PATTERN.match("A") is not None
        assert NAME_PATTERN.match("Test123") is not None
        assert NAME_PATTERN.match("Test_Name") is not None

    def test_invalid_patterns(self):
        """Test invalid name patterns."""
        assert NAME_PATTERN.match("123") is None
        assert NAME_PATTERN.match("_test") is None
        assert NAME_PATTERN.match("test-name") is None


class TestNamedRangeErrorHandling:
    """Tests for error handling in named range operations.

    These tests verify that proper errors are raised for invalid inputs.
    """

    def test_add_name_with_spaces_raises_error(self):
        """Test that adding a name with spaces raises ValueError."""
        mgr = NamedRangeManager()
        ref = CellReference(0, 0)
        with pytest.raises(ValueError) as excinfo:
            mgr.add("HELLO RANGE", ref)
        assert "Invalid name" in str(excinfo.value)

    def test_add_name_with_special_chars_raises_error(self):
        """Test that adding a name with special characters raises ValueError."""
        mgr = NamedRangeManager()
        ref = CellReference(0, 0)
        with pytest.raises(ValueError) as excinfo:
            mgr.add("HELLO-RANGE", ref)
        assert "Invalid name" in str(excinfo.value)

    def test_add_name_starting_with_number_raises_error(self):
        """Test that adding a name starting with a number raises ValueError."""
        mgr = NamedRangeManager()
        ref = CellReference(0, 0)
        with pytest.raises(ValueError) as excinfo:
            mgr.add("123ABC", ref)
        assert "Invalid name" in str(excinfo.value)

    def test_add_from_string_invalid_name_raises_error(self):
        """Test that add_from_string with invalid name raises ValueError."""
        mgr = NamedRangeManager()
        with pytest.raises(ValueError) as excinfo:
            mgr.add_from_string("HELLO RANGE", "A1:B10")
        assert "Invalid name" in str(excinfo.value)

    def test_add_from_string_cell_reference_name_raises_error(self):
        """Test that using a cell reference as a name raises ValueError."""
        mgr = NamedRangeManager()
        with pytest.raises(ValueError) as excinfo:
            mgr.add_from_string("A1", "B1:C10")
        assert "Invalid name" in str(excinfo.value) or "cell reference" in str(excinfo.value).lower()

    def test_add_duplicate_name_replaces(self):
        """Test that adding a duplicate name replaces the old one."""
        mgr = NamedRangeManager()
        mgr.add("Test", CellReference(0, 0))
        mgr.add("Test", CellReference(5, 5))  # Should replace
        ref = mgr.get_reference("Test")
        assert isinstance(ref, CellReference)
        assert ref.row == 5
        assert ref.col == 5

    def test_error_message_includes_reason(self):
        """Test that error messages explain why the name is invalid."""
        mgr = NamedRangeManager()
        ref = CellReference(0, 0)

        # Name with spaces
        with pytest.raises(ValueError) as excinfo:
            mgr.add("MY RANGE", ref)
        error_msg = str(excinfo.value).lower()
        # Should mention the name or the invalid pattern
        assert "my range" in error_msg or "invalid" in error_msg

    def test_empty_name_raises_error(self):
        """Test that empty name raises ValueError."""
        mgr = NamedRangeManager()
        ref = CellReference(0, 0)
        with pytest.raises(ValueError):
            mgr.add("", ref)

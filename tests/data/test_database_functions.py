"""Tests for database statistical functions."""


from lotus123.formula.functions.database import (
    DATABASE_FUNCTIONS,
    _get_field_index,
    _matches_criteria,
    _to_number,
    fn_davg,
    fn_dcount,
    fn_dcounta,
    fn_dget,
    fn_dmax,
    fn_dmin,
    fn_dstd,
    fn_dstdp,
    fn_dsum,
    fn_dvar,
    fn_dvarp,
)


class TestHelperFunctions:
    """Tests for helper functions."""

    def test_to_number_int(self):
        """Test converting int to number."""
        assert _to_number(42) == 42.0

    def test_to_number_float(self):
        """Test converting float to number."""
        assert _to_number(3.14) == 3.14

    def test_to_number_string(self):
        """Test converting numeric string to number."""
        assert _to_number("123.45") == 123.45

    def test_to_number_string_with_comma(self):
        """Test converting string with comma."""
        assert _to_number("1,234.56") == 1234.56

    def test_to_number_non_numeric_string(self):
        """Test non-numeric string returns None."""
        assert _to_number("hello") is None

    def test_to_number_none(self):
        """Test None returns None."""
        assert _to_number(None) is None

    def test_get_field_index_by_number(self):
        """Test getting field index by number."""
        db = [["Name", "Age", "Salary"], ["Alice", 30, 50000]]
        assert _get_field_index(db, 1) == 0
        assert _get_field_index(db, 2) == 1
        assert _get_field_index(db, 3) == 2

    def test_get_field_index_by_name(self):
        """Test getting field index by name."""
        db = [["Name", "Age", "Salary"], ["Alice", 30, 50000]]
        assert _get_field_index(db, "Name") == 0
        assert _get_field_index(db, "AGE") == 1  # Case insensitive
        assert _get_field_index(db, "salary") == 2

    def test_get_field_index_invalid(self):
        """Test invalid field returns None."""
        db = [["Name", "Age"], ["Alice", 30]]
        assert _get_field_index(db, "Invalid") is None
        assert _get_field_index(db, 10) is None

    def test_get_field_index_empty_db(self):
        """Test empty database returns None."""
        assert _get_field_index([], "Name") is None


class TestMatchesCriteria:
    """Tests for criteria matching."""

    def test_empty_criteria(self):
        """Test empty criteria matches all."""
        row = ["Alice", 30, 50000]
        headers = ["Name", "Age", "Salary"]
        assert _matches_criteria(row, headers, []) is True
        assert _matches_criteria(row, headers, [[]]) is True

    def test_exact_match(self):
        """Test exact match criteria."""
        row = ["Alice", 30, 50000]
        headers = ["Name", "Age", "Salary"]
        criteria = [["Name"], ["Alice"]]
        assert _matches_criteria(row, headers, criteria) is True

    def test_exact_match_case_insensitive(self):
        """Test exact match is case insensitive."""
        row = ["Alice", 30, 50000]
        headers = ["Name", "Age", "Salary"]
        criteria = [["Name"], ["ALICE"]]
        assert _matches_criteria(row, headers, criteria) is True

    def test_greater_than(self):
        """Test > comparison."""
        row = ["Alice", 30, 50000]
        headers = ["Name", "Age", "Salary"]
        criteria = [["Age"], [">25"]]
        assert _matches_criteria(row, headers, criteria) is True
        criteria = [["Age"], [">35"]]
        assert _matches_criteria(row, headers, criteria) is False

    def test_less_than(self):
        """Test < comparison."""
        row = ["Alice", 30, 50000]
        headers = ["Name", "Age", "Salary"]
        criteria = [["Age"], ["<35"]]
        assert _matches_criteria(row, headers, criteria) is True

    def test_greater_equal(self):
        """Test >= comparison."""
        row = ["Alice", 30, 50000]
        headers = ["Name", "Age", "Salary"]
        criteria = [["Age"], [">=30"]]
        assert _matches_criteria(row, headers, criteria) is True

    def test_less_equal(self):
        """Test <= comparison."""
        row = ["Alice", 30, 50000]
        headers = ["Name", "Age", "Salary"]
        criteria = [["Age"], ["<=30"]]
        assert _matches_criteria(row, headers, criteria) is True

    def test_not_equal(self):
        """Test <> comparison."""
        row = ["Alice", 30, 50000]
        headers = ["Name", "Age", "Salary"]
        criteria = [["Name"], ["<>Bob"]]
        assert _matches_criteria(row, headers, criteria) is True

    def test_equals_prefix(self):
        """Test = prefix for exact match."""
        row = ["Alice", 30, 50000]
        headers = ["Name", "Age", "Salary"]
        criteria = [["Name"], ["=Alice"]]
        assert _matches_criteria(row, headers, criteria) is True

    def test_wildcard_asterisk(self):
        """Test * wildcard."""
        row = ["Alice", 30, 50000]
        headers = ["Name", "Age", "Salary"]
        criteria = [["Name"], ["A*"]]
        assert _matches_criteria(row, headers, criteria) is True

    def test_wildcard_question(self):
        """Test ? wildcard."""
        row = ["Alice", 30, 50000]
        headers = ["Name", "Age", "Salary"]
        criteria = [["Name"], ["Alic?"]]
        assert _matches_criteria(row, headers, criteria) is True

    def test_and_conditions(self):
        """Test AND conditions (multiple columns in same row)."""
        row = ["Alice", 30, 50000]
        headers = ["Name", "Age", "Salary"]
        criteria = [["Name", "Age"], ["Alice", ">25"]]
        assert _matches_criteria(row, headers, criteria) is True

    def test_or_conditions(self):
        """Test OR conditions (multiple criteria rows)."""
        row = ["Bob", 25, 40000]
        headers = ["Name", "Age", "Salary"]
        criteria = [["Name"], ["Alice"], ["Bob"]]  # Alice OR Bob
        assert _matches_criteria(row, headers, criteria) is True


class TestDSUM:
    """Tests for DSUM function."""

    def setup_method(self):
        """Set up test database."""
        self.db = [
            ["Name", "Dept", "Salary"],
            ["Alice", "Sales", 50000],
            ["Bob", "Sales", 45000],
            ["Carol", "IT", 60000],
            ["Dave", "IT", 55000],
        ]

    def test_sum_all(self):
        """Test sum without criteria."""
        result = fn_dsum(self.db, "Salary", [])
        assert result == 210000

    def test_sum_with_criteria(self):
        """Test sum with criteria."""
        criteria = [["Dept"], ["Sales"]]
        result = fn_dsum(self.db, "Salary", criteria)
        assert result == 95000

    def test_sum_by_field_number(self):
        """Test sum using field number."""
        criteria = [["Dept"], ["IT"]]
        result = fn_dsum(self.db, 3, criteria)
        assert result == 115000

    def test_sum_non_list(self):
        """Test sum with non-list database."""
        result = fn_dsum("not a list", "Salary", [])
        assert result == 0.0


class TestDAVG:
    """Tests for DAVG function."""

    def setup_method(self):
        """Set up test database."""
        self.db = [
            ["Name", "Dept", "Salary"],
            ["Alice", "Sales", 50000],
            ["Bob", "Sales", 40000],
            ["Carol", "IT", 60000],
        ]

    def test_avg_all(self):
        """Test average without criteria."""
        result = fn_davg(self.db, "Salary", [])
        assert result == 50000

    def test_avg_with_criteria(self):
        """Test average with criteria."""
        criteria = [["Dept"], ["Sales"]]
        result = fn_davg(self.db, "Salary", criteria)
        assert result == 45000

    def test_avg_no_matches(self):
        """Test average with no matches."""
        criteria = [["Dept"], ["HR"]]
        result = fn_davg(self.db, "Salary", criteria)
        assert result == 0.0

    def test_avg_non_list(self):
        """Test avg with non-list database."""
        result = fn_davg("not a list", "Salary", [])
        assert result == 0.0


class TestDCOUNT:
    """Tests for DCOUNT function."""

    def setup_method(self):
        """Set up test database."""
        self.db = [
            ["Name", "Dept", "Salary"],
            ["Alice", "Sales", 50000],
            ["Bob", "Sales", 45000],
            ["Carol", "IT", 60000],
        ]

    def test_count_all(self):
        """Test count without criteria."""
        result = fn_dcount(self.db, "Salary", [])
        assert result == 3

    def test_count_with_criteria(self):
        """Test count with criteria."""
        criteria = [["Dept"], ["Sales"]]
        result = fn_dcount(self.db, "Salary", criteria)
        assert result == 2

    def test_count_non_list(self):
        """Test count with non-list database."""
        result = fn_dcount("not a list", "Salary", [])
        assert result == 0


class TestDCOUNTA:
    """Tests for DCOUNTA function."""

    def setup_method(self):
        """Set up test database."""
        self.db = [
            ["Name", "Dept", "Salary"],
            ["Alice", "Sales", 50000],
            ["Bob", "Sales", ""],
            ["Carol", "IT", 60000],
        ]

    def test_counta_all(self):
        """Test counta counts non-blank."""
        result = fn_dcounta(self.db, "Salary", [])
        assert result == 2  # Bob's salary is blank

    def test_counta_non_list(self):
        """Test counta with non-list database."""
        result = fn_dcounta("not a list", "Salary", [])
        assert result == 0


class TestDMINDMAX:
    """Tests for DMIN and DMAX functions."""

    def setup_method(self):
        """Set up test database."""
        self.db = [
            ["Name", "Dept", "Salary"],
            ["Alice", "Sales", 50000],
            ["Bob", "Sales", 45000],
            ["Carol", "IT", 60000],
        ]

    def test_min_all(self):
        """Test min without criteria."""
        result = fn_dmin(self.db, "Salary", [])
        assert result == 45000

    def test_max_all(self):
        """Test max without criteria."""
        result = fn_dmax(self.db, "Salary", [])
        assert result == 60000

    def test_min_with_criteria(self):
        """Test min with criteria."""
        criteria = [["Dept"], ["Sales"]]
        result = fn_dmin(self.db, "Salary", criteria)
        assert result == 45000

    def test_max_with_criteria(self):
        """Test max with criteria."""
        criteria = [["Dept"], ["IT"]]
        result = fn_dmax(self.db, "Salary", criteria)
        assert result == 60000

    def test_min_no_matches(self):
        """Test min with no matches."""
        criteria = [["Dept"], ["HR"]]
        result = fn_dmin(self.db, "Salary", criteria)
        assert result == 0.0

    def test_max_non_list(self):
        """Test max with non-list database."""
        result = fn_dmax("not a list", "Salary", [])
        assert result == 0.0


class TestDSTD:
    """Tests for DSTD function."""

    def setup_method(self):
        """Set up test database."""
        self.db = [
            ["Name", "Value"],
            ["A", 10],
            ["B", 20],
            ["C", 30],
        ]

    def test_std_sample(self):
        """Test sample standard deviation."""
        result = fn_dstd(self.db, "Value", [])
        assert abs(result - 10) < 0.01  # sqrt(100) = 10

    def test_std_population(self):
        """Test population standard deviation."""
        result = fn_dstdp(self.db, "Value", [])
        assert result < 10  # Population std is smaller

    def test_std_single_value(self):
        """Test std with single value returns 0."""
        db = [["Name", "Value"], ["A", 10]]
        result = fn_dstd(db, "Value", [])
        assert result == 0.0

    def test_std_non_list(self):
        """Test std with non-list database."""
        result = fn_dstd("not a list", "Value", [])
        assert result == 0.0


class TestDVAR:
    """Tests for DVAR function."""

    def setup_method(self):
        """Set up test database."""
        self.db = [
            ["Name", "Value"],
            ["A", 10],
            ["B", 20],
            ["C", 30],
        ]

    def test_var_sample(self):
        """Test sample variance."""
        result = fn_dvar(self.db, "Value", [])
        assert abs(result - 100) < 0.01

    def test_var_population(self):
        """Test population variance."""
        result = fn_dvarp(self.db, "Value", [])
        assert result < 100

    def test_var_single_value(self):
        """Test var with single value returns 0."""
        db = [["Name", "Value"], ["A", 10]]
        result = fn_dvar(db, "Value", [])
        assert result == 0.0


class TestDGET:
    """Tests for DGET function."""

    def setup_method(self):
        """Set up test database."""
        self.db = [
            ["Name", "Dept", "Salary"],
            ["Alice", "Sales", 50000],
            ["Bob", "IT", 45000],
            ["Carol", "IT", 60000],
        ]

    def test_get_single_match(self):
        """Test getting single matching value."""
        criteria = [["Name"], ["Alice"]]
        result = fn_dget(self.db, "Salary", criteria)
        assert result == 50000

    def test_get_multiple_matches(self):
        """Test error on multiple matches."""
        criteria = [["Dept"], ["IT"]]
        result = fn_dget(self.db, "Salary", criteria)
        assert result == "#NUM!"

    def test_get_no_match(self):
        """Test error on no match."""
        criteria = [["Name"], ["Dave"]]
        result = fn_dget(self.db, "Salary", criteria)
        assert result == "#VALUE!"

    def test_get_invalid_field(self):
        """Test error on invalid field."""
        criteria = [["Name"], ["Alice"]]
        result = fn_dget(self.db, "Invalid", criteria)
        assert result == "#VALUE!"

    def test_get_non_list(self):
        """Test dget with non-list database."""
        result = fn_dget("not a list", "Salary", [])
        assert result == "#VALUE!"


class TestFunctionRegistry:
    """Test the function registry."""

    def test_all_functions_registered(self):
        """Test that all functions are in the registry."""
        assert "DSUM" in DATABASE_FUNCTIONS
        assert "DAVG" in DATABASE_FUNCTIONS
        assert "DAVERAGE" in DATABASE_FUNCTIONS  # Alias
        assert "DCOUNT" in DATABASE_FUNCTIONS
        assert "DCOUNTA" in DATABASE_FUNCTIONS
        assert "DMIN" in DATABASE_FUNCTIONS
        assert "DMAX" in DATABASE_FUNCTIONS
        assert "DSTD" in DATABASE_FUNCTIONS
        assert "DSTDEV" in DATABASE_FUNCTIONS  # Alias
        assert "DSTDP" in DATABASE_FUNCTIONS
        assert "DVAR" in DATABASE_FUNCTIONS
        assert "DVARP" in DATABASE_FUNCTIONS
        assert "DGET" in DATABASE_FUNCTIONS

    def test_functions_callable(self):
        """Test that all registered functions are callable."""
        for name, func in DATABASE_FUNCTIONS.items():
            assert callable(func), f"{name} is not callable"

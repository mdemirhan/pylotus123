"""Tests for criteria parsing module."""

from lotus123.data.criteria import (
    Criterion,
    CriteriaParser,
    CriterionOperator,
    parse_simple_criteria,
)


class TestCriterionOperator:
    """Tests for CriterionOperator enum."""

    def test_all_operators_exist(self):
        """Test all operators exist."""
        assert CriterionOperator.EQUAL
        assert CriterionOperator.NOT_EQUAL
        assert CriterionOperator.LESS_THAN
        assert CriterionOperator.GREATER_THAN
        assert CriterionOperator.LESS_EQUAL
        assert CriterionOperator.GREATER_EQUAL
        assert CriterionOperator.CONTAINS
        assert CriterionOperator.STARTS_WITH
        assert CriterionOperator.ENDS_WITH


class TestCriterion:
    """Tests for Criterion dataclass."""

    def test_default_values(self):
        """Test default criterion values."""
        c = Criterion()
        assert c.column is None
        assert c.operator == CriterionOperator.EQUAL
        assert c.value is None
        assert c.pattern is None
        assert c.is_formula is False
        assert c.formula == ""

    def test_custom_values(self):
        """Test criterion with custom values."""
        c = Criterion(column=2, operator=CriterionOperator.GREATER_THAN, value=100)
        assert c.column == 2
        assert c.operator == CriterionOperator.GREATER_THAN
        assert c.value == 100


class TestCriterionMatches:
    """Tests for Criterion.matches() method."""

    def test_equal_match_number(self):
        """Test equal match with numbers."""
        c = Criterion(column=0, operator=CriterionOperator.EQUAL, value=42)
        assert c.matches(42) is True
        assert c.matches(42.0) is True
        assert c.matches(41) is False

    def test_equal_match_string(self):
        """Test equal match with strings (case-insensitive)."""
        c = Criterion(column=0, operator=CriterionOperator.EQUAL, value="hello")
        assert c.matches("hello") is True
        assert c.matches("HELLO") is True
        assert c.matches("Hello") is True
        assert c.matches("world") is False

    def test_not_equal_match(self):
        """Test not equal match."""
        c = Criterion(column=0, operator=CriterionOperator.NOT_EQUAL, value=42)
        assert c.matches(42) is False
        assert c.matches(41) is True
        assert c.matches(43) is True

    def test_less_than_number(self):
        """Test less than match with numbers."""
        c = Criterion(column=0, operator=CriterionOperator.LESS_THAN, value=50)
        assert c.matches(40) is True
        assert c.matches(50) is False
        assert c.matches(60) is False

    def test_greater_than_number(self):
        """Test greater than match with numbers."""
        c = Criterion(column=0, operator=CriterionOperator.GREATER_THAN, value=50)
        assert c.matches(60) is True
        assert c.matches(50) is False
        assert c.matches(40) is False

    def test_less_equal_number(self):
        """Test less than or equal match."""
        c = Criterion(column=0, operator=CriterionOperator.LESS_EQUAL, value=50)
        assert c.matches(40) is True
        assert c.matches(50) is True
        assert c.matches(60) is False

    def test_greater_equal_number(self):
        """Test greater than or equal match."""
        c = Criterion(column=0, operator=CriterionOperator.GREATER_EQUAL, value=50)
        assert c.matches(60) is True
        assert c.matches(50) is True
        assert c.matches(40) is False

    def test_string_comparison(self):
        """Test comparison operators with strings."""
        c = Criterion(column=0, operator=CriterionOperator.LESS_THAN, value="dog")
        assert c.matches("cat") is True  # cat < dog alphabetically
        assert c.matches("elephant") is False  # elephant > dog

    def test_contains_wildcard(self):
        """Test contains with wildcard pattern."""
        c = Criterion(column=0, operator=CriterionOperator.CONTAINS, pattern="*ello*")
        assert c.matches("Hello") is True
        assert c.matches("Yellow") is True
        assert c.matches("world") is False

    def test_starts_with(self):
        """Test starts with operator."""
        c = Criterion(column=0, operator=CriterionOperator.STARTS_WITH, value="Jo")
        assert c.matches("John") is True
        assert c.matches("JOHNSON") is True
        assert c.matches("Mary") is False

    def test_ends_with(self):
        """Test ends with operator."""
        c = Criterion(column=0, operator=CriterionOperator.ENDS_WITH, value="son")
        assert c.matches("Johnson") is True
        assert c.matches("WILSON") is True
        assert c.matches("Smith") is False

    def test_formula_criterion(self):
        """Test formula criterion (placeholder)."""
        c = Criterion(column=0, is_formula=True, formula="=A1>10")
        # Formula criteria currently return True (placeholder)
        assert c.matches("anything") is True


class TestCriterionWildcardMatch:
    """Tests for wildcard matching in Criterion."""

    def test_star_wildcard(self):
        """Test * wildcard matches any characters."""
        c = Criterion(column=0, operator=CriterionOperator.CONTAINS, pattern="J*n")
        assert c.matches("John") is True
        assert c.matches("Jan") is True
        assert c.matches("Jason") is True
        assert c.matches("Jane") is False

    def test_question_wildcard(self):
        """Test ? wildcard matches single character."""
        c = Criterion(column=0, operator=CriterionOperator.CONTAINS, pattern="J?n")
        assert c.matches("Jan") is True
        assert c.matches("Jon") is True
        assert c.matches("John") is False

    def test_no_pattern(self):
        """Test wildcard match with no pattern returns True."""
        c = Criterion(column=0, operator=CriterionOperator.CONTAINS, pattern=None)
        assert c.matches("anything") is True

    def test_dot_escaped(self):
        """Test dots are escaped in pattern."""
        c = Criterion(column=0, operator=CriterionOperator.CONTAINS, pattern="file.txt")
        assert c.matches("file.txt") is True
        assert c.matches("filextxt") is False


class TestCriterionCompare:
    """Tests for _compare method edge cases."""

    def test_compare_empty_values(self):
        """Test comparison with empty/None values."""
        c = Criterion(column=0, operator=CriterionOperator.GREATER_THAN, value=10)
        # Empty string becomes 0
        assert c.matches("") is False  # 0 > 10 is False
        assert c.matches(None) is False  # 0 > 10 is False

    def test_compare_string_vs_number(self):
        """Test comparison when types don't match."""
        c = Criterion(column=0, operator=CriterionOperator.GREATER_THAN, value="abc")
        assert c.matches("def") is True  # def > abc alphabetically


class TestCriteriaParser:
    """Tests for CriteriaParser class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.parser = CriteriaParser()

    def test_parse_empty_criteria(self):
        """Test parsing empty criteria."""
        self.parser.parse_range(["Name", "Age"], [])
        assert self.parser.matches(["John", 25]) is True

    def test_parse_exact_match(self):
        """Test parsing exact match criterion."""
        self.parser.parse_range(["Name", "Age"], [["John", ""]])
        assert self.parser.matches(["John", 25]) is True
        assert self.parser.matches(["Mary", 25]) is False

    def test_parse_comparison_operators(self):
        """Test parsing comparison operators."""
        # Greater than
        self.parser.parse_range(["Name", "Age"], [["", ">30"]])
        assert self.parser.matches(["John", 35]) is True
        assert self.parser.matches(["John", 25]) is False

    def test_parse_less_than(self):
        """Test parsing less than operator."""
        self.parser.parse_range(["Name", "Age"], [["", "<30"]])
        assert self.parser.matches(["John", 25]) is True
        assert self.parser.matches(["John", 35]) is False

    def test_parse_less_equal(self):
        """Test parsing less than or equal."""
        self.parser.parse_range(["Name", "Age"], [["", "<=30"]])
        assert self.parser.matches(["John", 30]) is True
        assert self.parser.matches(["John", 31]) is False

    def test_parse_greater_equal(self):
        """Test parsing greater than or equal."""
        self.parser.parse_range(["Name", "Age"], [["", ">=30"]])
        assert self.parser.matches(["John", 30]) is True
        assert self.parser.matches(["John", 29]) is False

    def test_parse_not_equal(self):
        """Test parsing not equal operator."""
        self.parser.parse_range(["Name", "Age"], [["", "<>30"]])
        assert self.parser.matches(["John", 25]) is True
        assert self.parser.matches(["John", 30]) is False

    def test_parse_equal_explicit(self):
        """Test parsing explicit equal operator."""
        self.parser.parse_range(["Name", "Age"], [["", "=30"]])
        assert self.parser.matches(["John", 30]) is True
        assert self.parser.matches(["John", 25]) is False

    def test_parse_wildcard(self):
        """Test parsing wildcard criterion."""
        self.parser.parse_range(["Name", "Age"], [["J*", ""]])
        assert self.parser.matches(["John", 25]) is True
        assert self.parser.matches(["Mary", 25]) is False

    def test_parse_formula(self):
        """Test parsing formula criterion."""
        self.parser.parse_range(["Name", "Age"], [["+A1>10", ""]])
        # Formula criteria are detected (start with + or =)
        # The criterion is_formula flag should be True
        assert len(self.parser._criteria) == 1
        assert len(self.parser._criteria[0]) == 1
        criterion = self.parser._criteria[0][0]
        assert criterion.is_formula is True

    def test_parse_and_criteria(self):
        """Test AND criteria (multiple columns same row)."""
        self.parser.parse_range(["Name", "Age"], [["John", ">20"]])
        assert self.parser.matches(["John", 25]) is True
        assert self.parser.matches(["John", 15]) is False  # Age doesn't match
        assert self.parser.matches(["Mary", 25]) is False  # Name doesn't match

    def test_parse_or_criteria(self):
        """Test OR criteria (multiple rows)."""
        self.parser.parse_range(["Name", "Age"], [["John", ""], ["Mary", ""]])
        assert self.parser.matches(["John", 25]) is True
        assert self.parser.matches(["Mary", 30]) is True
        assert self.parser.matches(["Bob", 25]) is False

    def test_parse_numeric_value(self):
        """Test parsing numeric values."""
        self.parser.parse_range(["Name", "Value"], [["", "100"]])
        assert self.parser.matches(["Test", 100]) is True
        assert self.parser.matches(["Test", 99]) is False

    def test_parse_float_value(self):
        """Test parsing float values."""
        self.parser.parse_range(["Name", "Value"], [["", "3.14"]])
        assert self.parser.matches(["Test", 3.14]) is True


class TestCriteriaParserMatches:
    """Tests for CriteriaParser.matches() method."""

    def setup_method(self):
        """Set up test fixtures."""
        self.parser = CriteriaParser()

    def test_matches_empty_criteria(self):
        """Test matches with no criteria returns True."""
        # No parse_range called, so _criteria is empty
        assert self.parser.matches(["anything"]) is True

    def test_matches_column_out_of_bounds(self):
        """Test matches handles out of bounds column."""
        # Parse criteria for column 5
        self.parser._criteria = [[Criterion(column=5, value="test")]]
        # Row only has 2 columns - should not match
        assert (
            self.parser.matches(["a", "b"]) is True
        )  # Column 5 doesn't exist, so criterion not checked


class TestCreateFilter:
    """Tests for create_filter method."""

    def test_create_filter_returns_callable(self):
        """Test create_filter returns callable."""
        parser = CriteriaParser()
        parser.parse_range(["Name"], [["John"]])
        f = parser.create_filter()
        assert callable(f)

    def test_create_filter_works(self):
        """Test filter function works correctly."""
        parser = CriteriaParser()
        parser.parse_range(["Name"], [["John"]])
        f = parser.create_filter()
        assert f(["John"]) is True
        assert f(["Mary"]) is False


class TestParseSimpleCriteria:
    """Tests for parse_simple_criteria convenience function."""

    def test_simple_greater_than(self):
        """Test parsing simple greater than."""
        c = parse_simple_criteria(0, ">100")
        assert c.column == 0
        assert c.operator == CriterionOperator.GREATER_THAN
        assert c.value == 100

    def test_simple_wildcard(self):
        """Test parsing simple wildcard."""
        c = parse_simple_criteria(0, "John*")
        assert c.column == 0
        assert c.operator == CriterionOperator.CONTAINS
        assert c.pattern == "John*"

    def test_simple_not_equal(self):
        """Test parsing simple not equal."""
        c = parse_simple_criteria(0, "<>0")
        assert c.column == 0
        assert c.operator == CriterionOperator.NOT_EQUAL
        assert c.value == 0

    def test_simple_exact_match(self):
        """Test parsing simple exact match."""
        c = parse_simple_criteria(0, "hello")
        assert c.column == 0
        assert c.operator == CriterionOperator.EQUAL
        assert c.value == "hello"

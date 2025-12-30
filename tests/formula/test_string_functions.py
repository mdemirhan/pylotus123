"""Tests for string functions."""

import pytest

from lotus123.formula.functions.string import (
    _to_int,
    _to_string,
    fn_char,
    fn_clean,
    fn_code,
    fn_concat,
    fn_concatenate,
    fn_dollar,
    fn_exact,
    fn_find,
    fn_fixed,
    fn_left,
    fn_len,
    fn_length,
    fn_lower,
    fn_mid,
    fn_n,
    fn_proper,
    fn_repeat,
    fn_replace,
    fn_rept,
    fn_right,
    fn_s,
    fn_search,
    fn_string,
    fn_substitute,
    fn_t,
    fn_text,
    fn_trim,
    fn_upper,
    fn_value,
)


class TestHelperFunctions:
    """Tests for helper functions."""

    def test_to_string_none(self):
        """Test converting None."""
        assert _to_string(None) == ""

    def test_to_string_empty(self):
        """Test converting empty string."""
        assert _to_string("") == ""

    def test_to_string_int_float(self):
        """Test converting float that equals int."""
        assert _to_string(5.0) == "5"

    def test_to_string_float(self):
        """Test converting float."""
        assert _to_string(3.14) == "3.14"

    def test_to_int_int(self):
        """Test converting int."""
        assert _to_int(42) == 42

    def test_to_int_float(self):
        """Test converting float."""
        assert _to_int(3.7) == 3

    def test_to_int_string(self):
        """Test converting string."""
        assert _to_int("42") == 42

    def test_to_int_invalid(self):
        """Test converting invalid returns 0."""
        assert _to_int("abc") == 0


class TestExtractionFunctions:
    """Tests for extraction functions."""

    def test_fn_left_basic(self):
        """Test LEFT basic usage."""
        assert fn_left("Hello", 3) == "Hel"

    def test_fn_left_default(self):
        """Test LEFT default num_chars."""
        assert fn_left("Hello") == "H"

    def test_fn_left_zero(self):
        """Test LEFT with zero."""
        assert fn_left("Hello", 0) == ""

    def test_fn_left_negative(self):
        """Test LEFT with negative."""
        assert fn_left("Hello", -1) == ""

    def test_fn_left_overflow(self):
        """Test LEFT with num > length."""
        assert fn_left("Hi", 10) == "Hi"

    def test_fn_right_basic(self):
        """Test RIGHT basic usage."""
        assert fn_right("Hello", 3) == "llo"

    def test_fn_right_default(self):
        """Test RIGHT default num_chars."""
        assert fn_right("Hello") == "o"

    def test_fn_right_zero(self):
        """Test RIGHT with zero."""
        assert fn_right("Hello", 0) == ""

    def test_fn_mid_basic(self):
        """Test MID basic usage."""
        assert fn_mid("Hello", 2, 3) == "ell"

    def test_fn_mid_start_at_one(self):
        """Test MID with start at 1."""
        assert fn_mid("Hello", 1, 2) == "He"

    def test_fn_mid_zero_length(self):
        """Test MID with zero length."""
        assert fn_mid("Hello", 2, 0) == ""


class TestLengthFunctions:
    """Tests for length functions."""

    def test_fn_length_basic(self):
        """Test LENGTH basic usage."""
        assert fn_length("Hello") == 5

    def test_fn_length_empty(self):
        """Test LENGTH with empty string."""
        assert fn_length("") == 0

    def test_fn_length_number(self):
        """Test LENGTH with number."""
        assert fn_length(12345) == 5

    def test_fn_len_alias(self):
        """Test LEN is alias for LENGTH."""
        assert fn_len("Hello") == fn_length("Hello")


class TestSearchFunctions:
    """Tests for search functions."""

    def test_fn_find_basic(self):
        """Test FIND basic usage."""
        assert fn_find("ll", "Hello") == 3

    def test_fn_find_not_found(self):
        """Test FIND when not found."""
        assert fn_find("x", "Hello") == 0

    def test_fn_find_case_sensitive(self):
        """Test FIND is case-sensitive."""
        assert fn_find("H", "Hello") == 1
        assert fn_find("h", "Hello") == 0

    def test_fn_find_with_start(self):
        """Test FIND with start position."""
        assert fn_find("l", "Hello", 4) == 4

    def test_fn_search_basic(self):
        """Test SEARCH basic usage."""
        assert fn_search("ll", "Hello") == 3

    def test_fn_search_case_insensitive(self):
        """Test SEARCH is case-insensitive."""
        assert fn_search("h", "Hello") == 1

    def test_fn_search_wildcard_question(self):
        """Test SEARCH with ? wildcard."""
        assert fn_search("H?llo", "Hello") == 1

    def test_fn_search_wildcard_star(self):
        """Test SEARCH with * wildcard."""
        assert fn_search("H*o", "Hello") == 1


class TestReplacementFunctions:
    """Tests for replacement functions."""

    def test_fn_replace_basic(self):
        """Test REPLACE basic usage."""
        assert fn_replace("Hello", 2, 3, "XY") == "HXYo"

    def test_fn_replace_at_start(self):
        """Test REPLACE at start."""
        assert fn_replace("Hello", 1, 2, "J") == "Jllo"

    def test_fn_replace_zero_chars(self):
        """Test REPLACE with zero chars (insert)."""
        assert fn_replace("Hello", 3, 0, "XXX") == "HeXXXllo"

    def test_fn_substitute_all(self):
        """Test SUBSTITUTE replaces all."""
        assert fn_substitute("banana", "a", "o") == "bonono"

    def test_fn_substitute_instance(self):
        """Test SUBSTITUTE specific instance."""
        assert fn_substitute("banana", "a", "o", 2) == "banona"

    def test_fn_substitute_invalid_instance(self):
        """Test SUBSTITUTE with invalid instance."""
        assert fn_substitute("banana", "a", "o", 0) == "banana"


class TestCaseFunctions:
    """Tests for case conversion functions."""

    def test_fn_upper(self):
        """Test UPPER function."""
        assert fn_upper("Hello") == "HELLO"

    def test_fn_lower(self):
        """Test LOWER function."""
        assert fn_lower("Hello") == "hello"

    def test_fn_proper(self):
        """Test PROPER function."""
        assert fn_proper("hello world") == "Hello World"


class TestCleaningFunctions:
    """Tests for cleaning functions."""

    def test_fn_trim_spaces(self):
        """Test TRIM removes extra spaces."""
        assert fn_trim("  Hello   World  ") == "Hello World"

    def test_fn_trim_multiple_internal(self):
        """Test TRIM reduces multiple internal spaces."""
        assert fn_trim("Hello    World") == "Hello World"

    def test_fn_clean_printable(self):
        """Test CLEAN removes non-printable."""
        assert fn_clean("Hello\x00World") == "HelloWorld"

    def test_fn_clean_keeps_newlines(self):
        """Test CLEAN keeps tabs and newlines."""
        assert fn_clean("Hello\tWorld\n") == "Hello\tWorld\n"


class TestConversionFunctions:
    """Tests for conversion functions."""

    def test_fn_value_basic(self):
        """Test VALUE basic conversion."""
        assert fn_value("123") == 123.0

    def test_fn_value_float(self):
        """Test VALUE with float string."""
        assert fn_value("3.14") == 3.14

    def test_fn_value_percentage(self):
        """Test VALUE with percentage."""
        assert fn_value("50%") == 0.5

    def test_fn_value_currency(self):
        """Test VALUE with currency."""
        assert fn_value("$1,234.56") == 1234.56

    def test_fn_value_invalid(self):
        """Test VALUE with invalid string."""
        assert fn_value("abc") == 0.0

    def test_fn_string_basic(self):
        """Test STRING basic usage."""
        assert fn_string(3.14159, 2) == "3.14"

    def test_fn_string_no_decimals(self):
        """Test STRING with no decimals."""
        assert fn_string(3.7, 0) == "4"

    def test_fn_text(self):
        """Test TEXT function."""
        assert fn_text(42) == "42"

    def test_fn_char_basic(self):
        """Test CHAR basic usage."""
        assert fn_char(65) == "A"
        assert fn_char(97) == "a"

    def test_fn_char_invalid(self):
        """Test CHAR with invalid code."""
        assert fn_char(-1) == ""

    def test_fn_code_basic(self):
        """Test CODE basic usage."""
        assert fn_code("A") == 65
        assert fn_code("a") == 97

    def test_fn_code_empty(self):
        """Test CODE with empty string."""
        assert fn_code("") == 0

    def test_fn_n_number(self):
        """Test N with number."""
        assert fn_n(42) == 42.0

    def test_fn_n_text(self):
        """Test N with text returns 0."""
        assert fn_n("hello") == 0.0

    def test_fn_s_text(self):
        """Test S with text."""
        assert fn_s("hello") == "hello"

    def test_fn_s_number(self):
        """Test S with number returns empty."""
        assert fn_s(42) == ""

    def test_fn_t_text(self):
        """Test T with text."""
        assert fn_t("hello") == "hello"

    def test_fn_t_number(self):
        """Test T with number returns empty."""
        assert fn_t(42) == ""


class TestRepetitionFunctions:
    """Tests for repetition functions."""

    def test_fn_repeat_basic(self):
        """Test REPEAT basic usage."""
        assert fn_repeat("ab", 3) == "ababab"

    def test_fn_repeat_zero(self):
        """Test REPEAT with zero."""
        assert fn_repeat("ab", 0) == ""

    def test_fn_repeat_negative(self):
        """Test REPEAT with negative."""
        assert fn_repeat("ab", -1) == ""

    def test_fn_rept_alias(self):
        """Test REPT is alias for REPEAT."""
        assert fn_rept("ab", 3) == fn_repeat("ab", 3)


class TestComparisonFunctions:
    """Tests for comparison functions."""

    def test_fn_exact_same(self):
        """Test EXACT with same strings."""
        assert fn_exact("Hello", "Hello") is True

    def test_fn_exact_different(self):
        """Test EXACT with different strings."""
        assert fn_exact("Hello", "hello") is False

    def test_fn_exact_empty(self):
        """Test EXACT with empty strings."""
        assert fn_exact("", "") is True


class TestConcatenationFunctions:
    """Tests for concatenation functions."""

    def test_fn_concatenate_basic(self):
        """Test CONCATENATE basic usage."""
        assert fn_concatenate("Hello", " ", "World") == "Hello World"

    def test_fn_concatenate_numbers(self):
        """Test CONCATENATE with numbers."""
        assert fn_concatenate("Value: ", 42) == "Value: 42"

    def test_fn_concat_alias(self):
        """Test CONCAT is alias."""
        assert fn_concat("a", "b") == fn_concatenate("a", "b")


class TestFormattingFunctions:
    """Tests for formatting functions."""

    def test_fn_fixed_basic(self):
        """Test FIXED basic usage."""
        assert fn_fixed(1234.567, 2) == "1,234.57"

    def test_fn_fixed_no_commas(self):
        """Test FIXED without commas."""
        assert fn_fixed(1234.567, 2, True) == "1234.57"

    def test_fn_dollar_basic(self):
        """Test DOLLAR basic usage."""
        assert fn_dollar(1234.567, 2) == "$1,234.57"

    def test_fn_dollar_no_decimals(self):
        """Test DOLLAR with no decimals."""
        assert fn_dollar(1234.567, 0) == "$1,235"

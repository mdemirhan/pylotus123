"""Tests for logical functions."""

from lotus123.formula.functions.logical import (
    _flatten_args,
    _to_bool,
    fn_and,
    fn_choose,
    fn_err,
    fn_false,
    fn_if,
    fn_iferror,
    fn_ifna,
    fn_isblank,
    fn_iserr,
    fn_iserror,
    fn_iseven,
    fn_islogical,
    fn_isna,
    fn_isnumber,
    fn_isodd,
    fn_isref,
    fn_isstring,
    fn_istext,
    fn_na,
    fn_not,
    fn_or,
    fn_switch,
    fn_true,
    fn_xor,
)


class TestHelperFunctions:
    """Tests for helper functions."""

    def test_flatten_args_simple(self):
        """Test flattening simple args."""
        result = _flatten_args((1, 2, 3))
        assert result == [1, 2, 3]

    def test_flatten_args_nested(self):
        """Test flattening nested lists."""
        result = _flatten_args((1, [2, 3], 4))
        assert result == [1, 2, 3, 4]

    def test_flatten_args_deeply_nested(self):
        """Test flattening deeply nested lists."""
        result = _flatten_args(([1, [2, 3]], 4))
        assert result == [1, 2, 3, 4]

    def test_to_bool_bool(self):
        """Test converting bool."""
        assert _to_bool(True) is True
        assert _to_bool(False) is False

    def test_to_bool_number(self):
        """Test converting numbers."""
        assert _to_bool(0) is False
        assert _to_bool(1) is True
        assert _to_bool(-1) is True
        assert _to_bool(0.5) is True

    def test_to_bool_string_true(self):
        """Test converting 'TRUE' string."""
        assert _to_bool("TRUE") is True
        assert _to_bool("true") is True
        assert _to_bool("True") is True

    def test_to_bool_string_false(self):
        """Test converting 'FALSE' string."""
        assert _to_bool("FALSE") is False
        assert _to_bool("false") is False

    def test_to_bool_numeric_string(self):
        """Test converting numeric string."""
        assert _to_bool("0") is False
        assert _to_bool("1") is True
        assert _to_bool("3.14") is True

    def test_to_bool_text_string(self):
        """Test converting text string."""
        assert _to_bool("hello") is True  # Non-empty string
        assert _to_bool("") is False  # Empty string

    def test_to_bool_none(self):
        """Test converting None."""
        assert _to_bool(None) is False


class TestCoreFunctions:
    """Tests for core logical functions."""

    def test_fn_if_true(self):
        """Test IF with true condition."""
        result = fn_if(True, "yes", "no")
        assert result == "yes"

    def test_fn_if_false(self):
        """Test IF with false condition."""
        result = fn_if(False, "yes", "no")
        assert result == "no"

    def test_fn_if_numeric_condition(self):
        """Test IF with numeric condition."""
        assert fn_if(1, "yes", "no") == "yes"
        assert fn_if(0, "yes", "no") == "no"

    def test_fn_if_default_false_value(self):
        """Test IF with default false value."""
        result = fn_if(False, "yes")
        assert result == ""

    def test_fn_true(self):
        """Test TRUE function."""
        assert fn_true() is True

    def test_fn_false(self):
        """Test FALSE function."""
        assert fn_false() is False

    def test_fn_and_all_true(self):
        """Test AND with all true."""
        assert fn_and(True, True, True) is True

    def test_fn_and_one_false(self):
        """Test AND with one false."""
        assert fn_and(True, False, True) is False

    def test_fn_and_empty(self):
        """Test AND with no args."""
        assert fn_and() is True

    def test_fn_and_nested_list(self):
        """Test AND with nested list."""
        assert fn_and([True, True], True) is True
        assert fn_and([True, False], True) is False

    def test_fn_or_all_false(self):
        """Test OR with all false."""
        assert fn_or(False, False, False) is False

    def test_fn_or_one_true(self):
        """Test OR with one true."""
        assert fn_or(False, True, False) is True

    def test_fn_or_empty(self):
        """Test OR with no args."""
        assert fn_or() is False

    def test_fn_or_nested_list(self):
        """Test OR with nested list."""
        assert fn_or([False, True], False) is True
        assert fn_or([False, False], False) is False

    def test_fn_not_true(self):
        """Test NOT with true."""
        assert fn_not(True) is False

    def test_fn_not_false(self):
        """Test NOT with false."""
        assert fn_not(False) is True

    def test_fn_not_number(self):
        """Test NOT with numbers."""
        assert fn_not(0) is True
        assert fn_not(1) is False

    def test_fn_xor_odd_true(self):
        """Test XOR with odd number of true."""
        assert fn_xor(True, False, False) is True
        assert fn_xor(True, True, True) is True

    def test_fn_xor_even_true(self):
        """Test XOR with even number of true."""
        assert fn_xor(True, True, False) is False
        assert fn_xor(False, False) is False


class TestErrorFunctions:
    """Tests for error-checking functions."""

    def test_fn_iserr_error(self):
        """Test ISERR with error."""
        assert fn_iserr("#ERR!") is True
        assert fn_iserr("#DIV/0!") is True
        assert fn_iserr("#CIRC!") is True

    def test_fn_iserr_na(self):
        """Test ISERR with #N/A (should be false)."""
        assert fn_iserr("#N/A") is False

    def test_fn_iserr_normal(self):
        """Test ISERR with normal value."""
        assert fn_iserr(123) is False
        assert fn_iserr("text") is False

    def test_fn_iserror_any_error(self):
        """Test ISERROR with any error."""
        assert fn_iserror("#ERR!") is True
        assert fn_iserror("#N/A") is True
        assert fn_iserror("#DIV/0!") is True

    def test_fn_iserror_normal(self):
        """Test ISERROR with normal value."""
        assert fn_iserror(123) is False
        assert fn_iserror("text") is False

    def test_fn_isna_na(self):
        """Test ISNA with #N/A."""
        assert fn_isna("#N/A") is True

    def test_fn_isna_other_error(self):
        """Test ISNA with other error."""
        assert fn_isna("#ERR!") is False

    def test_fn_isna_normal(self):
        """Test ISNA with normal value."""
        assert fn_isna(123) is False

    def test_fn_na(self):
        """Test NA function."""
        assert fn_na() == "#N/A"

    def test_fn_err(self):
        """Test ERR function."""
        assert fn_err() == "#ERR!"


class TestTypeFunctions:
    """Tests for type-checking functions."""

    def test_fn_isnumber_int(self):
        """Test ISNUMBER with int."""
        assert fn_isnumber(42) is True

    def test_fn_isnumber_float(self):
        """Test ISNUMBER with float."""
        assert fn_isnumber(3.14) is True

    def test_fn_isnumber_string(self):
        """Test ISNUMBER with numeric string."""
        assert fn_isnumber("123") is True
        assert fn_isnumber("3.14") is True

    def test_fn_isnumber_text(self):
        """Test ISNUMBER with text."""
        assert fn_isnumber("text") is False

    def test_fn_isnumber_bool(self):
        """Test ISNUMBER with bool (should be false)."""
        assert fn_isnumber(True) is False

    def test_fn_isstring_text(self):
        """Test ISSTRING with text."""
        assert fn_isstring("hello") is True

    def test_fn_isstring_error(self):
        """Test ISSTRING with error."""
        assert fn_isstring("#ERR!") is False

    def test_fn_isstring_numeric(self):
        """Test ISSTRING with numeric string."""
        assert fn_isstring("123") is False

    def test_fn_isstring_number(self):
        """Test ISSTRING with number."""
        assert fn_isstring(42) is False

    def test_fn_istext_alias(self):
        """Test ISTEXT is alias for ISSTRING."""
        assert fn_istext("hello") is True
        assert fn_istext(42) is False

    def test_fn_isblank_empty(self):
        """Test ISBLANK with empty."""
        assert fn_isblank(None) is True
        assert fn_isblank("") is True

    def test_fn_isblank_value(self):
        """Test ISBLANK with value."""
        assert fn_isblank("text") is False
        assert fn_isblank(0) is False

    def test_fn_islogical_bool(self):
        """Test ISLOGICAL with bool."""
        assert fn_islogical(True) is True
        assert fn_islogical(False) is True

    def test_fn_islogical_other(self):
        """Test ISLOGICAL with other types."""
        assert fn_islogical(1) is False
        assert fn_islogical("TRUE") is False

    def test_fn_iseven_even(self):
        """Test ISEVEN with even numbers."""
        assert fn_iseven(2) is True
        assert fn_iseven(0) is True
        assert fn_iseven(-4) is True

    def test_fn_iseven_odd(self):
        """Test ISEVEN with odd numbers."""
        assert fn_iseven(1) is False
        assert fn_iseven(3) is False

    def test_fn_iseven_invalid(self):
        """Test ISEVEN with invalid input."""
        assert fn_iseven("text") is False

    def test_fn_isodd_odd(self):
        """Test ISODD with odd numbers."""
        assert fn_isodd(1) is True
        assert fn_isodd(3) is True
        assert fn_isodd(-5) is True

    def test_fn_isodd_even(self):
        """Test ISODD with even numbers."""
        assert fn_isodd(2) is False
        assert fn_isodd(0) is False

    def test_fn_isodd_invalid(self):
        """Test ISODD with invalid input."""
        assert fn_isodd("text") is False

    def test_fn_isref(self):
        """Test ISREF (always returns False as placeholder)."""
        assert fn_isref("A1") is False


class TestConditionalFunctions:
    """Tests for conditional functions."""

    def test_fn_iferror_no_error(self):
        """Test IFERROR with no error."""
        assert fn_iferror(100, "error") == 100

    def test_fn_iferror_error(self):
        """Test IFERROR with error."""
        assert fn_iferror("#ERR!", "error") == "error"
        assert fn_iferror("#N/A", "error") == "error"

    def test_fn_ifna_no_na(self):
        """Test IFNA with no #N/A."""
        assert fn_ifna(100, "error") == 100
        assert fn_ifna("#ERR!", "error") == "#ERR!"

    def test_fn_ifna_na(self):
        """Test IFNA with #N/A."""
        assert fn_ifna("#N/A", "error") == "error"

    def test_fn_switch_match(self):
        """Test SWITCH with matching value."""
        result = fn_switch("B", "A", 1, "B", 2, "C", 3)
        assert result == 2

    def test_fn_switch_no_match(self):
        """Test SWITCH with no match and default."""
        result = fn_switch("X", "A", 1, "B", 2, "default")
        assert result == "default"

    def test_fn_switch_no_match_no_default(self):
        """Test SWITCH with no match and no default."""
        result = fn_switch("X", "A", 1, "B", 2)
        assert result == ""

    def test_fn_choose_valid_index(self):
        """Test CHOOSE with valid index."""
        assert fn_choose(1, "a", "b", "c") == "a"
        assert fn_choose(2, "a", "b", "c") == "b"
        assert fn_choose(3, "a", "b", "c") == "c"

    def test_fn_choose_out_of_range(self):
        """Test CHOOSE with out of range index."""
        assert fn_choose(0, "a", "b") == "#N/A"
        assert fn_choose(5, "a", "b") == "#N/A"

    def test_fn_choose_invalid_index(self):
        """Test CHOOSE with invalid index."""
        assert fn_choose("x", "a", "b") == "#N/A"

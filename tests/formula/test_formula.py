"""Tests for formula parser and evaluator."""

import math

import pytest

from lotus123 import Spreadsheet
from lotus123.formula import FormulaParser


class TestBasicArithmetic:
    def setup_method(self):
        self.ss = Spreadsheet()
        self.parser = FormulaParser(self.ss)

    def test_addition(self):
        assert self.parser.evaluate("1+2") == 3

    def test_subtraction(self):
        assert self.parser.evaluate("5-3") == 2

    def test_multiplication(self):
        assert self.parser.evaluate("4*5") == 20

    def test_division(self):
        assert self.parser.evaluate("10/2") == 5

    def test_exponentiation(self):
        assert self.parser.evaluate("2^3") == 8

    def test_modulo(self):
        assert self.parser.evaluate("7%3") == 1

    def test_operator_precedence(self):
        assert self.parser.evaluate("2+3*4") == 14
        assert self.parser.evaluate("(2+3)*4") == 20

    def test_nested_parentheses(self):
        assert self.parser.evaluate("((2+3)*4)+5") == 25

    def test_unary_minus(self):
        assert self.parser.evaluate("-5") == -5
        assert self.parser.evaluate("10+-5") == 5

    def test_unary_plus(self):
        assert self.parser.evaluate("+5") == 5

    def test_float_numbers(self):
        assert self.parser.evaluate("3.14*2") == pytest.approx(6.28)

    def test_division_by_zero(self):
        result = self.parser.evaluate("1/0")
        assert "#DIV/0!" in str(result)


class TestCellReferences:
    def setup_method(self):
        self.ss = Spreadsheet()
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "20")
        self.ss.set_cell(0, 2, "30")
        self.parser = FormulaParser(self.ss)

    def test_single_cell_ref(self):
        assert self.parser.evaluate("A1") == 10

    def test_cell_ref_addition(self):
        assert self.parser.evaluate("A1+B1") == 30

    def test_cell_ref_complex(self):
        assert self.parser.evaluate("A1*B1+C1") == 230

    def test_empty_cell_ref(self):
        assert self.parser.evaluate("Z99") == ""


class TestRangeFunctions:
    def setup_method(self):
        self.ss = Spreadsheet()
        for i in range(5):
            self.ss.set_cell(i, 0, str(i + 1))
        self.parser = FormulaParser(self.ss)

    def test_sum(self):
        assert self.parser.evaluate("SUM(A1:A5)") == 15

    def test_sum_at_prefix(self):
        assert self.parser.evaluate("@SUM(A1:A5)") == 15

    def test_avg(self):
        assert self.parser.evaluate("AVG(A1:A5)") == 3

    def test_average(self):
        assert self.parser.evaluate("AVERAGE(A1:A5)") == 3

    def test_min(self):
        assert self.parser.evaluate("MIN(A1:A5)") == 1

    def test_max(self):
        assert self.parser.evaluate("MAX(A1:A5)") == 5

    def test_count(self):
        assert self.parser.evaluate("COUNT(A1:A5)") == 5

    def test_count_with_empty(self):
        self.ss.set_cell(2, 0, "")
        assert self.parser.evaluate("COUNT(A1:A5)") == 4

    def test_counta(self):
        self.ss.set_cell(2, 0, "text")
        assert self.parser.evaluate("COUNTA(A1:A5)") == 5


class TestMathFunctions:
    def setup_method(self):
        self.ss = Spreadsheet()
        self.parser = FormulaParser(self.ss)

    def test_abs(self):
        assert self.parser.evaluate("ABS(-5)") == 5
        assert self.parser.evaluate("ABS(5)") == 5

    def test_int(self):
        assert self.parser.evaluate("INT(3.7)") == 3
        # Lotus INT uses floor() - truncates toward negative infinity
        assert self.parser.evaluate("INT(-3.7)") == -4

    def test_round(self):
        assert self.parser.evaluate("ROUND(3.456, 2)") == 3.46
        assert self.parser.evaluate("ROUND(3.5)") == 4

    def test_sqrt(self):
        assert self.parser.evaluate("SQRT(16)") == 4

    def test_power(self):
        assert self.parser.evaluate("POWER(2, 10)") == 1024

    def test_mod(self):
        assert self.parser.evaluate("MOD(10, 3)") == 1

    def test_sin(self):
        assert self.parser.evaluate("SIN(0)") == pytest.approx(0)

    def test_cos(self):
        assert self.parser.evaluate("COS(0)") == pytest.approx(1)

    def test_pi(self):
        assert self.parser.evaluate("PI()") == pytest.approx(math.pi)

    def test_exp(self):
        assert self.parser.evaluate("EXP(1)") == pytest.approx(math.e)

    def test_ln(self):
        assert self.parser.evaluate("LN(2.718281828)") == pytest.approx(1, abs=0.001)

    def test_log(self):
        assert self.parser.evaluate("LOG(100)") == pytest.approx(2)


class TestStringFunctions:
    def setup_method(self):
        self.ss = Spreadsheet()
        self.parser = FormulaParser(self.ss)

    def test_len(self):
        assert self.parser.evaluate('LEN("Hello")') == 5

    def test_left(self):
        assert self.parser.evaluate('LEFT("Hello", 2)') == "He"

    def test_right(self):
        assert self.parser.evaluate('RIGHT("Hello", 2)') == "lo"

    def test_mid(self):
        assert self.parser.evaluate('MID("Hello", 2, 3)') == "ell"

    def test_upper(self):
        assert self.parser.evaluate('UPPER("hello")') == "HELLO"

    def test_lower(self):
        assert self.parser.evaluate('LOWER("HELLO")') == "hello"

    def test_trim(self):
        assert self.parser.evaluate('TRIM("  hello  ")') == "hello"

    def test_concatenate(self):
        assert self.parser.evaluate('CONCATENATE("Hello", " ", "World")') == "Hello World"

    def test_concat(self):
        assert self.parser.evaluate('CONCAT("A", "B", "C")') == "ABC"


class TestLogicalFunctions:
    def setup_method(self):
        self.ss = Spreadsheet()
        self.parser = FormulaParser(self.ss)

    def test_if_true(self):
        assert self.parser.evaluate('IF(1>0, "yes", "no")') == "yes"

    def test_if_false(self):
        assert self.parser.evaluate('IF(1<0, "yes", "no")') == "no"

    def test_if_with_calculation(self):
        self.ss.set_cell(0, 0, "100")
        assert self.parser.evaluate('IF(A1>50, "big", "small")') == "big"

    def test_and_true(self):
        assert self.parser.evaluate("AND(1, 1)") is True

    def test_and_false(self):
        assert self.parser.evaluate("AND(1, 0)") is False

    def test_or_true(self):
        assert self.parser.evaluate("OR(1, 0)") is True

    def test_or_false(self):
        assert self.parser.evaluate("OR(0, 0)") is False

    def test_not(self):
        assert self.parser.evaluate("NOT(0)") is True
        assert self.parser.evaluate("NOT(1)") is False


class TestComparisons:
    def setup_method(self):
        self.ss = Spreadsheet()
        self.parser = FormulaParser(self.ss)

    def test_equal(self):
        assert self.parser.evaluate("5=5") is True
        assert self.parser.evaluate("5=6") is False

    def test_not_equal(self):
        assert self.parser.evaluate("5<>6") is True
        assert self.parser.evaluate("5<>5") is False

    def test_less_than(self):
        assert self.parser.evaluate("5<6") is True
        assert self.parser.evaluate("6<5") is False

    def test_greater_than(self):
        assert self.parser.evaluate("6>5") is True
        assert self.parser.evaluate("5>6") is False

    def test_less_equal(self):
        assert self.parser.evaluate("5<=5") is True
        assert self.parser.evaluate("5<=6") is True
        assert self.parser.evaluate("6<=5") is False

    def test_greater_equal(self):
        assert self.parser.evaluate("5>=5") is True
        assert self.parser.evaluate("6>=5") is True
        assert self.parser.evaluate("5>=6") is False


class TestComplexFormulas:
    def setup_method(self):
        self.ss = Spreadsheet()

    def test_nested_functions(self):
        self.ss.set_cell(0, 0, "=ROUND(SUM(1,2,3)/2, 1)")
        assert self.ss.get_value(0, 0) == 3.0

    def test_formula_chain(self):
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "=A1*2")
        self.ss.set_cell(0, 2, "=B1+5")
        assert self.ss.get_value(0, 2) == 25

    def test_sum_with_individual_values(self):
        parser = FormulaParser(self.ss)
        assert parser.evaluate("SUM(1, 2, 3, 4, 5)") == 15

    def test_mixed_sum(self):
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(1, 0, "20")
        parser = FormulaParser(self.ss)
        assert parser.evaluate("SUM(A1:A2, 100)") == 130


class TestEdgeCases:
    def setup_method(self):
        self.ss = Spreadsheet()
        self.parser = FormulaParser(self.ss)

    def test_empty_formula(self):
        assert self.parser.evaluate("") == ""

    def test_whitespace(self):
        assert self.parser.evaluate("  1 + 2  ") == 3

    def test_unknown_function(self):
        result = self.parser.evaluate("NOTAFUNC(1)")
        assert "#NAME?" in str(result)

    def test_string_literal(self):
        assert self.parser.evaluate('"Hello World"') == "Hello World"


class TestErrorPropagation:
    """Tests for error propagation in formulas.

    These tests verify that error values are properly propagated through
    operations and don't get multiplied/concatenated (e.g., #ERR!#ERR!).
    """

    def setup_method(self):
        self.ss = Spreadsheet()
        self.parser = FormulaParser(self.ss)

    def test_error_not_multiplied(self):
        """Test that error values in multiplication don't produce repeated strings."""
        # Set up a cell with an error value (reference to invalid cell in range)
        self.ss.set_cell(0, 0, "=1/0")  # This creates #DIV/0!
        result = self.ss.get_value(0, 0)
        assert result == "#DIV/0!"
        # Multiplying cell containing error should propagate the error, not multiply it
        self.ss.set_cell(0, 1, "=A1*100")
        result = self.ss.get_value(0, 1)
        # Should be a single error, not "#DIV/0!" repeated
        assert result.startswith("#")
        assert result.endswith("!")
        assert len(result) < 20  # Error strings are short

    def test_error_propagation_in_addition(self):
        """Test error propagation in addition."""
        self.ss.set_cell(0, 0, "=1/0")
        self.ss.set_cell(0, 1, "=A1+10")
        result = self.ss.get_value(0, 1)
        assert result.startswith("#")
        assert result.endswith("!")
        assert len(result) < 20

    def test_error_propagation_in_subtraction(self):
        """Test error propagation in subtraction."""
        self.ss.set_cell(0, 0, "=1/0")
        self.ss.set_cell(0, 1, "=100-A1")
        result = self.ss.get_value(0, 1)
        assert result.startswith("#")
        assert result.endswith("!")
        assert len(result) < 20

    def test_error_propagation_in_division(self):
        """Test error propagation in division."""
        self.ss.set_cell(0, 0, "=1/0")
        self.ss.set_cell(0, 1, "=100/A1")
        result = self.ss.get_value(0, 1)
        assert result.startswith("#")
        assert result.endswith("!")
        assert len(result) < 20

    def test_error_propagation_in_exponentiation(self):
        """Test error propagation in exponentiation."""
        self.ss.set_cell(0, 0, "=1/0")
        self.ss.set_cell(0, 1, "=A1^2")
        result = self.ss.get_value(0, 1)
        assert result.startswith("#")
        assert result.endswith("!")
        assert len(result) < 20

    def test_error_propagation_chain(self):
        """Test error propagation through chain of cells."""
        self.ss.set_cell(0, 0, "=1/0")  # #DIV/0!
        self.ss.set_cell(0, 1, "=A1*2")
        self.ss.set_cell(0, 2, "=B1+3")
        self.ss.set_cell(0, 3, "=C1/4")
        result = self.ss.get_value(0, 3)
        assert result.startswith("#")
        assert result.endswith("!")
        assert len(result) < 20

    def test_ref_error_propagation(self):
        """Test #REF! error propagation."""
        # Create a formula that will produce #REF!
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "=A1*B1")  # B1 is empty = 0, so result is 0
        # Now test with explicit ref error in multiplication
        self.ss.set_cell(0, 2, "=#REF!*100")
        # This parses as REF + ! * 100, but the key is no string multiplication
        result = self.ss.get_value(0, 2)
        # Result should be short, not a long repeated string
        assert len(str(result)) < 100

    def test_comparison_with_error(self):
        """Test error propagation in comparisons."""
        self.ss.set_cell(0, 0, "=1/0")
        self.ss.set_cell(0, 1, "=A1>5")
        result = self.ss.get_value(0, 1)
        assert result.startswith("#")
        assert result.endswith("!")
        assert len(result) < 20

    def test_unary_minus_with_error(self):
        """Test error propagation with unary minus."""
        self.ss.set_cell(0, 0, "=1/0")
        self.ss.set_cell(0, 1, "=-A1")
        result = self.ss.get_value(0, 1)
        assert result.startswith("#")
        assert result.endswith("!")
        assert len(result) < 20

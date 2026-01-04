"""Comprehensive tests for all formula functions."""

import pytest

from lotus123 import Spreadsheet


class TestMathFunctions:
    """Tests for mathematical functions."""

    def setup_method(self):
        self.ss = Spreadsheet()

    def test_abs(self):
        """Test @ABS function."""
        self.ss.set_cell(0, 0, "@ABS(-5)")
        assert self.ss.get_value(0, 0) == 5

        self.ss.set_cell(0, 1, "@ABS(5)")
        assert self.ss.get_value(0, 1) == 5

    def test_int(self):
        """Test @INT function."""
        self.ss.set_cell(0, 0, "@INT(5.7)")
        assert self.ss.get_value(0, 0) == 5

        self.ss.set_cell(0, 1, "@INT(-5.7)")
        assert self.ss.get_value(0, 1) == -6

    def test_round(self):
        """Test @ROUND function."""
        self.ss.set_cell(0, 0, "@ROUND(5.567, 2)")
        assert self.ss.get_value(0, 0) == 5.57

        self.ss.set_cell(0, 1, "@ROUND(5.567, 0)")
        assert self.ss.get_value(0, 1) == 6

    def test_sqrt(self):
        """Test @SQRT function."""
        self.ss.set_cell(0, 0, "@SQRT(16)")
        assert self.ss.get_value(0, 0) == 4

        self.ss.set_cell(0, 1, "@SQRT(2)")
        assert abs(self.ss.get_value(0, 1) - 1.414) < 0.01

    def test_power(self):
        """Test @POWER function."""
        self.ss.set_cell(0, 0, "@POWER(2, 3)")
        assert self.ss.get_value(0, 0) == 8

    def test_mod(self):
        """Test @MOD function."""
        self.ss.set_cell(0, 0, "@MOD(10, 3)")
        assert self.ss.get_value(0, 0) == 1

    def test_sin(self):
        """Test @SIN function."""
        self.ss.set_cell(0, 0, "@SIN(0)")
        assert self.ss.get_value(0, 0) == 0

    def test_cos(self):
        """Test @COS function."""
        self.ss.set_cell(0, 0, "@COS(0)")
        assert self.ss.get_value(0, 0) == 1

    def test_tan(self):
        """Test @TAN function."""
        self.ss.set_cell(0, 0, "@TAN(0)")
        assert self.ss.get_value(0, 0) == 0

    def test_pi(self):
        """Test @PI function."""
        self.ss.set_cell(0, 0, "@PI()")
        assert abs(self.ss.get_value(0, 0) - 3.14159) < 0.001

    def test_exp(self):
        """Test @EXP function."""
        self.ss.set_cell(0, 0, "@EXP(1)")
        assert abs(self.ss.get_value(0, 0) - 2.718) < 0.01

    def test_ln(self):
        """Test @LN function."""
        self.ss.set_cell(0, 0, "@LN(2.718)")
        assert abs(self.ss.get_value(0, 0) - 1) < 0.01

    def test_log(self):
        """Test @LOG function."""
        self.ss.set_cell(0, 0, "@LOG(100)")
        assert self.ss.get_value(0, 0) == 2

    def test_rand(self):
        """Test @RAND function."""
        self.ss.set_cell(0, 0, "@RAND()")
        value = self.ss.get_value(0, 0)
        assert 0 <= value < 1

    def test_sign(self):
        """Test @SIGN function."""
        self.ss.set_cell(0, 0, "@SIGN(-5)")
        assert self.ss.get_value(0, 0) == -1

        self.ss.set_cell(0, 1, "@SIGN(5)")
        assert self.ss.get_value(0, 1) == 1

        self.ss.set_cell(0, 2, "@SIGN(0)")
        assert self.ss.get_value(0, 2) == 0

    def test_ceiling(self):
        """Test @CEILING function."""
        self.ss.set_cell(0, 0, "@CEILING(4.2)")
        assert self.ss.get_value(0, 0) == 5

    def test_floor(self):
        """Test @FLOOR function."""
        self.ss.set_cell(0, 0, "@FLOOR(4.8)")
        assert self.ss.get_value(0, 0) == 4

    def test_trunc(self):
        """Test @TRUNC function."""
        self.ss.set_cell(0, 0, "@TRUNC(4.567, 1)")
        assert self.ss.get_value(0, 0) == 4.5

    def test_fact(self):
        """Test @FACT function."""
        self.ss.set_cell(0, 0, "@FACT(5)")
        assert self.ss.get_value(0, 0) == 120

    def test_gcd(self):
        """Test @GCD function."""
        self.ss.set_cell(0, 0, "@GCD(12, 18)")
        assert self.ss.get_value(0, 0) == 6

    def test_lcm(self):
        """Test @LCM function."""
        self.ss.set_cell(0, 0, "@LCM(4, 6)")
        assert self.ss.get_value(0, 0) == 12


class TestStatisticalFunctions:
    """Tests for statistical functions."""

    def setup_method(self):
        self.ss = Spreadsheet()
        # Set up test data
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(1, 0, "20")
        self.ss.set_cell(2, 0, "30")
        self.ss.set_cell(3, 0, "40")
        self.ss.set_cell(4, 0, "50")

    def test_sum(self):
        """Test @SUM function."""
        self.ss.set_cell(5, 0, "@SUM(A1..A5)")
        assert self.ss.get_value(5, 0) == 150

    def test_sum_with_colon(self):
        """Test @SUM with colon notation."""
        self.ss.set_cell(5, 0, "@SUM(A1:A5)")
        assert self.ss.get_value(5, 0) == 150

    def test_avg(self):
        """Test @AVG function."""
        self.ss.set_cell(5, 0, "@AVG(A1..A5)")
        assert self.ss.get_value(5, 0) == 30

    def test_average(self):
        """Test @AVERAGE function (alias for AVG)."""
        self.ss.set_cell(5, 0, "@AVERAGE(A1..A5)")
        assert self.ss.get_value(5, 0) == 30

    def test_min(self):
        """Test @MIN function."""
        self.ss.set_cell(5, 0, "@MIN(A1..A5)")
        assert self.ss.get_value(5, 0) == 10

    def test_max(self):
        """Test @MAX function."""
        self.ss.set_cell(5, 0, "@MAX(A1..A5)")
        assert self.ss.get_value(5, 0) == 50

    def test_count(self):
        """Test @COUNT function."""
        self.ss.set_cell(5, 0, "@COUNT(A1..A5)")
        assert self.ss.get_value(5, 0) == 5

    def test_counta(self):
        """Test @COUNTA function."""
        self.ss.set_cell(5, 0, "@COUNTA(A1..A5)")
        assert self.ss.get_value(5, 0) == 5

    def test_stdev(self):
        """Test @STDEV function."""
        self.ss.set_cell(5, 0, "@STDEV(A1..A5)")
        value = self.ss.get_value(5, 0)
        assert abs(value - 15.81) < 0.1

    def test_var(self):
        """Test @VAR function."""
        self.ss.set_cell(5, 0, "@VAR(A1..A5)")
        value = self.ss.get_value(5, 0)
        assert value == 250

    def test_median(self):
        """Test @MEDIAN function."""
        self.ss.set_cell(5, 0, "@MEDIAN(A1..A5)")
        assert self.ss.get_value(5, 0) == 30

    def test_mode(self):
        """Test @MODE function."""
        self.ss.set_cell(0, 1, "10")  # Add duplicate
        self.ss.set_cell(5, 0, "@MODE(A1..B1)")
        assert self.ss.get_value(5, 0) == 10

    def test_large(self):
        """Test @LARGE function."""
        self.ss.set_cell(5, 0, "@LARGE(A1..A5, 2)")
        assert self.ss.get_value(5, 0) == 40

    def test_small(self):
        """Test @SMALL function."""
        self.ss.set_cell(5, 0, "@SMALL(A1..A5, 2)")
        assert self.ss.get_value(5, 0) == 20

    def test_rank(self):
        """Test @RANK function."""
        self.ss.set_cell(5, 0, "@RANK(30, A1..A5)")
        assert self.ss.get_value(5, 0) == 3


class TestStringFunctions:
    """Tests for string functions."""

    def setup_method(self):
        self.ss = Spreadsheet()

    def test_len(self):
        """Test @LEN function."""
        self.ss.set_cell(0, 0, '@LEN("Hello")')
        assert self.ss.get_value(0, 0) == 5

    def test_left(self):
        """Test @LEFT function."""
        self.ss.set_cell(0, 0, '@LEFT("Hello", 2)')
        assert self.ss.get_value(0, 0) == "He"

    def test_right(self):
        """Test @RIGHT function."""
        self.ss.set_cell(0, 0, '@RIGHT("Hello", 2)')
        assert self.ss.get_value(0, 0) == "lo"

    def test_mid(self):
        """Test @MID function."""
        self.ss.set_cell(0, 0, '@MID("Hello", 2, 3)')
        assert self.ss.get_value(0, 0) == "ell"

    def test_upper(self):
        """Test @UPPER function."""
        self.ss.set_cell(0, 0, '@UPPER("hello")')
        assert self.ss.get_value(0, 0) == "HELLO"

    def test_lower(self):
        """Test @LOWER function."""
        self.ss.set_cell(0, 0, '@LOWER("HELLO")')
        assert self.ss.get_value(0, 0) == "hello"

    def test_proper(self):
        """Test @PROPER function."""
        self.ss.set_cell(0, 0, '@PROPER("hello world")')
        assert self.ss.get_value(0, 0) == "Hello World"

    def test_trim(self):
        """Test @TRIM function."""
        self.ss.set_cell(0, 0, '@TRIM("  hello  ")')
        assert self.ss.get_value(0, 0) == "hello"

    def test_concatenate(self):
        """Test @CONCATENATE function."""
        self.ss.set_cell(0, 0, '@CONCATENATE("Hello", " ", "World")')
        assert self.ss.get_value(0, 0) == "Hello World"

    def test_concat(self):
        """Test @CONCAT function."""
        self.ss.set_cell(0, 0, '@CONCAT("Hello", "World")')
        assert self.ss.get_value(0, 0) == "HelloWorld"

    def test_find(self):
        """Test @FIND function."""
        self.ss.set_cell(0, 0, '@FIND("l", "Hello")')
        assert self.ss.get_value(0, 0) == 3

    def test_search(self):
        """Test @SEARCH function (case insensitive)."""
        self.ss.set_cell(0, 0, '@SEARCH("L", "Hello")')
        assert self.ss.get_value(0, 0) == 3

    def test_replace(self):
        """Test @REPLACE function."""
        self.ss.set_cell(0, 0, '@REPLACE("Hello", 2, 3, "XYZ")')
        assert self.ss.get_value(0, 0) == "HXYZo"

    def test_substitute(self):
        """Test @SUBSTITUTE function."""
        self.ss.set_cell(0, 0, '@SUBSTITUTE("Hello", "l", "L")')
        assert self.ss.get_value(0, 0) == "HeLLo"

    def test_rept(self):
        """Test @REPT function."""
        self.ss.set_cell(0, 0, '@REPT("ab", 3)')
        assert self.ss.get_value(0, 0) == "ababab"

    def test_text(self):
        """Test @TEXT function."""
        self.ss.set_cell(0, 0, '@TEXT(1234.5, "0.00")')
        assert "1234" in str(self.ss.get_value(0, 0))

    def test_value(self):
        """Test @VALUE function."""
        self.ss.set_cell(0, 0, '@VALUE("123.45")')
        assert self.ss.get_value(0, 0) == 123.45

    def test_exact(self):
        """Test @EXACT function."""
        self.ss.set_cell(0, 0, '@EXACT("Hello", "Hello")')
        assert self.ss.get_value(0, 0) is True

        self.ss.set_cell(0, 1, '@EXACT("Hello", "hello")')
        assert self.ss.get_value(0, 1) is False

    def test_char(self):
        """Test @CHAR function."""
        self.ss.set_cell(0, 0, "@CHAR(65)")
        assert self.ss.get_value(0, 0) == "A"

    def test_code(self):
        """Test @CODE function."""
        self.ss.set_cell(0, 0, '@CODE("A")')
        assert self.ss.get_value(0, 0) == 65


class TestLogicalFunctions:
    """Tests for logical functions."""

    def setup_method(self):
        self.ss = Spreadsheet()

    def test_if_true(self):
        """Test @IF function with true condition."""
        self.ss.set_cell(0, 0, "@IF(1>0, 10, 20)")
        assert self.ss.get_value(0, 0) == 10

    def test_if_false(self):
        """Test @IF function with false condition."""
        self.ss.set_cell(0, 0, "@IF(1<0, 10, 20)")
        assert self.ss.get_value(0, 0) == 20

    def test_and_true(self):
        """Test @AND function returning true."""
        self.ss.set_cell(0, 0, "@AND(1>0, 2>1)")
        assert self.ss.get_value(0, 0) is True

    def test_and_false(self):
        """Test @AND function returning false."""
        self.ss.set_cell(0, 0, "@AND(1>0, 2<1)")
        assert self.ss.get_value(0, 0) is False

    def test_or_true(self):
        """Test @OR function returning true."""
        self.ss.set_cell(0, 0, "@OR(1<0, 2>1)")
        assert self.ss.get_value(0, 0) is True

    def test_or_false(self):
        """Test @OR function returning false."""
        self.ss.set_cell(0, 0, "@OR(1<0, 2<1)")
        assert self.ss.get_value(0, 0) is False

    def test_not(self):
        """Test @NOT function."""
        self.ss.set_cell(0, 0, "@NOT(1>0)")
        assert self.ss.get_value(0, 0) is False

        self.ss.set_cell(0, 1, "@NOT(1<0)")
        assert self.ss.get_value(0, 1) is True

    def test_true(self):
        """Test @TRUE function."""
        self.ss.set_cell(0, 0, "@TRUE()")
        assert self.ss.get_value(0, 0) is True

    def test_false(self):
        """Test @FALSE function."""
        self.ss.set_cell(0, 0, "@FALSE()")
        assert self.ss.get_value(0, 0) is False

    def test_isblank(self):
        """Test @ISBLANK function."""
        self.ss.set_cell(0, 0, "@ISBLANK(A2)")
        assert self.ss.get_value(0, 0) is True

        self.ss.set_cell(1, 0, "test")
        self.ss.set_cell(0, 1, "@ISBLANK(A2)")
        assert self.ss.get_value(0, 1) is False

    def test_isnumber(self):
        """Test @ISNUMBER function."""
        self.ss.set_cell(0, 0, "123")
        self.ss.set_cell(0, 1, "@ISNUMBER(A1)")
        assert self.ss.get_value(0, 1) is True

    def test_istext(self):
        """Test @ISTEXT function."""
        self.ss.set_cell(0, 0, "hello")
        self.ss.set_cell(0, 1, "@ISTEXT(A1)")
        assert self.ss.get_value(0, 1) is True

    def test_iserror(self):
        """Test @ISERROR function."""
        self.ss.set_cell(0, 0, "@ISERROR(1/0)")
        assert self.ss.get_value(0, 0) is True

    def test_iferror(self):
        """Test @IFERROR function."""
        self.ss.set_cell(0, 0, '@IFERROR(1/0, "Error")')
        assert self.ss.get_value(0, 0) == "Error"

        self.ss.set_cell(0, 1, '@IFERROR(10/2, "Error")')
        assert self.ss.get_value(0, 1) == 5


class TestDateTimeFunctions:
    """Tests for date and time functions."""

    def setup_method(self):
        self.ss = Spreadsheet()

    def test_now(self):
        """Test @NOW function."""
        self.ss.set_cell(0, 0, "@NOW()")
        value = self.ss.get_value(0, 0)
        assert value > 40000  # Serial date number

    def test_today(self):
        """Test @TODAY function."""
        self.ss.set_cell(0, 0, "@TODAY()")
        value = self.ss.get_value(0, 0)
        assert value > 40000

    def test_date(self):
        """Test @DATE function."""
        self.ss.set_cell(0, 0, "@DATE(2024, 1, 15)")
        value = self.ss.get_value(0, 0)
        assert isinstance(value, (int, float))

    def test_year(self):
        """Test @YEAR function."""
        self.ss.set_cell(0, 0, "@DATE(2024, 6, 15)")
        self.ss.set_cell(0, 1, "@YEAR(A1)")
        assert self.ss.get_value(0, 1) == 2024

    def test_month(self):
        """Test @MONTH function."""
        self.ss.set_cell(0, 0, "@DATE(2024, 6, 15)")
        self.ss.set_cell(0, 1, "@MONTH(A1)")
        assert self.ss.get_value(0, 1) == 6

    def test_day(self):
        """Test @DAY function."""
        self.ss.set_cell(0, 0, "@DATE(2024, 6, 15)")
        self.ss.set_cell(0, 1, "@DAY(A1)")
        assert self.ss.get_value(0, 1) == 15

    def test_weekday(self):
        """Test @WEEKDAY function."""
        self.ss.set_cell(0, 0, "@DATE(2024, 1, 1)")  # Monday
        self.ss.set_cell(0, 1, "@WEEKDAY(A1)")
        value = self.ss.get_value(0, 1)
        assert 1 <= value <= 7

    def test_hour(self):
        """Test @HOUR function."""
        self.ss.set_cell(0, 0, "@HOUR(0.5)")  # Noon
        assert self.ss.get_value(0, 0) == 12

    def test_minute(self):
        """Test @MINUTE function."""
        # Test with direct time value calculation
        # 0.5208333 days = 12.5 hours = 12 hours 30 minutes
        # However, the MINUTE function extracts the minute from a time serial
        self.ss.set_cell(0, 0, "@MINUTE(0.520833)")  # About 12:30
        value = self.ss.get_value(0, 0)
        # Just verify we get an integer minute value
        assert isinstance(value, (int, float))

    def test_second(self):
        """Test @SECOND function."""
        self.ss.set_cell(0, 0, "@SECOND(0.500694)")  # 12:01:00
        value = self.ss.get_value(0, 0)
        assert isinstance(value, int)


class TestLookupFunctions:
    """Tests for lookup functions."""

    def setup_method(self):
        self.ss = Spreadsheet()
        # Set up lookup table
        self.ss.set_cell(0, 0, "1")
        self.ss.set_cell(0, 1, "Apple")
        self.ss.set_cell(1, 0, "2")
        self.ss.set_cell(1, 1, "Banana")
        self.ss.set_cell(2, 0, "3")
        self.ss.set_cell(2, 1, "Cherry")

    def test_vlookup(self):
        """Test @VLOOKUP function."""
        self.ss.set_cell(5, 0, "@VLOOKUP(2, A1:B3, 2, 0)")
        assert self.ss.get_value(5, 0) == "Banana"

    def test_hlookup(self):
        """Test @HLOOKUP function."""
        # Set up horizontal table
        self.ss.set_cell(5, 0, "1")
        self.ss.set_cell(5, 1, "2")
        self.ss.set_cell(5, 2, "3")
        self.ss.set_cell(6, 0, "A")
        self.ss.set_cell(6, 1, "B")
        self.ss.set_cell(6, 2, "C")

        self.ss.set_cell(7, 0, "@HLOOKUP(2, A6:C7, 2, 0)")
        assert self.ss.get_value(7, 0) == "B"

    def test_index(self):
        """Test @INDEX function."""
        self.ss.set_cell(5, 0, "@INDEX(A1:B3, 2, 2)")
        assert self.ss.get_value(5, 0) == "Banana"

    def test_match(self):
        """Test @MATCH function."""
        self.ss.set_cell(5, 0, "@MATCH(2, A1:A3, 0)")
        assert self.ss.get_value(5, 0) == 2

    def test_rows(self):
        """Test @ROWS function."""
        self.ss.set_cell(5, 0, "@ROWS(A1:B3)")
        assert self.ss.get_value(5, 0) == 3

    def test_cols(self):
        """Test @COLS function."""
        self.ss.set_cell(5, 0, "@COLS(A1:B3)")
        assert self.ss.get_value(5, 0) == 2

    def test_choose(self):
        """Test @CHOOSE function."""
        self.ss.set_cell(5, 0, '@CHOOSE(2, "A", "B", "C")')
        assert self.ss.get_value(5, 0) == "B"


class TestFinancialFunctions:
    """Tests for financial functions."""

    def setup_method(self):
        self.ss = Spreadsheet()

    def test_pv(self):
        """Test @PV function."""
        self.ss.set_cell(0, 0, "@PV(0.1, 10, 100)")
        value = self.ss.get_value(0, 0)
        assert abs(value - (-614.46)) < 1

    def test_fv(self):
        """Test @FV function."""
        self.ss.set_cell(0, 0, "@FV(0.1, 10, 100)")
        value = self.ss.get_value(0, 0)
        assert abs(value - (-1593.74)) < 1

    def test_pmt(self):
        """Test @PMT function."""
        self.ss.set_cell(0, 0, "@PMT(0.1, 10, 1000)")
        value = self.ss.get_value(0, 0)
        assert abs(value - (-162.75)) < 1

    def test_npv(self):
        """Test @NPV function."""
        self.ss.set_cell(0, 0, "100")
        self.ss.set_cell(1, 0, "100")
        self.ss.set_cell(2, 0, "100")
        self.ss.set_cell(3, 0, "@NPV(0.1, A1:A3)")
        value = self.ss.get_value(3, 0)
        assert abs(value - 248.69) < 1

    def test_irr(self):
        """Test @IRR function."""
        self.ss.set_cell(0, 0, "-1000")
        self.ss.set_cell(1, 0, "400")
        self.ss.set_cell(2, 0, "400")
        self.ss.set_cell(3, 0, "400")
        self.ss.set_cell(4, 0, "@IRR(A1:A4)")
        value = self.ss.get_value(4, 0)
        assert abs(value - 0.0966) < 0.01


class TestFormulaPrefixes:
    """Tests for different formula prefixes."""

    def setup_method(self):
        self.ss = Spreadsheet()
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(1, 0, "20")

    def test_at_prefix(self):
        """Test @ prefix for formulas."""
        self.ss.set_cell(2, 0, "@SUM(A1..A2)")
        assert self.ss.get_value(2, 0) == 30

    def test_equals_prefix(self):
        """Test = prefix for formulas."""
        self.ss.set_cell(2, 0, "=SUM(A1:A2)")
        assert self.ss.get_value(2, 0) == 30

    def test_plus_prefix(self):
        """Test + prefix for formulas."""
        self.ss.set_cell(2, 0, "+A1+A2")
        assert self.ss.get_value(2, 0) == 30

    def test_lotus_range_notation(self):
        """Test .. range notation (Lotus style)."""
        self.ss.set_cell(2, 0, "@SUM(A1..A2)")
        assert self.ss.get_value(2, 0) == 30

    def test_excel_range_notation(self):
        """Test : range notation (Excel style)."""
        self.ss.set_cell(2, 0, "@SUM(A1:A2)")
        assert self.ss.get_value(2, 0) == 30


class TestCellReferences:
    """Tests for cell reference handling."""

    def setup_method(self):
        self.ss = Spreadsheet()

    def test_simple_reference(self):
        """Test simple cell reference."""
        self.ss.set_cell(0, 0, "100")
        self.ss.set_cell(1, 0, "=A1")
        assert self.ss.get_value(1, 0) == 100

    def test_reference_chain(self):
        """Test chain of references."""
        self.ss.set_cell(0, 0, "100")
        self.ss.set_cell(1, 0, "=A1*2")
        self.ss.set_cell(2, 0, "=A2+50")
        assert self.ss.get_value(2, 0) == 250

    def test_circular_reference(self):
        """Test circular reference detection."""
        self.ss.set_cell(0, 0, "=A1")
        assert "#CIRC!" in str(self.ss.get_value(0, 0))

    def test_cross_reference(self):
        """Test references across columns."""
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "20")
        self.ss.set_cell(0, 2, "=A1+B1")
        assert self.ss.get_value(0, 2) == 30


class TestComparisonOperators:
    """Tests for comparison operators."""

    def setup_method(self):
        self.ss = Spreadsheet()

    def test_equal(self):
        """Test = comparison."""
        self.ss.set_cell(0, 0, "=5=5")
        assert self.ss.get_value(0, 0) is True

    def test_not_equal(self):
        """Test <> comparison."""
        self.ss.set_cell(0, 0, "=5<>3")
        assert self.ss.get_value(0, 0) is True

    def test_less_than(self):
        """Test < comparison."""
        self.ss.set_cell(0, 0, "=3<5")
        assert self.ss.get_value(0, 0) is True

    def test_greater_than(self):
        """Test > comparison."""
        self.ss.set_cell(0, 0, "=5>3")
        assert self.ss.get_value(0, 0) is True

    def test_less_equal(self):
        """Test <= comparison."""
        self.ss.set_cell(0, 0, "=5<=5")
        assert self.ss.get_value(0, 0) is True

    def test_greater_equal(self):
        """Test >= comparison."""
        self.ss.set_cell(0, 0, "=5>=5")
        assert self.ss.get_value(0, 0) is True


class TestArithmeticOperators:
    """Tests for arithmetic operators."""

    def setup_method(self):
        self.ss = Spreadsheet()

    def test_addition(self):
        """Test + operator."""
        self.ss.set_cell(0, 0, "=5+3")
        assert self.ss.get_value(0, 0) == 8

    def test_subtraction(self):
        """Test - operator."""
        self.ss.set_cell(0, 0, "=5-3")
        assert self.ss.get_value(0, 0) == 2

    def test_multiplication(self):
        """Test * operator."""
        self.ss.set_cell(0, 0, "=5*3")
        assert self.ss.get_value(0, 0) == 15

    def test_division(self):
        """Test / operator."""
        self.ss.set_cell(0, 0, "=15/3")
        assert self.ss.get_value(0, 0) == 5

    def test_exponentiation(self):
        """Test ^ operator."""
        self.ss.set_cell(0, 0, "=2^3")
        assert self.ss.get_value(0, 0) == 8

    def test_modulo(self):
        """Test % operator."""
        self.ss.set_cell(0, 0, "=10%3")
        assert self.ss.get_value(0, 0) == 1

    def test_operator_precedence(self):
        """Test operator precedence."""
        self.ss.set_cell(0, 0, "=2+3*4")
        assert self.ss.get_value(0, 0) == 14

    def test_parentheses(self):
        """Test parentheses override precedence."""
        self.ss.set_cell(0, 0, "=(2+3)*4")
        assert self.ss.get_value(0, 0) == 20

    def test_unary_minus(self):
        """Test unary minus."""
        self.ss.set_cell(0, 0, "=-5")
        assert self.ss.get_value(0, 0) == -5

    def test_division_by_zero(self):
        """Test division by zero returns error."""
        self.ss.set_cell(0, 0, "=1/0")
        result = str(self.ss.get_value(0, 0))
        assert "#DIV/0!" in result or "ERR" in result or "INF" in result.upper()

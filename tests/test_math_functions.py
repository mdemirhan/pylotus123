"""Tests for mathematical functions."""

import math

import pytest

from lotus123.formula.functions.math import (
    _flatten_args,
    _to_number,
    fn_abs,
    fn_acos,
    fn_asin,
    fn_atan,
    fn_atan2,
    fn_ceiling,
    fn_cos,
    fn_degrees,
    fn_exp,
    fn_fact,
    fn_floor,
    fn_gcd,
    fn_int,
    fn_lcm,
    fn_ln,
    fn_log,
    fn_mod,
    fn_pi,
    fn_power,
    fn_radians,
    fn_rand,
    fn_round,
    fn_sign,
    fn_sin,
    fn_sqrt,
    fn_sum,
    fn_tan,
    fn_trunc,
)


class TestHelperFunctions:
    """Tests for helper functions."""

    def test_to_number_int(self):
        """Test converting int."""
        assert _to_number(42) == 42.0

    def test_to_number_float(self):
        """Test converting float."""
        assert _to_number(3.14) == 3.14

    def test_to_number_string(self):
        """Test converting numeric string."""
        assert _to_number("123") == 123.0

    def test_to_number_string_with_comma(self):
        """Test converting string with commas."""
        assert _to_number("1,234") == 1234.0

    def test_to_number_invalid(self):
        """Test invalid conversion returns 0."""
        assert _to_number("abc") == 0.0

    def test_to_number_none(self):
        """Test None returns 0."""
        assert _to_number(None) == 0.0

    def test_flatten_args_simple(self):
        """Test flattening simple args."""
        result = _flatten_args((1, 2, 3))
        assert result == [1, 2, 3]

    def test_flatten_args_nested(self):
        """Test flattening nested lists."""
        result = _flatten_args((1, [2, 3], 4))
        assert result == [1, 2, 3, 4]


class TestBasicMath:
    """Tests for basic math functions."""

    def test_fn_sum_basic(self):
        """Test basic SUM."""
        assert fn_sum(1, 2, 3) == 6

    def test_fn_sum_empty(self):
        """Test SUM with no args."""
        assert fn_sum() == 0

    def test_fn_sum_with_list(self):
        """Test SUM with list."""
        assert fn_sum([1, 2, 3]) == 6

    def test_fn_sum_ignores_text(self):
        """Test SUM ignores non-numeric text."""
        assert fn_sum(1, "text", 2) == 3

    def test_fn_abs_positive(self):
        """Test ABS with positive."""
        assert fn_abs(5) == 5

    def test_fn_abs_negative(self):
        """Test ABS with negative."""
        assert fn_abs(-5) == 5

    def test_fn_abs_zero(self):
        """Test ABS with zero."""
        assert fn_abs(0) == 0

    def test_fn_int_positive(self):
        """Test INT with positive."""
        assert fn_int(3.7) == 3

    def test_fn_int_negative(self):
        """Test INT with negative (floors toward -inf)."""
        assert fn_int(-3.7) == -4

    def test_fn_round_default(self):
        """Test ROUND with default decimals."""
        assert fn_round(3.7) == 4
        assert fn_round(3.4) == 3

    def test_fn_round_decimals(self):
        """Test ROUND with decimals."""
        assert fn_round(3.14159, 2) == 3.14

    def test_fn_mod_basic(self):
        """Test MOD basic usage."""
        assert fn_mod(10, 3) == 1

    def test_fn_mod_zero_divisor(self):
        """Test MOD with zero divisor."""
        assert math.isnan(fn_mod(10, 0))

    def test_fn_sqrt_positive(self):
        """Test SQRT with positive."""
        assert fn_sqrt(16) == 4

    def test_fn_sqrt_zero(self):
        """Test SQRT with zero."""
        assert fn_sqrt(0) == 0

    def test_fn_sqrt_negative(self):
        """Test SQRT with negative."""
        assert math.isnan(fn_sqrt(-1))

    def test_fn_power_basic(self):
        """Test POWER basic usage."""
        assert fn_power(2, 3) == 8

    def test_fn_power_fractional(self):
        """Test POWER with fractional exponent."""
        assert fn_power(4, 0.5) == 2

    def test_fn_sign_positive(self):
        """Test SIGN with positive."""
        assert fn_sign(5) == 1

    def test_fn_sign_negative(self):
        """Test SIGN with negative."""
        assert fn_sign(-5) == -1

    def test_fn_sign_zero(self):
        """Test SIGN with zero."""
        assert fn_sign(0) == 0

    def test_fn_trunc_positive(self):
        """Test TRUNC with positive."""
        assert fn_trunc(3.7) == 3

    def test_fn_trunc_negative(self):
        """Test TRUNC with negative."""
        assert fn_trunc(-3.7) == -3

    def test_fn_trunc_decimals(self):
        """Test TRUNC with decimals."""
        assert fn_trunc(3.14159, 2) == 3.14

    def test_fn_ceiling_positive(self):
        """Test CEILING with positive."""
        assert fn_ceiling(3.1) == 4

    def test_fn_ceiling_negative(self):
        """Test CEILING with negative."""
        assert fn_ceiling(-3.7) == -3

    def test_fn_floor_positive(self):
        """Test FLOOR with positive."""
        assert fn_floor(3.9) == 3

    def test_fn_floor_negative(self):
        """Test FLOOR with negative."""
        assert fn_floor(-3.1) == -4

    def test_fn_fact_basic(self):
        """Test FACT basic usage."""
        assert fn_fact(5) == 120

    def test_fn_fact_zero(self):
        """Test FACT with zero."""
        assert fn_fact(0) == 1

    def test_fn_fact_negative(self):
        """Test FACT with negative."""
        assert fn_fact(-1) == 0

    def test_fn_gcd_basic(self):
        """Test GCD basic usage."""
        assert fn_gcd(12, 8) == 4

    def test_fn_lcm_basic(self):
        """Test LCM basic usage."""
        assert fn_lcm(4, 6) == 12

    def test_fn_lcm_zero(self):
        """Test LCM with zero."""
        assert fn_lcm(4, 0) == 0


class TestExponentialLog:
    """Tests for exponential and logarithmic functions."""

    def test_fn_exp_basic(self):
        """Test EXP basic usage."""
        assert abs(fn_exp(1) - math.e) < 0.0001

    def test_fn_exp_zero(self):
        """Test EXP with zero."""
        assert fn_exp(0) == 1

    def test_fn_ln_basic(self):
        """Test LN basic usage."""
        assert abs(fn_ln(math.e) - 1) < 0.0001

    def test_fn_ln_one(self):
        """Test LN with one."""
        assert fn_ln(1) == 0

    def test_fn_ln_zero(self):
        """Test LN with zero or negative."""
        assert math.isnan(fn_ln(0))
        assert math.isnan(fn_ln(-1))

    def test_fn_log_basic(self):
        """Test LOG (base 10) basic usage."""
        assert fn_log(100) == 2

    def test_fn_log_zero(self):
        """Test LOG with zero or negative."""
        assert math.isnan(fn_log(0))
        assert math.isnan(fn_log(-1))


class TestTrigonometric:
    """Tests for trigonometric functions."""

    def test_fn_sin_zero(self):
        """Test SIN with zero."""
        assert fn_sin(0) == 0

    def test_fn_sin_pi_half(self):
        """Test SIN with pi/2."""
        assert abs(fn_sin(math.pi / 2) - 1) < 0.0001

    def test_fn_cos_zero(self):
        """Test COS with zero."""
        assert fn_cos(0) == 1

    def test_fn_cos_pi(self):
        """Test COS with pi."""
        assert abs(fn_cos(math.pi) - (-1)) < 0.0001

    def test_fn_tan_zero(self):
        """Test TAN with zero."""
        assert fn_tan(0) == 0

    def test_fn_asin_valid(self):
        """Test ASIN with valid value."""
        assert abs(fn_asin(1) - math.pi / 2) < 0.0001

    def test_fn_asin_invalid(self):
        """Test ASIN with invalid value."""
        assert math.isnan(fn_asin(2))

    def test_fn_acos_valid(self):
        """Test ACOS with valid value."""
        assert abs(fn_acos(1)) < 0.0001

    def test_fn_acos_invalid(self):
        """Test ACOS with invalid value."""
        assert math.isnan(fn_acos(2))

    def test_fn_atan_basic(self):
        """Test ATAN basic usage."""
        assert abs(fn_atan(1) - math.pi / 4) < 0.0001

    def test_fn_atan2_basic(self):
        """Test ATAN2 basic usage."""
        assert abs(fn_atan2(1, 1) - math.pi / 4) < 0.0001

    def test_fn_degrees_basic(self):
        """Test DEGREES basic usage."""
        assert abs(fn_degrees(math.pi) - 180) < 0.0001

    def test_fn_radians_basic(self):
        """Test RADIANS basic usage."""
        assert abs(fn_radians(180) - math.pi) < 0.0001


class TestConstantsRandom:
    """Tests for constants and random functions."""

    def test_fn_pi(self):
        """Test PI function."""
        assert fn_pi() == math.pi

    def test_fn_rand_range(self):
        """Test RAND returns value in [0, 1)."""
        for _ in range(10):
            r = fn_rand()
            assert 0 <= r < 1

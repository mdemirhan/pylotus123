"""Tests for statistical functions."""

import math

import pytest

from lotus123.formula.functions.statistical import (
    _flatten_args,
    _get_numbers,
    _to_number,
    fn_avg,
    fn_combin,
    fn_count,
    fn_counta,
    fn_countblank,
    fn_fact,
    fn_geomean,
    fn_harmean,
    fn_large,
    fn_max,
    fn_median,
    fn_min,
    fn_mode,
    fn_percentile,
    fn_permut,
    fn_product,
    fn_quartile,
    fn_rand,
    fn_randbetween,
    fn_rank,
    fn_small,
    fn_std,
    fn_stdp,
    fn_stds,
    fn_sum,
    fn_sumproduct,
    fn_sumsq,
    fn_var,
    fn_varp,
    fn_vars,
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
        """Test invalid conversion returns None."""
        assert _to_number("abc") is None

    def test_flatten_args_simple(self):
        """Test flattening simple args."""
        result = _flatten_args((1, 2, 3))
        assert result == [1, 2, 3]

    def test_flatten_args_nested(self):
        """Test flattening nested lists."""
        result = _flatten_args((1, [2, 3], 4))
        assert result == [1, 2, 3, 4]

    def test_get_numbers_basic(self):
        """Test extracting numbers."""
        result = _get_numbers((1, 2, "3", "text"))
        assert result == [1.0, 2.0, 3.0]


class TestBasicStatistics:
    """Tests for basic statistical functions."""

    def test_fn_sum_basic(self):
        """Test basic SUM."""
        assert fn_sum(1, 2, 3) == 6

    def test_fn_sum_empty(self):
        """Test SUM with no numbers."""
        assert fn_sum() == 0

    def test_fn_sum_with_text(self):
        """Test SUM ignores text."""
        assert fn_sum(1, "text", 2) == 3

    def test_fn_avg_basic(self):
        """Test basic AVG."""
        assert fn_avg(1, 2, 3) == 2

    def test_fn_avg_empty(self):
        """Test AVG with no numbers."""
        assert fn_avg() == 0

    def test_fn_count_basic(self):
        """Test basic COUNT."""
        assert fn_count(1, 2, 3) == 3

    def test_fn_count_with_text(self):
        """Test COUNT ignores text."""
        assert fn_count(1, "text", 2, None) == 2

    def test_fn_counta_basic(self):
        """Test COUNTA counts non-empty."""
        assert fn_counta(1, "text", 3) == 3

    def test_fn_counta_skips_empty(self):
        """Test COUNTA skips empty."""
        assert fn_counta(1, "", None, "text") == 2

    def test_fn_countblank_basic(self):
        """Test COUNTBLANK counts empty."""
        assert fn_countblank(1, "", None, "text") == 2

    def test_fn_min_basic(self):
        """Test basic MIN."""
        assert fn_min(3, 1, 2) == 1

    def test_fn_min_empty(self):
        """Test MIN with no numbers."""
        assert fn_min() == 0

    def test_fn_min_negative(self):
        """Test MIN with negative numbers."""
        assert fn_min(-5, 0, 5) == -5

    def test_fn_max_basic(self):
        """Test basic MAX."""
        assert fn_max(1, 3, 2) == 3

    def test_fn_max_empty(self):
        """Test MAX with no numbers."""
        assert fn_max() == 0

    def test_fn_product_basic(self):
        """Test basic PRODUCT."""
        assert fn_product(2, 3, 4) == 24

    def test_fn_product_empty(self):
        """Test PRODUCT with no numbers."""
        assert fn_product() == 0


class TestDispersionFunctions:
    """Tests for dispersion functions."""

    def test_fn_std_basic(self):
        """Test sample standard deviation."""
        result = fn_std(2, 4, 4, 4, 5, 5, 7, 9)
        assert abs(result - 2.138) < 0.01

    def test_fn_std_few_values(self):
        """Test STD with fewer than 2 values."""
        assert fn_std(5) == 0
        assert fn_std() == 0

    def test_fn_stds_alias(self):
        """Test STDS is alias for STD."""
        assert fn_stds(2, 4, 6) == fn_std(2, 4, 6)

    def test_fn_stdp_basic(self):
        """Test population standard deviation."""
        result = fn_stdp(2, 4, 4, 4, 5, 5, 7, 9)
        assert abs(result - 2.0) < 0.01

    def test_fn_stdp_empty(self):
        """Test STDP with no values."""
        assert fn_stdp() == 0

    def test_fn_var_basic(self):
        """Test sample variance."""
        result = fn_var(2, 4, 6)
        assert abs(result - 4.0) < 0.01

    def test_fn_var_few_values(self):
        """Test VAR with fewer than 2 values."""
        assert fn_var(5) == 0

    def test_fn_vars_alias(self):
        """Test VARS is alias for VAR."""
        assert fn_vars(2, 4, 6) == fn_var(2, 4, 6)

    def test_fn_varp_basic(self):
        """Test population variance."""
        result = fn_varp(2, 4, 6)
        assert abs(result - 2.666) < 0.01

    def test_fn_varp_empty(self):
        """Test VARP with no values."""
        assert fn_varp() == 0

    def test_fn_sumsq_basic(self):
        """Test sum of squares."""
        assert fn_sumsq(1, 2, 3) == 14  # 1 + 4 + 9


class TestPositionFunctions:
    """Tests for position functions."""

    def test_fn_median_odd(self):
        """Test MEDIAN with odd count."""
        assert fn_median(1, 3, 2) == 2

    def test_fn_median_even(self):
        """Test MEDIAN with even count."""
        assert fn_median(1, 2, 3, 4) == 2.5

    def test_fn_median_empty(self):
        """Test MEDIAN with no values."""
        assert fn_median() == 0

    def test_fn_mode_basic(self):
        """Test MODE returns most frequent."""
        assert fn_mode(1, 2, 2, 3) == 2

    def test_fn_mode_no_mode(self):
        """Test MODE with no repeated values."""
        assert fn_mode(1, 2, 3) == 1  # Returns first

    def test_fn_mode_empty(self):
        """Test MODE with no values."""
        assert fn_mode() == 0

    def test_fn_large_basic(self):
        """Test LARGE returns k-th largest."""
        assert fn_large(1, 2, 3, 4, 5, 1) == 5  # 1st largest
        assert fn_large(1, 2, 3, 4, 5, 2) == 4  # 2nd largest

    def test_fn_large_out_of_range(self):
        """Test LARGE with out of range k."""
        result = fn_large(1, 2, 3, 10)
        assert math.isnan(result)

    def test_fn_small_basic(self):
        """Test SMALL returns k-th smallest."""
        assert fn_small(5, 4, 3, 2, 1, 1) == 1  # 1st smallest
        assert fn_small(5, 4, 3, 2, 1, 2) == 2  # 2nd smallest

    def test_fn_small_out_of_range(self):
        """Test SMALL with out of range k."""
        result = fn_small(1, 2, 3, 10)
        assert math.isnan(result)

    def test_fn_rank_descending(self):
        """Test RANK with descending order."""
        assert fn_rank(5, [1, 2, 3, 4, 5], 0) == 1  # 5 is largest

    def test_fn_rank_ascending(self):
        """Test RANK with ascending order."""
        assert fn_rank(1, [1, 2, 3, 4, 5], 1) == 1  # 1 is smallest

    def test_fn_rank_not_found(self):
        """Test RANK when value not in list."""
        assert fn_rank(10, [1, 2, 3]) == 0

    def test_fn_percentile_basic(self):
        """Test PERCENTILE basic usage."""
        result = fn_percentile(1, 2, 3, 4, 5, 0.5)  # 50th percentile
        assert abs(result - 3) < 0.01

    def test_fn_percentile_edges(self):
        """Test PERCENTILE at edges."""
        assert fn_percentile(1, 2, 3, 0.0) == 1  # 0th percentile
        assert fn_percentile(1, 2, 3, 1.0) == 3  # 100th percentile

    def test_fn_quartile_basic(self):
        """Test QUARTILE basic usage."""
        result = fn_quartile(1, 2, 3, 4, 5, 2)  # 2nd quartile = median
        assert abs(result - 3) < 0.01

    def test_fn_quartile_out_of_range(self):
        """Test QUARTILE with invalid quartile."""
        result = fn_quartile(1, 2, 3, 5)  # Invalid quartile
        assert math.isnan(result)


class TestRandomFunctions:
    """Tests for random functions."""

    def test_fn_rand_range(self):
        """Test RAND returns value in [0, 1)."""
        for _ in range(10):
            result = fn_rand()
            assert 0 <= result < 1

    def test_fn_randbetween_range(self):
        """Test RANDBETWEEN returns value in range."""
        for _ in range(10):
            result = fn_randbetween(5, 10)
            assert 5 <= result <= 10
            assert isinstance(result, int)


class TestCombinatoricsFunctions:
    """Tests for combinatorics functions."""

    def test_fn_fact_basic(self):
        """Test FACT basic usage."""
        assert fn_fact(5) == 120  # 5!
        assert fn_fact(0) == 1  # 0!
        assert fn_fact(1) == 1  # 1!

    def test_fn_fact_negative(self):
        """Test FACT with negative."""
        assert fn_fact(-1) == 0

    def test_fn_permut_basic(self):
        """Test PERMUT basic usage."""
        assert fn_permut(5, 2) == 20  # 5!/(5-2)! = 5*4

    def test_fn_permut_invalid(self):
        """Test PERMUT with invalid inputs."""
        assert fn_permut(-1, 2) == 0
        assert fn_permut(5, 6) == 0  # k > n

    def test_fn_combin_basic(self):
        """Test COMBIN basic usage."""
        assert fn_combin(5, 2) == 10  # 5!/(2!*3!)

    def test_fn_combin_edges(self):
        """Test COMBIN at edges."""
        assert fn_combin(5, 0) == 1
        assert fn_combin(5, 5) == 1

    def test_fn_combin_invalid(self):
        """Test COMBIN with invalid inputs."""
        assert fn_combin(-1, 2) == 0
        assert fn_combin(5, 6) == 0  # k > n

    def test_fn_sumproduct_basic(self):
        """Test SUMPRODUCT basic usage."""
        result = fn_sumproduct([1, 2, 3], [4, 5, 6])
        assert result == 32  # 1*4 + 2*5 + 3*6

    def test_fn_sumproduct_empty(self):
        """Test SUMPRODUCT with no arrays."""
        assert fn_sumproduct() == 0


class TestMeanFunctions:
    """Tests for mean functions."""

    def test_fn_geomean_basic(self):
        """Test GEOMEAN basic usage."""
        result = fn_geomean(2, 8)
        assert abs(result - 4) < 0.01  # sqrt(16)

    def test_fn_geomean_with_zero(self):
        """Test GEOMEAN with zero."""
        assert fn_geomean(2, 0, 8) == 0

    def test_fn_geomean_with_negative(self):
        """Test GEOMEAN with negative."""
        assert fn_geomean(2, -1, 8) == 0

    def test_fn_geomean_empty(self):
        """Test GEOMEAN with no values."""
        assert fn_geomean() == 0

    def test_fn_harmean_basic(self):
        """Test HARMEAN basic usage."""
        result = fn_harmean(1, 2, 4)
        expected = 3 / (1 + 0.5 + 0.25)  # n / sum(1/x)
        assert abs(result - expected) < 0.01

    def test_fn_harmean_with_zero(self):
        """Test HARMEAN with zero."""
        assert fn_harmean(2, 0, 8) == 0

    def test_fn_harmean_empty(self):
        """Test HARMEAN with no values."""
        assert fn_harmean() == 0

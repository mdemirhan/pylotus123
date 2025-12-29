"""Tests for financial functions."""

import math

import pytest

from lotus123.formula.functions.financial import (
    FINANCIAL_FUNCTIONS,
    fn_cterm,
    fn_ddb,
    fn_fv,
    fn_ipmt,
    fn_irr,
    fn_nper,
    fn_npv,
    fn_pmt,
    fn_ppmt,
    fn_pv,
    fn_rate,
    fn_sln,
    fn_syd,
    fn_term,
)


class TestPMT:
    """Tests for PMT (payment) function."""

    def test_basic_loan(self):
        """Test basic loan payment calculation."""
        # $10000 loan at 10% for 12 periods
        result = fn_pmt(10000, 0.1, 12)
        assert abs(result - 1467.63) < 1

    def test_zero_rate(self):
        """Test PMT with zero interest rate."""
        result = fn_pmt(12000, 0, 12)
        assert result == 1000

    def test_high_rate(self):
        """Test PMT with high interest rate."""
        result = fn_pmt(10000, 0.5, 5)
        assert result > 0


class TestPV:
    """Tests for PV (present value) function."""

    def test_basic_present_value(self):
        """Test basic present value calculation."""
        # $100 payment, 10% rate, 10 periods
        result = fn_pv(100, 0.1, 10)
        assert abs(result - 614.46) < 1

    def test_zero_rate(self):
        """Test PV with zero interest rate."""
        result = fn_pv(100, 0, 10)
        assert result == 1000

    def test_single_period(self):
        """Test PV with single period."""
        result = fn_pv(110, 0.1, 1)
        assert abs(result - 100) < 1


class TestFV:
    """Tests for FV (future value) function."""

    def test_basic_future_value(self):
        """Test basic future value calculation."""
        # $100 payment, 10% rate, 10 periods
        result = fn_fv(100, 0.1, 10)
        assert abs(result - 1593.74) < 1

    def test_zero_rate(self):
        """Test FV with zero interest rate."""
        result = fn_fv(100, 0, 10)
        assert result == 1000


class TestNPV:
    """Tests for NPV (net present value) function."""

    def test_basic_npv(self):
        """Test basic NPV calculation."""
        result = fn_npv(0.1, 100, 100, 100)
        assert abs(result - 248.69) < 1

    def test_npv_with_list(self):
        """Test NPV with list of cash flows."""
        result = fn_npv(0.1, [100, 100, 100])
        assert abs(result - 248.69) < 1

    def test_npv_with_nested_list(self):
        """Test NPV with nested list."""
        result = fn_npv(0.1, [[100], [100], [100]])
        assert abs(result - 248.69) < 1

    def test_npv_with_negative_flows(self):
        """Test NPV with investment (negative) and returns."""
        result = fn_npv(0.1, -1000, 400, 400, 400)
        assert result > -200


class TestIRR:
    """Tests for IRR (internal rate of return) function."""

    def test_basic_irr(self):
        """Test basic IRR calculation."""
        result = fn_irr(0.1, -1000, 400, 400, 400)
        assert isinstance(result, float)
        assert abs(result - 0.0966) < 0.01

    def test_irr_with_list(self):
        """Test IRR with list of cash flows."""
        result = fn_irr(0.1, [-1000, 400, 400, 400])
        assert isinstance(result, float)

    def test_irr_empty_flows(self):
        """Test IRR with no valid flows."""
        result = fn_irr(0.1)
        assert result == "#ERR!"

    def test_irr_no_guess(self):
        """Test IRR with no guess provided."""
        result = fn_irr(None, -1000, 500, 500, 500)
        assert isinstance(result, float)


class TestRATE:
    """Tests for RATE function."""

    def test_basic_rate(self):
        """Test basic rate calculation."""
        result = fn_rate(10, -100, 1000)
        assert isinstance(result, float)

    def test_rate_with_fv(self):
        """Test rate with future value."""
        result = fn_rate(10, -100, 1000, 500)
        assert isinstance(result, float)


class TestNPER:
    """Tests for NPER (number of periods) function."""

    def test_basic_nper(self):
        """Test basic NPER calculation."""
        # Use values that don't cause division by zero
        # payment + present * rate must not equal zero
        result = fn_nper(0.1, -200, 1000)
        assert result > 0

    def test_nper_with_fv(self):
        """Test NPER with future value."""
        result = fn_nper(0.05, -200, 1000, 0)
        assert isinstance(result, float)

    def test_nper_zero_rate(self):
        """Test NPER with zero rate."""
        result = fn_nper(0, -100, 1000)
        assert result == 10


class TestCTERM:
    """Tests for CTERM (compound term) function."""

    def test_basic_cterm(self):
        """Test basic CTERM calculation."""
        result = fn_cterm(0.1, 2000, 1000)
        # Time to double at 10%
        assert abs(result - 7.27) < 0.1

    def test_cterm_error_negative_rate(self):
        """Test CTERM with invalid rate."""
        result = fn_cterm(-0.1, 2000, 1000)
        assert result == "#ERR!"

    def test_cterm_error_negative_pv(self):
        """Test CTERM with negative present value."""
        result = fn_cterm(0.1, 2000, -1000)
        assert result == "#ERR!"


class TestTERM:
    """Tests for TERM function."""

    def test_basic_term(self):
        """Test basic TERM calculation."""
        result = fn_term(100, 0.1, 1000)
        assert result > 0

    def test_term_zero_rate(self):
        """Test TERM with zero rate."""
        result = fn_term(100, 0, 1000)
        assert result == 10


class TestSLN:
    """Tests for SLN (straight-line depreciation) function."""

    def test_basic_sln(self):
        """Test basic straight-line depreciation."""
        result = fn_sln(10000, 1000, 10)
        assert result == 900

    def test_sln_zero_life(self):
        """Test SLN with zero life."""
        result = fn_sln(10000, 1000, 0)
        assert result == "#DIV/0!"


class TestSYD:
    """Tests for SYD (sum of years digits) function."""

    def test_basic_syd(self):
        """Test basic SYD calculation."""
        result = fn_syd(10000, 1000, 5, 1)
        # First year: (10000-1000) * 5/15
        assert abs(result - 3000) < 1

    def test_syd_last_year(self):
        """Test SYD for last year."""
        result = fn_syd(10000, 1000, 5, 5)
        # Last year: (10000-1000) * 1/15
        assert abs(result - 600) < 1

    def test_syd_invalid_period(self):
        """Test SYD with invalid period."""
        result = fn_syd(10000, 1000, 5, 6)
        assert result == "#ERR!"

    def test_syd_zero_life(self):
        """Test SYD with zero life."""
        result = fn_syd(10000, 1000, 0, 1)
        assert result == "#ERR!"


class TestDDB:
    """Tests for DDB (double declining balance) function."""

    def test_basic_ddb(self):
        """Test basic DDB calculation."""
        result = fn_ddb(10000, 1000, 5, 1)
        # First year at 2/5 rate
        assert abs(result - 4000) < 1

    def test_ddb_with_factor(self):
        """Test DDB with custom factor."""
        result = fn_ddb(10000, 1000, 5, 1, 1.5)
        assert result > 0

    def test_ddb_invalid_period(self):
        """Test DDB with invalid period."""
        result = fn_ddb(10000, 1000, 5, 0)
        assert result == "#ERR!"

    def test_ddb_invalid_life(self):
        """Test DDB with zero life."""
        result = fn_ddb(10000, 1000, 0, 1)
        assert result == "#ERR!"


class TestIPMT:
    """Tests for IPMT (interest payment) function."""

    def test_basic_ipmt(self):
        """Test basic interest payment."""
        result = fn_ipmt(0.1, 1, 12, 10000)
        # First period interest on 10000 at 10%
        assert abs(result - 1000) < 1

    def test_ipmt_later_period(self):
        """Test interest payment in later period."""
        result = fn_ipmt(0.1, 6, 12, 10000)
        assert isinstance(result, float)
        assert result < 1000  # Interest decreases over time

    def test_ipmt_invalid_period(self):
        """Test IPMT with invalid period."""
        result = fn_ipmt(0.1, 0, 12, 10000)
        assert result == "#ERR!"

    def test_ipmt_period_exceeds_nper(self):
        """Test IPMT when period exceeds nper."""
        result = fn_ipmt(0.1, 13, 12, 10000)
        assert result == "#ERR!"

    def test_ipmt_zero_rate(self):
        """Test IPMT with zero rate."""
        result = fn_ipmt(0, 1, 12, 10000)
        assert result == 0


class TestPPMT:
    """Tests for PPMT (principal payment) function."""

    def test_basic_ppmt(self):
        """Test basic principal payment."""
        result = fn_ppmt(0.1, 1, 12, 10000)
        assert isinstance(result, float)
        assert result > 0

    def test_ppmt_later_period(self):
        """Test principal payment in later period."""
        result = fn_ppmt(0.1, 6, 12, 10000)
        assert isinstance(result, float)

    def test_ppmt_invalid_period(self):
        """Test PPMT with invalid period."""
        result = fn_ppmt(0.1, 0, 12, 10000)
        assert result == "#ERR!"

    def test_ppmt_zero_rate(self):
        """Test PPMT with zero rate."""
        result = fn_ppmt(0, 1, 12, 12000)
        assert result == 1000


class TestFunctionRegistry:
    """Test the function registry."""

    def test_all_functions_registered(self):
        """Test that all functions are in the registry."""
        assert "PMT" in FINANCIAL_FUNCTIONS
        assert "PV" in FINANCIAL_FUNCTIONS
        assert "FV" in FINANCIAL_FUNCTIONS
        assert "NPV" in FINANCIAL_FUNCTIONS
        assert "IRR" in FINANCIAL_FUNCTIONS
        assert "RATE" in FINANCIAL_FUNCTIONS
        assert "NPER" in FINANCIAL_FUNCTIONS
        assert "CTERM" in FINANCIAL_FUNCTIONS
        assert "TERM" in FINANCIAL_FUNCTIONS
        assert "SLN" in FINANCIAL_FUNCTIONS
        assert "SYD" in FINANCIAL_FUNCTIONS
        assert "DDB" in FINANCIAL_FUNCTIONS
        assert "IPMT" in FINANCIAL_FUNCTIONS
        assert "PPMT" in FINANCIAL_FUNCTIONS

    def test_functions_callable(self):
        """Test that all registered functions are callable."""
        for name, func in FINANCIAL_FUNCTIONS.items():
            assert callable(func), f"{name} is not callable"

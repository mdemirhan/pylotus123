"""Financial functions for the formula engine.

Implements Lotus 1-2-3 compatible financial functions:
@PMT, @PV, @FV, @NPV, @IRR, @RATE, @NPER, @CTERM
@SLN, @SYD, @DDB, @IPMT, @PPMT, @TERM
"""

from __future__ import annotations

import math
from typing import Any

from ...core.errors import FormulaError


def fn_pmt(principal: Any, rate: Any, nper: Any) -> float:
    """@PMT - Calculate loan payment.

    Usage: @PMT(principal, interest_rate, num_periods)

    Returns the periodic payment for a loan based on constant payments
    and a constant interest rate (ordinary annuity).
    """
    pv = float(principal)
    r = float(rate)
    n = float(nper)

    if r == 0:
        return float(pv / n)

    return float(pv * (r * (1 + r) ** n) / ((1 + r) ** n - 1))


def fn_pv(payment: Any, rate: Any, nper: Any) -> float:
    """@PV - Calculate present value.

    Usage: @PV(payment, interest_rate, num_periods)

    Returns the present value of an investment based on periodic,
    constant payments and a constant interest rate.
    """
    pmt = float(payment)
    r = float(rate)
    n = float(nper)

    if r == 0:
        return float(pmt * n)

    return float(pmt * ((1 - (1 + r) ** -n) / r))


def fn_fv(payment: Any, rate: Any, nper: Any) -> float:
    """@FV - Calculate future value.

    Usage: @FV(payment, interest_rate, num_periods)

    Returns the future value of an investment based on periodic,
    constant payments and a constant interest rate.
    """
    pmt = float(payment)
    r = float(rate)
    n = float(nper)

    if r == 0:
        return float(pmt * n)

    return float(pmt * (((1 + r) ** n - 1) / r))


def fn_npv(rate: Any, *cash_flows: Any) -> float:
    """@NPV - Calculate net present value.

    Usage: @NPV(discount_rate, value1, value2, ...)

    Returns the net present value of an investment based on a discount
    rate and a series of future payments (negative) and income (positive).
    """
    r = float(rate)

    # Flatten any lists in cash_flows
    flows = []
    for cf in cash_flows:
        if isinstance(cf, list):
            for item in cf:
                if isinstance(item, list):
                    flows.extend(item)
                else:
                    flows.append(item)
        else:
            flows.append(cf)

    npv = 0.0
    for i, cf in enumerate(flows):
        try:
            npv += float(cf) / ((1 + r) ** (i + 1))
        except (ValueError, TypeError):
            pass

    return npv


def fn_irr(guess: Any, *cash_flows: Any) -> float | str:
    """@IRR - Calculate internal rate of return.

    Usage: @IRR(guess, value1, value2, ...)

    Returns the internal rate of return for a series of cash flows.
    Uses Newton-Raphson iteration.
    """
    # Flatten cash flows
    flows: list[Any] = []
    for cf in cash_flows:
        if isinstance(cf, list):
            for item in cf:
                if isinstance(item, list):
                    flows.extend(item)
                else:
                    flows.append(item)
        else:
            flows.append(cf)

    float_flows = [
        float(f)
        for f in flows
        if isinstance(f, (int, float))
        or (isinstance(f, str) and f.replace(".", "").replace("-", "").isdigit())
    ]

    if not float_flows:
        return FormulaError.ERR

    rate = float(guess) if guess else 0.1

    # Newton-Raphson iteration
    for _ in range(100):
        npv = sum(cf / ((1 + rate) ** i) for i, cf in enumerate(float_flows))
        npv_deriv = sum(-i * cf / ((1 + rate) ** (i + 1)) for i, cf in enumerate(float_flows))

        if abs(npv_deriv) < 1e-10:
            break

        new_rate = rate - npv / npv_deriv

        if abs(new_rate - rate) < 1e-10:
            return new_rate

        rate = new_rate

    return rate


def fn_rate(nper: Any, pmt: Any, pv: Any, fv: Any = 0, guess: Any = 0.1) -> float:
    """@RATE - Calculate interest rate per period.

    Usage: @RATE(num_periods, payment, present_value, future_value, guess)

    Returns the interest rate per period of an annuity.
    """
    n = float(nper)
    payment = float(pmt)
    present = float(pv)
    future = float(fv) if fv else 0
    rate = float(guess) if guess else 0.1

    # Newton-Raphson iteration
    for _ in range(100):
        if rate == 0:
            y = present + payment * n + future
            dy = 0
        else:
            r1 = (1 + rate) ** n
            y = present * r1 + payment * (r1 - 1) / rate + future
            dy = present * n * (1 + rate) ** (n - 1) + payment * (
                (n * (1 + rate) ** (n - 1) * rate - (r1 - 1)) / (rate**2)
            )

        if abs(dy) < 1e-10:
            break

        new_rate = rate - y / dy

        if abs(new_rate - rate) < 1e-10:
            return new_rate

        rate = new_rate

    return rate


def fn_nper(rate: Any, pmt: Any, pv: Any, fv: Any = 0) -> float:
    """@NPER - Calculate number of periods.

    Usage: @NPER(interest_rate, payment, present_value, future_value)

    Returns the number of periods for an investment.
    """
    r = float(rate)
    payment = float(pmt)
    present = float(pv)
    future = float(fv) if fv else 0

    if r == 0:
        return -(present + future) / payment

    return math.log((payment - future * r) / (payment + present * r)) / math.log(1 + r)


def fn_cterm(rate: Any, fv: Any, pv: Any) -> float | str:
    """@CTERM - Calculate compounding term.

    Usage: @CTERM(interest_rate, future_value, present_value)

    Returns the number of compounding periods required for an investment
    to grow from present value to future value at a given interest rate.
    """
    r = float(rate)
    future = float(fv)
    present = float(pv)

    if r <= 0 or present <= 0 or future <= 0:
        return FormulaError.ERR

    return math.log(future / present) / math.log(1 + r)


def fn_term(pmt: Any, rate: Any, fv: Any) -> float:
    """@TERM - Calculate term of investment.

    Usage: @TERM(payment, interest_rate, future_value)

    Returns the number of payment periods needed to reach a future value.
    """
    payment = float(pmt)
    r = float(rate)
    future = float(fv)

    if r == 0:
        return future / payment

    return math.log(1 + (future * r / payment)) / math.log(1 + r)


def fn_sln(cost: Any, salvage: Any, life: Any) -> float | str:
    """@SLN - Straight-line depreciation.

    Usage: @SLN(cost, salvage_value, useful_life)

    Returns the straight-line depreciation of an asset for one period.
    """
    c = float(cost)
    s = float(salvage)
    n = float(life)

    if n == 0:
        return FormulaError.DIV_ZERO

    return (c - s) / n


def fn_syd(cost: Any, salvage: Any, life: Any, period: Any) -> float | str:
    """@SYD - Sum-of-years-digits depreciation.

    Usage: @SYD(cost, salvage_value, useful_life, period)

    Returns the sum-of-years-digits depreciation of an asset for a specified period.
    """
    c = float(cost)
    s = float(salvage)
    n = int(life)
    per = int(period)

    if n <= 0 or per <= 0 or per > n:
        return FormulaError.ERR

    sum_years = n * (n + 1) / 2
    return (c - s) * (n - per + 1) / sum_years


def fn_ddb(cost: Any, salvage: Any, life: Any, period: Any, factor: Any = 2) -> float | str:
    """@DDB - Double-declining balance depreciation.

    Usage: @DDB(cost, salvage_value, useful_life, period, factor)

    Returns the depreciation of an asset using the double-declining balance method.
    """
    c = float(cost)
    s = float(salvage)
    n = float(life)
    per = int(period)
    f = float(factor) if factor else 2

    if n <= 0 or per <= 0:
        return FormulaError.ERR

    # Calculate book value at start of period
    rate = f / n
    book_value = c
    for _ in range(per - 1):
        depreciation = book_value * rate
        book_value -= depreciation
        if book_value < s:
            book_value = s
            break

    # Calculate depreciation for this period
    depreciation = book_value * rate
    if book_value - depreciation < s:
        depreciation = book_value - s

    return max(0.0, depreciation)


def fn_ipmt(rate: Any, period: Any, nper: Any, pv: Any) -> float | str:
    """@IPMT - Interest portion of payment.

    Usage: @IPMT(interest_rate, period, num_periods, present_value)

    Returns the interest payment for a given period of an investment.
    """
    r = float(rate)
    per = int(period)
    n = int(nper)
    present = float(pv)

    if per < 1 or per > n:
        return FormulaError.ERR

    # Calculate payment
    if r == 0:
        pmt = present / n
    else:
        pmt = present * (r * (1 + r) ** n) / ((1 + r) ** n - 1)

    # Calculate remaining principal at start of period
    if r == 0:
        remaining = present - (per - 1) * pmt
    else:
        remaining = present * (1 + r) ** (per - 1) - pmt * ((1 + r) ** (per - 1) - 1) / r

    return remaining * r


def fn_ppmt(rate: Any, period: Any, nper: Any, pv: Any) -> float | str:
    """@PPMT - Principal portion of payment.

    Usage: @PPMT(interest_rate, period, num_periods, present_value)

    Returns the principal payment for a given period of an investment.
    """
    r = float(rate)
    per = int(period)
    n = int(nper)
    present = float(pv)

    if per < 1 or per > n:
        return FormulaError.ERR

    # Calculate payment
    if r == 0:
        pmt = present / n
        return pmt
    else:
        pmt = present * (r * (1 + r) ** n) / ((1 + r) ** n - 1)

    # Principal = Payment - Interest
    ipmt_result = fn_ipmt(rate, period, nper, pv)
    if isinstance(ipmt_result, str):
        return ipmt_result
    return pmt - ipmt_result


# Function registry for this module
FINANCIAL_FUNCTIONS = {
    "PMT": fn_pmt,
    "PV": fn_pv,
    "FV": fn_fv,
    "NPV": fn_npv,
    "IRR": fn_irr,
    "RATE": fn_rate,
    "NPER": fn_nper,
    "CTERM": fn_cterm,
    "TERM": fn_term,
    "SLN": fn_sln,
    "SYD": fn_syd,
    "DDB": fn_ddb,
    "IPMT": fn_ipmt,
    "PPMT": fn_ppmt,
}

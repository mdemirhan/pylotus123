"""Mathematical functions for the formula engine.

Implements Lotus 1-2-3 compatible math functions:
@SUM, @ABS, @INT, @ROUND, @MOD, @SQRT, @EXP, @LN, @LOG
Trigonometric: @SIN, @COS, @TAN, @ASIN, @ACOS, @ATAN, @ATAN2, @PI, @RAND
"""

import math
import random
from typing import Any


def _to_number(value: Any) -> float:
    """Convert value to number, returning 0 for non-numeric."""
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.replace(",", ""))
        except ValueError:
            return 0.0
    return 0.0


def _flatten_args(args: tuple) -> list:
    """Flatten nested lists in arguments."""
    result = []
    for arg in args:
        if isinstance(arg, list):
            result.extend(_flatten_args(tuple(arg)))
        else:
            result.append(arg)
    return result


def fn_sum(*args) -> float:
    """@SUM - Sum of all numeric values.

    Can take individual values, ranges, or mixed arguments.
    """
    values = _flatten_args(args)
    return sum(_to_number(v) for v in values if isinstance(v, (int, float, str)) and v != "")


def fn_abs(value: Any) -> float:
    """@ABS - Absolute value."""
    return abs(_to_number(value))


def fn_int(value: Any) -> int:
    """@INT - Integer portion (truncates toward negative infinity)."""
    n = _to_number(value)
    return int(math.floor(n))


def fn_round(value: Any, decimals: Any = 0) -> float:
    """@ROUND - Round to specified decimal places."""
    n = _to_number(value)
    d = int(_to_number(decimals))
    return round(n, d)


def fn_mod(dividend: Any, divisor: Any) -> float:
    """@MOD - Modulo (remainder after division)."""
    n = _to_number(dividend)
    d = _to_number(divisor)
    if d == 0:
        return float("nan")  # Will be converted to #DIV/0!
    return n % d


def fn_sqrt(value: Any) -> float:
    """@SQRT - Square root."""
    n = _to_number(value)
    if n < 0:
        return float("nan")  # Will be converted to #ERR!
    return math.sqrt(n)


def fn_exp(value: Any) -> float:
    """@EXP - e raised to power."""
    return math.exp(_to_number(value))


def fn_ln(value: Any) -> float:
    """@LN - Natural logarithm."""
    n = _to_number(value)
    if n <= 0:
        return float("nan")
    return math.log(n)


def fn_log(value: Any) -> float:
    """@LOG - Base-10 logarithm."""
    n = _to_number(value)
    if n <= 0:
        return float("nan")
    return math.log10(n)


def fn_sin(value: Any) -> float:
    """@SIN - Sine (argument in radians)."""
    return math.sin(_to_number(value))


def fn_cos(value: Any) -> float:
    """@COS - Cosine (argument in radians)."""
    return math.cos(_to_number(value))


def fn_tan(value: Any) -> float:
    """@TAN - Tangent (argument in radians)."""
    return math.tan(_to_number(value))


def fn_asin(value: Any) -> float:
    """@ASIN - Arc sine (result in radians)."""
    n = _to_number(value)
    if n < -1 or n > 1:
        return float("nan")
    return math.asin(n)


def fn_acos(value: Any) -> float:
    """@ACOS - Arc cosine (result in radians)."""
    n = _to_number(value)
    if n < -1 or n > 1:
        return float("nan")
    return math.acos(n)


def fn_atan(value: Any) -> float:
    """@ATAN - Arc tangent (result in radians)."""
    return math.atan(_to_number(value))


def fn_atan2(y: Any, x: Any) -> float:
    """@ATAN2 - Arc tangent of y/x (result in radians).

    Returns angle in correct quadrant.
    """
    return math.atan2(_to_number(y), _to_number(x))


def fn_pi() -> float:
    """@PI - The constant pi."""
    return math.pi


def fn_rand() -> float:
    """@RAND - Random number between 0 and 1."""
    return random.random()


def fn_power(base: Any, exponent: Any) -> float:
    """@POWER - Raise base to exponent (alternative to ^ operator)."""
    return float(_to_number(base) ** _to_number(exponent))


def fn_sign(value: Any) -> int:
    """@SIGN - Returns -1, 0, or 1 based on sign of value."""
    n = _to_number(value)
    if n > 0:
        return 1
    elif n < 0:
        return -1
    return 0


def fn_trunc(value: Any, decimals: Any = 0) -> float:
    """@TRUNC - Truncate to specified decimal places."""
    n = _to_number(value)
    d = int(_to_number(decimals))
    factor = 10.0**d
    return float(int(n * factor) / factor)


def fn_ceiling(value: Any) -> int:
    """@CEILING - Round up to nearest integer."""
    return int(math.ceil(_to_number(value)))


def fn_floor(value: Any) -> int:
    """@FLOOR - Round down to nearest integer."""
    return int(math.floor(_to_number(value)))


def fn_degrees(radians: Any) -> float:
    """@DEGREES - Convert radians to degrees."""
    return math.degrees(_to_number(radians))


def fn_radians(degrees: Any) -> float:
    """@RADIANS - Convert degrees to radians."""
    return math.radians(_to_number(degrees))


def fn_fact(n: Any) -> int:
    """@FACT - Factorial."""
    num = int(_to_number(n))
    if num < 0:
        return 0  # Error
    return math.factorial(num)


def fn_gcd(a: Any, b: Any) -> int:
    """@GCD - Greatest common divisor."""
    return math.gcd(int(_to_number(a)), int(_to_number(b)))


def fn_lcm(a: Any, b: Any) -> int:
    """@LCM - Least common multiple."""
    x, y = int(_to_number(a)), int(_to_number(b))
    return abs(x * y) // math.gcd(x, y) if x and y else 0


# Function registry for this module
MATH_FUNCTIONS = {
    # Basic math
    "SUM": fn_sum,
    "ABS": fn_abs,
    "INT": fn_int,
    "ROUND": fn_round,
    "MOD": fn_mod,
    "SQRT": fn_sqrt,
    "POWER": fn_power,
    "SIGN": fn_sign,
    "TRUNC": fn_trunc,
    "CEILING": fn_ceiling,
    "FLOOR": fn_floor,
    "FACT": fn_fact,
    "GCD": fn_gcd,
    "LCM": fn_lcm,
    # Exponential and logarithmic
    "EXP": fn_exp,
    "LN": fn_ln,
    "LOG": fn_log,
    # Trigonometric
    "SIN": fn_sin,
    "COS": fn_cos,
    "TAN": fn_tan,
    "ASIN": fn_asin,
    "ACOS": fn_acos,
    "ATAN": fn_atan,
    "ATAN2": fn_atan2,
    "DEGREES": fn_degrees,
    "RADIANS": fn_radians,
    # Constants and random
    "PI": fn_pi,
    "RAND": fn_rand,
}

"""Statistical functions for the formula engine.

Implements Lotus 1-2-3 compatible statistical functions:
@AVG, @COUNT, @MIN, @MAX, @STD, @VAR
"""
from __future__ import annotations

import math
from typing import Any


def _to_number(value: Any) -> float | None:
    """Convert value to number, returning None for non-numeric."""
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.replace(",", ""))
        except ValueError:
            return None
    return None


def _flatten_args(args: tuple) -> list:
    """Flatten nested lists in arguments."""
    result = []
    for arg in args:
        if isinstance(arg, list):
            result.extend(_flatten_args(tuple(arg)))
        else:
            result.append(arg)
    return result


def _get_numbers(args: tuple) -> list[float]:
    """Extract all numeric values from arguments."""
    values = _flatten_args(args)
    numbers = []
    for v in values:
        n = _to_number(v)
        if n is not None:
            numbers.append(n)
    return numbers


def fn_avg(*args) -> float:
    """@AVG - Arithmetic mean of numeric values.

    Also known as @AVERAGE.
    """
    numbers = _get_numbers(args)
    if not numbers:
        return 0.0
    return sum(numbers) / len(numbers)


def fn_count(*args) -> int:
    """@COUNT - Count of numeric values.

    Only counts cells containing numbers.
    """
    numbers = _get_numbers(args)
    return len(numbers)


def fn_counta(*args) -> int:
    """@COUNTA - Count of non-empty values.

    Counts all non-empty cells including text.
    """
    values = _flatten_args(args)
    return sum(1 for v in values if v != "" and v is not None)


def fn_countblank(*args) -> int:
    """@COUNTBLANK - Count of empty cells."""
    values = _flatten_args(args)
    return sum(1 for v in values if v == "" or v is None)


def fn_min(*args) -> float:
    """@MIN - Minimum numeric value."""
    numbers = _get_numbers(args)
    if not numbers:
        return 0.0
    return min(numbers)


def fn_max(*args) -> float:
    """@MAX - Maximum numeric value."""
    numbers = _get_numbers(args)
    if not numbers:
        return 0.0
    return max(numbers)


def fn_std(*args) -> float:
    """@STD - Sample standard deviation.

    Uses n-1 denominator (sample std dev).
    """
    numbers = _get_numbers(args)
    if len(numbers) < 2:
        return 0.0

    mean = sum(numbers) / len(numbers)
    variance = sum((x - mean) ** 2 for x in numbers) / (len(numbers) - 1)
    return math.sqrt(variance)


def fn_stds(*args) -> float:
    """@STDS - Alias for sample standard deviation."""
    return fn_std(*args)


def fn_stdp(*args) -> float:
    """@STDP - Population standard deviation.

    Uses n denominator (population std dev).
    """
    numbers = _get_numbers(args)
    if not numbers:
        return 0.0

    mean = sum(numbers) / len(numbers)
    variance = sum((x - mean) ** 2 for x in numbers) / len(numbers)
    return math.sqrt(variance)


def fn_var(*args) -> float:
    """@VAR - Sample variance.

    Uses n-1 denominator.
    """
    numbers = _get_numbers(args)
    if len(numbers) < 2:
        return 0.0

    mean = sum(numbers) / len(numbers)
    return sum((x - mean) ** 2 for x in numbers) / (len(numbers) - 1)


def fn_vars(*args) -> float:
    """@VARS - Alias for sample variance."""
    return fn_var(*args)


def fn_varp(*args) -> float:
    """@VARP - Population variance.

    Uses n denominator.
    """
    numbers = _get_numbers(args)
    if not numbers:
        return 0.0

    mean = sum(numbers) / len(numbers)
    return sum((x - mean) ** 2 for x in numbers) / len(numbers)


def fn_median(*args) -> float:
    """@MEDIAN - Middle value when sorted."""
    numbers = sorted(_get_numbers(args))
    if not numbers:
        return 0.0

    n = len(numbers)
    mid = n // 2
    if n % 2 == 0:
        return (numbers[mid - 1] + numbers[mid]) / 2
    return numbers[mid]


def fn_mode(*args) -> float:
    """@MODE - Most frequent value."""
    numbers = _get_numbers(args)
    if not numbers:
        return 0.0

    from collections import Counter
    counts = Counter(numbers)
    max_count = max(counts.values())
    if max_count == 1:
        return numbers[0]  # No mode, return first value
    modes = [k for k, v in counts.items() if v == max_count]
    return min(modes)  # Return smallest mode


def fn_large(*args) -> float:
    """@LARGE - k-th largest value.

    Usage: @LARGE(range, k)
    """
    if len(args) < 2:
        return float('nan')

    *range_args, k = args
    numbers = sorted(_get_numbers(tuple(range_args)), reverse=True)
    k_val = int(_to_number(k) or 1)

    if k_val < 1 or k_val > len(numbers):
        return float('nan')
    return numbers[k_val - 1]


def fn_small(*args) -> float:
    """@SMALL - k-th smallest value.

    Usage: @SMALL(range, k)
    """
    if len(args) < 2:
        return float('nan')

    *range_args, k = args
    numbers = sorted(_get_numbers(tuple(range_args)))
    k_val = int(_to_number(k) or 1)

    if k_val < 1 or k_val > len(numbers):
        return float('nan')
    return numbers[k_val - 1]


def fn_rank(*args) -> int:
    """@RANK - Rank of a value in a list.

    Usage: @RANK(value, range, order)
    order: 0 = descending (largest is 1), 1 = ascending
    """
    if len(args) < 2:
        return 0

    value = _to_number(args[0])
    if value is None:
        return 0

    numbers = _get_numbers(args[1:-1] if len(args) > 2 else (args[1],))
    order = int(_to_number(args[-1])) if len(args) > 2 else 0

    if order == 0:
        # Descending - largest is rank 1
        sorted_nums = sorted(numbers, reverse=True)
    else:
        # Ascending - smallest is rank 1
        sorted_nums = sorted(numbers)

    try:
        return sorted_nums.index(value) + 1
    except ValueError:
        return 0


def fn_percentile(*args) -> float:
    """@PERCENTILE - Value at given percentile.

    Usage: @PERCENTILE(range, k) where k is 0-1
    """
    if len(args) < 2:
        return float('nan')

    *range_args, k = args
    numbers = sorted(_get_numbers(tuple(range_args)))
    k_val = _to_number(k) or 0

    if not numbers or k_val < 0 or k_val > 1:
        return float('nan')

    n = len(numbers)
    idx = k_val * (n - 1)
    lower = int(idx)
    upper = lower + 1

    if upper >= n:
        return numbers[-1]

    frac = idx - lower
    return numbers[lower] * (1 - frac) + numbers[upper] * frac


def fn_quartile(*args) -> float:
    """@QUARTILE - Value at given quartile.

    Usage: @QUARTILE(range, quart) where quart is 0-4
    """
    if len(args) < 2:
        return float('nan')

    *range_args, quart = args
    q = int(_to_number(quart) or 0)

    if q < 0 or q > 4:
        return float('nan')

    # Convert quartile to percentile
    percentiles = {0: 0, 1: 0.25, 2: 0.5, 3: 0.75, 4: 1}
    return fn_percentile(*range_args, percentiles[q])


def fn_rand() -> float:
    """@RAND - Random number between 0 and 1."""
    import random
    return random.random()


def fn_randbetween(bottom: Any, top: Any) -> int:
    """@RANDBETWEEN - Random integer between two values."""
    import random
    b = int(_to_number(bottom) if bottom is not None else 0)
    t = int(_to_number(top) if top is not None else 1)
    return random.randint(b, t)


def fn_sumproduct(*args) -> float:
    """@SUMPRODUCT - Sum of products of corresponding elements.

    Usage: @SUMPRODUCT(array1, array2, ...)
    """
    arrays = []
    for arg in args:
        if isinstance(arg, list):
            flat = _flatten_args((arg,))
            arrays.append([_to_number(v) or 0 for v in flat])
        else:
            n = _to_number(arg)
            arrays.append([n if n is not None else 0])

    if not arrays:
        return 0.0

    # All arrays must be same length
    min_len = min(len(a) for a in arrays)

    result = 0.0
    for i in range(min_len):
        product = 1.0
        for arr in arrays:
            product *= arr[i]
        result += product

    return result


def fn_permut(n: Any, k: Any) -> int:
    """@PERMUT - Number of permutations.

    Usage: @PERMUT(n, k)
    Returns n! / (n-k)!
    """
    n_val = int(_to_number(n) if n is not None else 0)
    k_val = int(_to_number(k) if k is not None else 0)

    if n_val < 0 or k_val < 0 or k_val > n_val:
        return 0

    result = 1
    for i in range(n_val - k_val + 1, n_val + 1):
        result *= i
    return result


def fn_combin(n: Any, k: Any) -> int:
    """@COMBIN - Number of combinations.

    Usage: @COMBIN(n, k)
    Returns n! / (k! * (n-k)!)
    """
    n_val = int(_to_number(n) if n is not None else 0)
    k_val = int(_to_number(k) if k is not None else 0)

    if n_val < 0 or k_val < 0 or k_val > n_val:
        return 0

    # Use more efficient calculation
    if k_val > n_val - k_val:
        k_val = n_val - k_val

    result = 1
    for i in range(k_val):
        result = result * (n_val - i) // (i + 1)
    return result


def fn_fact(n: Any) -> int:
    """@FACT - Factorial."""
    n_val = int(_to_number(n) if n is not None else 0)
    if n_val < 0:
        return 0
    result = 1
    for i in range(2, n_val + 1):
        result *= i
    return result


def fn_sum(*args) -> float:
    """@SUM - Sum of numeric values."""
    numbers = _get_numbers(args)
    return sum(numbers)


def fn_sumsq(*args) -> float:
    """@SUMSQ - Sum of squares."""
    numbers = _get_numbers(args)
    return sum(x * x for x in numbers)


def fn_product(*args) -> float:
    """@PRODUCT - Product of numeric values."""
    numbers = _get_numbers(args)
    if not numbers:
        return 0.0
    result = 1.0
    for n in numbers:
        result *= n
    return result


def fn_geomean(*args) -> float:
    """@GEOMEAN - Geometric mean."""
    numbers = _get_numbers(args)
    if not numbers:
        return 0.0
    product = 1.0
    for n in numbers:
        if n <= 0:
            return 0.0
        product *= n
    return product ** (1 / len(numbers))


def fn_harmean(*args) -> float:
    """@HARMEAN - Harmonic mean."""
    numbers = _get_numbers(args)
    if not numbers:
        return 0.0
    reciprocal_sum = 0.0
    for n in numbers:
        if n == 0:
            return 0.0
        reciprocal_sum += 1 / n
    return len(numbers) / reciprocal_sum


# Function registry for this module
STATISTICAL_FUNCTIONS = {
    # Basic statistics
    "SUM": fn_sum,
    "AVG": fn_avg,
    "AVERAGE": fn_avg,
    "COUNT": fn_count,
    "COUNTA": fn_counta,
    "COUNTBLANK": fn_countblank,
    "MIN": fn_min,
    "MAX": fn_max,
    "PRODUCT": fn_product,

    # Dispersion
    "STD": fn_std,
    "STDS": fn_stds,
    "STDP": fn_stdp,
    "STDEV": fn_std,
    "VAR": fn_var,
    "VARS": fn_vars,
    "VARP": fn_varp,
    "SUMSQ": fn_sumsq,

    # Position
    "MEDIAN": fn_median,
    "MODE": fn_mode,
    "LARGE": fn_large,
    "SMALL": fn_small,
    "RANK": fn_rank,
    "PERCENTILE": fn_percentile,
    "QUARTILE": fn_quartile,

    # Random
    "RAND": fn_rand,
    "RANDBETWEEN": fn_randbetween,

    # Combinatorics
    "SUMPRODUCT": fn_sumproduct,
    "PERMUT": fn_permut,
    "COMBIN": fn_combin,
    "FACT": fn_fact,

    # Means
    "GEOMEAN": fn_geomean,
    "HARMEAN": fn_harmean,
}

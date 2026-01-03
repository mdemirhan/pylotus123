"""Formula error constants for the Lotus 1-2-3 clone.

These constants define all error values used throughout the formula system,
matching the Lotus 1-2-3 error conventions.
"""

from __future__ import annotations

from typing import Final


class FormulaError:
    """Standard formula error values.

    These match Lotus 1-2-3 error conventions and are used throughout
    the formula evaluation system.
    """

    DIV_ZERO: Final[str] = "#DIV/0!"  # Division by zero
    ERR: Final[str] = "#ERR!"  # Generic error
    CIRC: Final[str] = "#CIRC!"  # Circular reference
    REF: Final[str] = "#REF!"  # Invalid reference
    NAME: Final[str] = "#NAME?"  # Unknown function/name
    NA: Final[str] = "#N/A"  # Value not available
    NUM: Final[str] = "#NUM!"  # Invalid numeric value
    NULL: Final[str] = "#NULL!"  # Null intersection
    VALUE: Final[str] = "#VALUE!"  # Wrong value type

    # Set of all error values for checking
    ALL_ERRORS: Final[frozenset[str]] = frozenset(
        {
            "#DIV/0!",
            "#ERR!",
            "#CIRC!",
            "#REF!",
            "#NAME?",
            "#N/A",
            "#NUM!",
            "#NULL!",
            "#VALUE!",
        }
    )

    @classmethod
    def is_error(cls, value: object) -> bool:
        """Check if a value is an error string."""
        return isinstance(value, str) and value in cls.ALL_ERRORS


# Error type mappings for ERROR.TYPE function (Lotus 1-2-3 compatible)
ERROR_TYPE_MAP: Final[dict[str, int]] = {
    FormulaError.NULL: 1,
    FormulaError.DIV_ZERO: 2,
    FormulaError.VALUE: 3,
    FormulaError.REF: 4,
    FormulaError.NAME: 5,
    FormulaError.NUM: 6,
    FormulaError.NA: 7,
    FormulaError.CIRC: 8,
    FormulaError.ERR: 3,  # Maps to VALUE for consistency
}

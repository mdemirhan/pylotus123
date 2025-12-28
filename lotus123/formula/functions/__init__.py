"""Formula function implementations organized by category."""
from typing import Any, Callable
from .math import MATH_FUNCTIONS
from .statistical import STATISTICAL_FUNCTIONS
from .string import STRING_FUNCTIONS
from .logical import LOGICAL_FUNCTIONS
from .lookup import LOOKUP_FUNCTIONS
from .datetime import DATETIME_FUNCTIONS
from .info import INFO_FUNCTIONS
from .financial import FINANCIAL_FUNCTIONS
from .database import DATABASE_FUNCTIONS


class FunctionRegistry:
    """Registry of all available formula functions.

    Functions are registered by name and can be looked up for evaluation.
    Supports both Excel-style names and Lotus @-prefix style.
    """

    def __init__(self) -> None:
        self._functions: dict[str, Callable] = {}
        self._register_all()

    def _register_all(self) -> None:
        """Register all built-in functions."""
        for funcs in [
            MATH_FUNCTIONS,
            STATISTICAL_FUNCTIONS,
            STRING_FUNCTIONS,
            LOGICAL_FUNCTIONS,
            LOOKUP_FUNCTIONS,
            DATETIME_FUNCTIONS,
            INFO_FUNCTIONS,
            FINANCIAL_FUNCTIONS,
            DATABASE_FUNCTIONS,
        ]:
            self._functions.update(funcs)

    def get(self, name: str) -> Callable | None:
        """Get a function by name.

        Args:
            name: Function name (case-insensitive)

        Returns:
            Function callable or None if not found
        """
        return self._functions.get(name.upper())

    def exists(self, name: str) -> bool:
        """Check if a function exists."""
        return name.upper() in self._functions

    def register(self, name: str, func: Callable) -> None:
        """Register a custom function.

        Args:
            name: Function name
            func: Callable implementing the function
        """
        self._functions[name.upper()] = func

    def list_all(self) -> list[str]:
        """Get sorted list of all function names."""
        return sorted(self._functions.keys())

    def __contains__(self, name: str) -> bool:
        return self.exists(name)


def get_all_functions() -> dict[str, Callable]:
    """Get dictionary of all registered functions."""
    registry = FunctionRegistry()
    return dict(registry._functions)


__all__ = [
    "FunctionRegistry",
    "get_all_functions",
    "MATH_FUNCTIONS",
    "STATISTICAL_FUNCTIONS",
    "STRING_FUNCTIONS",
    "LOGICAL_FUNCTIONS",
    "LOOKUP_FUNCTIONS",
    "DATETIME_FUNCTIONS",
    "INFO_FUNCTIONS",
    "FINANCIAL_FUNCTIONS",
    "DATABASE_FUNCTIONS",
]

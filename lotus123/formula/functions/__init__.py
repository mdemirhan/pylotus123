"""Formula function implementations organized by category."""

from typing import Callable

from .database import DATABASE_FUNCTIONS
from .datetime import DATETIME_FUNCTIONS
from .financial import FINANCIAL_FUNCTIONS
from .info import INFO_FUNCTIONS
from .logical import LOGICAL_FUNCTIONS
from .lookup import LOOKUP_FUNCTIONS
from .math import MATH_FUNCTIONS
from .statistical import STATISTICAL_FUNCTIONS
from .string import STRING_FUNCTIONS


class FunctionRegistry:
    """Registry of all available formula functions.

    Functions are registered by name and can be looked up for evaluation.
    Supports both Excel-style names and Lotus @-prefix style.
    """

    def __init__(self) -> None:
        self._functions: dict = {}
        self._register_all()

    def _register_all(self) -> None:
        """Register all built-in functions."""
        self._functions.update(MATH_FUNCTIONS)
        self._functions.update(STATISTICAL_FUNCTIONS)
        self._functions.update(STRING_FUNCTIONS)
        self._functions.update(LOGICAL_FUNCTIONS)
        self._functions.update(LOOKUP_FUNCTIONS)
        self._functions.update(DATETIME_FUNCTIONS)
        self._functions.update(INFO_FUNCTIONS)
        self._functions.update(FINANCIAL_FUNCTIONS)
        self._functions.update(DATABASE_FUNCTIONS)

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

    @property
    def functions(self) -> dict[str, Callable]:
        """Read-only access to the functions dictionary."""
        return self._functions


def get_all_functions() -> dict:
    """Get dictionary of all registered functions."""
    registry = FunctionRegistry()
    return dict(registry.functions)


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
    "REGISTRY",
]

# Global singleton registry
REGISTRY = FunctionRegistry()

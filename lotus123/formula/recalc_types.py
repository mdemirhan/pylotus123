"""Recalculation types and enums."""

from dataclasses import dataclass
from enum import Enum, auto


class RecalcMode(Enum):
    """Recalculation mode."""

    AUTOMATIC = auto()  # Recalculate immediately when cells change
    MANUAL = auto()  # Only recalculate when requested (F9)


class RecalcOrder(Enum):
    """Order of recalculation for cells."""

    NATURAL = auto()  # Dependency-based (smart recalc)
    COLUMN_WISE = auto()  # Left to right, top to bottom (A1, A2, ..., B1, B2, ...)
    ROW_WISE = auto()  # Top to bottom, left to right (A1, B1, ..., A2, B2, ...)


@dataclass
class RecalcStats:
    """Statistics from a recalculation run."""

    cells_evaluated: int = 0
    circular_refs_found: int = 0
    errors_found: int = 0
    elapsed_ms: float = 0.0

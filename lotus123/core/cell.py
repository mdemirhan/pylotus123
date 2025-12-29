"""Cell data model with support for various data types and alignment."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Any


class CellType(Enum):
    """Types of cell content."""

    EMPTY = auto()
    NUMBER = auto()
    TEXT = auto()
    FORMULA = auto()
    DATE = auto()
    TIME = auto()
    DATETIME = auto()
    ERROR = auto()


class TextAlignment(Enum):
    """Text alignment options in Lotus 1-2-3 style.

    Lotus 1-2-3 uses prefix characters:
    - ' (apostrophe) = left aligned
    - " (quote) = right aligned
    - ^ (caret) = centered
    - \\ (backslash) = repeating character
    """

    DEFAULT = auto()  # Based on content type (numbers right, text left)
    LEFT = auto()  # Apostrophe prefix (')
    RIGHT = auto()  # Quotation mark prefix (")
    CENTER = auto()  # Caret prefix (^)
    REPEAT = auto()  # Backslash prefix (\\) for separator lines


# Prefix characters for alignment
ALIGNMENT_PREFIXES = {
    "'": TextAlignment.LEFT,
    '"': TextAlignment.RIGHT,
    "^": TextAlignment.CENTER,
    "\\": TextAlignment.REPEAT,
}

PREFIX_CHARS = {
    TextAlignment.LEFT: "'",
    TextAlignment.RIGHT: '"',
    TextAlignment.CENTER: "^",
    TextAlignment.REPEAT: "\\",
}


@dataclass
class Cell:
    """Represents a single cell in the spreadsheet.

    Attributes:
        raw_value: The raw string value as entered by user
        format_code: Format code for display (e.g., 'F2' for 2 decimal fixed)
        is_protected: Whether cell is protected from editing
        _cached_value: Cached computed value
        _cached_display: Cached display string
    """

    raw_value: str = ""
    format_code: str = "G"  # General format by default
    is_protected: bool = False
    _cached_value: Any = field(default=None, repr=False, compare=False)
    _cached_display: str | None = field(default=None, repr=False, compare=False)

    def __post_init__(self) -> None:
        """Clear cache on initialization."""
        self._cached_value = None
        self._cached_display = None

    @property
    def is_empty(self) -> bool:
        """Check if cell has no content."""
        return not self.raw_value

    @property
    def is_formula(self) -> bool:
        """Check if cell contains a formula.

        Lotus 1-2-3 formulas can start with:
        - = (standard)
        - @ (Lotus 1-2-3 function prefix)
        - + or - (followed by non-numeric content)
        """
        if not self.raw_value:
            return False
        first = self.raw_value[0]
        # = and @ always indicate a formula
        if first in "=@":
            return True
        # + or - are formulas only if followed by non-numeric content
        if first in "+-" and len(self.raw_value) > 1:
            rest = self.raw_value[1:]
            try:
                float(rest)
                return False
            except ValueError:
                return True
        return False

    @property
    def formula(self) -> str:
        """Return the formula without the leading prefix (=, @)."""
        if self.is_formula:
            first = self.raw_value[0]
            if first in "=@":
                return self.raw_value[1:]
            return self.raw_value  # For + formulas, keep the +
        return ""

    @property
    def alignment(self) -> TextAlignment:
        """Determine text alignment from content."""
        if not self.raw_value:
            return TextAlignment.DEFAULT
        first_char = self.raw_value[0]
        return ALIGNMENT_PREFIXES.get(first_char, TextAlignment.DEFAULT)

    @property
    def display_value(self) -> str:
        """Get the value for display, stripping alignment prefixes."""
        if not self.raw_value:
            return ""
        if self.raw_value[0] in ALIGNMENT_PREFIXES:
            return self.raw_value[1:]
        return self.raw_value

    @property
    def cell_type(self) -> CellType:
        """Determine the type of cell content."""
        if not self.raw_value:
            return CellType.EMPTY
        if self.is_formula:
            return CellType.FORMULA

        # Check for alignment prefix and get actual value
        value = self.display_value
        if not value:
            return CellType.TEXT

        # Check for error values
        if value.startswith("#"):
            return CellType.ERROR

        # Check for numeric
        try:
            float(value.replace(",", ""))
            return CellType.NUMBER
        except ValueError:
            pass

        return CellType.TEXT

    def get_aligned_display(self, width: int) -> str:
        """Get display value aligned to specified width.

        Args:
            width: Column width in characters

        Returns:
            String padded/truncated to width with proper alignment
        """
        value = self.display_value
        alignment = self.alignment

        # Handle repeating character
        if alignment == TextAlignment.REPEAT and value:
            return (value * (width // len(value) + 1))[:width]

        # Truncate if too long
        if len(value) > width:
            return value[:width]

        # Apply alignment
        if alignment == TextAlignment.LEFT:
            return value.ljust(width)
        elif alignment == TextAlignment.RIGHT:
            return value.rjust(width)
        elif alignment == TextAlignment.CENTER:
            return value.center(width)
        else:
            # Default: numbers right, text left
            if self.cell_type == CellType.NUMBER:
                return value.rjust(width)
            return value.ljust(width)

    def invalidate_cache(self) -> None:
        """Clear cached computed values."""
        self._cached_value = None
        self._cached_display = None

    def set_value(self, value: str) -> None:
        """Set raw value and invalidate cache."""
        self.raw_value = value
        self.invalidate_cache()

    def to_dict(self) -> dict:
        """Serialize cell to dictionary."""
        data: dict = {"raw_value": self.raw_value}
        if self.format_code != "G":
            data["format_code"] = self.format_code
        if self.is_protected:
            data["is_protected"] = True
        return data

    @classmethod
    def from_dict(cls, data: dict) -> Cell:
        """Deserialize cell from dictionary.

        Accepts both 'format_code' and legacy 'format_str' keys for file compatibility.
        """
        # Accept both format_code and legacy format_str
        format_code = data.get("format_code") or data.get("format_str") or "G"
        return cls(
            raw_value=data.get("raw_value", ""),
            format_code=format_code,
            is_protected=data.get("is_protected", False),
        )

    def copy(self) -> Cell:
        """Create a copy of this cell."""
        return Cell(
            raw_value=self.raw_value, format_code=self.format_code, is_protected=self.is_protected
        )

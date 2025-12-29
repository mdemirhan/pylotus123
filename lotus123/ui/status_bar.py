"""Status bar with Lotus 1-2-3 style mode indicators.

Status indicators include:
- Mode indicator (READY, VALUE, LABEL, EDIT, MENU, etc.)
- Cell address and contents
- Memory available
- Calculation indicator (CALC when recalculation needed)
- Lock indicators (CAPS, NUM, SCROLL)
- Circular reference warning (CIRC)
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum, auto
from typing import TYPE_CHECKING, Any

from textual.widgets import Static

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


class Mode(Enum):
    """Spreadsheet operation modes."""

    READY = auto()  # Ready for input
    VALUE = auto()  # Entering numeric value
    LABEL = auto()  # Entering text label
    EDIT = auto()  # Editing cell
    MENU = auto()  # In menu system
    POINT = auto()  # Pointing to cell for formula
    WAIT = auto()  # Processing
    ERROR = auto()  # Error state
    HELP = auto()  # Help mode
    FILES = auto()  # File operations
    NAMES = auto()  # Named range operations
    STAT = auto()  # Status display


@dataclass
class ModeIndicator:
    """Mode indicator display text."""

    mode: Mode = Mode.READY
    text: str = "READY"

    TEXTS = {
        Mode.READY: "READY",
        Mode.VALUE: "VALUE",
        Mode.LABEL: "LABEL",
        Mode.EDIT: "EDIT",
        Mode.MENU: "MENU",
        Mode.POINT: "POINT",
        Mode.WAIT: "WAIT",
        Mode.ERROR: "ERROR",
        Mode.HELP: "HELP",
        Mode.FILES: "FILES",
        Mode.NAMES: "NAMES",
        Mode.STAT: "STAT",
    }

    def set_mode(self, mode: Mode) -> None:
        """Set the current mode."""
        self.mode = mode
        self.text = self.TEXTS.get(mode, "READY")


@dataclass
class LockIndicators:
    """Keyboard lock state indicators."""

    caps_lock: bool = False
    num_lock: bool = False
    scroll_lock: bool = False

    def as_string(self) -> str:
        """Get indicators as string."""
        parts = []
        if self.caps_lock:
            parts.append("CAPS")
        if self.num_lock:
            parts.append("NUM")
        if self.scroll_lock:
            parts.append("SCRL")
        return " ".join(parts)


class StatusBar:
    """Status bar information manager.

    Provides all the status information for the bottom status bar:
    - Current mode
    - Cell reference and contents
    - Memory usage
    - Calculation status
    - Lock indicators
    - Warnings
    """

    def __init__(self, spreadsheet: Spreadsheet | None = None) -> None:
        self.spreadsheet = spreadsheet
        self.mode = ModeIndicator()
        self.locks = LockIndicators()
        self.current_cell_ref: str = "A1"
        self.current_cell_value: str = ""
        self.current_cell_formula: str = ""
        self.memory_available: int = 0
        self.needs_recalc: bool = False
        self.has_circular_ref: bool = False
        self.modified: bool = False
        self.message: str = ""

    def set_mode(self, mode: Mode) -> None:
        """Set the current mode."""
        self.mode.set_mode(mode)

    def update_cell(self, row: int, col: int) -> None:
        """Update status for current cell."""
        if not self.spreadsheet:
            return

        from ..core.reference import make_cell_ref

        self.current_cell_ref = make_cell_ref(row, col)

        cell = self.spreadsheet.get_cell_if_exists(row, col)
        if cell:
            if cell.is_formula:
                self.current_cell_formula = cell.raw_value
                self.current_cell_value = self.spreadsheet.get_display_value(row, col)
            else:
                self.current_cell_formula = ""
                self.current_cell_value = cell.raw_value
        else:
            self.current_cell_formula = ""
            self.current_cell_value = ""

    def update_from_spreadsheet(self) -> None:
        """Update status from spreadsheet state."""
        if not self.spreadsheet:
            return

        self.needs_recalc = self.spreadsheet.needs_recalc
        self.has_circular_ref = self.spreadsheet.has_circular_refs
        self.modified = self.spreadsheet.modified

    def get_cell_display(self) -> str:
        """Get cell reference and value display."""
        if self.current_cell_formula:
            return (
                f"{self.current_cell_ref}: {self.current_cell_formula} = {self.current_cell_value}"
            )
        elif self.current_cell_value:
            return f"{self.current_cell_ref}: {self.current_cell_value}"
        else:
            return f"{self.current_cell_ref}:"

    def get_indicators(self) -> str:
        """Get indicator string for status bar."""
        parts = []

        # Mode
        parts.append(self.mode.text)

        # Calculation needed
        if self.needs_recalc:
            parts.append("CALC")

        # Circular reference
        if self.has_circular_ref:
            parts.append("CIRC")

        # Modified
        if self.modified:
            parts.append("*")

        # Lock indicators
        lock_str = self.locks.as_string()
        if lock_str:
            parts.append(lock_str)

        return " ".join(parts)

    def get_full_status(self, width: int = 80) -> str:
        """Get full status bar string.

        Args:
            width: Available width

        Returns:
            Formatted status string
        """
        left = self.get_cell_display()
        right = self.get_indicators()

        # Add message if present
        if self.message:
            left = f"{left}  [{self.message}]"

        # Calculate padding
        padding = width - len(left) - len(right) - 2
        if padding < 1:
            padding = 1

        return f" {left}{' ' * padding}{right} "

    def set_message(self, message: str, duration_ms: int = 3000) -> None:
        """Set a temporary message.

        Args:
            message: Message to display
            duration_ms: How long to show (for UI to handle)
        """
        self.message = message

    def clear_message(self) -> None:
        """Clear the temporary message."""
        self.message = ""

    def format_memory(self, bytes_available: int) -> str:
        """Format memory display.

        Args:
            bytes_available: Available memory in bytes

        Returns:
            Formatted string like "256K" or "1.2M"
        """
        if bytes_available < 1024:
            return f"{bytes_available}B"
        elif bytes_available < 1024 * 1024:
            return f"{bytes_available // 1024}K"
        else:
            return f"{bytes_available / (1024 * 1024):.1f}M"


class StatusBarWidget(Static):
    """Textual widget for the Lotus 1-2-3 status bar.

    This widget wraps the StatusBar data class and renders it
    as a Textual Static widget with automatic updates.
    """

    def __init__(self, spreadsheet: Spreadsheet | None = None, **kwargs: Any) -> None:
        super().__init__(
            " A1:                                                                    READY ",
            **kwargs,
        )
        self._status = StatusBar(spreadsheet)

    def on_mount(self) -> None:
        """Initialize status bar content on mount."""
        self.refresh_status()

    @property
    def status(self) -> StatusBar:
        """Get the underlying StatusBar data object."""
        return self._status

    def set_spreadsheet(self, spreadsheet: Spreadsheet) -> None:
        """Set the spreadsheet reference."""
        self._status.spreadsheet = spreadsheet

    def set_mode(self, mode: Mode) -> None:
        """Set the current mode and refresh display."""
        self._status.set_mode(mode)
        self.refresh_status()

    def update_cell(self, row: int, col: int) -> None:
        """Update status for the current cell."""
        self._status.update_cell(row, col)
        self.refresh_status()

    def update_from_spreadsheet(self) -> None:
        """Update indicators from spreadsheet state."""
        self._status.update_from_spreadsheet()
        self.refresh_status()

    def set_modified(self, modified: bool) -> None:
        """Set the modified indicator."""
        self._status.modified = modified
        self.refresh_status()

    def set_message(self, message: str) -> None:
        """Set a temporary message."""
        self._status.set_message(message)
        self.refresh_status()

    def clear_message(self) -> None:
        """Clear the temporary message."""
        self._status.clear_message()
        self.refresh_status()

    def refresh_status(self) -> None:
        """Refresh the status bar display."""
        # Get available width (default to 80 if not mounted)
        try:
            width = self.size.width or 80
        except Exception:
            width = 80

        self.update(self._status.get_full_status(width))

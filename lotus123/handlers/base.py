"""Base protocol and handler class for app handlers."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Callable, Literal, Protocol

if TYPE_CHECKING:
    from textual.screen import Screen

    from ..charting import Chart
    from ..core import Spreadsheet
    from ..ui import SpreadsheetGrid
    from ..utils.undo import UndoManager


class AppProtocol(Protocol):
    """Protocol defining what handlers can access from the app.

    This provides type safety for handler classes by explicitly
    defining the interface they depend on.
    """

    # Core data
    spreadsheet: "Spreadsheet"
    chart: "Chart"
    undo_manager: "UndoManager"
    config: Any  # AppConfig

    # State flags
    editing: bool
    _dirty: bool
    _menu_active: bool

    # Global settings (public - shared across handlers)
    global_format_code: str
    global_label_prefix: str
    global_col_width: int
    global_zero_display: bool
    recalc_mode: str

    # Theme
    current_theme_type: Any  # ThemeType
    color_theme: Any  # Theme

    # App title/subtitle
    sub_title: str

    # App size (from Textual) - read via property
    @property
    def size(self) -> Any:
        """Get the app size."""
        ...

    # Methods from Textual App
    def push_screen(
        self,
        screen: "Screen[Any] | str",
        callback: Callable[[Any], None] | Callable[[Any], Any] | None = ...,
        wait_for_dismiss: bool = ...,
    ) -> Any:
        """Push a screen onto the screen stack."""
        ...

    def notify(
        self,
        message: str,
        *,
        title: str = "",
        severity: Literal["information", "warning", "error"] = "information",
        timeout: float | None = None,
    ) -> None:
        """Show a notification toast."""
        ...

    def query_one(
        self, selector: str | type, expect_type: type | None = None
    ) -> Any:
        """Query for a single widget."""
        ...

    def exit(self, result: Any = None) -> None:
        """Exit the application."""
        ...

    # Methods from LotusApp that handlers need
    def _update_status(self) -> None:
        """Update the status bar and cell reference display."""
        ...

    def _mark_dirty(self) -> None:
        """Mark the spreadsheet as having unsaved changes."""
        ...

    def _update_title(self) -> None:
        """Update the window title."""
        ...

    def _apply_theme(self) -> None:
        """Apply the current theme to all widgets."""
        ...


class BaseHandler:
    """Base class for all handlers providing common functionality."""

    def __init__(self, app: AppProtocol) -> None:
        self._app = app

    @property
    def spreadsheet(self) -> "Spreadsheet":
        """Access the spreadsheet data model."""
        return self._app.spreadsheet

    @property
    def undo_manager(self) -> "UndoManager":
        """Access the undo manager."""
        return self._app.undo_manager

    def notify(
        self,
        message: str,
        *,
        severity: Literal["information", "warning", "error"] = "information",
    ) -> None:
        """Show a notification to the user."""
        self._app.notify(message, severity=severity)

    def get_grid(self) -> "SpreadsheetGrid":
        """Get the spreadsheet grid widget."""
        from ..ui import SpreadsheetGrid

        return self._app.query_one("#grid", SpreadsheetGrid)

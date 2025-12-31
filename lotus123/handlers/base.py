"""Base protocol and handler class for app handlers."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Callable, Literal, Protocol, TypeVar, overload

if TYPE_CHECKING:
    from textual.screen import Screen

    from ..charting import Chart, TextChartRenderer
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

    # Chart renderer
    _chart_renderer: "TextChartRenderer"

    # State flags
    editing: bool
    _dirty: bool
    _menu_active: bool
    _recalc_mode: str

    # Clipboard state
    _cell_clipboard: tuple[int, int, str] | None
    _range_clipboard: list[list[str]] | None
    _clipboard_is_cut: bool
    _clipboard_origin: tuple[int, int]

    # Query state
    _query_input_range: tuple[int, int, int, int] | None
    _query_criteria_range: tuple[int, int, int, int] | None
    _query_output_range: tuple[int, int] | None
    _query_find_results: list[int] | None
    _query_find_index: int

    # Global settings
    _global_format_code: str
    _global_label_prefix: str
    _global_col_width: int
    _global_zero_display: bool

    # Pending operation state
    _pending_source_range: tuple[int, int, int, int]
    _pending_range: str

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

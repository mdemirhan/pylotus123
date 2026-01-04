"""Chart view screen.

Provides a modal screen for displaying rendered ASCII/Unicode charts.
"""

from __future__ import annotations

from typing import Any

from textual.app import ComposeResult
from textual.binding import Binding
from textual.containers import Container
from textual.screen import ModalScreen
from textual.widgets import Static


class ChartViewScreen(ModalScreen[None]):
    """Modal screen for viewing a chart."""

    BINDINGS = [
        Binding("escape", "dismiss_screen", "Close"),
        Binding("enter", "dismiss_screen", "Close"),
        Binding("q", "dismiss_screen", "Close"),
    ]

    CSS = """
    ChartViewScreen {
        align: center middle;
    }

    #chart-container {
        width: 80%;
        height: 80%;
        padding: 1;
    }

    #chart-content {
        width: 100%;
        height: 1fr;
        content-align: center middle;
    }

    #chart-footer {
        height: 1;
        dock: bottom;
        text-align: center;
    }
    """

    def __init__(self, chart_lines: list[str], **kwargs: Any) -> None:
        super().__init__(**kwargs)
        self.chart_lines = chart_lines

    def compose(self) -> ComposeResult:
        with Container(id="chart-container"):
            yield Static("\n".join(self.chart_lines), id="chart-content")
            yield Static("Press ESC or Enter to close", id="chart-footer")

    def on_mount(self) -> None:
        # Get theme fresh from the app's current setting
        from ..themes import THEMES

        theme_type = self.app.current_theme_type  # type: ignore[attr-defined]
        theme = THEMES[theme_type]

        container = self.query_one("#chart-container")
        container.styles.background = theme.cell_bg
        container.styles.border = ("thick", theme.accent)

    def action_dismiss_screen(self) -> None:
        self.dismiss(None)

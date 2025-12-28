"""Chart view screen.

Provides a modal screen for displaying rendered ASCII/Unicode charts.
"""
from __future__ import annotations

from textual.binding import Binding
from textual.containers import Container
from textual.screen import ModalScreen
from textual.widgets import Static
from textual.app import ComposeResult


class ChartViewScreen(ModalScreen):
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
        border: thick $accent;
        background: $surface;
        padding: 1;
    }

    #chart-content {
        width: 100%;
        height: 100%;
    }

    #chart-footer {
        height: 1;
        dock: bottom;
        text-align: center;
    }
    """

    def __init__(self, chart_lines: list[str], **kwargs):
        super().__init__(**kwargs)
        self.chart_lines = chart_lines

    def compose(self) -> ComposeResult:
        with Container(id="chart-container"):
            yield Static("\n".join(self.chart_lines), id="chart-content")
            yield Static("Press ESC or Enter to close", id="chart-footer")

    def action_dismiss_screen(self) -> None:
        self.dismiss(None)

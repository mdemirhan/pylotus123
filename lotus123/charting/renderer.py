"""Chart rendering for terminal display.

Renders charts using ASCII/Unicode characters for display in the TUI.
This module provides the main ChartRenderer interface and TextChartRenderer
implementation that delegates to type-specific renderers.
"""

from abc import ABC, abstractmethod
from typing import override

from ..core.spreadsheet_protocol import SpreadsheetProtocol
from .chart import Chart
from .renderers import RenderContext, get_renderer


class ChartRenderer(ABC):
    """Abstract base class for chart renderers."""

    @abstractmethod
    def render(self, chart: Chart, width: int, height: int) -> list[str]:
        """Render a chart to lines of text.

        Args:
            chart: Chart to render
            width: Width in characters
            height: Height in lines

        Returns:
            List of strings, one per line
        """
        pass


class TextChartRenderer(ChartRenderer):
    """Renders charts using ASCII/Unicode characters.

    This class delegates to type-specific renderers for each chart type.
    """

    def __init__(self, spreadsheet: SpreadsheetProtocol | None = None) -> None:
        """Initialize the text chart renderer.

        Args:
            spreadsheet: SpreadsheetProtocol instance for resolving range references
        """
        self.spreadsheet = spreadsheet

    @override
    def render(self, chart: Chart, width: int, height: int) -> list[str]:
        """Render a chart to text.

        Delegates to the appropriate type-specific renderer based on chart type.

        Args:
            chart: Chart configuration to render
            width: Width in characters
            height: Height in lines

        Returns:
            List of strings representing the rendered chart
        """
        # Create render context
        ctx = RenderContext(chart=chart, width=width, height=height, spreadsheet=self.spreadsheet)

        # Get the appropriate renderer for this chart type
        renderer = get_renderer(chart.chart_type)

        # Render and return
        return renderer.render(ctx)

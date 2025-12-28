"""Chart rendering for terminal display.

Renders charts using ASCII/Unicode characters for display in the TUI.
This module provides the main ChartRenderer interface and TextChartRenderer
implementation that delegates to type-specific renderers.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from typing import TYPE_CHECKING

from .chart import Chart, ChartType
from .renderers import get_renderer, RenderContext

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


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
    It maintains backward compatibility while using the new modular architecture.
    """

    def __init__(self, spreadsheet: Spreadsheet = None) -> None:
        """Initialize the text chart renderer.

        Args:
            spreadsheet: Spreadsheet instance for resolving range references
        """
        self.spreadsheet = spreadsheet

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
        ctx = RenderContext(
            chart=chart,
            width=width,
            height=height,
            spreadsheet=self.spreadsheet,
        )

        # Get the appropriate renderer for this chart type
        renderer = get_renderer(chart.chart_type)

        # Render and return
        return renderer.render(ctx)

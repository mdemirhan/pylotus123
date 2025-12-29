"""Base class and protocols for chart type renderers.

This module defines the interface that all chart type renderers must implement.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ...core.spreadsheet import Spreadsheet
    from ..chart import Chart, ChartSeries


# Box drawing characters shared across renderers
BOX_HORIZONTAL = "\u2500"
BOX_VERTICAL = "\u2502"
BOX_CORNER_TL = "\u250c"
BOX_CORNER_TR = "\u2510"
BOX_CORNER_BL = "\u2514"
BOX_CORNER_BR = "\u2518"
BOX_CROSS = "\u253c"
BOX_T_DOWN = "\u252c"
BOX_T_UP = "\u2534"
BOX_T_RIGHT = "\u251c"
BOX_T_LEFT = "\u2524"

# Bar characters
BAR_FULL = "\u2588"
BAR_LIGHT = "\u2591"
BAR_MEDIUM = "\u2592"
BAR_DARK = "\u2593"

# Plot characters
POINT = "*"
LINE_H = "-"
LINE_V = "|"


@dataclass
class RenderContext:
    """Context for rendering a chart.

    Contains all the information a renderer needs to produce output.
    """

    chart: Chart
    width: int
    height: int
    spreadsheet: Spreadsheet | None = None

    # Computed values (populated by prepare())
    all_values: list[float] = field(default_factory=list)
    min_val: float = 0.0
    max_val: float = 1.0
    plot_height: int = 0
    plot_width: int = 0

    def prepare(self, title_lines: int = 0, y_axis_width: int = 8) -> None:
        """Prepare computed values for rendering.

        Args:
            title_lines: Number of lines used for title
            y_axis_width: Width reserved for Y-axis labels
        """
        # Gather all values from series
        self.all_values = []
        for series in self.chart.series:
            self.all_values.extend(get_series_values(series, self.spreadsheet))

        # Calculate scale
        if self.all_values:
            self.min_val = (
                self.chart.y_axis.min_value
                if self.chart.y_axis.min_value is not None
                else min(self.all_values)
            )
            self.max_val = (
                self.chart.y_axis.max_value
                if self.chart.y_axis.max_value is not None
                else max(self.all_values)
            )
        else:
            self.min_val = 0.0
            self.max_val = 1.0

        if self.min_val == self.max_val:
            self.max_val = self.min_val + 1

        # Calculate plot area
        self.plot_height = self.height - title_lines - 4  # Reserve for axes/legend
        self.plot_width = self.width - y_axis_width


def get_series_values(series: ChartSeries, spreadsheet: Spreadsheet | None) -> list[float]:
    """Get numeric values from a chart series.

    Handles both direct values and spreadsheet range references.

    Args:
        series: The chart series to get values from
        spreadsheet: Spreadsheet instance for range lookups

    Returns:
        List of float values
    """
    if series.values:
        return series.values

    if series.data_range and spreadsheet:
        values = []
        if ":" in series.data_range:
            parts = series.data_range.split(":")
            flat_values = spreadsheet.get_range_flat(parts[0], parts[1])
            for v in flat_values:
                if isinstance(v, (int, float)):
                    values.append(float(v))
                elif isinstance(v, str):
                    try:
                        values.append(float(v))
                    except ValueError:
                        pass
        return values

    return []


class ChartTypeRenderer(ABC):
    """Abstract base class for chart type-specific renderers.

    Each chart type (line, bar, pie, etc.) implements this interface
    to provide its own rendering logic.
    """

    @abstractmethod
    def render(self, ctx: RenderContext) -> list[str]:
        """Render a chart to lines of text.

        Args:
            ctx: Render context containing chart data and dimensions

        Returns:
            List of strings, one per line
        """
        pass

    def render_title(self, ctx: RenderContext) -> list[str]:
        """Render chart title and subtitle.

        Args:
            ctx: Render context

        Returns:
            List of title lines
        """
        lines = []
        if ctx.chart.options.title:
            lines.append(ctx.chart.options.title.center(ctx.width))
        if ctx.chart.options.subtitle:
            lines.append(ctx.chart.options.subtitle.center(ctx.width))
        return lines

    def render_no_data(self, ctx: RenderContext) -> list[str]:
        """Render a no-data message.

        Args:
            ctx: Render context

        Returns:
            Lines with no-data message
        """
        lines = self.render_title(ctx)
        lines.append("No data to display".center(ctx.width))
        return lines

    def render_too_small(self, ctx: RenderContext) -> list[str]:
        """Render a too-small message.

        Args:
            ctx: Render context

        Returns:
            Lines with too-small message
        """
        lines = self.render_title(ctx)
        lines.append("Chart area too small".center(ctx.width))
        return lines

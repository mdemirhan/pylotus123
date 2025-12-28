"""Chart rendering for terminal display.

Renders charts using ASCII/Unicode characters for display in the TUI.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from typing import TYPE_CHECKING

from .chart import Chart, ChartType, ChartSeries

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
    """Renders charts using ASCII/Unicode characters."""

    # Box drawing characters
    BOX_HORIZONTAL = '\u2500'
    BOX_VERTICAL = '\u2502'
    BOX_CORNER_TL = '\u250c'
    BOX_CORNER_TR = '\u2510'
    BOX_CORNER_BL = '\u2514'
    BOX_CORNER_BR = '\u2518'
    BOX_CROSS = '\u253c'
    BOX_T_DOWN = '\u252c'
    BOX_T_UP = '\u2534'
    BOX_T_RIGHT = '\u251c'
    BOX_T_LEFT = '\u2524'

    # Bar characters
    BAR_FULL = '\u2588'
    BAR_LIGHT = '\u2591'
    BAR_MEDIUM = '\u2592'
    BAR_DARK = '\u2593'

    # Plot characters
    POINT = '*'
    LINE_H = '-'
    LINE_V = '|'

    def __init__(self, spreadsheet: Spreadsheet = None) -> None:
        self.spreadsheet = spreadsheet

    def render(self, chart: Chart, width: int, height: int) -> list[str]:
        """Render a chart to text."""
        if chart.chart_type == ChartType.LINE:
            return self._render_line_chart(chart, width, height)
        elif chart.chart_type == ChartType.BAR:
            return self._render_bar_chart(chart, width, height)
        elif chart.chart_type == ChartType.PIE:
            return self._render_pie_chart(chart, width, height)
        elif chart.chart_type == ChartType.XY_SCATTER:
            return self._render_scatter_chart(chart, width, height)
        else:
            return self._render_bar_chart(chart, width, height)

    def _get_series_values(self, series: ChartSeries) -> list[float]:
        """Get values for a series, loading from spreadsheet if needed."""
        if series.values:
            return series.values

        if series.data_range and self.spreadsheet:
            values = []
            # Parse range and get values
            if ':' in series.data_range:
                parts = series.data_range.split(':')
                flat_values = self.spreadsheet.get_range_flat(parts[0], parts[1])
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

    def _render_line_chart(self, chart: Chart, width: int, height: int) -> list[str]:
        """Render a line chart with data points connected by lines."""
        lines = []

        # Title
        if chart.options.title:
            lines.append(chart.options.title.center(width))
        if chart.options.subtitle:
            lines.append(chart.options.subtitle.center(width))

        # Get all values to determine scale
        all_values = []
        for series in chart.series:
            all_values.extend(self._get_series_values(series))

        if not all_values:
            lines.append("No data to display".center(width))
            return lines

        min_val = chart.y_axis.min_value if chart.y_axis.min_value is not None else min(all_values)
        max_val = chart.y_axis.max_value if chart.y_axis.max_value is not None else max(all_values)

        if min_val == max_val:
            max_val = min_val + 1

        # Calculate plot area
        plot_height = height - len(lines) - 4  # Reserve for axes and legend
        plot_width = width - 8  # Reserve for Y-axis labels

        if plot_height < 3 or plot_width < 10:
            lines.append("Chart area too small".center(width))
            return lines

        # Create plot grid
        plot = [[' ' for _ in range(plot_width)] for _ in range(plot_height)]

        # Plot each series - just show data points (cleanest for text mode)
        symbols = ['*', '+', 'o', 'x', '#', '@']

        for series_idx, series in enumerate(chart.series):
            values = self._get_series_values(series)
            if not values:
                continue

            symbol = symbols[series_idx % len(symbols)]

            # Plot each data point
            for i, val in enumerate(values):
                x = int((i / max(1, len(values) - 1)) * (plot_width - 1)) if len(values) > 1 else plot_width // 2
                y_ratio = (val - min_val) / (max_val - min_val)
                y = plot_height - 1 - round(y_ratio * (plot_height - 1))

                if 0 <= x < plot_width and 0 <= y < plot_height:
                    plot[y][x] = symbol

        # Add Y-axis labels and plot to lines
        for i, row in enumerate(plot):
            # Y-axis label
            if i == 0:
                label = f"{max_val:6.1f}"
            elif i == plot_height - 1:
                label = f"{min_val:6.1f}"
            elif i == plot_height // 2:
                mid = (max_val + min_val) / 2
                label = f"{mid:6.1f}"
            else:
                label = "      "

            lines.append(f"{label}{self.BOX_VERTICAL}{''.join(row)}")

        # X-axis
        lines.append("      " + self.BOX_CORNER_BL + self.BOX_HORIZONTAL * plot_width)

        # Legend
        if chart.options.show_legend and chart.series:
            legend_parts = []
            for i, series in enumerate(chart.series):
                symbol = symbols[i % len(symbols)]
                name = series.name or f"Series {i+1}"
                legend_parts.append(f"{symbol} {name}")
            lines.append(" | ".join(legend_parts).center(width))

        return lines

    def _render_bar_chart(self, chart: Chart, width: int, height: int) -> list[str]:
        """Render a vertical bar chart with clean, solid bars."""
        lines = []

        # Title
        if chart.options.title:
            lines.append(chart.options.title.center(width))

        # Get values
        all_values = []
        for series in chart.series:
            all_values.extend(self._get_series_values(series))

        if not all_values:
            lines.append("No data to display".center(width))
            return lines

        max_val = chart.y_axis.max_value if chart.y_axis.max_value is not None else max(all_values)
        min_val = chart.y_axis.min_value if chart.y_axis.min_value is not None else 0

        if max_val == min_val:
            max_val = min_val + 1

        # Calculate dimensions
        plot_height = height - len(lines) - 3
        num_bars = len(all_values)

        # Calculate bar width ensuring consistent spacing
        available_width = width - 8  # Reserve for Y-axis labels
        # Each bar needs bar_width + 2 (for spacing between bars)
        bar_width = max(3, (available_width - num_bars * 2) // max(1, num_bars))

        if plot_height < 3:
            lines.append("Chart area too small".center(width))
            return lines

        # Pre-calculate all bar heights for consistency
        bar_heights = []
        for val in all_values:
            ratio = (val - min_val) / (max_val - min_val)
            bar_heights.append(round(ratio * plot_height))

        # Use simple ASCII character for solid fill - works better across terminals
        FILL_CHAR = '#'

        # Create plot rows
        for row in range(plot_height):
            # Y-axis label - only show at top, middle, bottom
            if row == 0:
                y_label = f"{max_val:6.1f}"
            elif row == plot_height - 1:
                y_label = f"{min_val:6.1f}"
            elif row == plot_height // 2:
                mid_val = (max_val + min_val) / 2
                y_label = f"{mid_val:6.1f}"
            else:
                y_label = "      "

            line = y_label + self.BOX_VERTICAL

            row_from_bottom = plot_height - row - 1

            for bar_height in bar_heights:
                line += ' '  # Space before bar
                if row_from_bottom < bar_height:
                    line += FILL_CHAR * bar_width
                else:
                    line += ' ' * bar_width
                line += ' '  # Space after bar

            lines.append(line)

        # X-axis
        axis_width = num_bars * (bar_width + 2)
        lines.append("      " + self.BOX_CORNER_BL + self.BOX_HORIZONTAL * axis_width)

        return lines

    def _render_pie_chart(self, chart: Chart, width: int, height: int) -> list[str]:
        """Render a pie chart (simplified ASCII representation)."""
        lines = []

        # Title
        if chart.options.title:
            lines.append(chart.options.title.center(width))

        # Get values from first series
        if not chart.series:
            lines.append("No data to display".center(width))
            return lines

        values = self._get_series_values(chart.series[0])
        if not values:
            lines.append("No data to display".center(width))
            return lines

        total = sum(values)
        if total == 0:
            lines.append("All values are zero".center(width))
            return lines

        # Calculate percentages
        percentages = [v / total * 100 for v in values]

        # Simple text representation
        lines.append("")
        symbols = ['#', '@', '*', '+', '=', '-']
        for i, (val, pct) in enumerate(zip(values, percentages)):
            symbol = symbols[i % len(symbols)]
            bar_len = int(pct / 100 * (width - 30))
            label = f"  {symbol} {pct:5.1f}% {symbol * bar_len}"
            lines.append(label)

        lines.append("")
        lines.append(f"Total: {total}".center(width))

        return lines

    def _render_scatter_chart(self, chart: Chart, width: int, height: int) -> list[str]:
        """Render an XY scatter plot."""
        lines = []

        # Title
        if chart.options.title:
            lines.append(chart.options.title.center(width))

        # For scatter, we need X and Y values
        # X from x_range, Y from first series
        if not chart.series:
            lines.append("No data to display".center(width))
            return lines

        y_values = self._get_series_values(chart.series[0])
        if not y_values:
            lines.append("No data to display".center(width))
            return lines

        # X values: either from range or just indices
        x_values = list(range(len(y_values)))
        if chart.x_range and self.spreadsheet and ':' in chart.x_range:
            parts = chart.x_range.split(':')
            flat = self.spreadsheet.get_range_flat(parts[0], parts[1])
            x_values = []
            for v in flat:
                if isinstance(v, (int, float)):
                    x_values.append(float(v))

        if len(x_values) != len(y_values):
            x_values = list(range(len(y_values)))

        # Calculate bounds
        x_min, x_max = min(x_values), max(x_values)
        y_min, y_max = min(y_values), max(y_values)

        if x_min == x_max:
            x_max = x_min + 1
        if y_min == y_max:
            y_max = y_min + 1

        # Plot area
        plot_height = height - len(lines) - 3
        plot_width = width - 10

        if plot_height < 3 or plot_width < 10:
            lines.append("Chart area too small".center(width))
            return lines

        # Create grid
        plot = [[' ' for _ in range(plot_width)] for _ in range(plot_height)]

        # Plot points
        for x, y in zip(x_values, y_values):
            px = int(((x - x_min) / (x_max - x_min)) * (plot_width - 1))
            py = plot_height - 1 - int(((y - y_min) / (y_max - y_min)) * (plot_height - 1))

            if 0 <= px < plot_width and 0 <= py < plot_height:
                plot[py][px] = '*'

        # Output
        for i, row in enumerate(plot):
            if i == 0:
                label = f"{y_max:7.1f}"
            elif i == plot_height - 1:
                label = f"{y_min:7.1f}"
            else:
                label = "       "
            lines.append(f"{label}{self.BOX_VERTICAL}{''.join(row)}")

        lines.append("       " + self.BOX_CORNER_BL + self.BOX_HORIZONTAL * plot_width)

        return lines

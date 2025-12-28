"""Line chart renderer.

Renders line charts with data points and optional connecting lines.
"""
from __future__ import annotations

from .base import (
    ChartTypeRenderer,
    RenderContext,
    BOX_VERTICAL,
    BOX_CORNER_BL,
    BOX_HORIZONTAL,
    get_series_values,
)


class LineChartRenderer(ChartTypeRenderer):
    """Renders line charts with data points connected by lines."""

    # Symbols for different data series
    SYMBOLS = ['*', '+', 'o', 'x', '#', '@']

    def render(self, ctx: RenderContext) -> list[str]:
        """Render a line chart.

        Args:
            ctx: Render context with chart data and dimensions

        Returns:
            List of lines representing the chart
        """
        lines = self.render_title(ctx)

        # Prepare context with title line count
        ctx.prepare(title_lines=len(lines))

        if not ctx.all_values:
            return self.render_no_data(ctx)

        # Calculate Y-axis label width dynamically
        y_label_width = max(
            len(f"{ctx.max_val:.1f}"),
            len(f"{ctx.min_val:.1f}"),
            len(f"{(ctx.max_val + ctx.min_val) / 2:.1f}")
        )

        # Recalculate plot width with dynamic Y-axis
        ctx.plot_width = ctx.width - y_label_width - 1

        if ctx.plot_height < 3 or ctx.plot_width < 10:
            return self.render_too_small(ctx)

        # Create plot grid
        plot = [[' ' for _ in range(ctx.plot_width)] for _ in range(ctx.plot_height)]

        # Plot each series
        for series_idx, series in enumerate(ctx.chart.series):
            values = get_series_values(series, ctx.spreadsheet)
            if not values:
                continue

            symbol = self.SYMBOLS[series_idx % len(self.SYMBOLS)]
            self._plot_series(plot, values, symbol, ctx)

        # Build output with Y-axis labels
        lines.extend(self._build_plot_output(plot, ctx, y_label_width))

        # X-axis
        lines.append(" " * y_label_width + BOX_CORNER_BL + BOX_HORIZONTAL * ctx.plot_width)

        # X-axis title
        if ctx.chart.x_axis.title:
            lines.append("")
            lines.append(ctx.chart.x_axis.title.center(ctx.width))

        # Y-axis title
        if ctx.chart.y_axis.title:
            lines.append("")
            lines.append(f"Y: {ctx.chart.y_axis.title}".center(ctx.width))

        # Legend
        if ctx.chart.options.show_legend and ctx.chart.series:
            lines.append("")
            legend = self._build_legend(ctx)
            lines.append(legend.center(ctx.width))

        return lines

    def _plot_series(
        self,
        plot: list[list[str]],
        values: list[float],
        symbol: str,
        ctx: RenderContext
    ) -> None:
        """Plot a single data series onto the plot grid.

        Args:
            plot: 2D grid to plot onto
            values: Data values to plot
            symbol: Symbol to use for data points
            ctx: Render context
        """
        for i, val in enumerate(values):
            # Calculate X position
            if len(values) > 1:
                x = int((i / (len(values) - 1)) * (ctx.plot_width - 1))
            else:
                x = ctx.plot_width // 2

            # Calculate Y position (inverted - 0 at top)
            y_ratio = (val - ctx.min_val) / (ctx.max_val - ctx.min_val)
            y = ctx.plot_height - 1 - round(y_ratio * (ctx.plot_height - 1))

            # Plot if within bounds
            if 0 <= x < ctx.plot_width and 0 <= y < ctx.plot_height:
                plot[y][x] = symbol

    def _build_plot_output(
        self,
        plot: list[list[str]],
        ctx: RenderContext,
        y_label_width: int
    ) -> list[str]:
        """Build output lines with Y-axis labels.

        Args:
            plot: 2D plot grid
            ctx: Render context
            y_label_width: Width for Y-axis labels

        Returns:
            List of formatted output lines
        """
        lines = []
        for i, row in enumerate(plot):
            # Y-axis label at top, middle, bottom
            if i == 0:
                label = f"{ctx.max_val:.1f}".rjust(y_label_width)
            elif i == ctx.plot_height - 1:
                label = f"{ctx.min_val:.1f}".rjust(y_label_width)
            elif i == ctx.plot_height // 2:
                mid = (ctx.max_val + ctx.min_val) / 2
                label = f"{mid:.1f}".rjust(y_label_width)
            else:
                label = " " * y_label_width

            lines.append(f"{label}{BOX_VERTICAL}{''.join(row)}")

        return lines

    def _build_legend(self, ctx: RenderContext) -> str:
        """Build legend string.

        Args:
            ctx: Render context

        Returns:
            Formatted legend string
        """
        legend_parts = []
        for i, series in enumerate(ctx.chart.series):
            symbol = self.SYMBOLS[i % len(self.SYMBOLS)]
            name = series.name or f"Series {i+1}"
            legend_parts.append(f"{symbol} {name}")
        return " | ".join(legend_parts)

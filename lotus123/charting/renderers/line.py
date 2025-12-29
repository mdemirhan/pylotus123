"""Line chart renderer.

Renders line charts with data points and optional connecting lines.
"""

from __future__ import annotations

from .base import (
    BOX_CORNER_BL,
    BOX_HORIZONTAL,
    BOX_VERTICAL,
    ChartTypeRenderer,
    RenderContext,
    get_series_values,
    get_x_labels,
)


class LineChartRenderer(ChartTypeRenderer):
    """Renders line charts with data points connected by lines."""

    # Symbols for different data series
    SYMBOLS = ["*", "+", "o", "x", "#", "@"]

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
            len(f"{(ctx.max_val + ctx.min_val) / 2:.1f}"),
        )

        # Recalculate plot width with dynamic Y-axis
        ctx.plot_width = ctx.width - y_label_width - 1

        if ctx.plot_height < 3 or ctx.plot_width < 10:
            return self.render_too_small(ctx)

        # Create plot grid
        plot = [[" " for _ in range(ctx.plot_width)] for _ in range(ctx.plot_height)]

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

        # X-axis labels (from x_range) - aligned with data points
        x_labels = get_x_labels(ctx.chart, ctx.spreadsheet)
        if x_labels:
            num_labels = len(x_labels)
            # Build label line character by character
            label_chars = [" "] * ctx.plot_width

            # Calculate x positions using same formula as _plot_series
            for i, label in enumerate(x_labels):
                if num_labels > 1:
                    x = int((i / (num_labels - 1)) * (ctx.plot_width - 1))
                else:
                    x = ctx.plot_width // 2

                # Position label: left-align first, right-align last, center middle
                if i == 0:
                    # First label: left-align at x position
                    label_start = x
                elif i == num_labels - 1:
                    # Last label: right-align to ensure it fits
                    label_start = max(0, x - len(label) + 1)
                else:
                    # Middle labels: center around x position
                    label_start = max(0, x - len(label) // 2)

                for j, char in enumerate(label):
                    pos = label_start + j
                    if 0 <= pos < ctx.plot_width:
                        label_chars[pos] = char

            label_line = " " * (y_label_width + 1) + "".join(label_chars)
            lines.append(label_line[:ctx.width])

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
        self, plot: list[list[str]], values: list[float], symbol: str, ctx: RenderContext
    ) -> None:
        """Plot a single data series onto the plot grid with connecting lines.

        Args:
            plot: 2D grid to plot onto
            values: Data values to plot
            symbol: Symbol to use for data points
            ctx: Render context
        """
        points = []  # Store (x, y) positions for drawing lines

        for i, val in enumerate(values):
            # Calculate X position
            if len(values) > 1:
                x = int((i / (len(values) - 1)) * (ctx.plot_width - 1))
            else:
                x = ctx.plot_width // 2

            # Calculate Y position (inverted - 0 at top)
            y_ratio = (val - ctx.min_val) / (ctx.max_val - ctx.min_val)
            y = ctx.plot_height - 1 - round(y_ratio * (ctx.plot_height - 1))

            points.append((x, y))

        # Draw connecting lines between consecutive points
        for i in range(len(points) - 1):
            x1, y1 = points[i]
            x2, y2 = points[i + 1]
            self._draw_line(plot, x1, y1, x2, y2, ctx)

        # Plot data point symbols on top of lines
        for x, y in points:
            if 0 <= x < ctx.plot_width and 0 <= y < ctx.plot_height:
                plot[y][x] = symbol

    def _draw_line(
        self,
        plot: list[list[str]],
        x1: int,
        y1: int,
        x2: int,
        y2: int,
        ctx: RenderContext,
    ) -> None:
        """Draw a line between two points using ASCII characters.

        Args:
            plot: 2D grid to draw onto
            x1, y1: Start point
            x2, y2: End point
            ctx: Render context
        """
        dx = x2 - x1
        dy = y2 - y1

        if dx == 0:
            # Vertical line
            for y in range(min(y1, y2), max(y1, y2) + 1):
                if 0 <= y < ctx.plot_height and 0 <= x1 < ctx.plot_width:
                    if plot[y][x1] == " ":
                        plot[y][x1] = "|"
            return

        # Use Bresenham-like approach for diagonal lines
        steps = max(abs(dx), abs(dy))
        if steps == 0:
            return

        x_inc = dx / steps
        y_inc = dy / steps

        x, y = float(x1), float(y1)
        for _ in range(steps + 1):
            ix, iy = int(round(x)), int(round(y))
            if 0 <= ix < ctx.plot_width and 0 <= iy < ctx.plot_height:
                if plot[iy][ix] == " ":
                    # Choose character based on line direction
                    if abs(dy) < abs(dx) * 0.3:
                        # Mostly horizontal
                        plot[iy][ix] = "-"
                    elif dy < 0:
                        # Going up (remember y is inverted)
                        plot[iy][ix] = "/"
                    else:
                        # Going down
                        plot[iy][ix] = "\\"
            x += x_inc
            y += y_inc

    def _build_plot_output(
        self, plot: list[list[str]], ctx: RenderContext, y_label_width: int
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
        """Build legend string for series that have data.

        Args:
            ctx: Render context

        Returns:
            Formatted legend string
        """
        legend_parts = []
        for i, series in enumerate(ctx.chart.series):
            # Only include series that have actual data
            values = get_series_values(series, ctx.spreadsheet)
            if not values:
                continue
            symbol = self.SYMBOLS[i % len(self.SYMBOLS)]
            name = series.name or f"Series {i + 1}"
            legend_parts.append(f"{symbol} {name}")
        return " | ".join(legend_parts)

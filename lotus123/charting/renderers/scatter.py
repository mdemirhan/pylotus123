"""Scatter chart renderer.

Renders XY scatter plots with data points.
"""

from .base import (
    BOX_CORNER_BL,
    BOX_HORIZONTAL,
    BOX_VERTICAL,
    ChartTypeRenderer,
    RenderContext,
    get_series_values,
)


class ScatterChartRenderer(ChartTypeRenderer):
    """Renders XY scatter plots."""

    # Default point marker
    POINT_MARKER = "*"

    def render(self, ctx: RenderContext) -> list[str]:
        """Render an XY scatter plot.

        Uses X values from x_range (or indices) and Y values from first series.

        Args:
            ctx: Render context with chart data and dimensions

        Returns:
            List of lines representing the chart
        """
        lines = self.render_title(ctx)

        # Get Y values from first series
        if not ctx.chart.series:
            return self.render_no_data(ctx)

        y_values = get_series_values(ctx.chart.series[0], ctx.spreadsheet)
        if not y_values:
            return self.render_no_data(ctx)

        # Get X values from x_range or use indices
        x_values = self._get_x_values(ctx, len(y_values))

        # Ensure X and Y have same length
        if len(x_values) != len(y_values):
            x_values = list(range(len(y_values)))

        # Calculate bounds
        x_min, x_max = min(x_values), max(x_values)
        y_min, y_max = min(y_values), max(y_values)

        if x_min == x_max:
            x_max = x_min + 1
        if y_min == y_max:
            y_max = y_min + 1

        # Calculate Y-axis label width dynamically
        y_label_width = max(
            len(f"{y_max:.1f}"), len(f"{y_min:.1f}"), len(f"{(y_max + y_min) / 2:.1f}")
        )

        # Calculate plot area
        plot_height = ctx.height - len(lines) - 3
        plot_width = ctx.width - y_label_width - 1

        if plot_height < 3 or plot_width < 10:
            return self.render_too_small(ctx)

        # Create plot grid
        plot = [[" " for _ in range(plot_width)] for _ in range(plot_height)]

        # Plot points
        for x, y in zip(x_values, y_values):
            px = int(((x - x_min) / (x_max - x_min)) * (plot_width - 1))
            py = plot_height - 1 - int(((y - y_min) / (y_max - y_min)) * (plot_height - 1))

            if 0 <= px < plot_width and 0 <= py < plot_height:
                plot[py][px] = self.POINT_MARKER

        # Build output with Y-axis labels
        for i, row in enumerate(plot):
            if i == 0:
                label = f"{y_max:.1f}".rjust(y_label_width)
            elif i == plot_height - 1:
                label = f"{y_min:.1f}".rjust(y_label_width)
            else:
                label = " " * y_label_width
            lines.append(f"{label}{BOX_VERTICAL}{''.join(row)}")

        # X-axis
        lines.append(" " * y_label_width + BOX_CORNER_BL + BOX_HORIZONTAL * plot_width)

        # X-axis numeric labels
        x_label_line = " " * (y_label_width + 1)
        x_min_str = f"{x_min:.1f}"
        x_max_str = f"{x_max:.1f}"
        # Place min at left, max at right
        remaining = plot_width - len(x_min_str) - len(x_max_str)
        x_label_line += x_min_str + " " * max(0, remaining) + x_max_str
        lines.append(x_label_line[: ctx.width])

        # X-axis title
        if ctx.chart.x_axis.title:
            lines.append("")
            lines.append(ctx.chart.x_axis.title.center(ctx.width))

        # Y-axis title
        if ctx.chart.y_axis.title:
            lines.append("")
            lines.append(f"Y: {ctx.chart.y_axis.title}".center(ctx.width))

        # Legend for scatter - only show series with data
        if ctx.chart.options.show_legend and ctx.chart.series:
            lines.append("")
            legend_parts = []
            for i, series in enumerate(ctx.chart.series):
                vals = get_series_values(series, ctx.spreadsheet)
                if vals:
                    name = series.name or f"Series {i + 1}"
                    legend_parts.append(f"[*] {name}")
            if legend_parts:
                lines.append("  ".join(legend_parts).center(ctx.width))

        return lines

    def _get_x_values(self, ctx: RenderContext, default_count: int) -> list[float]:
        """Get X values from chart x_range or generate indices.

        Args:
            ctx: Render context
            default_count: Number of indices to generate if no x_range

        Returns:
            List of X values
        """
        # Default to indices
        x_values: list[float] = [float(i) for i in range(default_count)]

        if ctx.chart.x_range and ctx.spreadsheet and ":" in ctx.chart.x_range:
            parts = ctx.chart.x_range.split(":")
            flat = ctx.spreadsheet.get_range_flat(parts[0], parts[1])
            parsed_values = []
            for v in flat:
                if isinstance(v, (int, float)):
                    parsed_values.append(float(v))
                elif isinstance(v, str):
                    # Try to parse string as number
                    try:
                        parsed_values.append(float(v))
                    except ValueError:
                        pass
            if parsed_values:
                x_values = parsed_values

        return x_values

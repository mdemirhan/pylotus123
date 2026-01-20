"""Bar chart renderer.

Renders vertical bar charts with solid bars.
"""

from typing import override

from .base import (
    BOX_CORNER_BL,
    BOX_HORIZONTAL,
    BOX_VERTICAL,
    ChartTypeRenderer,
    RenderContext,
    get_series_values,
    get_x_labels,
)


class BarChartRenderer(ChartTypeRenderer):
    """Renders vertical bar charts with clean, solid bars."""

    # Fill characters for different series
    FILL_CHARS = ["#", "@", "*", "+", "=", "%"]

    @override
    def render(self, ctx: RenderContext) -> list[str]:
        """Render a bar chart.

        Args:
            ctx: Render context with chart data and dimensions

        Returns:
            List of lines representing the chart
        """
        lines = self.render_title(ctx)

        # Prepare context
        ctx.prepare(title_lines=len(lines))

        if not ctx.all_values:
            return self.render_no_data(ctx)

        # Override min to 0 for bar charts (typical behavior)
        if ctx.chart.y_axis.min_value is None:
            ctx.min_val = 0

        if ctx.max_val == ctx.min_val:
            ctx.max_val = ctx.min_val + 1

        # Calculate Y-axis label width dynamically
        y_label_width = max(
            len(f"{ctx.max_val:.1f}"),
            len(f"{ctx.min_val:.1f}"),
            len(f"{(ctx.max_val + ctx.min_val) / 2:.1f}"),
        )
        y_axis_width = y_label_width + 1  # +1 for vertical line

        # Get values per series for grouped bars
        series_values = []
        for series in ctx.chart.series:
            vals = get_series_values(series, ctx.spreadsheet)
            if vals:
                series_values.append(vals)

        if not series_values:
            return self.render_no_data(ctx)

        # Calculate bar dimensions
        num_groups = max(len(vals) for vals in series_values)
        num_series = len(series_values)
        available_width = ctx.width - y_axis_width - 4  # Reserve for Y-axis + padding

        # Each group has num_series bars + spacing
        group_width = max(num_series * 2, (available_width - num_groups * 2) // max(1, num_groups))
        bar_width = max(1, (group_width - 1) // max(1, num_series))

        if ctx.plot_height < 3:
            return self.render_too_small(ctx)

        # Build plot rows
        for row in range(ctx.plot_height):
            line = self._build_row(row, series_values, bar_width, ctx, y_label_width)
            lines.append(line)

        # X-axis line
        axis_width = num_groups * (num_series * (bar_width + 1) + 2)
        lines.append(" " * y_label_width + BOX_CORNER_BL + BOX_HORIZONTAL * axis_width)

        # X-axis labels (from x_range)
        x_labels = get_x_labels(ctx.chart, ctx.spreadsheet)
        if x_labels:
            label_line = " " * (y_label_width + 1)  # Offset for Y-axis
            group_width = num_series * (bar_width + 1) + 2
            for i, label in enumerate(x_labels[:num_groups]):
                # Truncate label to fit group width
                truncated = label[: group_width - 1] if len(label) >= group_width else label
                # Center the label under each group
                label_line += truncated.center(group_width)
            lines.append(label_line)

        # X-axis title
        if ctx.chart.x_axis.title:
            lines.append("")
            lines.append(ctx.chart.x_axis.title.center(ctx.width))

        # Y-axis title (shown at left of chart)
        if ctx.chart.y_axis.title:
            lines.append("")
            lines.append(f"Y: {ctx.chart.y_axis.title}".center(ctx.width))

        # Legend
        if ctx.chart.options.show_legend and ctx.chart.series:
            lines.append("")
            legend = self._build_legend(ctx)
            lines.append(legend.center(ctx.width))

        return lines

    def _build_row(
        self,
        row: int,
        series_values: list[list[float]],
        bar_width: int,
        ctx: RenderContext,
        y_label_width: int,
    ) -> str:
        """Build a single row of the bar chart.

        Args:
            row: Row index (0 = top)
            series_values: Values for each series
            bar_width: Width of each bar
            ctx: Render context
            y_label_width: Width for Y-axis labels

        Returns:
            Formatted row string
        """
        # Y-axis label
        if row == 0:
            y_label = f"{ctx.max_val:.1f}".rjust(y_label_width)
        elif row == ctx.plot_height - 1:
            y_label = f"{ctx.min_val:.1f}".rjust(y_label_width)
        elif row == ctx.plot_height // 2:
            mid_val = (ctx.max_val + ctx.min_val) / 2
            y_label = f"{mid_val:.1f}".rjust(y_label_width)
        else:
            y_label = " " * y_label_width

        line = y_label + BOX_VERTICAL

        row_from_bottom = ctx.plot_height - row - 1
        num_groups = max(len(vals) for vals in series_values)

        for group_idx in range(num_groups):
            line += " "  # Space before group
            for series_idx, vals in enumerate(series_values):
                fill_char = self.FILL_CHARS[series_idx % len(self.FILL_CHARS)]
                if group_idx < len(vals):
                    val = vals[group_idx]
                    ratio = (val - ctx.min_val) / (ctx.max_val - ctx.min_val)
                    bar_height = round(ratio * ctx.plot_height)
                    if row_from_bottom < bar_height:
                        line += fill_char * bar_width
                    else:
                        line += " " * bar_width
                else:
                    line += " " * bar_width
                line += " "  # Space between bars in group
            line += " "  # Extra space between groups

        return line

    def _build_legend(self, ctx: RenderContext) -> str:
        """Build legend string showing series names with their symbols.

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
            fill_char = self.FILL_CHARS[i % len(self.FILL_CHARS)]
            name = series.name or f"Series {i + 1}"
            legend_parts.append(f"[{fill_char}] {name}")
        return "  ".join(legend_parts)

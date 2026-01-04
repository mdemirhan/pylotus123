"""Stacked bar chart renderer.

Renders vertical bar charts with series values stacked on top of each other.
"""

from .base import (
    BOX_CORNER_BL,
    BOX_HORIZONTAL,
    BOX_VERTICAL,
    ChartTypeRenderer,
    RenderContext,
    get_series_values,
    get_x_labels,
)


class StackedBarChartRenderer(ChartTypeRenderer):
    """Renders stacked bar charts with values stacked vertically."""

    # Fill characters for different series
    FILL_CHARS = ["#", "@", "*", "+", "=", "%"]

    def render(self, ctx: RenderContext) -> list[str]:
        """Render a stacked bar chart.

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

        # Get values per series
        series_values: list[list[float]] = []
        for series in ctx.chart.series:
            vals = get_series_values(series, ctx.spreadsheet)
            if vals:
                series_values.append(vals)

        if not series_values:
            return self.render_no_data(ctx)

        # Calculate stacked totals per group for proper scaling
        num_groups = max(len(vals) for vals in series_values)
        stacked_totals = []
        for group_idx in range(num_groups):
            total = 0.0
            for vals in series_values:
                if group_idx < len(vals):
                    total += vals[group_idx]
            stacked_totals.append(total)

        # Override min to 0 for stacked bar charts
        ctx.min_val = 0
        ctx.max_val = max(stacked_totals) if stacked_totals else 1

        if ctx.max_val == ctx.min_val:
            ctx.max_val = ctx.min_val + 1

        # Calculate Y-axis label width dynamically
        y_label_width = max(
            len(f"{ctx.max_val:.1f}"),
            len(f"{ctx.min_val:.1f}"),
            len(f"{(ctx.max_val + ctx.min_val) / 2:.1f}"),
        )
        y_axis_width = y_label_width + 1

        # Calculate bar dimensions
        available_width = ctx.width - y_axis_width - 4

        # Each group has ONE stacked bar + spacing
        bar_width = max(2, (available_width - num_groups * 2) // max(1, num_groups))

        if ctx.plot_height < 3:
            return self.render_too_small(ctx)

        # Build plot rows
        for row in range(ctx.plot_height):
            line = self._build_row(row, series_values, bar_width, ctx, y_label_width)
            lines.append(line)

        # X-axis line
        axis_width = num_groups * (bar_width + 2)
        lines.append(" " * y_label_width + BOX_CORNER_BL + BOX_HORIZONTAL * axis_width)

        # X-axis labels
        x_labels = get_x_labels(ctx.chart, ctx.spreadsheet)
        if x_labels:
            label_line = " " * (y_label_width + 1)
            group_width = bar_width + 2
            for i, label in enumerate(x_labels[:num_groups]):
                truncated = label[: group_width - 1] if len(label) >= group_width else label
                label_line += truncated.center(group_width)
            lines.append(label_line)

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
            legend = self._build_legend(ctx, series_values)
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
        """Build a single row of the stacked bar chart.

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
            line += " "  # Space before bar

            # Calculate cumulative heights for stacked bars
            cumulative = 0.0
            fill_char = " "

            # Find which series this row belongs to (from bottom to top)
            for series_idx, vals in enumerate(series_values):
                if group_idx < len(vals):
                    val = vals[group_idx]
                    prev_cumulative = cumulative
                    cumulative += val

                    # Calculate height ranges
                    prev_ratio = prev_cumulative / (ctx.max_val - ctx.min_val)
                    curr_ratio = cumulative / (ctx.max_val - ctx.min_val)
                    prev_height = round(prev_ratio * ctx.plot_height)
                    curr_height = round(curr_ratio * ctx.plot_height)

                    # Check if this row falls within this series' portion
                    if prev_height <= row_from_bottom < curr_height:
                        fill_char = self.FILL_CHARS[series_idx % len(self.FILL_CHARS)]
                        break

            line += fill_char * bar_width
            line += " "  # Space after bar

        return line

    def _build_legend(self, ctx: RenderContext, series_values: list[list[float]]) -> str:
        """Build legend string showing series names with their symbols.

        Args:
            ctx: Render context
            series_values: List of values for each series (to track which have data)

        Returns:
            Formatted legend string
        """
        legend_parts = []
        series_with_data_idx = 0
        for i, series in enumerate(ctx.chart.series):
            values = get_series_values(series, ctx.spreadsheet)
            if not values:
                continue
            fill_char = self.FILL_CHARS[series_with_data_idx % len(self.FILL_CHARS)]
            name = series.name or f"Series {i + 1}"
            legend_parts.append(f"[{fill_char}] {name}")
            series_with_data_idx += 1
        return "  ".join(legend_parts)

"""Bar chart renderer.

Renders vertical bar charts with solid bars.
"""
from __future__ import annotations

from .base import (
    ChartTypeRenderer,
    RenderContext,
    BOX_VERTICAL,
    BOX_CORNER_BL,
    BOX_HORIZONTAL,
)


class BarChartRenderer(ChartTypeRenderer):
    """Renders vertical bar charts with clean, solid bars."""

    # Fill character for bars (ASCII # works well across terminals)
    FILL_CHAR = '#'

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
            len(f"{(ctx.max_val + ctx.min_val) / 2:.1f}")
        )
        y_axis_width = y_label_width + 1  # +1 for vertical line

        # Calculate bar dimensions
        num_bars = len(ctx.all_values)
        available_width = ctx.width - y_axis_width - 2  # Reserve for Y-axis labels + padding

        # Each bar needs bar_width + 2 (for spacing between bars)
        bar_width = max(3, (available_width - num_bars * 2) // max(1, num_bars))

        if ctx.plot_height < 3:
            return self.render_too_small(ctx)

        # Calculate bar heights
        bar_heights = self._calculate_bar_heights(ctx.all_values, ctx)

        # Build plot rows
        for row in range(ctx.plot_height):
            line = self._build_row(row, bar_heights, bar_width, ctx, y_label_width)
            lines.append(line)

        # X-axis
        axis_width = num_bars * (bar_width + 2)
        lines.append(" " * y_label_width + BOX_CORNER_BL + BOX_HORIZONTAL * axis_width)

        return lines

    def _calculate_bar_heights(
        self,
        values: list[float],
        ctx: RenderContext
    ) -> list[int]:
        """Calculate the height of each bar in rows.

        Args:
            values: Data values
            ctx: Render context

        Returns:
            List of bar heights in rows
        """
        bar_heights = []
        for val in values:
            ratio = (val - ctx.min_val) / (ctx.max_val - ctx.min_val)
            bar_heights.append(round(ratio * ctx.plot_height))
        return bar_heights

    def _build_row(
        self,
        row: int,
        bar_heights: list[int],
        bar_width: int,
        ctx: RenderContext,
        y_label_width: int
    ) -> str:
        """Build a single row of the bar chart.

        Args:
            row: Row index (0 = top)
            bar_heights: Heights of each bar
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

        for bar_height in bar_heights:
            line += ' '  # Space before bar
            if row_from_bottom < bar_height:
                line += self.FILL_CHAR * bar_width
            else:
                line += ' ' * bar_width
            line += ' '  # Space after bar

        return line

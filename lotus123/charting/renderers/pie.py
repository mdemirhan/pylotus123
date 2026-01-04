"""Pie chart renderer.

Renders pie charts using ASCII representation with percentage bars.
"""

from .base import ChartTypeRenderer, RenderContext, get_series_values, get_x_labels


class PieChartRenderer(ChartTypeRenderer):
    """Renders pie charts as ASCII percentage bars."""

    # Symbols for different slices
    SYMBOLS = ["#", "@", "*", "+", "=", "-"]

    def render(self, ctx: RenderContext) -> list[str]:
        """Render a pie chart.

        Since true circular pie charts are difficult in ASCII,
        this renders a horizontal bar representation showing
        each value's percentage of the total.

        Args:
            ctx: Render context with chart data and dimensions

        Returns:
            List of lines representing the chart
        """
        lines = self.render_title(ctx)

        # Get values from first series only (pie charts use one series)
        if not ctx.chart.series:
            return self.render_no_data(ctx)

        values = get_series_values(ctx.chart.series[0], ctx.spreadsheet)
        if not values:
            return self.render_no_data(ctx)

        total = sum(values)
        if total == 0:
            lines.append("All values are zero".center(ctx.width))
            return lines

        # Get labels from x_range
        labels = get_x_labels(ctx.chart, ctx.spreadsheet)

        # Calculate percentages
        percentages = [v / total * 100 for v in values]

        # Render each slice as a horizontal bar
        lines.append("")
        for i, (val, pct) in enumerate(zip(values, percentages)):
            label = labels[i] if i < len(labels) else f"Slice {i + 1}"
            line = self._render_slice(i, label, pct, ctx)
            lines.append(line)

        lines.append("")
        lines.append(f"Total: {total}".center(ctx.width))

        return lines

    def _render_slice(self, index: int, label: str, percentage: float, ctx: RenderContext) -> str:
        """Render a single pie slice as a horizontal bar.

        Args:
            index: Slice index (for symbol selection)
            label: The slice label
            percentage: Percentage of total
            ctx: Render context

        Returns:
            Formatted line for this slice
        """
        symbol = self.SYMBOLS[index % len(self.SYMBOLS)]
        # Truncate label to max 10 chars for alignment
        label_display = label[:10].ljust(10)
        bar_len = int(percentage / 100 * (ctx.width - 35))
        return f"  {symbol} {label_display} {percentage:5.1f}% {symbol * bar_len}"

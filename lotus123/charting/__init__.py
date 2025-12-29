"""Charting module for creating graphs and visualizations.

The charting system uses a modular architecture where each chart type
has its own renderer class. New chart types can be added by:

1. Creating a new renderer class in charting/renderers/
2. Implementing the ChartTypeRenderer interface
3. Registering it in the RENDERER_REGISTRY

Available chart types:
- LINE: Line charts with data points
- BAR: Vertical bar charts
- PIE: Pie charts (as horizontal percentage bars)
- XY_SCATTER: XY scatter plots
- STACKED_BAR, AREA, HLBAR: Fall back to bar chart rendering
"""

from .chart import Chart, ChartAxis, ChartOptions, ChartSeries, ChartType, LineStyle, ScaleType
from .renderer import ChartRenderer, TextChartRenderer
from .renderers import (
    RENDERER_REGISTRY,
    BarChartRenderer,
    ChartTypeRenderer,
    LineChartRenderer,
    PieChartRenderer,
    RenderContext,
    ScatterChartRenderer,
    get_renderer,
)

__all__ = [
    # Data model
    "Chart",
    "ChartType",
    "ChartSeries",
    "ChartOptions",
    "ChartAxis",
    "LineStyle",
    "ScaleType",
    # Main renderer interface
    "ChartRenderer",
    "TextChartRenderer",
    # Type-specific renderers
    "ChartTypeRenderer",
    "RenderContext",
    "LineChartRenderer",
    "BarChartRenderer",
    "PieChartRenderer",
    "ScatterChartRenderer",
    "RENDERER_REGISTRY",
    "get_renderer",
]

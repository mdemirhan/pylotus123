"""Chart type renderers.

Each chart type has its own renderer class implementing the ChartTypeRenderer protocol.
"""
from .base import ChartTypeRenderer, RenderContext
from .line import LineChartRenderer
from .bar import BarChartRenderer
from .pie import PieChartRenderer
from .scatter import ScatterChartRenderer

# Registry mapping ChartType to renderer class
from ..chart import ChartType

RENDERER_REGISTRY: dict[ChartType, type[ChartTypeRenderer]] = {
    ChartType.LINE: LineChartRenderer,
    ChartType.BAR: BarChartRenderer,
    ChartType.STACKED_BAR: BarChartRenderer,  # Falls back to bar
    ChartType.AREA: BarChartRenderer,  # Falls back to bar
    ChartType.HLBAR: BarChartRenderer,  # Falls back to bar
    ChartType.PIE: PieChartRenderer,
    ChartType.XY_SCATTER: ScatterChartRenderer,
}


def get_renderer(chart_type: ChartType) -> ChartTypeRenderer:
    """Get a renderer instance for the given chart type.

    Args:
        chart_type: The type of chart to render

    Returns:
        A renderer instance for that chart type
    """
    renderer_class = RENDERER_REGISTRY.get(chart_type, BarChartRenderer)
    return renderer_class()


__all__ = [
    "ChartTypeRenderer",
    "RenderContext",
    "LineChartRenderer",
    "BarChartRenderer",
    "PieChartRenderer",
    "ScatterChartRenderer",
    "RENDERER_REGISTRY",
    "get_renderer",
]

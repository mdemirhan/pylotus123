"""Tests for charting module."""

import pytest

from lotus123.charting.chart import (
    Chart,
    ChartAxis,
    ChartOptions,
    ChartSeries,
    ChartType,
    LineStyle,
    ScaleType,
)
from lotus123.charting.renderer import TextChartRenderer
from lotus123.charting.renderers import get_renderer
from lotus123.charting.renderers.bar import BarChartRenderer
from lotus123.charting.renderers.base import ChartTypeRenderer, RenderContext
from lotus123.charting.renderers.line import LineChartRenderer
from lotus123.charting.renderers.pie import PieChartRenderer
from lotus123.charting.renderers.scatter import ScatterChartRenderer


class TestChartType:
    """Tests for ChartType enum."""

    def test_all_types_exist(self):
        """Test all chart types exist."""
        assert ChartType.LINE
        assert ChartType.BAR
        assert ChartType.STACKED_BAR
        assert ChartType.XY_SCATTER
        assert ChartType.PIE
        assert ChartType.AREA
        assert ChartType.HLBAR


class TestLineStyle:
    """Tests for LineStyle enum."""

    def test_all_styles_exist(self):
        """Test all line styles exist."""
        assert LineStyle.LINES
        assert LineStyle.SYMBOLS
        assert LineStyle.BOTH
        assert LineStyle.NEITHER


class TestScaleType:
    """Tests for ScaleType enum."""

    def test_all_scale_types_exist(self):
        """Test all scale types exist."""
        assert ScaleType.LINEAR
        assert ScaleType.LOGARITHMIC


class TestChartSeries:
    """Tests for ChartSeries class."""

    def test_basic_series(self):
        """Test basic series creation."""
        series = ChartSeries(name="Sales", values=[10, 20, 30])
        assert series.name == "Sales"
        assert series.values == [10, 20, 30]

    def test_series_with_range(self):
        """Test series with data_range."""
        series = ChartSeries(name="XY", data_range="A1:A10")
        assert series.data_range == "A1:A10"

    def test_series_line_style(self):
        """Test series line style."""
        series = ChartSeries(name="Test", line_style=LineStyle.SYMBOLS)
        assert series.line_style == LineStyle.SYMBOLS


class TestChartAxis:
    """Tests for ChartAxis class."""

    def test_default_axis(self):
        """Test default axis settings."""
        axis = ChartAxis()
        assert axis.title == ""
        assert axis.scale_type == ScaleType.LINEAR
        assert axis.min_value is None

    def test_custom_axis(self):
        """Test custom axis settings."""
        axis = ChartAxis(title="Revenue", min_value=0, max_value=100, scale_type=ScaleType.LOGARITHMIC)
        assert axis.title == "Revenue"
        assert axis.min_value == 0
        assert axis.max_value == 100


class TestChartOptions:
    """Tests for ChartOptions class."""

    def test_default_options(self):
        """Test default options."""
        opts = ChartOptions()
        assert opts.show_legend is True
        assert opts.title == ""

    def test_custom_options(self):
        """Test custom options."""
        opts = ChartOptions(title="My Chart", show_legend=False, legend_position="right")
        assert opts.title == "My Chart"
        assert opts.show_legend is False
        assert opts.legend_position == "right"


class TestChart:
    """Tests for Chart class."""

    def test_basic_chart(self):
        """Test basic chart creation."""
        chart = Chart(chart_type=ChartType.LINE)
        assert chart.chart_type == ChartType.LINE
        assert len(chart.series) == 0

    def test_add_series(self):
        """Test adding series to chart."""
        chart = Chart(chart_type=ChartType.BAR)
        series = ChartSeries(name="Data", values=[1, 2, 3])
        chart.series.append(series)
        assert len(chart.series) == 1

    def test_chart_with_axes(self):
        """Test chart with custom axes."""
        x_axis = ChartAxis(title="X")
        y_axis = ChartAxis(title="Y")
        chart = Chart(chart_type=ChartType.LINE, x_axis=x_axis, y_axis=y_axis)
        assert chart.x_axis.title == "X"
        assert chart.y_axis.title == "Y"

    def test_chart_with_options(self):
        """Test chart with options."""
        opts = ChartOptions(title="Test")
        chart = Chart(chart_type=ChartType.PIE, options=opts)
        assert chart.options.title == "Test"


class TestTextChartRenderer:
    """Tests for TextChartRenderer class."""

    def setup_method(self):
        """Set up test chart."""
        self.series = ChartSeries(name="Data", values=[10, 20, 30, 40])
        self.chart = Chart(chart_type=ChartType.LINE)
        self.chart.series.append(self.series)

    def test_render_line_chart(self):
        """Test rendering line chart."""
        renderer = TextChartRenderer()
        result = renderer.render(self.chart, width=40, height=10)
        assert isinstance(result, list)
        assert len(result) > 0

    def test_render_bar_chart(self):
        """Test rendering bar chart."""
        self.chart.chart_type = ChartType.BAR
        renderer = TextChartRenderer()
        result = renderer.render(self.chart, width=40, height=10)
        assert isinstance(result, list)

    def test_render_pie_chart(self):
        """Test rendering pie chart."""
        self.chart.chart_type = ChartType.PIE
        renderer = TextChartRenderer()
        result = renderer.render(self.chart, width=40, height=10)
        assert isinstance(result, list)

    def test_render_scatter_chart(self):
        """Test rendering scatter chart."""
        self.chart.chart_type = ChartType.XY_SCATTER
        renderer = TextChartRenderer()
        result = renderer.render(self.chart, width=40, height=10)
        assert isinstance(result, list)


class TestRenderContext:
    """Tests for RenderContext class."""

    def test_render_context_creation(self):
        """Test creating render context."""
        chart = Chart(chart_type=ChartType.LINE)
        ctx = RenderContext(chart=chart, width=80, height=24)
        assert ctx.width == 80
        assert ctx.height == 24


class TestGetRenderer:
    """Tests for get_renderer function."""

    def test_get_line_renderer(self):
        """Test getting line chart renderer."""
        renderer = get_renderer(ChartType.LINE)
        assert isinstance(renderer, LineChartRenderer)

    def test_get_bar_renderer(self):
        """Test getting bar chart renderer."""
        renderer = get_renderer(ChartType.BAR)
        assert isinstance(renderer, BarChartRenderer)

    def test_get_pie_renderer(self):
        """Test getting pie chart renderer."""
        renderer = get_renderer(ChartType.PIE)
        assert isinstance(renderer, PieChartRenderer)

    def test_get_scatter_renderer(self):
        """Test getting scatter chart renderer."""
        renderer = get_renderer(ChartType.XY_SCATTER)
        assert isinstance(renderer, ScatterChartRenderer)


class TestLineChartRenderer:
    """Tests for LineChartRenderer class."""

    def test_render(self):
        """Test rendering line chart."""
        chart = Chart(chart_type=ChartType.LINE)
        chart.series.append(ChartSeries(name="Data", values=[10, 20, 15, 25]))
        ctx = RenderContext(chart=chart, width=40, height=10)

        renderer = LineChartRenderer()
        lines = renderer.render(ctx)  # Chart is in context
        assert isinstance(lines, list)


class TestBarChartRenderer:
    """Tests for BarChartRenderer class."""

    def test_render(self):
        """Test rendering bar chart."""
        chart = Chart(chart_type=ChartType.BAR)
        chart.series.append(ChartSeries(name="Data", values=[10, 20, 30]))
        ctx = RenderContext(chart=chart, width=40, height=10)

        renderer = BarChartRenderer()
        lines = renderer.render(ctx)  # Chart is in context
        assert isinstance(lines, list)


class TestPieChartRenderer:
    """Tests for PieChartRenderer class."""

    def test_render(self):
        """Test rendering pie chart."""
        chart = Chart(chart_type=ChartType.PIE)
        chart.series.append(ChartSeries(name="Data", values=[30, 20, 50]))
        ctx = RenderContext(chart=chart, width=40, height=10)

        renderer = PieChartRenderer()
        lines = renderer.render(ctx)  # Chart is in context
        assert isinstance(lines, list)


class TestScatterChartRenderer:
    """Tests for ScatterChartRenderer class."""

    def test_render(self):
        """Test rendering scatter chart."""
        chart = Chart(chart_type=ChartType.XY_SCATTER)
        series = ChartSeries(name="XY", values=[1, 4, 9, 16])
        chart.series.append(series)
        ctx = RenderContext(chart=chart, width=40, height=10)

        renderer = ScatterChartRenderer()
        lines = renderer.render(ctx)  # Chart is in context
        assert isinstance(lines, list)


class TestChartMethods:
    """Tests for Chart class methods."""

    def test_add_series_with_values(self):
        """Test add_series with values."""
        chart = Chart(chart_type=ChartType.LINE)
        series = chart.add_series("Sales", values=[10, 20, 30])
        assert series.name == "Sales"
        assert series.values == [10, 20, 30]
        assert len(chart.series) == 1

    def test_add_series_with_range(self):
        """Test add_series with data range."""
        chart = Chart(chart_type=ChartType.BAR)
        series = chart.add_series("Data", data_range="A1:A10")
        assert series.data_range == "A1:A10"

    def test_set_x_range(self):
        """Test set_x_range method."""
        chart = Chart()
        chart.set_x_range("B1:B10")
        assert chart.x_range == "B1:B10"

    def test_set_type(self):
        """Test set_type method."""
        chart = Chart()
        chart.set_type(ChartType.PIE)
        assert chart.chart_type == ChartType.PIE

    def test_set_title(self):
        """Test set_title method."""
        chart = Chart()
        chart.set_title("Sales Report", "Q4 2023")
        assert chart.options.title == "Sales Report"
        assert chart.options.subtitle == "Q4 2023"

    def test_set_axis_titles(self):
        """Test set_axis_titles method."""
        chart = Chart()
        chart.set_axis_titles("Time", "Revenue")
        assert chart.x_axis.title == "Time"
        assert chart.y_axis.title == "Revenue"

    def test_set_scale_y_axis(self):
        """Test set_scale for y axis."""
        chart = Chart()
        chart.set_scale("y", min_val=0, max_val=100, scale_type=ScaleType.LINEAR)
        assert chart.y_axis.min_value == 0
        assert chart.y_axis.max_value == 100
        assert chart.y_axis.scale_type == ScaleType.LINEAR

    def test_set_scale_x_axis(self):
        """Test set_scale for x axis."""
        chart = Chart()
        chart.set_scale("x", min_val=1, max_val=10, scale_type=ScaleType.LOGARITHMIC)
        assert chart.x_axis.min_value == 1
        assert chart.x_axis.max_value == 10
        assert chart.x_axis.scale_type == ScaleType.LOGARITHMIC

    def test_reset(self):
        """Test reset method."""
        chart = Chart(chart_type=ChartType.PIE)
        chart.add_series("Data", values=[1, 2, 3])
        chart.set_title("Test")
        chart.reset()
        assert chart.chart_type == ChartType.LINE
        assert len(chart.series) == 0
        assert chart.options.title == ""

    def test_to_dict(self):
        """Test to_dict serialization."""
        chart = Chart(chart_type=ChartType.BAR)
        chart.add_series("Sales", values=[10, 20, 30])
        chart.set_title("Revenue")
        chart.x_range = "A1:A3"

        data = chart.to_dict()
        assert data["chart_type"] == "BAR"
        assert data["x_range"] == "A1:A3"
        assert len(data["series"]) == 1
        assert data["series"][0]["name"] == "Sales"
        assert data["options"]["title"] == "Revenue"

    def test_from_dict(self):
        """Test from_dict deserialization."""
        data = {
            "chart_type": "PIE",
            "x_range": "A1:A5",
            "series": [
                {
                    "name": "Data",
                    "data_range": "B1:B5",
                    "values": [10, 20, 30, 40, 50],
                    "line_style": "SYMBOLS",
                    "color": "blue"
                }
            ],
            "x_axis": {
                "title": "X Axis",
                "min_value": 0,
                "max_value": 100,
                "scale_type": "LOGARITHMIC"
            },
            "y_axis": {
                "title": "Y Axis",
                "min_value": 0,
                "max_value": 1000
            },
            "options": {
                "title": "My Chart",
                "subtitle": "Subtitle",
                "show_legend": False,
                "legend_position": "right",
                "grid_horizontal": True,
                "grid_vertical": True
            }
        }

        chart = Chart.from_dict(data)
        assert chart.chart_type == ChartType.PIE
        assert chart.x_range == "A1:A5"
        assert len(chart.series) == 1
        assert chart.series[0].name == "Data"
        assert chart.series[0].line_style == LineStyle.SYMBOLS
        assert chart.x_axis.title == "X Axis"
        assert chart.x_axis.scale_type == ScaleType.LOGARITHMIC
        assert chart.y_axis.title == "Y Axis"
        assert chart.options.title == "My Chart"
        assert chart.options.show_legend is False
        assert chart.options.grid_horizontal is True

    def test_from_dict_defaults(self):
        """Test from_dict with minimal data uses defaults."""
        chart = Chart.from_dict({})
        assert chart.chart_type == ChartType.LINE
        assert len(chart.series) == 0

    def test_roundtrip_serialization(self):
        """Test that to_dict/from_dict roundtrip works."""
        original = Chart(chart_type=ChartType.STACKED_BAR)
        original.add_series("Series 1", values=[1, 2, 3])
        original.add_series("Series 2", values=[4, 5, 6])
        original.set_title("Test Chart", "Subtitle")
        original.set_axis_titles("X", "Y")
        original.set_scale("y", 0, 100)

        data = original.to_dict()
        restored = Chart.from_dict(data)

        assert restored.chart_type == original.chart_type
        assert len(restored.series) == len(original.series)
        assert restored.options.title == original.options.title


class TestChartSeriesAttributes:
    """Tests for ChartSeries attributes."""

    def test_default_symbol(self):
        """Test default symbol."""
        series = ChartSeries()
        assert series.symbol == "*"

    def test_custom_symbol(self):
        """Test custom symbol."""
        series = ChartSeries(symbol="o")
        assert series.symbol == "o"

    def test_default_color(self):
        """Test default color."""
        series = ChartSeries()
        assert series.color == "white"


class TestChartAxisAttributes:
    """Tests for ChartAxis attributes."""

    def test_grid_lines_default(self):
        """Test grid_lines default."""
        axis = ChartAxis()
        assert axis.grid_lines is False

    def test_grid_lines_enabled(self):
        """Test grid_lines enabled."""
        axis = ChartAxis(grid_lines=True)
        assert axis.grid_lines is True

    def test_labels(self):
        """Test custom labels."""
        axis = ChartAxis(labels=["Jan", "Feb", "Mar"])
        assert axis.labels == ["Jan", "Feb", "Mar"]


class TestChartOptionsAttributes:
    """Tests for ChartOptions attributes."""

    def test_subtitle(self):
        """Test subtitle."""
        opts = ChartOptions(subtitle="Q4 Results")
        assert opts.subtitle == "Q4 Results"

    def test_grid_options(self):
        """Test grid options."""
        opts = ChartOptions(grid_horizontal=True, grid_vertical=True)
        assert opts.grid_horizontal is True
        assert opts.grid_vertical is True

    def test_color_mode(self):
        """Test color mode."""
        opts = ChartOptions(color_mode="bw")
        assert opts.color_mode == "bw"

    def test_data_labels(self):
        """Test data labels option."""
        opts = ChartOptions(data_labels=True)
        assert opts.data_labels is True


class TestRendererEdgeCases:
    """Tests for renderer edge cases."""

    def test_render_empty_chart(self):
        """Test rendering chart with no data."""
        chart = Chart(chart_type=ChartType.LINE)
        ctx = RenderContext(chart=chart, width=40, height=10)
        renderer = LineChartRenderer()
        result = renderer.render(ctx)
        assert isinstance(result, list)

    def test_render_single_value(self):
        """Test rendering chart with single value."""
        chart = Chart(chart_type=ChartType.BAR)
        chart.series.append(ChartSeries(name="Data", values=[50]))
        ctx = RenderContext(chart=chart, width=40, height=10)
        renderer = BarChartRenderer()
        result = renderer.render(ctx)
        assert isinstance(result, list)

    def test_render_negative_values(self):
        """Test rendering chart with negative values."""
        chart = Chart(chart_type=ChartType.LINE)
        chart.series.append(ChartSeries(name="Data", values=[-10, 0, 10, 20]))
        ctx = RenderContext(chart=chart, width=40, height=10)
        renderer = LineChartRenderer()
        result = renderer.render(ctx)
        assert isinstance(result, list)

    def test_render_multiple_series(self):
        """Test rendering chart with multiple series."""
        chart = Chart(chart_type=ChartType.LINE)
        chart.series.append(ChartSeries(name="Series 1", values=[10, 20, 30]))
        chart.series.append(ChartSeries(name="Series 2", values=[15, 25, 35]))
        ctx = RenderContext(chart=chart, width=40, height=10)
        renderer = LineChartRenderer()
        result = renderer.render(ctx)
        assert isinstance(result, list)

    def test_render_small_dimensions(self):
        """Test rendering with small dimensions."""
        chart = Chart(chart_type=ChartType.BAR)
        chart.series.append(ChartSeries(name="Data", values=[10, 20, 30]))
        ctx = RenderContext(chart=chart, width=10, height=5)
        renderer = BarChartRenderer()
        result = renderer.render(ctx)
        assert isinstance(result, list)

    def test_render_stacked_bar(self):
        """Test rendering stacked bar chart."""
        chart = Chart(chart_type=ChartType.STACKED_BAR)
        chart.series.append(ChartSeries(name="A", values=[10, 20, 30]))
        chart.series.append(ChartSeries(name="B", values=[5, 10, 15]))
        ctx = RenderContext(chart=chart, width=40, height=10)
        renderer = BarChartRenderer()
        result = renderer.render(ctx)
        assert isinstance(result, list)

    def test_pie_with_zero_values(self):
        """Test pie chart with zero values."""
        chart = Chart(chart_type=ChartType.PIE)
        chart.series.append(ChartSeries(name="Data", values=[0, 0, 100]))
        ctx = RenderContext(chart=chart, width=40, height=10)
        renderer = PieChartRenderer()
        result = renderer.render(ctx)
        assert isinstance(result, list)

    def test_scatter_with_x_values(self):
        """Test scatter chart with x values."""
        chart = Chart(chart_type=ChartType.XY_SCATTER)
        chart.series.append(ChartSeries(name="XY", values=[1, 4, 9, 16, 25]))
        chart.x_range = "test"
        ctx = RenderContext(chart=chart, width=40, height=10)
        renderer = ScatterChartRenderer()
        result = renderer.render(ctx)
        assert isinstance(result, list)

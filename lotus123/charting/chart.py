"""Chart data model and configuration.

Supports Lotus 1-2-3 style charts:
- Line graphs
- Bar charts (vertical)
- Stacked bar charts
- XY scatter plots
- Pie charts
"""

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Any


class ChartType(Enum):
    """Types of charts available."""

    LINE = auto()
    BAR = auto()
    STACKED_BAR = auto()
    XY_SCATTER = auto()
    PIE = auto()
    AREA = auto()
    HLBAR = auto()  # Horizontal bar


class LineStyle(Enum):
    """Line/symbol styles for data series."""

    LINES = auto()  # Lines only
    SYMBOLS = auto()  # Symbols only
    BOTH = auto()  # Lines and symbols
    NEITHER = auto()  # No display (for hidden series)


class ScaleType(Enum):
    """Scale type for axes."""

    LINEAR = auto()
    LOGARITHMIC = auto()


@dataclass
class ChartSeries:
    """A data series in a chart.

    Attributes:
        name: Series name/legend label
        data_range: Range reference for data (e.g., 'B1:B10')
        line_style: How to display the series
        color: Color name or index
    """

    name: str = ""
    data_range: str = ""
    values: list[float] = field(default_factory=list)
    line_style: LineStyle = LineStyle.BOTH
    color: str = "white"
    symbol: str = "*"


@dataclass
class ChartAxis:
    """Configuration for a chart axis.

    Attributes:
        title: Axis title
        min_value: Minimum value (None for auto)
        max_value: Maximum value (None for auto)
        scale_type: Linear or logarithmic
        grid_lines: Show grid lines
        labels: Custom labels for categories
    """

    title: str = ""
    min_value: float | None = None
    max_value: float | None = None
    scale_type: ScaleType = ScaleType.LINEAR
    grid_lines: bool = False
    labels: list[str] = field(default_factory=list)


@dataclass
class ChartOptions:
    """Chart display options.

    Attributes:
        title: Main chart title
        subtitle: Second line of title
        show_legend: Display legend
        legend_position: 'bottom', 'right', 'top', 'left'
        grid_horizontal: Show horizontal grid lines
        grid_vertical: Show vertical grid lines
        color_mode: 'color' or 'bw' (black/white)
        data_labels: Show values on data points
    """

    title: str = ""
    subtitle: str = ""
    show_legend: bool = True
    legend_position: str = "bottom"
    grid_horizontal: bool = False
    grid_vertical: bool = False
    color_mode: str = "color"
    data_labels: bool = False


@dataclass
class Chart:
    """A chart/graph configuration.

    Lotus 1-2-3 charts support up to 6 data ranges (A through F).
    """

    chart_type: ChartType = ChartType.LINE
    x_range: str = ""  # X-axis values/labels
    series: list[ChartSeries] = field(default_factory=list)
    x_axis: ChartAxis = field(default_factory=ChartAxis)
    y_axis: ChartAxis = field(default_factory=ChartAxis)
    options: ChartOptions = field(default_factory=ChartOptions)

    def add_series(
        self, name: str, data_range: str = "", values: list[float] | None = None
    ) -> ChartSeries:
        """Add a data series to the chart.

        Args:
            name: Series name
            data_range: Cell range for data
            values: Direct values (alternative to range)

        Returns:
            The created series
        """
        series = ChartSeries(name=name, data_range=data_range, values=values or [])
        self.series.append(series)
        return series

    def set_x_range(self, range_ref: str) -> None:
        """Set the X-axis data range."""
        self.x_range = range_ref

    def set_type(self, chart_type: ChartType) -> None:
        """Set the chart type."""
        self.chart_type = chart_type

    def set_title(self, title: str, subtitle: str = "") -> None:
        """Set chart titles."""
        self.options.title = title
        self.options.subtitle = subtitle

    def set_axis_titles(self, x_title: str = "", y_title: str = "") -> None:
        """Set axis titles."""
        self.x_axis.title = x_title
        self.y_axis.title = y_title

    def set_scale(
        self,
        axis: str = "y",
        min_val: float | None = None,
        max_val: float | None = None,
        scale_type: ScaleType = ScaleType.LINEAR,
    ) -> None:
        """Set axis scale options.

        Args:
            axis: 'x' or 'y'
            min_val: Minimum value
            max_val: Maximum value
            scale_type: Linear or logarithmic
        """
        ax = self.y_axis if axis.lower() == "y" else self.x_axis
        if min_val is not None:
            ax.min_value = min_val
        if max_val is not None:
            ax.max_value = max_val
        ax.scale_type = scale_type

    def reset(self) -> None:
        """Reset chart to default settings."""
        self.chart_type = ChartType.LINE
        self.x_range = ""
        self.series.clear()
        self.x_axis = ChartAxis()
        self.y_axis = ChartAxis()
        self.options = ChartOptions()

    def to_dict(self) -> dict[str, Any]:
        """Serialize chart to dictionary."""
        return {
            "chart_type": self.chart_type.name,
            "x_range": self.x_range,
            "series": [
                {
                    "name": s.name,
                    "data_range": s.data_range,
                    "values": s.values,
                    "line_style": s.line_style.name,
                    "color": s.color,
                }
                for s in self.series
            ],
            "x_axis": {
                "title": self.x_axis.title,
                "min_value": self.x_axis.min_value,
                "max_value": self.x_axis.max_value,
                "scale_type": self.x_axis.scale_type.name,
                "grid_lines": self.x_axis.grid_lines,
            },
            "y_axis": {
                "title": self.y_axis.title,
                "min_value": self.y_axis.min_value,
                "max_value": self.y_axis.max_value,
                "scale_type": self.y_axis.scale_type.name,
                "grid_lines": self.y_axis.grid_lines,
            },
            "options": {
                "title": self.options.title,
                "subtitle": self.options.subtitle,
                "show_legend": self.options.show_legend,
                "legend_position": self.options.legend_position,
                "grid_horizontal": self.options.grid_horizontal,
                "grid_vertical": self.options.grid_vertical,
            },
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> Chart:
        """Deserialize chart from dictionary."""
        chart = cls()
        chart.chart_type = ChartType[data.get("chart_type", "LINE")]
        chart.x_range = data.get("x_range", "")

        for s_data in data.get("series", []):
            series = ChartSeries(
                name=s_data.get("name", ""),
                data_range=s_data.get("data_range", ""),
                values=s_data.get("values", []),
                line_style=LineStyle[s_data.get("line_style", "BOTH")],
                color=s_data.get("color", "white"),
            )
            chart.series.append(series)

        if "x_axis" in data:
            ax_data = data["x_axis"]
            chart.x_axis.title = ax_data.get("title", "")
            chart.x_axis.min_value = ax_data.get("min_value")
            chart.x_axis.max_value = ax_data.get("max_value")
            if "scale_type" in ax_data:
                chart.x_axis.scale_type = ScaleType[ax_data["scale_type"]]

        if "y_axis" in data:
            ax_data = data["y_axis"]
            chart.y_axis.title = ax_data.get("title", "")
            chart.y_axis.min_value = ax_data.get("min_value")
            chart.y_axis.max_value = ax_data.get("max_value")
            if "scale_type" in ax_data:
                chart.y_axis.scale_type = ScaleType[ax_data["scale_type"]]

        if "options" in data:
            opt_data = data["options"]
            chart.options.title = opt_data.get("title", "")
            chart.options.subtitle = opt_data.get("subtitle", "")
            chart.options.show_legend = opt_data.get("show_legend", True)
            chart.options.legend_position = opt_data.get("legend_position", "bottom")
            chart.options.grid_horizontal = opt_data.get("grid_horizontal", False)
            chart.options.grid_vertical = opt_data.get("grid_vertical", False)

        return chart

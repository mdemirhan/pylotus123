"""Chart-related handler methods for LotusApp."""

import json
from typing import Callable

from ..charting import Chart, ChartType, TextChartRenderer
from ..core import make_cell_ref
from ..ui import ChartViewScreen, CommandInput, FileDialog
from .base import AppProtocol, BaseHandler


class ChartHandler(BaseHandler):
    """Handler for chart-related operations."""

    def __init__(self, app: AppProtocol) -> None:
        super().__init__(app)
        # Chart renderer - owned by this handler
        self.chart_renderer = TextChartRenderer(self.spreadsheet)

    @property
    def chart(self):
        """Access the chart object."""
        return self._app.chart

    def set_chart_type(self, chart_type: ChartType) -> None:
        """Set the chart type."""
        self.chart.set_type(chart_type)
        type_names = {
            ChartType.LINE: "Line",
            ChartType.BAR: "Bar",
            ChartType.XY_SCATTER: "XY Scatter",
            ChartType.STACKED_BAR: "Stacked Bar",
            ChartType.PIE: "Pie",
        }
        self.notify(f"Chart type set to {type_names.get(chart_type, 'Unknown')}")

    def set_x_range(self) -> None:
        """Set the X-axis range for the chart."""

        def apply_range(range_str: str) -> None:
            self.chart.set_x_range(range_str)
            self.notify(f"X-Range set to {range_str}")

        self._set_range_from_selection("X-Range (e.g., A1:A10):", apply_range)

    def set_a_range(self) -> None:
        """Set the A data range."""
        self._set_series_range(0, "A", "A-Range (e.g., B1:B10):")

    def set_b_range(self) -> None:
        """Set the B data range."""
        self._set_series_range(1, "B", "B-Range (e.g., C1:C10):")

    def set_c_range(self) -> None:
        """Set the C data range."""
        self._set_series_range(2, "C", "C-Range (e.g., D1:D10):")

    def set_d_range(self) -> None:
        """Set the D data range."""
        self._set_series_range(3, "D", "D-Range (e.g., E1:E10):")

    def set_e_range(self) -> None:
        """Set the E data range."""
        self._set_series_range(4, "E", "E-Range (e.g., F1:F10):")

    def set_f_range(self) -> None:
        """Set the F data range."""
        self._set_series_range(5, "F", "F-Range (e.g., G1:G10):")

    def _set_series_range(self, index: int, name: str, prompt: str) -> None:
        """Set a series range by index."""
        self._set_range_from_selection(
            prompt, lambda range_str: self._add_or_update_series(index, name, range_str)
        )

    def _set_range_from_selection(self, prompt: str, on_set: Callable[[str], None]) -> None:
        """Handle range selection or prompt input."""
        grid = self.get_grid()
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            on_set(range_str)
        else:
            self._app.push_screen(
                CommandInput(prompt),
                lambda result: self._do_set_range(result, on_set),
            )

    def _do_set_range(self, result: str | None, on_set: Callable[[str], None]) -> None:
        """Handle prompted range input."""
        if result:
            on_set(result.upper())

    def _add_or_update_series(self, index: int, name: str, range_str: str) -> None:
        """Add or update a data series at the given index."""
        while len(self.chart.series) <= index:
            self.chart.add_series(f"Series {len(self.chart.series) + 1}")
        self.chart.series[index].name = name
        self.chart.series[index].data_range = range_str
        self.notify(f"{name}-Range set to {range_str}")

    def view_chart(self) -> None:
        """Display the chart in a modal screen."""
        if not self.chart.series:
            self.notify("No data series defined. Use A-Range to set data.")
            return
        self.chart_renderer.spreadsheet = self.spreadsheet
        # Use ~75% of terminal size for the chart
        chart_width = int(self._app.size.width * 0.75)
        chart_height = int(self._app.size.height * 0.70)
        chart_lines = self.chart_renderer.render(self.chart, width=chart_width, height=chart_height)
        self._app.push_screen(ChartViewScreen(chart_lines))

    def reset_chart(self) -> None:
        """Reset the chart to default state."""
        self.chart.reset()
        self.notify("Chart reset")

    def save_chart(self) -> None:
        """Save the current chart configuration to a file."""
        if not self.chart.series:
            self.notify("No chart data to save. Define data ranges first.")
            return
        self._app.push_screen(
            FileDialog(mode="save", file_extensions=[".chart"]),
            self._do_save_chart,
        )

    def _do_save_chart(self, result: str | None) -> None:
        """Handle the save chart dialog result."""
        if result:
            if not result.endswith(".chart"):
                result += ".chart"
            self.confirm_overwrite(result, self._perform_chart_save)

    def _perform_chart_save(self, filepath: str) -> None:
        """Actually save the chart to the file."""
        try:
            chart_data = self.chart.to_dict()
            with open(filepath, "w") as f:
                json.dump(chart_data, f, indent=2)
            self.notify(f"Chart saved: {filepath}")
        except Exception as e:
            self.notify(f"Error saving chart: {e}", severity="error")

    def load_chart(self) -> None:
        """Load a chart configuration from a file."""
        self._app.push_screen(
            FileDialog(mode="open", file_extensions=[".chart"]),
            self._do_load_chart,
        )

    def _do_load_chart(self, result: str | None) -> None:
        """Handle the load chart dialog result."""
        if result:
            try:
                with open(result) as f:
                    chart_data = json.load(f)
                loaded_chart = Chart.from_dict(chart_data)
                # Copy loaded chart properties to the app's chart
                self.chart.chart_type = loaded_chart.chart_type
                self.chart.x_range = loaded_chart.x_range
                self.chart.series = loaded_chart.series
                self.chart.x_axis = loaded_chart.x_axis
                self.chart.y_axis = loaded_chart.y_axis
                self.chart.options = loaded_chart.options
                self.notify(f"Chart loaded: {result}")
            except Exception as e:
                self.notify(f"Error loading chart: {e}", severity="error")

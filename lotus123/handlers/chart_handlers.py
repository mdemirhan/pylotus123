"""Chart-related handler methods for LotusApp."""

from __future__ import annotations

import json
from pathlib import Path
from typing import TYPE_CHECKING

from ..charting import Chart, ChartType
from ..core import make_cell_ref
from ..ui import ChartViewScreen, CommandInput, FileDialog
from .base import BaseHandler

if TYPE_CHECKING:
    from .base import AppProtocol


class ChartHandler(BaseHandler):
    """Handler for chart-related operations."""

    def __init__(self, app: "AppProtocol") -> None:
        super().__init__(app)

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
        grid = self.get_grid()
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self.chart.set_x_range(range_str)
            self.notify(f"X-Range set to {range_str}")
        else:
            self._app.push_screen(
                CommandInput("X-Range (e.g., A1:A10):"), self._do_set_x_range
            )

    def _do_set_x_range(self, result: str | None) -> None:
        if result:
            self.chart.set_x_range(result.upper())
            self.notify(f"X-Range set to {result.upper()}")

    def set_a_range(self) -> None:
        """Set the A data range."""
        grid = self.get_grid()
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self._add_or_update_series(0, "A", range_str)
        else:
            self._app.push_screen(
                CommandInput("A-Range (e.g., B1:B10):"), self._do_set_a_range
            )

    def _do_set_a_range(self, result: str | None) -> None:
        if result:
            self._add_or_update_series(0, "A", result.upper())

    def set_b_range(self) -> None:
        """Set the B data range."""
        grid = self.get_grid()
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self._add_or_update_series(1, "B", range_str)
        else:
            self._app.push_screen(
                CommandInput("B-Range (e.g., C1:C10):"), self._do_set_b_range
            )

    def _do_set_b_range(self, result: str | None) -> None:
        if result:
            self._add_or_update_series(1, "B", result.upper())

    def set_c_range(self) -> None:
        """Set the C data range."""
        grid = self.get_grid()
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self._add_or_update_series(2, "C", range_str)
        else:
            self._app.push_screen(
                CommandInput("C-Range (e.g., D1:D10):"), self._do_set_c_range
            )

    def _do_set_c_range(self, result: str | None) -> None:
        if result:
            self._add_or_update_series(2, "C", result.upper())

    def set_d_range(self) -> None:
        """Set the D data range."""
        grid = self.get_grid()
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self._add_or_update_series(3, "D", range_str)
        else:
            self._app.push_screen(
                CommandInput("D-Range (e.g., E1:E10):"), self._do_set_d_range
            )

    def _do_set_d_range(self, result: str | None) -> None:
        if result:
            self._add_or_update_series(3, "D", result.upper())

    def set_e_range(self) -> None:
        """Set the E data range."""
        grid = self.get_grid()
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self._add_or_update_series(4, "E", range_str)
        else:
            self._app.push_screen(
                CommandInput("E-Range (e.g., F1:F10):"), self._do_set_e_range
            )

    def _do_set_e_range(self, result: str | None) -> None:
        if result:
            self._add_or_update_series(4, "E", result.upper())

    def set_f_range(self) -> None:
        """Set the F data range."""
        grid = self.get_grid()
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self._add_or_update_series(5, "F", range_str)
        else:
            self._app.push_screen(
                CommandInput("F-Range (e.g., G1:G10):"), self._do_set_f_range
            )

    def _do_set_f_range(self, result: str | None) -> None:
        if result:
            self._add_or_update_series(5, "F", result.upper())

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
        self._app._chart_renderer.spreadsheet = self.spreadsheet
        # Use ~75% of terminal size for the chart
        chart_width = int(self._app.size.width * 0.75)
        chart_height = int(self._app.size.height * 0.70)
        chart_lines = self._app._chart_renderer.render(
            self.chart, width=chart_width, height=chart_height
        )
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
        self._app.push_screen(FileDialog(mode="save"), self._do_save_chart)

    def _do_save_chart(self, result: str | None) -> None:
        """Handle the save chart dialog result."""
        if result:
            if not result.endswith(".chart"):
                result += ".chart"
            if Path(result).exists():
                self._pending_chart_path = result
                self._app.push_screen(
                    CommandInput(f"File '{result}' exists. Overwrite? (Y/N):"),
                    self._do_save_chart_confirm,
                )
            else:
                self._perform_chart_save(result)

    def _do_save_chart_confirm(self, result: str | None) -> None:
        """Handle overwrite confirmation."""
        if result and result.strip().upper().startswith("Y"):
            self._perform_chart_save(self._pending_chart_path)
        else:
            self.notify("Save cancelled", severity="warning")

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
        self._app.push_screen(FileDialog(mode="open"), self._do_load_chart)

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

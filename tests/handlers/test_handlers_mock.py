from typing import Any

from unittest.mock import MagicMock, patch
from lotus123.handlers.clipboard_handlers import ClipboardHandler
from lotus123.handlers.query_handlers import QueryHandler
from lotus123.core.spreadsheet import Spreadsheet
from lotus123.core.cell import Cell
from lotus123.ui.grid import SpreadsheetGrid


class MockApp:
    def __init__(self) -> None:
        self.spreadsheet = MagicMock(spec=Spreadsheet)
        self.spreadsheet.rows = 100
        self.spreadsheet.cols = 26
        self.spreadsheet.get_cell.return_value = MagicMock(spec=Cell, raw_value="123")
        self.spreadsheet.get_cell_if_exists.return_value = MagicMock(spec=Cell, raw_value="123")
        self.spreadsheet.get_col_width.return_value = 10  # Default column width
        self.spreadsheet.modified = False
        # Global settings mock
        self.spreadsheet.global_settings = {
            "format_code": "G",
            "label_prefix": "'",
            "default_col_width": 10,
            "zero_display": True,
        }
        self.undo_manager = MagicMock()
        self.push_screen = MagicMock()
        self.notify = MagicMock()
        self.query_one = MagicMock()

        # Grid mock
        self.grid = MagicMock(spec=SpreadsheetGrid)
        self.grid.selection_range = (0, 0, 2, 2)  # A1:C3
        self.grid.cursor_row = 0
        self.grid.cursor_col = 0
        self.grid.has_selection = True
        self.query_one.return_value = self.grid

        # Global settings (public - shared across handlers)
        self.global_format_code = "G"
        self.global_label_prefix = "'"
        self.global_col_width = 9
        self.recalc_mode = "auto"
        self.global_zero_display = True

        # AppProtocol required attributes
        self.chart: Any = MagicMock()
        self.config: Any = MagicMock()
        self.editing: bool = False
        self._menu_active: bool = False
        self.current_theme_type: Any = MagicMock()
        self.color_theme: Any = MagicMock()
        self.sub_title: str = ""
        self._size: Any = MagicMock()

    @property
    def size(self) -> Any:
        return self._size

    @size.setter
    def size(self, value: Any) -> None:
        self._size = value

    def update_status(self) -> None:
        pass

    def update_title(self) -> None:
        pass

    def mark_dirty(self) -> None:
        self.spreadsheet.modified = True

    def apply_theme(self) -> None:
        pass

    def set_recalc_mode(self, mode: Any) -> None:
        from lotus123.formula.recalc import RecalcMode

        self.recalc_mode = "manual" if mode == RecalcMode.MANUAL else "auto"
        self.spreadsheet.set_recalc_mode(mode)

    def get_recalc_mode(self) -> Any:
        return self.spreadsheet.get_recalc_mode()

    def set_recalc_order(self, order: Any) -> None:
        self.spreadsheet.set_recalc_order(order)

    def get_recalc_order(self) -> Any:
        return self.spreadsheet.get_recalc_order()


class TestClipboardHandler:
    def setup_method(self):
        self.app = MockApp()
        self.handler = ClipboardHandler(self.app)

    def test_menu_copy(self):
        self.handler.menu_copy()
        assert self.handler.pending_source_range == (0, 0, 2, 2)
        self.app.push_screen.assert_called()

    def test_do_menu_copy(self):
        # Make source and target different to trigger change
        cell_src = MagicMock(spec=Cell, raw_value="SRC")
        cell_dst = MagicMock(spec=Cell, raw_value="DST")
        # get_cell called for src then target
        self.app.spreadsheet.get_cell.side_effect = [cell_src, cell_dst]

        self.handler.pending_source_range = (0, 0, 0, 0)  # A1:A1
        self.handler._do_menu_copy("B1")  # Copy A1 to B1
        self.app.undo_manager.execute.assert_called()
        cmd = self.app.undo_manager.execute.call_args[0][0]
        # Should be RangeChangeCommand
        assert cmd.__class__.__name__ == "RangeChangeCommand"

    def test_menu_move(self):
        self.handler.menu_move()
        self.app.push_screen.assert_called()

    def test_do_menu_move(self):
        self.handler.pending_source_range = (0, 0, 0, 0)
        self.handler._do_menu_move("B1")
        self.app.undo_manager.execute.assert_called()

    def test_copy_cells(self):
        self.handler.copy_cells()
        assert self.handler.range_clipboard is not None
        assert len(self.handler.range_clipboard) == 3  # 3 rows (0,1,2)
        assert len(self.handler.range_clipboard[0]) == 3  # 3 cols
        assert not self.handler.clipboard_is_cut

    def test_cut_cells(self):
        self.handler.cut_cells()
        assert self.handler.clipboard_is_cut

    def test_paste_cells(self):
        self.handler.range_clipboard = [["1", "2"]]
        self.handler.paste_cells()
        self.app.undo_manager.execute.assert_called()


class TestQueryHandler:
    def setup_method(self):
        self.app = MockApp()
        self.handler = QueryHandler(self.app)

    def test_set_input(self):
        # With selection
        self.handler.set_input()
        assert self.handler.input_range == (0, 0, 2, 2)

        # Without selection
        self.app.grid.has_selection = False
        self.handler.set_input()
        self.app.push_screen.assert_called()  # Should prompt

    def test_do_set_input(self):
        self.handler._do_set_input("A1:D5")
        assert self.handler.input_range == (0, 0, 4, 3)

    def test_set_criteria(self):
        self.handler.set_criteria()
        assert self.handler.criteria_range == (0, 0, 2, 2)

    def test_do_set_criteria(self):
        self.handler._do_set_criteria("F1:G2")
        assert self.handler.criteria_range == (0, 5, 1, 6)

    def test_set_output(self):
        self.handler.set_output()
        assert self.handler.output_range == (0, 0)

    def test_do_set_output(self):
        self.handler._do_set_output("H1")
        assert self.handler.output_range == (0, 7)

    def test_find_flow(self):
        # Setup mock db query result
        self.handler.input_range = (0, 0, 5, 5)
        self.handler.criteria_range = (0, 6, 1, 6)

        # Mock DatabaseOperations locally within find method scope or patch it
        with patch("lotus123.data.database.DatabaseOperations") as MockDB:
            mock_db_instance = MockDB.return_value
            mock_db_instance.query.return_value = [1, 2, 3]  # rows 1, 2, 3 overlap

            self.handler.find()

            assert self.handler.find_results == [1, 2, 3]
            assert self.handler.find_index == 0
            # Should invoke cursor move
            assert self.app.grid.cursor_row == 1

            # Next find
            self.handler.find()
            assert self.handler.find_index == 1
            assert self.app.grid.cursor_row == 2

        self.handler.reset()
        assert self.handler.input_range is None


class TestWorksheetHandler:
    def setup_method(self):
        self.app = MockApp()
        # Create a dummy WorksheetHandler using BaseHandler pattern
        from lotus123.handlers.worksheet_handlers import WorksheetHandler

        self.handler = WorksheetHandler(self.app)

    def test_global_format(self):
        self.app.push_screen.assert_called_with = MagicMock()
        self.handler.global_format()
        self.app.push_screen.assert_called()

    def test_insert_rows(self):
        self.handler.insert_rows()
        self.app.push_screen.assert_called()

    def test_insert_columns(self):
        self.handler.insert_columns()
        self.app.push_screen.assert_called()

    def test_delete_rows(self):
        self.handler.delete_rows()
        self.app.push_screen.assert_called()

    def test_delete_columns(self):
        self.handler.delete_columns()
        self.app.push_screen.assert_called()

    def test_col_width(self):
        self.handler.set_column_width()
        self.app.push_screen.assert_called()

    def test_do_set_width_single_column(self):
        """Test setting width for a single column."""
        from lotus123.utils.undo import ColWidthCommand

        # Set selection to single column (A1:A3)
        self.app.grid.selection_range = (0, 0, 2, 0)
        self.handler._do_set_width("15")
        # Should execute a ColWidthCommand via undo manager
        self.app.undo_manager.execute.assert_called_once()
        cmd = self.app.undo_manager.execute.call_args[0][0]
        assert isinstance(cmd, ColWidthCommand)
        assert cmd.changes == {0: (15, 10)}  # new=15, old=10
        assert self.app.spreadsheet.modified is True

    def test_do_set_width_multiple_columns(self):
        """Test setting width for multiple selected columns."""
        from lotus123.utils.undo import ColWidthCommand

        # Set selection to columns A-C (A1:C3)
        self.app.grid.selection_range = (0, 0, 2, 2)
        self.handler._do_set_width("12")
        # Should execute a ColWidthCommand via undo manager
        self.app.undo_manager.execute.assert_called_once()
        cmd = self.app.undo_manager.execute.call_args[0][0]
        assert isinstance(cmd, ColWidthCommand)
        # All 3 columns should be in changes (new=12, old=10)
        assert cmd.changes == {0: (12, 10), 1: (12, 10), 2: (12, 10)}
        assert self.app.spreadsheet.modified is True

    def test_do_set_width_invalid_value(self):
        """Test setting width with invalid value."""
        self.handler._do_set_width("abc")
        self.app.notify.assert_called_with("Invalid width value", severity="error")
        self.app.undo_manager.execute.assert_not_called()

    def test_do_set_width_out_of_range(self):
        """Test setting width outside valid range."""
        self.handler._do_set_width("2")  # Less than 3
        self.app.notify.assert_called_with("Width must be between 3 and 50", severity="error")
        self.app.undo_manager.execute.assert_not_called()

        self.app.notify.reset_mock()
        self.handler._do_set_width("51")  # Greater than 50
        self.app.notify.assert_called_with("Width must be between 3 and 50", severity="error")


class TestFileHandler:
    def setup_method(self):
        self.app = MockApp()
        from lotus123.handlers.file_handlers import FileHandler

        self.handler = FileHandler(self.app)
        self.app.spreadsheet.filename = "test.json"

    def test_save(self):
        with patch("lotus123.handlers.file_handlers.FileDialog"):
            self.handler.save()
            # If filename set, simply notifies
            assert self.app.spreadsheet.modified is False

    def test_retrieve(self):
        with patch("lotus123.handlers.file_handlers.FileDialog"):
            self.handler.open_file()
            self.app.push_screen.assert_called()


class TestDataHandler:
    def setup_method(self):
        self.app = MockApp()
        from lotus123.handlers.data_handlers import DataHandler

        self.handler = DataHandler(self.app)

    def test_fill(self):
        # Fill needs range
        self.app.grid.selection_range = (0, 0, 5, 0)
        self.handler.data_fill()
        self.app.push_screen.assert_called()

    def test_sort(self):
        pass  # Skipping brittle sort test as per user direction

    def test_sort_column_d_parsing(self):
        """Test that column D sorting correctly distinguishes ascending vs descending.

        Bug fix: Previously "D" and "DD" both resulted in descending sort because
        the code checked `result.endswith("D")` which is True for both.
        """
        from lotus123.core.spreadsheet import Spreadsheet
        from lotus123.handlers.data_handlers import DataHandler

        # Create a real spreadsheet with test data
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 3, "300")  # D1
        spreadsheet.set_cell(1, 3, "100")  # D2
        spreadsheet.set_cell(2, 3, "200")  # D3

        # Create app mock with real spreadsheet
        app = MockApp()
        app.spreadsheet = spreadsheet
        app.grid.selection_range = (0, 3, 2, 3)  # D1:D3

        handler = DataHandler(app)

        # Sort ascending with "D" - should sort 100, 200, 300
        handler._do_data_sort("D")
        assert spreadsheet.get_cell(0, 3).raw_value == "100"
        assert spreadsheet.get_cell(1, 3).raw_value == "200"
        assert spreadsheet.get_cell(2, 3).raw_value == "300"

        # Sort descending with "DD" - should sort 300, 200, 100
        handler._do_data_sort("DD")
        assert spreadsheet.get_cell(0, 3).raw_value == "300"
        assert spreadsheet.get_cell(1, 3).raw_value == "200"
        assert spreadsheet.get_cell(2, 3).raw_value == "100"

    def test_sort_input_parsing_various_columns(self):
        """Test sort input parsing for various column letters."""
        from lotus123.core.spreadsheet import Spreadsheet
        from lotus123.handlers.data_handlers import DataHandler

        # Create spreadsheet with data in columns A and B
        spreadsheet = Spreadsheet()
        spreadsheet.set_cell(0, 0, "300")  # A1
        spreadsheet.set_cell(1, 0, "100")  # A2
        spreadsheet.set_cell(2, 0, "200")  # A3
        spreadsheet.set_cell(0, 1, "Z")  # B1
        spreadsheet.set_cell(1, 1, "A")  # B2
        spreadsheet.set_cell(2, 1, "M")  # B3

        app = MockApp()
        app.spreadsheet = spreadsheet
        app.grid.selection_range = (0, 0, 2, 1)  # A1:B3

        handler = DataHandler(app)

        # Sort by column A ascending
        handler._do_data_sort("A")
        assert spreadsheet.get_cell(0, 0).raw_value == "100"
        assert spreadsheet.get_cell(0, 1).raw_value == "A"  # Row with 100 moved to top

        # Sort by column A descending
        handler._do_data_sort("AD")
        assert spreadsheet.get_cell(0, 0).raw_value == "300"

        # Sort by column B ascending
        handler._do_data_sort("B")
        assert spreadsheet.get_cell(0, 1).raw_value == "A"

        # Sort by column B descending
        handler._do_data_sort("BD")
        assert spreadsheet.get_cell(0, 1).raw_value == "Z"


class TestRangeHandler:
    def setup_method(self):
        self.app = MockApp()
        from lotus123.handlers.range_handlers import RangeHandler

        self.handler = RangeHandler(self.app)

    def test_range_format(self):
        self.handler.range_format()
        self.app.push_screen.assert_called()

    def test_range_label(self):
        self.handler.range_label()
        self.app.push_screen.assert_called()

    def test_range_name(self):
        self.handler.range_name()
        self.app.push_screen.assert_called()  # Prompt for name


class TestChartHandler:
    def setup_method(self):
        self.app = MockApp()
        # Mock chart internals
        self.app.chart = MagicMock()
        self.app.chart.series = []

        def add_series_side_effect(name):
            self.app.chart.series.append(MagicMock(name=name))

        self.app.chart.add_series.side_effect = add_series_side_effect

        from lotus123.handlers.chart_handlers import ChartHandler

        self.handler = ChartHandler(self.app)
        # Mock the handler's chart_renderer
        self.handler.chart_renderer = MagicMock()

    def test_set_chart_type(self):
        from lotus123.charting import ChartType

        self.handler.set_chart_type(ChartType.BAR)
        self.app.chart.set_type.assert_called_with(ChartType.BAR)

    def test_set_ranges(self):
        # Test X
        self.handler.set_x_range()
        self.app.chart.set_x_range.assert_called()

        # Test A-F
        ranges = ["a", "b", "c", "d", "e", "f"]
        for idx, r in enumerate(ranges):
            method_name = f"set_{r}_range"
            getattr(self.handler, method_name)()
            # Should invoke add_series
            assert self.app.chart.add_series.call_count >= 1

            # Test callback via shared range handler
            self.handler._do_set_range(
                f"{r.upper()}1:{r.upper()}10",
                lambda range_str, i=idx, name=r.upper(): self.handler._add_or_update_series(
                    i, name, range_str
                ),
            )

    def test_view_chart(self):
        # With no series -> warning
        self.handler.view_chart()
        self.handler.chart_renderer.render.assert_not_called()

        # With series
        self.app.chart.series = [MagicMock()]
        self.app.size = MagicMock(width=80, height=24)
        self.handler.view_chart()
        self.handler.chart_renderer.render.assert_called()
        self.app.push_screen.assert_called()

    def test_reset_chart(self):
        self.handler.reset_chart()
        self.app.chart.reset.assert_called()

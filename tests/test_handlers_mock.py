
import pytest
from unittest.mock import MagicMock, patch
from lotus123.handlers.clipboard_handlers import ClipboardHandler
from lotus123.handlers.query_handlers import QueryHandler
from lotus123.core.spreadsheet import Spreadsheet
from lotus123.core.cell import Cell
from lotus123.ui.grid import SpreadsheetGrid

class MockApp:
    def __init__(self):
        self.spreadsheet = MagicMock(spec=Spreadsheet)
        self.spreadsheet.rows = 100
        self.spreadsheet.cols = 26
        self.spreadsheet.get_cell.return_value = MagicMock(spec=Cell, raw_value="123")
        self.spreadsheet.get_cell_if_exists.return_value = MagicMock(spec=Cell, raw_value="123")
        self.undo_manager = MagicMock()
        self.push_screen = MagicMock()
        self.notify = MagicMock()
        self.query_one = MagicMock()
        
        # Grid mock
        self.grid = MagicMock(spec=SpreadsheetGrid)
        self.grid.selection_range = (0, 0, 2, 2) # A1:C3
        self.grid.cursor_row = 0
        self.grid.cursor_col = 0
        self.grid.has_selection = True
        self.query_one.return_value = self.grid
        
        # App state
        self._pending_source_range = None
        self._range_clipboard = []
        self._cell_clipboard = None
        self._clipboard_is_cut = False
        self._clipboard_origin = (0, 0)
        
        # Query state
        self._query_input_range = None
        self._query_criteria_range = None
        self._query_output_range = None
        self._query_find_results = None
        self._query_find_index = 0

        # Global settings
        self._global_format_code = "G"
        self._global_label_prefix = "'"
        self._global_col_width = 9
        self._recalc_mode = "auto"
        self._global_protection = False
        self._global_zero_display = True
        self._dirty = False

    def _update_status(self): pass
    def _update_title(self): pass
    def _mark_dirty(self): self._dirty = True

class TestClipboardHandler:
    def setup_method(self):
        self.app = MockApp()
        self.handler = ClipboardHandler(self.app)

    def test_menu_copy(self):
        self.handler.menu_copy()
        assert self.app._pending_source_range == (0, 0, 2, 2)
        self.app.push_screen.assert_called()
        
    def test_do_menu_copy(self):
        # Make source and target different to trigger change
        cell_src = MagicMock(spec=Cell, raw_value="SRC")
        cell_dst = MagicMock(spec=Cell, raw_value="DST")
        # get_cell called for src then target
        self.app.spreadsheet.get_cell.side_effect = [cell_src, cell_dst]
        
        self.app._pending_source_range = (0, 0, 0, 0) # A1:A1
        self.handler._do_menu_copy("B1") # Copy A1 to B1
        self.app.undo_manager.execute.assert_called()
        cmd = self.app.undo_manager.execute.call_args[0][0]
        # Should be RangeChangeCommand
        assert cmd.__class__.__name__ == "RangeChangeCommand"

    def test_menu_move(self):
        self.handler.menu_move()
        self.app.push_screen.assert_called()

    def test_do_menu_move(self):
        self.app._pending_source_range = (0, 0, 0, 0)
        self.handler._do_menu_move("B1")
        self.app.undo_manager.execute.assert_called()
    
    def test_copy_cells(self):
        self.handler.copy_cells()
        assert len(self.app._range_clipboard) == 3 # 3 rows (0,1,2)
        assert len(self.app._range_clipboard[0]) == 3 # 3 cols
        assert not self.app._clipboard_is_cut
        
    def test_cut_cells(self):
        self.handler.cut_cells()
        assert self.app._clipboard_is_cut
        
    def test_paste_cells(self):
        self.app._range_clipboard = [["1", "2"]]
        self.handler.paste_cells()
        self.app.undo_manager.execute.assert_called()

class TestQueryHandler:
    def setup_method(self):
        self.app = MockApp()
        self.handler = QueryHandler(self.app)

    def test_set_input(self):
        # With selection
        self.handler.set_input()
        assert self.app._query_input_range == (0, 0, 2, 2)
        
        # Without selection
        self.app.grid.has_selection = False
        self.handler.set_input()
        self.app.push_screen.assert_called() # Should prompt
        
    def test_do_set_input(self):
        self.handler._do_set_input("A1:D5")
        assert self.app._query_input_range == (0, 0, 4, 3)

    def test_set_criteria(self):
        self.handler.set_criteria()
        assert self.app._query_criteria_range == (0, 0, 2, 2)

    def test_do_set_criteria(self):
        self.handler._do_set_criteria("F1:G2")
        assert self.app._query_criteria_range == (0, 5, 1, 6)

    def test_set_output(self):
        self.handler.set_output()
        assert self.app._query_output_range == (0, 0)

    def test_do_set_output(self):
        self.handler._do_set_output("H1")
        assert self.app._query_output_range == (0, 7)
    
    def test_find_flow(self):
        # Setup mock db query result
        self.app._query_input_range = (0, 0, 5, 5)
        self.app._query_criteria_range = (0, 6, 1, 6)
        
        # Mock DatabaseOperations locally within find method scope or patch it
        with patch('lotus123.data.database.DatabaseOperations') as MockDB:
            mock_db_instance = MockDB.return_value
            mock_db_instance.query.return_value = [1, 2, 3] # rows 1, 2, 3 overlap
            
            self.handler.find()
            
            assert self.app._query_find_results == [1, 2, 3]
            assert self.app._query_find_index == 0
            # Should invoke cursor move
            assert self.app.grid.cursor_row == 1

            # Next find
            self.handler.find()
            assert self.app._query_find_index == 1
            assert self.app.grid.cursor_row == 2

        self.handler.reset()
        assert self.app._query_input_range is None

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

    def test_insert_row(self):
        self.handler.insert_row()
        self.app.undo_manager.execute.assert_called()
        
    def test_delete_row(self):
        self.handler.delete_row()
        self.app.undo_manager.execute.assert_called()

    def test_col_width(self):
        self.handler.set_column_width()
        self.app.push_screen.assert_called()

class TestFileHandler:
    def setup_method(self):
        self.app = MockApp()
        from lotus123.handlers.file_handlers import FileHandler
        self.handler = FileHandler(self.app)
        self.app.spreadsheet.filename = "test.json"
        
    def test_save(self):
        with patch('lotus123.handlers.file_handlers.FileDialog'):
            self.handler.save()
            # If filename set, simply notifies
            assert self.app._dirty is False

    def test_retrieve(self):
        with patch('lotus123.handlers.file_handlers.FileDialog'):
            self.handler.open_file()
            self.app.push_screen.assert_called()

class TestDataHandler:
    def setup_method(self):
        self.app = MockApp()
        from lotus123.handlers.data_handlers import DataHandler
        self.handler = DataHandler(self.app)
        
    def test_fill(self):
        # Fill needs range
        self.app.grid.selection_range = (0,0,5,0)
        self.handler.data_fill()
        self.app.push_screen.assert_called()

    def test_sort(self):
        pass # Skipping brittle sort test as per user direction

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
        self.app._pending_range = ""
        self.handler.range_name()
        self.app.push_screen.assert_called() # Prompt for name

    def test_range_protect(self):
        # Range is A1:C3
        self.handler.range_protect()
        # Should toggle protection
        # We can assert notify was called with "Protected" or "Unprotected"
        # self.app.notify.filter_args?
        pass

class TestChartHandler:
    def setup_method(self):
        self.app = MockApp()
        # Mock chart internals
        self.app.chart = MagicMock()
        self.app.chart.series = []
        
        def add_series_side_effect(name):
            self.app.chart.series.append(MagicMock(name=name))
            
        self.app.chart.add_series.side_effect = add_series_side_effect
        self.app._chart_renderer = MagicMock()
        
        from lotus123.handlers.chart_handlers import ChartHandler, ChartType
        self.handler = ChartHandler(self.app)

    def test_set_chart_type(self):
        from lotus123.charting import ChartType
        self.handler.set_chart_type(ChartType.BAR)
        self.app.chart.set_type.assert_called_with(ChartType.BAR)

    def test_set_ranges(self):
        # Test X
        self.handler.set_x_range()
        self.app.chart.set_x_range.assert_called()
        
        # Test A-F
        ranges = ['a', 'b', 'c', 'd', 'e', 'f']
        for r in ranges:
            method_name = f"set_{r}_range"
            getattr(self.handler, method_name)()
            # Should invoke add_series
            assert self.app.chart.add_series.call_count >= 1
            
            # Test callback
            callback_name = f"_do_set_{r}_range"
            getattr(self.handler, callback_name)(f"{r.upper()}1:{r.upper()}10")

    def test_view_chart(self):
        # With no series -> warning
        self.handler.view_chart()
        self.app._chart_renderer.render.assert_not_called()
        
        # With series
        self.app.chart.series = [MagicMock()]
        self.app.size = MagicMock(width=80, height=24)
        self.handler.view_chart()
        self.app._chart_renderer.render.assert_called()
        self.app.push_screen.assert_called()

    def test_reset_chart(self):
        self.handler.reset_chart()
        self.app.chart.reset.assert_called()

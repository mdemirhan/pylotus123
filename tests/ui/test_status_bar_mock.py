
from unittest.mock import MagicMock, PropertyMock, patch
from lotus123.ui.status_bar import StatusBar, StatusBarWidget, Mode
from lotus123.core.cell import Cell

class TestStatusBar:
    def setup_method(self):
        self.mock_sheet = MagicMock()
        self.status = StatusBar(self.mock_sheet)
    
    def test_initial_state(self):
        assert self.status.mode.mode == Mode.READY
        assert self.status.current_cell_ref == "A1"
        assert not self.status.needs_recalc
        assert not self.status.has_circular_ref
        
    def test_mode_indicator(self):
        self.status.set_mode(Mode.EDIT)
        assert self.status.mode.mode == Mode.EDIT
        assert self.status.mode.text == "EDIT"
        
        self.status.set_mode(Mode.ERROR)
        assert self.status.mode.text == "ERROR"

    def test_lock_indicators(self):
        self.status.locks.caps_lock = True
        assert "CAPS" in self.status.locks.as_string()
        
        self.status.locks.num_lock = True
        assert "CAPS NUM" in self.status.locks.as_string()

    def test_update_from_spreadsheet(self):
        self.mock_sheet.needs_recalc = True
        self.mock_sheet.has_circular_refs = True
        self.mock_sheet.modified = True
        self.mock_sheet.filename = "test.json"

        self.status.update_from_spreadsheet()

        assert self.status.needs_recalc
        assert self.status.has_circular_ref
        assert self.status.modified

        assert "CALC" in self.status.get_indicators()
        assert "CIRC" in self.status.get_indicators()
        assert "(modified)" in self.status.get_indicators()

    def test_filename_display(self):
        # New file (no filename) - should show **New File**
        self.status.filename = ""
        self.status.modified = False
        assert self.status.get_filename_display() == "**New File**"

        # New file modified - still shows **New File** (no modified suffix)
        self.status.modified = True
        assert self.status.get_filename_display() == "**New File**"

        # Existing file - shows filename
        self.status.filename = "/path/to/budget.json"
        self.status.modified = False
        assert self.status.get_filename_display() == "budget.json"

        # Existing file modified - shows filename (modified)
        self.status.modified = True
        assert self.status.get_filename_display() == "budget.json (modified)"

    def test_update_cell(self):
        # Mock cell
        cell = MagicMock(spec=Cell)
        cell.is_formula = False
        cell.raw_value = "123"
        self.mock_sheet.get_cell_if_exists.return_value = cell
        self.mock_sheet.get_display_value.return_value = "123"
        
        self.status.update_cell(0, 0)
        
        assert self.status.current_cell_ref == "A1"
        assert self.status.current_cell_value == "123"
        assert "A1: 123" in self.status.get_cell_display()
        
        # Test formula
        cell.is_formula = True
        cell.raw_value = "+A1+1"
        self.status.update_cell(0, 0)
        assert "+A1+1" in self.status.current_cell_formula
        assert "A1: +A1+1 = 123" in self.status.get_cell_display()

    def test_full_status_formatting(self):
        self.status.current_cell_ref = "A1"
        self.status.current_cell_value = "Test"
        self.status.set_mode(Mode.READY)
        
        full = self.status.get_full_status(width=40)
        assert "A1: Test" in full
        assert "READY" in full
        assert len(full) > 20

    def test_memory_format(self):
        assert self.status.format_memory(500) == "500B"
        assert self.status.format_memory(2048) == "2K"
        assert self.status.format_memory(1500000) == "1.4M"

class TestStatusBarWidget:
    def setup_method(self):
        self.mock_sheet = MagicMock()
        self.widget = StatusBarWidget(self.mock_sheet)
        
        # Mock textual internals
        self.widget.update = MagicMock()
        self.size_patcher = patch.object(StatusBarWidget, 'size', new_callable=PropertyMock)
        self.mock_size = self.size_patcher.start()
        self.mock_size.return_value = MagicMock(width=80)

    def teardown_method(self):
        self.size_patcher.stop()

    def test_refresh_delegation(self):
        self.widget.set_mode(Mode.EDIT)
        self.widget.update.assert_called()
        call_arg = self.widget.update.call_args[0][0]
        assert "EDIT" in call_arg

    def test_update_wrappers(self):
        self.widget.set_message("Hello")
        assert self.widget.status.message == "Hello"
        self.widget.update.assert_called()
        
        self.widget.clear_message()
        assert self.widget.status.message == ""
        
        self.widget.set_modified(True)
        assert self.widget.status.modified

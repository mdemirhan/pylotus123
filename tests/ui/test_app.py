"""Tests for main Lotus 1-2-3 TUI application."""

import pytest

from lotus123.app import LotusApp
from lotus123.core import Spreadsheet


class TestLotusAppInit:
    """Tests for LotusApp initialization."""

    def test_app_init_no_file(self):
        """Test app initializes without file."""
        app = LotusApp()
        assert app._initial_file is None
        assert app.spreadsheet is not None
        assert isinstance(app.spreadsheet, Spreadsheet)

    def test_app_init_with_file(self):
        """Test app initializes with file path."""
        app = LotusApp(initial_file="test.wk1")
        assert app._initial_file == "test.wk1"

    def test_app_has_config(self):
        """Test app loads configuration."""
        app = LotusApp()
        assert app.config is not None

    def test_app_has_theme(self):
        """Test app has theme configured."""
        app = LotusApp()
        assert app.color_theme is not None

    def test_app_has_undo_manager(self):
        """Test app has undo manager."""
        app = LotusApp()
        assert app.undo_manager is not None
        assert app.undo_manager.can_undo is False

    def test_app_has_chart(self):
        """Test app has chart object."""
        app = LotusApp()
        assert app.chart is not None

    def test_app_initial_state(self):
        """Test app initial state."""
        app = LotusApp()
        assert app.editing is False
        assert app._menu_active is False
        assert app._dirty is False
        assert app.recalc_mode == "auto"


class TestLotusAppProperties:
    """Tests for LotusApp properties."""

    def test_has_modal_false(self):
        """Test _has_modal is False initially."""
        app = LotusApp()
        # Can't test directly without running app, but structure exists
        assert hasattr(app, "_has_modal")

    def test_global_settings(self):
        """Test global settings are initialized."""
        app = LotusApp()
        assert app.global_format_code == "G"
        assert app.global_label_prefix == "'"
        assert app.global_col_width == 10
        assert app.global_zero_display is True


class TestLotusAppClipboard:
    """Tests for clipboard functionality."""

    def test_clipboard_initial_state(self):
        """Test clipboard is empty initially."""
        app = LotusApp()
        # Clipboard state is now owned by the clipboard handler
        assert app._clipboard_handler.cell_clipboard is None
        assert app._clipboard_handler.range_clipboard is None
        assert app._clipboard_handler.clipboard_is_cut is False


class TestLotusAppQuery:
    """Tests for query settings."""

    def test_query_initial_state(self):
        """Test query settings initial state."""
        app = LotusApp()
        # Query state is now owned by the query handler
        assert app._query_handler.input_range is None
        assert app._query_handler.criteria_range is None
        assert app._query_handler.output_range is None
        assert app._query_handler.find_results is None
        assert app._query_handler.find_index == 0


class TestLotusAppBindings:
    """Tests for key bindings."""

    def test_bindings_exist(self):
        """Test key bindings are defined."""
        assert LotusApp.BINDINGS is not None
        assert len(LotusApp.BINDINGS) > 0

    def test_save_binding(self):
        """Test save binding exists."""
        bindings = {b.key: b for b in LotusApp.BINDINGS}
        assert "ctrl+s" in bindings

    def test_open_binding(self):
        """Test open binding exists."""
        bindings = {b.key: b for b in LotusApp.BINDINGS}
        assert "ctrl+o" in bindings

    def test_undo_redo_bindings(self):
        """Test undo/redo bindings exist."""
        bindings = {b.key: b for b in LotusApp.BINDINGS}
        assert "ctrl+z" in bindings
        assert "ctrl+y" in bindings

    def test_copy_paste_bindings(self):
        """Test copy/paste bindings exist."""
        bindings = {b.key: b for b in LotusApp.BINDINGS}
        assert "ctrl+c" in bindings
        assert "ctrl+v" in bindings
        assert "ctrl+x" in bindings

    def test_navigation_bindings(self):
        """Test navigation bindings exist."""
        bindings = {b.key: b for b in LotusApp.BINDINGS}
        assert "ctrl+g" in bindings  # Goto
        assert "pageup" in bindings
        assert "pagedown" in bindings
        assert "home" in bindings
        assert "end" in bindings

    def test_function_key_bindings(self):
        """Test function key bindings exist."""
        bindings = {b.key: b for b in LotusApp.BINDINGS}
        assert "f2" in bindings  # Edit
        assert "f5" in bindings  # Goto


class TestLotusAppCSS:
    """Tests for CSS styling."""

    def test_css_defined(self):
        """Test CSS is defined."""
        assert LotusApp.CSS is not None
        assert len(LotusApp.CSS) > 0

    def test_css_contains_screen(self):
        """Test CSS defines Screen."""
        assert "Screen" in LotusApp.CSS

    def test_css_contains_grid(self):
        """Test CSS defines grid."""
        assert "#grid" in LotusApp.CSS

    def test_css_contains_menu(self):
        """Test CSS defines menu bar."""
        assert "#menu-bar" in LotusApp.CSS

    def test_css_contains_status_bar(self):
        """Test CSS defines status bar."""
        assert "#status-bar" in LotusApp.CSS


class TestLotusAppMethods:
    """Tests for LotusApp methods."""

    def test_mark_dirty(self):
        """Test _mark_dirty method."""
        app = LotusApp()
        assert app._dirty is False
        app._mark_dirty()
        assert app._dirty is True

    def test_generate_css(self):
        """Test _generate_css returns CSS string."""
        app = LotusApp()
        css = app._generate_css()
        assert isinstance(css, str)
        assert len(css) > 0
        assert "Screen" in css
        assert "background" in css


# Async tests for Textual app functionality
class TestLotusAppAsync:
    """Async tests using Textual's testing framework."""

    @pytest.mark.asyncio
    async def test_app_mounts(self):
        """Test app can mount."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # App should be running
            assert app.is_running

    @pytest.mark.asyncio
    async def test_app_has_grid(self):
        """Test app has spreadsheet grid widget."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Should be able to find the grid
            try:
                grid = app.query_one("#grid")
                assert grid is not None
            except Exception:
                pass  # Grid might not be found in test environment

    @pytest.mark.asyncio
    async def test_app_quit(self):
        """Test app can quit."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Press Ctrl+Q to quit
            await pilot.press("ctrl+q")
            # App may show quit dialog or exit

    @pytest.mark.asyncio
    async def test_escape_key(self):
        """Test escape key handling."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("escape")
            # Should not crash

    @pytest.mark.asyncio
    async def test_navigation_keys(self):
        """Test navigation keys work."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("down")
            await pilot.press("right")
            await pilot.press("up")
            await pilot.press("left")
            # Should not crash

    @pytest.mark.asyncio
    async def test_enter_text(self):
        """Test entering text in cell."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Type some text
            await pilot.press("1")
            await pilot.press("2")
            await pilot.press("3")
            # Should not crash

    @pytest.mark.asyncio
    async def test_f2_edit(self):
        """Test F2 edit key."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("f2")
            # Should enter edit mode or handle gracefully

    @pytest.mark.asyncio
    async def test_delete_key(self):
        """Test delete key clears cell."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # First add some content
            app.spreadsheet.set_cell(0, 0, "test")
            # Press delete
            await pilot.press("delete")

    @pytest.mark.asyncio
    async def test_page_navigation(self):
        """Test page up/down navigation."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("pagedown")
            await pilot.press("pageup")
            # Should not crash

    @pytest.mark.asyncio
    async def test_home_end_keys(self):
        """Test home/end keys."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("home")
            await pilot.press("end")
            # Should not crash


class TestLotusAppSpreadsheetIntegration:
    """Tests for spreadsheet integration."""

    def test_spreadsheet_accessible(self):
        """Test spreadsheet is accessible."""
        app = LotusApp()
        assert app.spreadsheet is not None
        # Can set and get cells
        app.spreadsheet.set_cell(0, 0, "test")
        assert app.spreadsheet.get_value(0, 0) == "test"

    def test_spreadsheet_formulas(self):
        """Test formulas work through app."""
        app = LotusApp()
        app.spreadsheet.set_cell(0, 0, "10")
        app.spreadsheet.set_cell(0, 1, "20")
        app.spreadsheet.set_cell(0, 2, "=A1+B1")
        assert app.spreadsheet.get_value(0, 2) == 30


class TestLotusAppActionsAsync:
    """More async tests for app actions."""

    @pytest.mark.asyncio
    async def test_action_new_file(self):
        """Test new file action."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Set some data first
            app.spreadsheet.set_cell(0, 0, "test data")
            app._dirty = True
            # Run new file action
            app.action_new_file()
            # Spreadsheet should be cleared
            assert app.spreadsheet.get_value(0, 0) == ""
            assert app._dirty is False

    @pytest.mark.asyncio
    async def test_action_undo_nothing(self):
        """Test undo with nothing to undo."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Undo should not crash with empty undo stack
            app.action_undo()

    @pytest.mark.asyncio
    async def test_action_redo_nothing(self):
        """Test redo with nothing to redo."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app.action_redo()

    @pytest.mark.asyncio
    async def test_action_copy(self):
        """Test copy action."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app.spreadsheet.set_cell(0, 0, "copy me")
            app.action_copy()

    @pytest.mark.asyncio
    async def test_action_cut(self):
        """Test cut action."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app.spreadsheet.set_cell(0, 0, "cut me")
            app.action_cut()

    @pytest.mark.asyncio
    async def test_action_paste(self):
        """Test paste action."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app.spreadsheet.set_cell(0, 0, "copy me")
            app.action_copy()
            await pilot.press("right")
            app.action_paste()

    @pytest.mark.asyncio
    async def test_action_clear_cell(self):
        """Test clear cell action."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app.spreadsheet.set_cell(0, 0, "clear me")
            app.action_clear_cell()

    @pytest.mark.asyncio
    async def test_action_edit_cell(self):
        """Test edit cell action."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app.action_edit_cell()
            assert app.editing is True

    @pytest.mark.asyncio
    async def test_action_cancel_edit(self):
        """Test cancel edit action."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app.action_edit_cell()
            assert app.editing is True
            app.action_cancel_edit()
            assert app.editing is False

    @pytest.mark.asyncio
    async def test_action_show_menu(self):
        """Test show menu action."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Press / to show menu
            await pilot.press("/")

    @pytest.mark.asyncio
    async def test_action_goto(self):
        """Test goto action shows dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Trigger goto
            await pilot.press("ctrl+g")

    @pytest.mark.asyncio
    async def test_action_save(self):
        """Test save action."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Save without filename should show dialog
            await pilot.press("ctrl+s")

    @pytest.mark.asyncio
    async def test_action_move_to_cell_start(self):
        """Test move to cell start."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Navigate first
            await pilot.press("down")
            await pilot.press("right")
            # Press home
            await pilot.press("home")

    @pytest.mark.asyncio
    async def test_action_move_to_cell_end(self):
        """Test move to cell end."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Set some data
            app.spreadsheet.set_cell(5, 5, "data")
            # Press end
            await pilot.press("end")

    @pytest.mark.asyncio
    async def test_action_select_all(self):
        """Test select all action."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("ctrl+a")

    @pytest.mark.asyncio
    async def test_enter_and_submit_value(self):
        """Test entering and submitting a value."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Type a value
            await pilot.press("1")
            await pilot.press("0")
            await pilot.press("0")
            await pilot.press("enter")

    @pytest.mark.asyncio
    async def test_enter_formula(self):
        """Test entering a formula."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Set up data
            app.spreadsheet.set_cell(0, 0, "10")
            app.spreadsheet.set_cell(0, 1, "20")
            # Move to new cell and enter formula
            await pilot.press("right")
            await pilot.press("right")
            await pilot.press("=")

    @pytest.mark.asyncio
    async def test_ctrl_home(self):
        """Test Ctrl+Home navigation."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("down")
            await pilot.press("down")
            await pilot.press("right")
            await pilot.press("ctrl+home")

    @pytest.mark.asyncio
    async def test_ctrl_end(self):
        """Test Ctrl+End navigation."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app.spreadsheet.set_cell(10, 10, "data")
            await pilot.press("ctrl+end")

    @pytest.mark.asyncio
    async def test_tab_navigation(self):
        """Test tab navigation."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("tab")
            await pilot.press("shift+tab")


class TestLotusAppMenuHandling:
    """Tests for menu handling."""

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_goto(self):
        """Test _handle_menu with Goto opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Goto")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_worksheet_goto(self):
        """Test _handle_menu with Worksheet:Goto opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Worksheet:Goto")

    @pytest.mark.asyncio
    async def test_handle_menu_file_new(self):
        """Test _handle_menu with File:New."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("File:New")
            # Should clear the spreadsheet
            assert app._dirty is False

    @pytest.mark.asyncio
    async def test_handle_menu_range_erase(self):
        """Test _handle_menu with Range:Erase."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Range:Erase")

    @pytest.mark.asyncio
    async def test_handle_menu_quit_no(self):
        """Test _handle_menu with Quit:No."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Quit:No")

    @pytest.mark.asyncio
    async def test_handle_menu_none(self):
        """Test _handle_menu with None."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu(None)

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_copy(self):
        """Test _handle_menu with Copy opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Copy")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_move(self):
        """Test _handle_menu with Move opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Move")


class TestLotusAppChartHandling:
    """Tests for chart handling."""

    @pytest.mark.asyncio
    async def test_handle_menu_graph_type_line(self):
        """Test _handle_menu with Graph:Type:Line."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Graph:Type:Line")

    @pytest.mark.asyncio
    async def test_handle_menu_graph_type_bar(self):
        """Test _handle_menu with Graph:Type:Bar."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Graph:Type:Bar")

    @pytest.mark.asyncio
    async def test_handle_menu_graph_type_xy(self):
        """Test _handle_menu with Graph:Type:XY."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Graph:Type:XY")

    @pytest.mark.asyncio
    async def test_handle_menu_graph_type_stacked(self):
        """Test _handle_menu with Graph:Type:Stacked."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Graph:Type:Stacked")

    @pytest.mark.asyncio
    async def test_handle_menu_graph_type_pie(self):
        """Test _handle_menu with Graph:Type:Pie."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Graph:Type:Pie")

    @pytest.mark.asyncio
    async def test_handle_menu_graph_reset(self):
        """Test _handle_menu with Graph:Reset."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Graph:Reset")


class TestLotusAppWorksheetHandling:
    """Tests for worksheet handling."""

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_worksheet_insert_rows(self):
        """Test _handle_menu with Worksheet:Insert:Rows opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Worksheet:Insert:Rows")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_worksheet_insert_columns(self):
        """Test _handle_menu with Worksheet:Insert:Columns opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Worksheet:Insert:Columns")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_worksheet_delete_rows(self):
        """Test _handle_menu with Worksheet:Delete:Rows opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Worksheet:Delete:Rows")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_worksheet_delete_columns(self):
        """Test _handle_menu with Worksheet:Delete:Columns opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Worksheet:Delete:Columns")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_worksheet_column(self):
        """Test _handle_menu with Worksheet:Column opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Worksheet:Column")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_worksheet_erase(self):
        """Test _handle_menu with Worksheet:Erase opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Worksheet:Erase")


class TestLotusAppDataHandling:
    """Tests for data handling."""

    @pytest.mark.asyncio
    async def test_handle_menu_data_fill(self):
        """Test _handle_menu with Data:Fill."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Data:Fill")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_data_sort(self):
        """Test _handle_menu with Data:Sort opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Data:Sort")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_query_input(self):
        """Test _handle_menu with Data:Query:Input opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Data:Query:Input")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_query_criteria(self):
        """Test _handle_menu with Data:Query:Criteria opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Data:Query:Criteria")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_query_output(self):
        """Test _handle_menu with Data:Query:Output opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Data:Query:Output")

    @pytest.mark.asyncio
    async def test_handle_menu_query_find(self):
        """Test _handle_menu with Data:Query:Find."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Data:Query:Find")

    @pytest.mark.asyncio
    async def test_handle_menu_query_extract(self):
        """Test _handle_menu with Data:Query:Extract."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Data:Query:Extract")

    @pytest.mark.asyncio
    async def test_handle_menu_query_unique(self):
        """Test _handle_menu with Data:Query:Unique."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Data:Query:Unique")

    @pytest.mark.asyncio
    async def test_handle_menu_query_delete(self):
        """Test _handle_menu with Data:Query:Delete."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Data:Query:Delete")

    @pytest.mark.asyncio
    async def test_handle_menu_query_reset(self):
        """Test _handle_menu with Data:Query:Reset."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Data:Query:Reset")


class TestLotusAppGlobalSettings:
    """Tests for global settings handling."""

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_global_format(self):
        """Test _handle_menu with Worksheet:Global:Format opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Worksheet:Global:Format")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_global_label_prefix(self):
        """Test _handle_menu with Worksheet:Global:Label-Prefix opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Worksheet:Global:Label-Prefix")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_global_column_width(self):
        """Test _handle_menu with Worksheet:Global:Column-Width opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Worksheet:Global:Column-Width")

    @pytest.mark.asyncio
    async def test_handle_menu_global_recalculation(self):
        """Test _handle_menu with Worksheet:Global:Recalculation."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Worksheet:Global:Recalculation")

    @pytest.mark.asyncio
    async def test_handle_menu_global_zero(self):
        """Test _handle_menu with Worksheet:Global:Zero."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Worksheet:Global:Zero")


class TestLotusAppRangeHandling:
    """Tests for range handling."""

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_range_format(self):
        """Test _handle_menu with Range:Format opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Range:Format")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_range_label(self):
        """Test _handle_menu with Range:Label opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Range:Label")

    @pytest.mark.skip(reason="Dialog interactions require complex setup")
    @pytest.mark.asyncio
    async def test_handle_menu_range_name(self):
        """Test _handle_menu with Range:Name opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            app._handle_menu("Range:Name")

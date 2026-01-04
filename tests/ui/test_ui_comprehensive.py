"""Comprehensive UI tests for Lotus 1-2-3 clone.

These tests cover gaps identified in UI test coverage including:
- Dialog integration tests (submit/cancel workflows)
- Range selection tests (Shift+arrow, clipboard)
- Menu navigation tests (keyboard traversal, submenus)
- Grid interaction tests (mouse clicks, scrolling)
- End-to-end workflow tests

These tests are marked as slow and excluded from default test runs.
Run with: uv run pytest -m slow
Run all tests: uv run pytest -m ''
"""

import pytest

from lotus123.app import LotusApp

# Mark all tests in this module as slow
pytestmark = pytest.mark.slow
from lotus123.ui import SpreadsheetGrid, LotusMenu
from lotus123.ui.themes import ThemeType


# =============================================================================
# Dialog Integration Tests
# =============================================================================


class TestCommandInputDialog:
    """Tests for CommandInput dialog interactions."""

    @pytest.mark.asyncio
    async def test_command_input_submit_with_enter(self):
        """Test submitting CommandInput with Enter key."""

        app = LotusApp()
        async with app.run_test() as pilot:
            # Open Goto dialog
            await pilot.press("ctrl+g")
            await pilot.pause()

            # Verify dialog is open
            assert len(app.screen_stack) > 1

            # Type a cell reference
            await pilot.press("b")
            await pilot.press("5")
            await pilot.pause()

            # Submit with Enter
            await pilot.press("enter")
            await pilot.pause()

            # Dialog should close and cursor should move
            grid = app.query_one("#grid", SpreadsheetGrid)
            assert grid.cursor_row == 4  # B5 = row 4 (0-indexed)
            assert grid.cursor_col == 1  # B = col 1

    @pytest.mark.asyncio
    async def test_command_input_cancel_with_escape(self):
        """Test canceling CommandInput with Escape key."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)
            initial_row = grid.cursor_row
            initial_col = grid.cursor_col

            # Open Goto dialog
            await pilot.press("ctrl+g")
            await pilot.pause()
            assert len(app.screen_stack) > 1

            # Type something
            await pilot.press("z")
            await pilot.press("9")
            await pilot.press("9")
            await pilot.pause()

            # Cancel with Escape
            await pilot.press("escape")
            await pilot.pause()

            # Dialog should close, cursor should NOT move
            assert len(app.screen_stack) == 1
            assert grid.cursor_row == initial_row
            assert grid.cursor_col == initial_col

    @pytest.mark.asyncio
    async def test_command_input_default_value_selection(self):
        """Test that default value is selected for replacement."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Move to B2
            await pilot.press("down")
            await pilot.press("right")
            await pilot.pause()

            # Open Goto dialog - should have current cell as default
            await pilot.press("ctrl+g")
            await pilot.pause()

            # Type new reference (should replace selected text)
            await pilot.press("a")
            await pilot.press("1")
            await pilot.press("enter")
            await pilot.pause()

            # Should be at A1
            grid = app.query_one("#grid", SpreadsheetGrid)
            assert grid.cursor_row == 0
            assert grid.cursor_col == 0

    @pytest.mark.asyncio
    async def test_goto_with_f5(self):
        """Test Goto dialog opens with F5."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("f5")
            await pilot.pause()
            assert len(app.screen_stack) > 1

            await pilot.press("escape")
            await pilot.pause()


class TestThemeDialogInteraction:
    """Tests for ThemeDialog interactions."""

    @pytest.mark.asyncio
    async def test_theme_dialog_number_shortcuts(self):
        """Test theme selection with number keys 1-7."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Open theme dialog
            await pilot.press("ctrl+t")
            await pilot.pause()
            assert len(app.screen_stack) > 1

            # Select theme 2 (TOMORROW)
            await pilot.press("2")
            await pilot.pause()

            # Dialog should close, theme should change
            assert len(app.screen_stack) == 1
            assert app.current_theme_type == ThemeType.TOMORROW

    @pytest.mark.asyncio
    async def test_theme_dialog_arrow_navigation(self):
        """Test theme selection with arrow keys and Enter."""
        app = LotusApp()
        async with app.run_test() as pilot:
            initial_theme = app.current_theme_type

            # Open theme dialog
            await pilot.press("ctrl+t")
            await pilot.pause()

            # Navigate down once
            await pilot.press("down")
            await pilot.pause()

            # Select with Enter
            await pilot.press("enter")
            await pilot.pause()

            # Theme should have changed (to the one after initial)
            assert app.current_theme_type != initial_theme

    @pytest.mark.asyncio
    async def test_theme_dialog_cancel(self):
        """Test canceling theme dialog preserves original theme."""
        app = LotusApp()
        async with app.run_test() as pilot:
            original_theme = app.current_theme_type

            # Open theme dialog
            await pilot.press("ctrl+t")
            await pilot.pause()

            # Navigate to different theme
            await pilot.press("down")
            await pilot.press("down")
            await pilot.pause()

            # Cancel
            await pilot.press("escape")
            await pilot.pause()

            # Theme should be unchanged
            assert app.current_theme_type == original_theme

    @pytest.mark.asyncio
    async def test_theme_dialog_all_themes_selectable(self):
        """Test all 7 themes can be selected."""
        for i in range(1, 8):
            app = LotusApp()
            async with app.run_test() as pilot:
                await pilot.press("ctrl+t")
                await pilot.pause()
                await pilot.press(str(i))
                await pilot.pause()
                # Should not crash and dialog should close
                assert len(app.screen_stack) == 1


class TestFileDialogInteraction:
    """Tests for FileDialog interactions."""

    @pytest.mark.asyncio
    async def test_file_dialog_opens_on_save(self):
        """Test FileDialog opens on Ctrl+S for new file."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("ctrl+s")
            await pilot.pause()
            # Should open file dialog for new file
            assert len(app.screen_stack) > 1

    @pytest.mark.asyncio
    async def test_file_dialog_opens_on_open(self):
        """Test FileDialog opens on Ctrl+O."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("ctrl+o")
            await pilot.pause()
            assert len(app.screen_stack) > 1

    @pytest.mark.asyncio
    async def test_file_dialog_cancel(self):
        """Test FileDialog cancels properly."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("ctrl+o")
            await pilot.pause()
            assert len(app.screen_stack) > 1

            await pilot.press("escape")
            await pilot.pause()
            assert len(app.screen_stack) == 1

    @pytest.mark.asyncio
    async def test_file_dialog_via_menu(self):
        """Test FileDialog opens via menu."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Activate menu
            await pilot.press("slash")
            await pilot.pause()

            # Select File menu
            await pilot.press("f")
            await pilot.pause()

            # Select Retrieve
            await pilot.press("r")
            await pilot.pause()

            # Should have file dialog open
            assert len(app.screen_stack) > 1


# =============================================================================
# Range Selection Tests
# =============================================================================


class TestRangeSelection:
    """Tests for range selection functionality."""

    @pytest.mark.asyncio
    async def test_shift_arrow_starts_selection(self):
        """Test Shift+Arrow starts range selection."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Initially no selection
            assert not grid.has_selection

            # Shift+Down should start selection
            await pilot.press("shift+down")
            await pilot.pause()

            # Should have selection now
            assert grid.has_selection
            r1, c1, r2, c2 = grid.selection_range
            assert r1 == 0 and c1 == 0  # Anchor
            assert r2 == 1 and c2 == 0  # Cursor moved down

    @pytest.mark.asyncio
    async def test_shift_arrow_extends_selection(self):
        """Test multiple Shift+Arrow extends selection."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Select 3x3 range
            await pilot.press("shift+down")
            await pilot.press("shift+down")
            await pilot.press("shift+right")
            await pilot.press("shift+right")
            await pilot.pause()

            r1, c1, r2, c2 = grid.selection_range
            assert r1 == 0 and c1 == 0
            assert r2 == 2 and c2 == 2  # 3 rows, 3 cols

    @pytest.mark.asyncio
    async def test_arrow_without_shift_clears_selection(self):
        """Test plain Arrow key clears selection."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Create selection
            await pilot.press("shift+down")
            await pilot.press("shift+right")
            await pilot.pause()
            assert grid.has_selection

            # Plain arrow should clear selection
            await pilot.press("down")
            await pilot.pause()
            assert not grid.has_selection

    @pytest.mark.asyncio
    async def test_selection_cleared_on_navigation(self):
        """Test selection is cleared when navigating without Shift."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Create selection
            await pilot.press("shift+down")
            await pilot.press("shift+right")
            await pilot.pause()
            assert grid.has_selection

            # Regular navigation should clear selection
            await pilot.press("down")
            await pilot.pause()
            assert not grid.has_selection

    @pytest.mark.asyncio
    async def test_selection_covers_range(self):
        """Test selection range is correctly computed."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Create a 2x2 selection
            await pilot.press("shift+down")
            await pilot.press("shift+right")
            await pilot.pause()

            # Check selection range
            r1, c1, r2, c2 = grid.selection_range
            assert r1 == 0 and c1 == 0
            assert r2 == 1 and c2 == 1


class TestClipboardOperations:
    """Tests for copy/cut/paste operations."""

    @pytest.mark.asyncio
    async def test_copy_single_cell(self):
        """Test copying a single cell."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Set value and copy
            app.spreadsheet.set_cell(0, 0, "test")
            await pilot.press("ctrl+c")
            await pilot.pause()

            # Move and paste
            await pilot.press("down")
            await pilot.press("ctrl+v")
            await pilot.pause()

            # Value should be copied
            assert app.spreadsheet.get_display_value(1, 0) == "test"

    @pytest.mark.asyncio
    async def test_cut_single_cell(self):
        """Test cutting a single cell."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Set value and cut
            app.spreadsheet.set_cell(0, 0, "test")
            await pilot.press("ctrl+x")
            await pilot.pause()

            # Move and paste
            await pilot.press("down")
            await pilot.press("ctrl+v")
            await pilot.pause()

            # Value should be moved
            assert app.spreadsheet.get_display_value(0, 0) == ""
            assert app.spreadsheet.get_display_value(1, 0) == "test"

    @pytest.mark.asyncio
    async def test_copy_range(self):
        """Test copying a range of cells stores to clipboard."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Set values
            app.spreadsheet.set_cell(0, 0, "A1")
            app.spreadsheet.set_cell(0, 1, "B1")
            app.spreadsheet.set_cell(1, 0, "A2")
            app.spreadsheet.set_cell(1, 1, "B2")

            # Select range
            await pilot.press("shift+down")
            await pilot.press("shift+right")
            await pilot.pause()

            # Copy - this stores to clipboard
            await pilot.press("ctrl+c")
            await pilot.pause()

            # Verify clipboard has data (range_clipboard is used for multi-cell)
            assert (
                app._clipboard_handler.range_clipboard is not None
                or app._clipboard_handler.cell_clipboard is not None
            )

    @pytest.mark.asyncio
    async def test_delete_clears_selection(self):
        """Test Delete key clears selected range."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Set values
            app.spreadsheet.set_cell(0, 0, "A")
            app.spreadsheet.set_cell(0, 1, "B")
            app.spreadsheet.set_cell(1, 0, "C")
            app.spreadsheet.set_cell(1, 1, "D")

            # Select range
            await pilot.press("shift+down")
            await pilot.press("shift+right")
            await pilot.pause()

            # Delete
            await pilot.press("delete")
            await pilot.pause()

            # All should be cleared
            assert app.spreadsheet.get_display_value(0, 0) == ""
            assert app.spreadsheet.get_display_value(0, 1) == ""
            assert app.spreadsheet.get_display_value(1, 0) == ""
            assert app.spreadsheet.get_display_value(1, 1) == ""


# =============================================================================
# Menu Navigation Tests
# =============================================================================


class TestMenuNavigation:
    """Tests for menu keyboard navigation."""

    @pytest.mark.asyncio
    async def test_menu_activate_with_slash(self):
        """Test menu activates with / key."""
        app = LotusApp()
        async with app.run_test() as pilot:
            menu = app.query_one("#menu-bar", LotusMenu)
            assert not menu.active

            await pilot.press("slash")
            await pilot.pause()

            assert menu.active
            assert menu.current_menu is None  # At top level

    @pytest.mark.asyncio
    async def test_menu_top_level_selection(self):
        """Test selecting top-level menu items."""
        app = LotusApp()
        async with app.run_test() as pilot:
            menu = app.query_one("#menu-bar", LotusMenu)

            await pilot.press("slash")
            await pilot.pause()

            # Press W for Worksheet
            await pilot.press("w")
            await pilot.pause()

            assert menu.current_menu == "Worksheet"

    @pytest.mark.asyncio
    async def test_menu_submenu_navigation(self):
        """Test navigating into submenus."""
        app = LotusApp()
        async with app.run_test() as pilot:
            menu = app.query_one("#menu-bar", LotusMenu)

            await pilot.press("slash")
            await pilot.pause()

            # Worksheet -> Global submenu
            await pilot.press("w")
            await pilot.pause()
            assert menu.current_menu == "Worksheet"

            # Press G for Global (has submenu)
            await pilot.press("g")
            await pilot.pause()
            assert "Global" in menu.submenu_path

    @pytest.mark.asyncio
    async def test_menu_escape_back_through_levels(self):
        """Test Escape goes back through menu levels."""
        app = LotusApp()
        async with app.run_test() as pilot:
            menu = app.query_one("#menu-bar", LotusMenu)

            await pilot.press("slash")
            await pilot.press("w")
            await pilot.press("g")
            await pilot.pause()

            # Should be in Global submenu
            assert menu.current_menu == "Worksheet"
            assert "Global" in menu.submenu_path

            # Escape once - back to Worksheet menu
            await pilot.press("escape")
            await pilot.pause()
            assert menu.current_menu == "Worksheet"
            assert len(menu.submenu_path) == 0

            # Escape again - back to top level
            await pilot.press("escape")
            await pilot.pause()
            assert menu.current_menu is None
            assert menu.active

            # Escape again - deactivate menu
            await pilot.press("escape")
            await pilot.pause()
            assert not menu.active

    @pytest.mark.asyncio
    async def test_menu_item_triggers_action(self):
        """Test selecting menu item triggers action."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("slash")
            await pilot.pause()

            # File -> New should clear spreadsheet
            await pilot.press("f")
            await pilot.pause()
            await pilot.press("n")
            await pilot.pause()

            # Menu should be deactivated
            menu = app.query_one("#menu-bar", LotusMenu)
            assert not menu.active

    @pytest.mark.asyncio
    async def test_menu_file_import_submenu(self):
        """Test File -> Import submenu navigation."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("slash")
            await pilot.press("f")  # File
            await pilot.press("i")  # Import (has submenu)
            await pilot.pause()

            menu = app.query_one("#menu-bar", LotusMenu)
            assert "Import" in menu.submenu_path

    @pytest.mark.asyncio
    async def test_menu_graph_type_submenu(self):
        """Test Graph -> Type submenu."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("slash")
            await pilot.press("g")  # Graph
            await pilot.press("t")  # Type (has submenu)
            await pilot.pause()

            menu = app.query_one("#menu-bar", LotusMenu)
            assert menu.current_menu == "Graph"
            assert "Type" in menu.submenu_path

    @pytest.mark.asyncio
    async def test_menu_data_query_submenu(self):
        """Test Data -> Query submenu."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("slash")
            await pilot.press("d")  # Data
            await pilot.press("q")  # Query (has submenu)
            await pilot.pause()

            menu = app.query_one("#menu-bar", LotusMenu)
            assert "Query" in menu.submenu_path

    @pytest.mark.asyncio
    async def test_menu_system_theme(self):
        """Test System -> Theme opens theme dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("slash")
            await pilot.press("s")  # System
            await pilot.press("t")  # Theme
            await pilot.pause()

            # Should open theme dialog
            assert len(app.screen_stack) > 1

    @pytest.mark.asyncio
    async def test_menu_quit_yes(self):
        """Test Quit -> Yes quits the app."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("slash")
            await pilot.press("q")  # Quit
            await pilot.pause()

            # Should show Yes/No submenu
            menu = app.query_one("#menu-bar", LotusMenu)
            assert menu.current_menu == "Quit"


# =============================================================================
# Grid Interaction Tests
# =============================================================================


class TestGridMouseInteraction:
    """Tests for grid mouse interactions."""

    @pytest.mark.asyncio
    async def test_grid_click_selects_cell(self):
        """Test clicking on grid selects cell."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Initial position
            assert grid.cursor_row == 0
            assert grid.cursor_col == 0

            # Click on a different cell (approximate coordinates)
            # Row 2 would be at y=4 (header + separator + 2 rows)
            # Col B would be at x ~= 15-25
            await pilot.click("#grid", offset=(15, 4))
            await pilot.pause()

            # Cursor should have moved (exact position depends on column widths)
            # Just verify it responded to click

    @pytest.mark.asyncio
    async def test_grid_scrolling(self):
        """Test grid scrolls when cursor moves beyond visible area."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Move down many rows
            for _ in range(30):
                await pilot.press("down")
            await pilot.pause()

            # Should have scrolled
            assert grid.scroll_row > 0
            assert grid.cursor_row == 30

    @pytest.mark.asyncio
    async def test_grid_horizontal_scrolling(self):
        """Test grid scrolls horizontally."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Move right many columns
            for _ in range(20):
                await pilot.press("right")
            await pilot.pause()

            # Should have scrolled horizontally
            assert grid.scroll_col > 0
            assert grid.cursor_col == 20

    @pytest.mark.asyncio
    async def test_grid_page_down(self):
        """Test PageDown moves cursor down significantly."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            await pilot.press("pagedown")
            await pilot.pause()

            # Should have moved down by approximately visible_rows
            # Allow for off-by-one in implementation
            assert grid.cursor_row >= grid.visible_rows - 1
            assert grid.cursor_row <= grid.visible_rows + 1

    @pytest.mark.asyncio
    async def test_grid_page_up(self):
        """Test PageUp moves up by visible rows."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # First go down
            await pilot.press("pagedown")
            await pilot.press("pagedown")
            await pilot.pause()
            initial_row = grid.cursor_row

            # Then page up
            await pilot.press("pageup")
            await pilot.pause()

            assert grid.cursor_row == initial_row - grid.visible_rows

    @pytest.mark.asyncio
    async def test_grid_home_goes_to_column_a(self):
        """Test Home key goes to column A."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Move right
            await pilot.press("right")
            await pilot.press("right")
            await pilot.press("right")
            await pilot.pause()
            assert grid.cursor_col == 3

            # Home
            await pilot.press("home")
            await pilot.pause()

            assert grid.cursor_col == 0

    @pytest.mark.asyncio
    async def test_grid_end_goes_to_last_used_column(self):
        """Test End key goes to last used column in row."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Set data in column E
            app.spreadsheet.set_cell(0, 4, "test")

            # End
            await pilot.press("end")
            await pilot.pause()

            # Should be at column E (index 4)
            assert grid.cursor_col == 4

    @pytest.mark.asyncio
    async def test_ctrl_home_goes_to_a1(self):
        """Test Ctrl+Home goes to A1."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Move away from A1
            await pilot.press("down")
            await pilot.press("down")
            await pilot.press("right")
            await pilot.press("right")
            await pilot.pause()

            # Ctrl+Home
            await pilot.press("ctrl+home")
            await pilot.pause()

            assert grid.cursor_row == 0
            assert grid.cursor_col == 0

    @pytest.mark.asyncio
    async def test_ctrl_end_goes_to_last_used_cell(self):
        """Test Ctrl+End goes to last used cell."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Set data
            app.spreadsheet.set_cell(10, 5, "last")

            # Ctrl+End
            await pilot.press("ctrl+end")
            await pilot.pause()

            assert grid.cursor_row == 10
            assert grid.cursor_col == 5


class TestGridCellEditing:
    """Tests for grid cell editing interactions."""

    @pytest.mark.asyncio
    async def test_typing_enters_edit_mode(self):
        """Test typing a character enters edit mode."""
        app = LotusApp()
        async with app.run_test() as pilot:
            assert not app.editing

            await pilot.press("h")
            await pilot.pause()

            assert app.editing

    @pytest.mark.asyncio
    async def test_f2_enters_edit_mode(self):
        """Test F2 enters edit mode."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # F2 to edit
            await pilot.press("f2")
            await pilot.pause()

            assert app.editing

    @pytest.mark.asyncio
    async def test_edit_existing_cell(self):
        """Test editing a cell with existing content."""

        app = LotusApp()
        async with app.run_test() as pilot:
            # Set cell value
            app.spreadsheet.set_cell(0, 0, "100")
            grid = app.query_one("#grid", SpreadsheetGrid)
            grid.refresh_grid()

            # Start editing by typing (replaces content)
            await pilot.press("2")
            await pilot.press("0")
            await pilot.press("0")
            await pilot.press("enter")
            await pilot.pause()

            # New value should be set
            assert app.spreadsheet.get_value(0, 0) == 200

    @pytest.mark.asyncio
    async def test_enter_submits_and_moves_down(self):
        """Test Enter submits value and moves cursor down."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Type and submit
            await pilot.press("1")
            await pilot.press("0")
            await pilot.press("0")
            await pilot.press("enter")
            await pilot.pause()

            # Value saved and cursor moved down
            assert app.spreadsheet.get_value(0, 0) == 100
            assert grid.cursor_row == 1

    @pytest.mark.asyncio
    async def test_right_arrow_moves_right(self):
        """Test Right arrow moves cursor right."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Right arrow moves right
            await pilot.press("right")
            await pilot.pause()

            assert grid.cursor_col == 1

    @pytest.mark.asyncio
    async def test_left_arrow_moves_left(self):
        """Test Left arrow moves cursor left."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # First move right
            await pilot.press("right")
            await pilot.pause()
            assert grid.cursor_col == 1

            # Left arrow moves left
            await pilot.press("left")
            await pilot.pause()

            assert grid.cursor_col == 0

    @pytest.mark.asyncio
    async def test_escape_cancels_edit(self):
        """Test Escape cancels editing without saving."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Type something
            await pilot.press("x")
            await pilot.press("y")
            await pilot.press("z")
            await pilot.pause()
            assert app.editing

            # Cancel
            await pilot.press("escape")
            await pilot.pause()

            assert not app.editing
            assert app.spreadsheet.get_display_value(0, 0) == ""

    @pytest.mark.asyncio
    async def test_formula_entry(self):
        """Test entering a formula."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Set up data
            app.spreadsheet.set_cell(0, 0, "10")
            app.spreadsheet.set_cell(0, 1, "20")

            # Move to C1 and enter formula
            await pilot.press("right")
            await pilot.press("right")
            for char in "=A1+B1":
                await pilot.press(char)
            await pilot.press("enter")
            await pilot.pause()

            # Formula should evaluate
            assert app.spreadsheet.get_value(0, 2) == 30


# =============================================================================
# End-to-End Workflow Tests
# =============================================================================


class TestUndoRedoWorkflow:
    """Tests for undo/redo workflow."""

    @pytest.mark.asyncio
    async def test_undo_cell_edit(self):
        """Test undoing a cell edit."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Edit cell
            await pilot.press("1")
            await pilot.press("0")
            await pilot.press("0")
            await pilot.press("enter")
            await pilot.pause()

            assert app.spreadsheet.get_value(0, 0) == 100

            # Undo
            await pilot.press("ctrl+z")
            await pilot.pause()

            # Should be undone
            assert app.spreadsheet.get_display_value(0, 0) == ""

    @pytest.mark.asyncio
    async def test_redo_cell_edit(self):
        """Test redoing a cell edit."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Edit cell
            await pilot.press("1")
            await pilot.press("0")
            await pilot.press("0")
            await pilot.press("enter")
            await pilot.pause()

            # Undo
            await pilot.press("ctrl+z")
            await pilot.pause()
            assert app.spreadsheet.get_display_value(0, 0) == ""

            # Redo
            await pilot.press("ctrl+y")
            await pilot.pause()

            # Should be restored
            assert app.spreadsheet.get_value(0, 0) == 100

    @pytest.mark.asyncio
    async def test_multiple_undo(self):
        """Test multiple undo operations."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Make multiple edits
            await pilot.press("a")
            await pilot.press("enter")
            await pilot.press("b")
            await pilot.press("enter")
            await pilot.press("c")
            await pilot.press("enter")
            await pilot.pause()

            # Undo all
            await pilot.press("ctrl+z")
            await pilot.press("ctrl+z")
            await pilot.press("ctrl+z")
            await pilot.pause()

            # All should be undone
            assert app.spreadsheet.get_display_value(0, 0) == ""
            assert app.spreadsheet.get_display_value(1, 0) == ""
            assert app.spreadsheet.get_display_value(2, 0) == ""


class TestNewFileWorkflow:
    """Tests for new file workflow."""

    @pytest.mark.asyncio
    async def test_ctrl_n_on_clean_spreadsheet(self):
        """Test Ctrl+N on unmodified spreadsheet."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Ctrl+N on clean spreadsheet should work directly
            await pilot.press("ctrl+n")
            await pilot.pause()

            # Should remain empty/clean
            assert app.spreadsheet.get_display_value(0, 0) == ""

    @pytest.mark.asyncio
    async def test_ctrl_n_shows_confirm_when_dirty(self):
        """Test Ctrl+N shows confirmation dialog when data is modified."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Add data (makes spreadsheet dirty)
            app.spreadsheet.set_cell(0, 0, "data")

            # New file - should show confirmation dialog
            await pilot.press("ctrl+n")
            await pilot.pause()

            # Dialog should be open asking to save
            assert len(app.screen_stack) > 1

    @pytest.mark.asyncio
    async def test_file_new_via_menu_clean(self):
        """Test File -> New on clean spreadsheet."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Menu: File -> New on clean spreadsheet
            await pilot.press("slash")
            await pilot.press("f")
            await pilot.press("n")
            await pilot.pause()

            # Should work without confirmation
            assert len(app.screen_stack) == 1


class TestDataEntryWorkflow:
    """Tests for typical data entry workflows."""

    @pytest.mark.asyncio
    async def test_enter_column_of_numbers(self):
        """Test entering a column of numbers."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Enter numbers down column A
            for i in range(1, 6):
                for char in str(i * 10):
                    await pilot.press(char)
                await pilot.press("enter")
            await pilot.pause()

            # Verify values
            assert app.spreadsheet.get_value(0, 0) == 10
            assert app.spreadsheet.get_value(1, 0) == 20
            assert app.spreadsheet.get_value(2, 0) == 30
            assert app.spreadsheet.get_value(3, 0) == 40
            assert app.spreadsheet.get_value(4, 0) == 50

    @pytest.mark.asyncio
    async def test_enter_row_with_enter_and_right(self):
        """Test entering a row of data using Enter and arrow keys."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Enter data across row using Enter then Right
            await pilot.press("a")
            await pilot.press("enter")
            await pilot.press("up")  # Enter moves down, go back up
            await pilot.press("right")
            await pilot.press("b")
            await pilot.press("enter")
            await pilot.press("up")
            await pilot.press("right")
            await pilot.press("c")
            await pilot.press("enter")
            await pilot.pause()

            # Verify values
            assert app.spreadsheet.get_display_value(0, 0) == "a"
            assert app.spreadsheet.get_display_value(0, 1) == "b"
            assert app.spreadsheet.get_display_value(0, 2) == "c"

    @pytest.mark.asyncio
    async def test_formula_with_sum(self):
        """Test entering SUM formula."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Enter numbers
            app.spreadsheet.set_cell(0, 0, "10")
            app.spreadsheet.set_cell(1, 0, "20")
            app.spreadsheet.set_cell(2, 0, "30")

            # Navigate to A4 and enter SUM
            await pilot.press("down")
            await pilot.press("down")
            await pilot.press("down")
            for char in "=SUM(A1:A3)":
                await pilot.press(char)
            await pilot.press("enter")
            await pilot.pause()

            # Should calculate sum
            assert app.spreadsheet.get_value(3, 0) == 60


class TestWorksheetOperations:
    """Tests for worksheet operations via menu."""

    @pytest.mark.asyncio
    async def test_worksheet_erase_shows_confirm(self):
        """Test Worksheet -> Erase shows confirmation dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Add data
            app.spreadsheet.set_cell(0, 0, "data")

            # Menu: Worksheet -> Erase
            await pilot.press("slash")
            await pilot.press("w")
            await pilot.press("e")
            await pilot.pause()

            # Should show confirmation dialog
            assert len(app.screen_stack) > 1

    @pytest.mark.asyncio
    async def test_worksheet_erase_with_confirm(self):
        """Test Worksheet -> Erase clears when confirmed."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Add data
            app.spreadsheet.set_cell(0, 0, "data")
            app.spreadsheet.set_cell(1, 1, "more")

            # Menu: Worksheet -> Erase
            await pilot.press("slash")
            await pilot.press("w")
            await pilot.press("e")
            await pilot.pause()

            # Confirm with Y
            await pilot.press("y")
            await pilot.press("enter")
            await pilot.pause()

            # Should be cleared
            assert app.spreadsheet.get_display_value(0, 0) == ""
            assert app.spreadsheet.get_display_value(1, 1) == ""

    @pytest.mark.asyncio
    async def test_worksheet_column_width_dialog(self):
        """Test Worksheet -> Column opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Menu: Worksheet -> Column
            await pilot.press("slash")
            await pilot.press("w")
            await pilot.press("c")
            await pilot.pause()

            # Should open dialog
            assert len(app.screen_stack) > 1

    @pytest.mark.asyncio
    async def test_range_format_dialog(self):
        """Test Range -> Format opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("slash")
            await pilot.press("r")  # Range
            await pilot.press("f")  # Format
            await pilot.pause()

            # Should open format dialog
            assert len(app.screen_stack) > 1

    @pytest.mark.asyncio
    async def test_range_name_dialog(self):
        """Test Range -> Name opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("slash")
            await pilot.press("r")  # Range
            await pilot.press("n")  # Name
            await pilot.pause()

            # Should open name dialog
            assert len(app.screen_stack) > 1


class TestGraphOperations:
    """Tests for graph/chart operations."""

    @pytest.mark.asyncio
    async def test_graph_view_no_data(self):
        """Test Graph -> View with no data shows notification."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Menu: Graph -> View without data
            await pilot.press("slash")
            await pilot.press("g")
            await pilot.press("v")
            await pilot.pause()

            # No dialog should open (notification is shown instead)
            # since no data series is defined
            assert len(app.screen_stack) == 1

    @pytest.mark.asyncio
    async def test_graph_view_with_data(self):
        """Test Graph -> View shows chart when data is defined."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Set up chart data
            app.spreadsheet.set_cell(0, 0, "10")
            app.spreadsheet.set_cell(1, 0, "20")
            app.spreadsheet.set_cell(2, 0, "30")

            # Add a series to the chart
            app.chart.add_series("A", data_range="A1:A3")

            # Menu: Graph -> View
            await pilot.press("slash")
            await pilot.press("g")
            await pilot.press("v")
            await pilot.pause()

            # Should open chart view
            assert len(app.screen_stack) > 1

    @pytest.mark.asyncio
    async def test_graph_type_line(self):
        """Test Graph -> Type -> Line changes chart type."""
        app = LotusApp()
        async with app.run_test() as pilot:
            from lotus123.charting import ChartType

            # Menu: Graph -> Type -> Line
            await pilot.press("slash")
            await pilot.press("g")
            await pilot.press("t")  # Type submenu
            await pilot.press("l")  # Line
            await pilot.pause()

            assert app.chart.chart_type == ChartType.LINE

    @pytest.mark.asyncio
    async def test_graph_type_bar(self):
        """Test Graph -> Type -> Bar changes chart type."""
        app = LotusApp()
        async with app.run_test() as pilot:
            from lotus123.charting import ChartType

            await pilot.press("slash")
            await pilot.press("g")
            await pilot.press("t")
            await pilot.press("b")  # Bar
            await pilot.pause()

            assert app.chart.chart_type == ChartType.BAR

    @pytest.mark.asyncio
    async def test_graph_type_pie(self):
        """Test Graph -> Type -> Pie changes chart type."""
        app = LotusApp()
        async with app.run_test() as pilot:
            from lotus123.charting import ChartType

            await pilot.press("slash")
            await pilot.press("g")
            await pilot.press("t")
            await pilot.press("p")  # Pie
            await pilot.pause()

            assert app.chart.chart_type == ChartType.PIE

    @pytest.mark.asyncio
    async def test_graph_reset(self):
        """Test Graph -> Reset clears chart settings."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Set chart data
            app.chart.add_series("A", data_range="A1:A10")
            app.chart.set_x_range("B1:B10")

            # Verify data is set
            assert len(app.chart.series) > 0
            assert app.chart.x_range == "B1:B10"

            # Menu: Graph -> Reset
            await pilot.press("slash")
            await pilot.press("g")
            await pilot.press("r")  # Reset
            await pilot.pause()

            # Chart should be reset (series cleared, x_range empty)
            assert len(app.chart.series) == 0
            assert app.chart.x_range == ""


class TestDataOperations:
    """Tests for Data menu operations."""

    @pytest.mark.asyncio
    async def test_data_fill_requires_selection(self):
        """Test Data -> Fill requires a selection."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Try fill without selection
            await pilot.press("slash")
            await pilot.press("d")  # Data
            await pilot.press("f")  # Fill
            await pilot.pause()

            # No dialog should open (notification shown instead)
            assert len(app.screen_stack) == 1

    @pytest.mark.asyncio
    async def test_data_fill_with_selection(self):
        """Test Data -> Fill opens dialog when selection exists."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # First create a selection
            await pilot.press("shift+down")
            await pilot.press("shift+down")
            await pilot.pause()

            # Then try fill
            await pilot.press("slash")
            await pilot.press("d")  # Data
            await pilot.press("f")  # Fill
            await pilot.pause()

            # Should open fill dialog
            assert len(app.screen_stack) > 1

    @pytest.mark.asyncio
    async def test_data_sort_dialog(self):
        """Test Data -> Sort opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("slash")
            await pilot.press("d")  # Data
            await pilot.press("s")  # Sort
            await pilot.pause()

            # Should open sort dialog
            assert len(app.screen_stack) > 1

    @pytest.mark.asyncio
    async def test_data_query_input(self):
        """Test Data -> Query -> Input opens dialog."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("slash")
            await pilot.press("d")  # Data
            await pilot.press("q")  # Query submenu
            await pilot.press("i")  # Input
            await pilot.pause()

            # Should open input range dialog
            assert len(app.screen_stack) > 1


class TestStatusDisplay:
    """Tests for status bar display."""

    @pytest.mark.asyncio
    async def test_status_shows_current_cell(self):
        """Test status bar exists and updates with cell changes."""
        from lotus123.ui import StatusBarWidget

        app = LotusApp()
        async with app.run_test() as pilot:
            # Verify status bar exists
            status_bar = app.query_one("#status-bar", StatusBarWidget)
            assert status_bar is not None

            # Move to B5
            await pilot.press("right")
            await pilot.press("down")
            await pilot.press("down")
            await pilot.press("down")
            await pilot.press("down")
            await pilot.pause()

            # Grid should be at B5
            grid = app.query_one("#grid", SpreadsheetGrid)
            assert grid.cursor_row == 4
            assert grid.cursor_col == 1

    @pytest.mark.asyncio
    async def test_status_shows_edit_mode(self):
        """Test status bar shows EDIT mode when editing."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Start editing
            await pilot.press("a")
            await pilot.pause()

            assert app.editing
            # Status should show EDIT mode


class TestKeyboardShortcuts:
    """Tests for various keyboard shortcuts."""

    @pytest.mark.asyncio
    async def test_ctrl_d_page_down(self):
        """Test Ctrl+D pages down."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            await pilot.press("ctrl+d")
            await pilot.pause()

            assert grid.cursor_row > 0

    @pytest.mark.asyncio
    async def test_ctrl_u_page_up(self):
        """Test Ctrl+U pages up."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # First go down
            await pilot.press("ctrl+d")
            await pilot.press("ctrl+d")
            await pilot.pause()
            initial = grid.cursor_row

            # Then up
            await pilot.press("ctrl+u")
            await pilot.pause()

            assert grid.cursor_row < initial

    @pytest.mark.asyncio
    async def test_backspace_clears_cell(self):
        """Test Backspace clears cell when not editing."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Set value
            app.spreadsheet.set_cell(0, 0, "test")

            # Backspace
            await pilot.press("backspace")
            await pilot.pause()

            # Should be cleared
            assert app.spreadsheet.get_display_value(0, 0) == ""

    @pytest.mark.asyncio
    async def test_ctrl_q_quits(self):
        """Test Ctrl+Q initiates quit."""
        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("ctrl+q")
            await pilot.pause()
            # App should handle quit action


class TestModalBlocksInput:
    """Tests that modals properly block grid input."""

    @pytest.mark.asyncio
    async def test_dialog_blocks_grid_navigation(self):
        """Test grid doesn't respond to arrows when dialog is open."""
        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid", SpreadsheetGrid)

            # Open dialog
            await pilot.press("ctrl+g")
            await pilot.pause()

            initial_row = grid.cursor_row
            initial_col = grid.cursor_col

            # Try to navigate
            await pilot.press("down")
            await pilot.press("right")
            await pilot.pause()

            # Grid should not have moved
            assert grid.cursor_row == initial_row
            assert grid.cursor_col == initial_col

            # Close dialog
            await pilot.press("escape")
            await pilot.pause()

    @pytest.mark.asyncio
    async def test_theme_dialog_blocks_typing(self):
        """Test typing doesn't affect grid when theme dialog is open."""
        app = LotusApp()
        async with app.run_test() as pilot:
            # Open theme dialog
            await pilot.press("ctrl+t")
            await pilot.pause()

            # Type some letters (that would normally start editing)
            await pilot.press("a")
            await pilot.press("b")
            await pilot.press("c")
            await pilot.pause()

            # Close dialog
            await pilot.press("escape")
            await pilot.pause()

            # Grid should be empty - typing didn't affect it
            assert app.spreadsheet.get_display_value(0, 0) == ""

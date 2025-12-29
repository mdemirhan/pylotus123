"""Comprehensive UI tests for Lotus 1-2-3 clone."""

import pytest


class TestLotusMenu:
    """Tests for the LotusMenu widget."""

    def test_menu_initialization(self):
        """Test menu initializes with correct state."""
        from lotus123.app import THEMES, LotusMenu, ThemeType

        theme = THEMES[ThemeType.LOTUS]
        menu = LotusMenu(theme)

        assert menu.active is False
        assert menu.current_menu is None
        assert menu.theme == theme

    def test_menu_structure(self):
        """Test menu structure contains expected items."""
        from lotus123.app import THEMES, LotusMenu, ThemeType

        theme = THEMES[ThemeType.LOTUS]
        menu = LotusMenu(theme)

        assert "Worksheet" in menu.MENU_STRUCTURE
        assert "Range" in menu.MENU_STRUCTURE
        assert "Copy" in menu.MENU_STRUCTURE
        assert "Move" in menu.MENU_STRUCTURE
        assert "File" in menu.MENU_STRUCTURE
        assert "Print" in menu.MENU_STRUCTURE
        assert "Graph" in menu.MENU_STRUCTURE
        assert "Data" in menu.MENU_STRUCTURE
        assert "Quit" in menu.MENU_STRUCTURE

    def test_menu_keys(self):
        """Test menu items have correct shortcut keys."""
        from lotus123.app import THEMES, LotusMenu, ThemeType

        theme = THEMES[ThemeType.LOTUS]
        menu = LotusMenu(theme)

        assert menu.MENU_STRUCTURE["Worksheet"]["key"] == "W"
        assert menu.MENU_STRUCTURE["Range"]["key"] == "R"
        assert menu.MENU_STRUCTURE["File"]["key"] == "F"
        assert menu.MENU_STRUCTURE["Quit"]["key"] == "Q"

    def test_file_menu_items(self):
        """Test File menu has correct sub-items."""
        from lotus123.app import THEMES, LotusMenu, ThemeType

        theme = THEMES[ThemeType.LOTUS]
        menu = LotusMenu(theme)

        file_items = menu.MENU_STRUCTURE["File"]["items"]
        item_keys = [item[0] for item in file_items]

        assert "R" in item_keys  # Retrieve
        assert "S" in item_keys  # Save
        assert "N" in item_keys  # New
        assert "Q" in item_keys  # Quit


class TestThemeSystem:
    """Tests for the theme system."""

    def test_all_themes_exist(self):
        """Test all expected themes are defined."""
        from lotus123.app import THEMES, ThemeType

        assert ThemeType.LOTUS in THEMES
        assert ThemeType.TOMORROW in THEMES
        assert ThemeType.MOCHA in THEMES

    def test_theme_has_all_colors(self):
        """Test each theme has all required color properties."""
        from lotus123.app import THEMES, ThemeType

        required_colors = [
            "name",
            "background",
            "foreground",
            "header_bg",
            "header_fg",
            "cell_bg",
            "cell_fg",
            "selected_bg",
            "selected_fg",
            "border",
            "menu_bg",
            "menu_fg",
            "menu_highlight",
            "status_bg",
            "status_fg",
            "input_bg",
            "input_fg",
            "accent",
        ]

        for theme_type in ThemeType:
            theme = THEMES[theme_type]
            for color in required_colors:
                assert hasattr(theme, color), f"{theme_type} missing {color}"

    def test_lotus_theme_colors(self):
        """Test Lotus theme has classic blue colors."""
        from lotus123.app import THEMES, ThemeType

        lotus = THEMES[ThemeType.LOTUS]
        assert lotus.name == "Lotus 1-2-3"
        assert lotus.background == "#000080"  # Classic blue

    def test_get_theme_type(self):
        """Test theme type lookup by name."""
        from lotus123.app import ThemeType, get_theme_type

        assert get_theme_type("LOTUS") == ThemeType.LOTUS
        assert get_theme_type("TOMORROW") == ThemeType.TOMORROW
        assert get_theme_type("MOCHA") == ThemeType.MOCHA
        assert get_theme_type("INVALID") == ThemeType.LOTUS  # Default


class TestAppConfig:
    """Tests for application configuration."""

    def test_config_defaults(self):
        """Test config has correct defaults."""
        from lotus123.app import AppConfig

        config = AppConfig()
        assert config.theme == "LOTUS"
        assert config.default_col_width == 10
        assert config.recent_files == []

    def test_config_serialization(self):
        """Test config can be serialized and deserialized."""
        import json
        from dataclasses import asdict

        from lotus123.app import AppConfig

        config = AppConfig(theme="MOCHA", default_col_width=15)
        data = asdict(config)
        json_str = json.dumps(data)
        loaded = json.loads(json_str)

        assert loaded["theme"] == "MOCHA"
        assert loaded["default_col_width"] == 15


class TestSpreadsheetGrid:
    """Tests for the SpreadsheetGrid widget.

    Note: Direct instantiation tests are skipped because Textual widgets
    require an active app context for reactive attributes. Grid functionality
    is tested via async app tests below.
    """

    @pytest.mark.skip(reason="SpreadsheetGrid requires app context for reactive attributes")
    def test_grid_initialization(self):
        """Test grid initializes with correct state."""
        from lotus123 import Spreadsheet
        from lotus123.app import THEMES, SpreadsheetGrid, ThemeType

        ss = Spreadsheet()
        theme = THEMES[ThemeType.LOTUS]
        grid = SpreadsheetGrid(ss, theme)

        assert grid.cursor_row == 0
        assert grid.cursor_col == 0
        assert grid.scroll_row == 0
        assert grid.scroll_col == 0

    @pytest.mark.skip(reason="SpreadsheetGrid requires app context for reactive attributes")
    def test_grid_move_cursor(self):
        """Test cursor movement."""
        from lotus123 import Spreadsheet
        from lotus123.app import THEMES, SpreadsheetGrid, ThemeType

        ss = Spreadsheet()
        theme = THEMES[ThemeType.LOTUS]
        grid = SpreadsheetGrid(ss, theme)

        grid.move_cursor(1, 0)
        assert grid.cursor_row == 1
        assert grid.cursor_col == 0

        grid.move_cursor(0, 1)
        assert grid.cursor_row == 1
        assert grid.cursor_col == 1

    @pytest.mark.skip(reason="SpreadsheetGrid requires app context for reactive attributes")
    def test_grid_cursor_bounds(self):
        """Test cursor respects grid bounds."""
        from lotus123 import Spreadsheet
        from lotus123.app import THEMES, SpreadsheetGrid, ThemeType

        ss = Spreadsheet(rows=100, cols=26)
        theme = THEMES[ThemeType.LOTUS]
        grid = SpreadsheetGrid(ss, theme)

        # Can't go negative
        grid.move_cursor(-1, 0)
        assert grid.cursor_row == 0

        grid.move_cursor(0, -1)
        assert grid.cursor_col == 0


class TestLotusAppAsync:
    """Async tests for the main Lotus app using Textual's test framework."""

    @pytest.mark.asyncio
    async def test_app_startup(self):
        """Test app starts correctly."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as _pilot:
            # App should start with grid focused
            assert app.editing is False
            assert app._menu_active is False

    @pytest.mark.asyncio
    async def test_menu_activation_with_slash(self):
        """Test menu activates with / key."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            menu = app.query_one("#menu-bar")

            assert menu.active is False

            await pilot.press("slash")
            await pilot.pause()

            assert menu.active is True

    @pytest.mark.asyncio
    async def test_menu_deactivation_with_escape(self):
        """Test menu deactivates with Escape."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            menu = app.query_one("#menu-bar")

            await pilot.press("slash")
            await pilot.pause()
            assert menu.active is True

            await pilot.press("escape")
            await pilot.pause()
            assert menu.active is False

    @pytest.mark.asyncio
    async def test_arrow_key_navigation(self):
        """Test arrow keys move cursor."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid")

            assert grid.cursor_row == 0
            assert grid.cursor_col == 0

            await pilot.press("down")
            await pilot.pause()
            assert grid.cursor_row == 1

            await pilot.press("right")
            await pilot.pause()
            assert grid.cursor_col == 1

            await pilot.press("up")
            await pilot.pause()
            assert grid.cursor_row == 0

            await pilot.press("left")
            await pilot.pause()
            assert grid.cursor_col == 0

    @pytest.mark.asyncio
    async def test_typing_starts_edit_mode(self):
        """Test typing a character starts edit mode."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            assert app.editing is False

            await pilot.press("a")
            await pilot.pause()

            assert app.editing is True

    @pytest.mark.asyncio
    async def test_typing_text_correctly(self):
        """Test typing multiple characters works correctly."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            for char in "Hello":
                await pilot.press(char)
                await pilot.pause()

            cell_input = app.query_one("#cell-input")
            assert cell_input.value == "Hello"

    @pytest.mark.asyncio
    async def test_enter_submits_cell(self):
        """Test Enter key submits cell value."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            # Type a value
            for char in "100":
                await pilot.press(char)
                await pilot.pause()

            # Press Enter to submit
            await pilot.press("enter")
            await pilot.pause()

            # Should move to next row
            grid = app.query_one("#grid")
            assert grid.cursor_row == 1

            # Value should be in spreadsheet
            assert app.spreadsheet.get_value(0, 0) == 100

    @pytest.mark.asyncio
    async def test_escape_cancels_edit(self):
        """Test Escape cancels editing."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            # Start editing
            await pilot.press("a")
            await pilot.pause()
            assert app.editing is True

            # Cancel with Escape
            await pilot.press("escape")
            await pilot.pause()
            assert app.editing is False

    @pytest.mark.asyncio
    async def test_f2_enters_edit_mode(self):
        """Test F2 enters edit mode."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            assert app.editing is False

            await pilot.press("f2")
            await pilot.pause()

            assert app.editing is True

    @pytest.mark.asyncio
    async def test_delete_clears_cell(self):
        """Test Delete key clears cell."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            # Set a value
            app.spreadsheet.set_cell(0, 0, "test")

            # Delete it
            await pilot.press("delete")
            await pilot.pause()

            assert app.spreadsheet.get_value(0, 0) == ""

    @pytest.mark.asyncio
    async def test_page_navigation(self):
        """Test PageUp/PageDown navigation."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid")

            await pilot.press("pagedown")
            await pilot.pause()

            # Should move by visible_rows
            assert grid.cursor_row > 0

    @pytest.mark.asyncio
    async def test_theme_dialog_opens(self):
        """Test theme dialog opens with Ctrl+T."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            assert not app._has_modal

            await pilot.press("ctrl+t")
            await pilot.pause()

            assert app._has_modal

    @pytest.mark.asyncio
    async def test_arrow_keys_blocked_during_modal(self):
        """Test arrow keys don't affect grid when modal is open."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            grid = app.query_one("#grid")
            initial_row = grid.cursor_row

            # Open modal
            await pilot.press("ctrl+t")
            await pilot.pause()

            # Try moving
            await pilot.press("down")
            await pilot.pause()
            await pilot.press("down")
            await pilot.pause()

            # Grid should not have moved
            assert grid.cursor_row == initial_row

            # Close modal
            await pilot.press("escape")
            await pilot.pause()

    @pytest.mark.asyncio
    async def test_goto_dialog(self):
        """Test Goto dialog with Ctrl+G."""
        from lotus123.app import LotusApp

        app = LotusApp()
        async with app.run_test() as pilot:
            await pilot.press("ctrl+g")
            await pilot.pause()

            assert app._has_modal


class TestDialogs:
    """Tests for dialog screens."""

    def test_theme_dialog_initialization(self):
        """Test ThemeDialog initializes correctly."""
        from lotus123.app import ThemeDialog, ThemeType

        dialog = ThemeDialog(current=ThemeType.LOTUS)
        assert dialog.current == ThemeType.LOTUS

    def test_command_input_initialization(self):
        """Test CommandInput initializes correctly."""
        from lotus123.app import CommandInput

        dialog = CommandInput(prompt="Enter cell:")
        assert dialog.prompt == "Enter cell:"

    def test_file_dialog_modes(self):
        """Test FileDialog supports open and save modes."""
        from lotus123.app import FileDialog

        open_dialog = FileDialog(mode="open")
        assert open_dialog.mode == "open"

        save_dialog = FileDialog(mode="save")
        assert save_dialog.mode == "save"

"""Tests for menu system module."""

import pytest

from lotus123 import Spreadsheet
from lotus123.ui.menu.menu_system import (
    Menu,
    MenuAction,
    MenuContext,
    MenuItem,
    MenuState,
    MenuSystem,
)


class TestMenuAction:
    """Tests for MenuAction enum."""

    def test_actions_exist(self):
        """Test all actions exist."""
        assert MenuAction.SUBMENU
        assert MenuAction.COMMAND
        assert MenuAction.INPUT
        assert MenuAction.RANGE
        assert MenuAction.CONFIRM


class TestMenuState:
    """Tests for MenuState enum."""

    def test_states_exist(self):
        """Test all states exist."""
        assert MenuState.INACTIVE
        assert MenuState.ACTIVE
        assert MenuState.AWAITING_INPUT
        assert MenuState.AWAITING_RANGE
        assert MenuState.AWAITING_CONFIRM


class TestMenuItem:
    """Tests for MenuItem dataclass."""

    def test_basic_item(self):
        """Test creating basic menu item."""
        item = MenuItem(key="F", label="File", action=MenuAction.SUBMENU)
        assert item.key == "F"
        assert item.label == "File"
        assert item.action == MenuAction.SUBMENU
        assert item.submenu is None
        assert item.handler is None
        assert item.help_text == ""

    def test_item_with_handler(self):
        """Test menu item with handler."""
        handler = lambda: None
        item = MenuItem(
            key="S",
            label="Save",
            action=MenuAction.COMMAND,
            handler=handler,
            help_text="Save file"
        )
        assert item.handler is handler
        assert item.help_text == "Save file"

    def test_display_key_in_label(self):
        """Test display when key is in label."""
        item = MenuItem(key="F", label="File", action=MenuAction.SUBMENU)
        display = item.display
        assert "[F]" in display

    def test_display_key_not_in_label(self):
        """Test display when key is not in label."""
        item = MenuItem(key="X", label="File", action=MenuAction.SUBMENU)
        display = item.display
        assert "[X]" in display


class TestMenu:
    """Tests for Menu dataclass."""

    def test_basic_menu(self):
        """Test creating basic menu."""
        menu = Menu(name="Test")
        assert menu.name == "Test"
        assert menu.items == []
        assert menu.parent is None

    def test_menu_with_items(self):
        """Test menu with items."""
        items = [
            MenuItem(key="A", label="Action1", action=MenuAction.COMMAND),
            MenuItem(key="B", label="Action2", action=MenuAction.COMMAND),
        ]
        menu = Menu(name="Test", items=items)
        assert len(menu.items) == 2

    def test_get_item_found(self):
        """Test getting item by key."""
        items = [
            MenuItem(key="A", label="Action1", action=MenuAction.COMMAND),
            MenuItem(key="B", label="Action2", action=MenuAction.COMMAND),
        ]
        menu = Menu(name="Test", items=items)
        item = menu.get_item("A")
        assert item is not None
        assert item.label == "Action1"

    def test_get_item_case_insensitive(self):
        """Test get_item is case-insensitive."""
        items = [MenuItem(key="A", label="Action1", action=MenuAction.COMMAND)]
        menu = Menu(name="Test", items=items)
        assert menu.get_item("a") is not None
        assert menu.get_item("A") is not None

    def test_get_item_not_found(self):
        """Test getting non-existent item."""
        menu = Menu(name="Test")
        assert menu.get_item("X") is None

    def test_get_display_line(self):
        """Test getting display line."""
        items = [
            MenuItem(key="F", label="File", action=MenuAction.SUBMENU),
            MenuItem(key="E", label="Edit", action=MenuAction.SUBMENU),
        ]
        menu = Menu(name="Test", items=items)
        line = menu.get_display_line()
        assert "File" in line
        assert "Edit" in line

    def test_get_keys(self):
        """Test getting all keys."""
        items = [
            MenuItem(key="F", label="File", action=MenuAction.SUBMENU),
            MenuItem(key="E", label="Edit", action=MenuAction.SUBMENU),
            MenuItem(key="H", label="Help", action=MenuAction.SUBMENU),
        ]
        menu = Menu(name="Test", items=items)
        keys = menu.get_keys()
        assert "F" in keys
        assert "E" in keys
        assert "H" in keys


class TestMenuContext:
    """Tests for MenuContext dataclass."""

    def test_default_values(self):
        """Test default context values."""
        ctx = MenuContext()
        assert ctx.current_menu is None
        assert ctx.menu_path == []
        assert ctx.state == MenuState.INACTIVE
        assert ctx.pending_action is None
        assert ctx.input_buffer == ""
        assert ctx.input_prompt == ""
        assert ctx.error_message == ""

    def test_get_path_string_empty(self):
        """Test path string when empty."""
        ctx = MenuContext()
        assert ctx.get_path_string() == ""

    def test_get_path_string_with_path(self):
        """Test path string with items."""
        ctx = MenuContext(menu_path=["File", "Save"])
        path = ctx.get_path_string()
        assert "File" in path
        assert "Save" in path
        assert ">" in path


class TestMenuSystem:
    """Tests for MenuSystem class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.menu_system = MenuSystem(self.ss)

    def test_init(self):
        """Test initialization."""
        assert self.menu_system.spreadsheet is self.ss
        assert self.menu_system.context is not None
        assert self.menu_system.main_menu is not None

    def test_init_without_spreadsheet(self):
        """Test initialization without spreadsheet."""
        ms = MenuSystem()
        assert ms.spreadsheet is None
        assert ms.main_menu is not None

    def test_main_menu_structure(self):
        """Test main menu has expected items."""
        main = self.menu_system.main_menu
        assert main.name == "Main"

        # Check main menu items exist
        keys = main.get_keys()
        assert "W" in keys  # Worksheet
        assert "R" in keys  # Range
        assert "C" in keys  # Copy
        assert "M" in keys  # Move
        assert "F" in keys  # File
        assert "G" in keys  # Graph
        assert "D" in keys  # Data
        assert "S" in keys  # System
        assert "Q" in keys  # Quit

    def test_worksheet_menu_exists(self):
        """Test worksheet submenu exists."""
        worksheet_item = self.menu_system.main_menu.get_item("W")
        assert worksheet_item is not None
        assert worksheet_item.submenu is not None
        assert worksheet_item.submenu.name == "Worksheet"

    def test_worksheet_submenu_items(self):
        """Test worksheet submenu has items."""
        worksheet_item = self.menu_system.main_menu.get_item("W")
        ws_menu = worksheet_item.submenu
        keys = ws_menu.get_keys()
        assert "G" in keys  # Global
        assert "I" in keys  # Insert
        assert "D" in keys  # Delete
        assert "C" in keys  # Column

    def test_range_menu_exists(self):
        """Test range submenu exists."""
        range_item = self.menu_system.main_menu.get_item("R")
        assert range_item is not None
        assert range_item.submenu is not None
        assert range_item.submenu.name == "Range"

    def test_file_menu_exists(self):
        """Test file submenu exists."""
        file_item = self.menu_system.main_menu.get_item("F")
        assert file_item is not None
        assert file_item.submenu is not None
        assert file_item.submenu.name == "File"

    def test_graph_menu_exists(self):
        """Test graph submenu exists."""
        graph_item = self.menu_system.main_menu.get_item("G")
        assert graph_item is not None
        assert graph_item.submenu is not None
        assert graph_item.submenu.name == "Graph"

    def test_data_menu_exists(self):
        """Test data submenu exists."""
        data_item = self.menu_system.main_menu.get_item("D")
        assert data_item is not None
        assert data_item.submenu is not None
        assert data_item.submenu.name == "Data"

    def test_quit_menu_exists(self):
        """Test quit submenu exists."""
        quit_item = self.menu_system.main_menu.get_item("Q")
        assert quit_item is not None
        assert quit_item.submenu is not None
        assert quit_item.submenu.name == "Quit"


class TestMenuSystemSubmenus:
    """Tests for submenu structure."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_worksheet_global_submenu(self):
        """Test worksheet global submenu structure."""
        ws_item = self.menu_system.main_menu.get_item("W")
        ws_menu = ws_item.submenu
        global_item = ws_menu.get_item("G")
        assert global_item is not None
        assert global_item.submenu is not None

        # Check global submenu items
        global_menu = global_item.submenu
        keys = global_menu.get_keys()
        assert "F" in keys  # Format
        assert "L" in keys  # Label-Prefix
        assert "C" in keys  # Column-Width
        assert "R" in keys  # Recalculation

    def test_range_format_submenu(self):
        """Test range format submenu structure."""
        range_item = self.menu_system.main_menu.get_item("R")
        range_menu = range_item.submenu
        format_item = range_menu.get_item("F")
        assert format_item is not None
        assert format_item.submenu is not None

    def test_file_submenu_items(self):
        """Test file submenu has expected items."""
        file_item = self.menu_system.main_menu.get_item("F")
        file_menu = file_item.submenu
        keys = file_menu.get_keys()
        assert "R" in keys  # Retrieve
        assert "S" in keys  # Save
        assert "C" in keys  # Combine
        assert "L" in keys  # List
        assert "I" in keys  # Import
        assert "D" in keys  # Directory

    def test_graph_submenu_items(self):
        """Test graph submenu has expected items."""
        graph_item = self.menu_system.main_menu.get_item("G")
        graph_menu = graph_item.submenu
        keys = graph_menu.get_keys()
        assert "T" in keys  # Type
        assert "X" in keys  # X-axis
        assert "A" in keys  # Data range A
        assert "V" in keys  # View
        assert "R" in keys  # Reset
        assert "S" in keys  # Save

    def test_data_submenu_items(self):
        """Test data submenu has expected items."""
        data_item = self.menu_system.main_menu.get_item("D")
        data_menu = data_item.submenu
        keys = data_menu.get_keys()
        assert "F" in keys  # Fill
        assert "T" in keys  # Table
        assert "S" in keys  # Sort
        assert "Q" in keys  # Query
        assert "D" in keys  # Distribution
        assert "M" in keys  # Matrix


class TestMenuParentLinks:
    """Tests for menu parent links."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_worksheet_parent_is_main(self):
        """Test worksheet menu parent is main."""
        ws_item = self.menu_system.main_menu.get_item("W")
        assert ws_item.submenu.parent is self.menu_system.main_menu

    def test_nested_submenu_parent(self):
        """Test nested submenu has correct parent."""
        ws_item = self.menu_system.main_menu.get_item("W")
        ws_menu = ws_item.submenu
        global_item = ws_menu.get_item("G")
        if global_item and global_item.submenu:
            assert global_item.submenu.parent is ws_menu


class TestMenuActionTypes:
    """Tests for different menu action types."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_submenu_action(self):
        """Test submenu action type."""
        ws_item = self.menu_system.main_menu.get_item("W")
        assert ws_item.action == MenuAction.SUBMENU
        assert ws_item.submenu is not None

    def test_command_action(self):
        """Test command action type."""
        # Find a command action in menus
        ws_item = self.menu_system.main_menu.get_item("W")
        ws_menu = ws_item.submenu
        erase_item = ws_menu.get_item("E")
        assert erase_item.action == MenuAction.COMMAND
        assert erase_item.handler is not None

    def test_input_action(self):
        """Test input action type."""
        file_item = self.menu_system.main_menu.get_item("F")
        file_menu = file_item.submenu
        save_item = file_menu.get_item("S")
        assert save_item.action == MenuAction.INPUT

    def test_range_action(self):
        """Test range action type."""
        range_item = self.menu_system.main_menu.get_item("R")
        range_menu = range_item.submenu
        erase_item = range_menu.get_item("E")
        assert erase_item.action == MenuAction.RANGE


class TestMenuItemHelp:
    """Tests for menu item help text."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_main_items_have_help(self):
        """Test main menu items have help text."""
        for item in self.menu_system.main_menu.items:
            assert item.help_text != "", f"Item {item.key} has no help text"

    def test_worksheet_items_have_help(self):
        """Test worksheet menu items have help text."""
        ws_item = self.menu_system.main_menu.get_item("W")
        for item in ws_item.submenu.items:
            assert item.help_text != "", f"Item {item.key} has no help text"


class TestDisplayLineFormatting:
    """Tests for display line formatting."""

    def test_display_line_key_highlighted(self):
        """Test key is highlighted in display."""
        items = [MenuItem(key="F", label="File", action=MenuAction.SUBMENU)]
        menu = Menu(name="Test", items=items)
        line = menu.get_display_line()
        # The key should appear in the line
        assert "F" in line

    def test_display_line_multiple_items(self):
        """Test multiple items are separated."""
        items = [
            MenuItem(key="A", label="Alpha", action=MenuAction.COMMAND),
            MenuItem(key="B", label="Beta", action=MenuAction.COMMAND),
        ]
        menu = Menu(name="Test", items=items)
        line = menu.get_display_line()
        assert "Alpha" in line
        assert "Beta" in line

    def test_display_key_outside_label(self):
        """Test display when key not in label uses format key:label."""
        items = [MenuItem(key="X", label="File", action=MenuAction.SUBMENU)]
        menu = Menu(name="Test", items=items)
        line = menu.get_display_line()
        assert "X" in line
        assert "File" in line


class TestMenuSystemCommands:
    """Tests for menu system command handlers."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.menu_system = MenuSystem(self.ss)

    def test_cmd_worksheet_erase(self):
        """Test worksheet erase command."""
        result = self.menu_system._cmd_worksheet_erase()
        assert isinstance(result, str)

    def test_cmd_worksheet_status(self):
        """Test worksheet status command."""
        result = self.menu_system._cmd_worksheet_status()
        assert isinstance(result, str)

    def test_cmd_worksheet_page(self):
        """Test worksheet page command."""
        result = self.menu_system._cmd_worksheet_page()
        assert isinstance(result, str)

    def test_cmd_insert_row(self):
        """Test insert row command."""
        result = self.menu_system._cmd_insert_row()
        assert isinstance(result, str)

    def test_cmd_insert_column(self):
        """Test insert column command."""
        result = self.menu_system._cmd_insert_column()
        assert isinstance(result, str)

    def test_cmd_delete_row(self):
        """Test delete row command."""
        result = self.menu_system._cmd_delete_row()
        assert isinstance(result, str)

    def test_cmd_delete_column(self):
        """Test delete column command."""
        result = self.menu_system._cmd_delete_column()
        assert isinstance(result, str)

    def test_cmd_column_setwidth(self):
        """Test column set width command."""
        result = self.menu_system._cmd_column_setwidth("15")
        assert isinstance(result, str)

    def test_cmd_column_resetwidth(self):
        """Test column reset width command."""
        result = self.menu_system._cmd_column_resetwidth()
        assert isinstance(result, str)

    def test_cmd_column_hide(self):
        """Test column hide command."""
        result = self.menu_system._cmd_column_hide()
        assert isinstance(result, str)

    def test_cmd_column_display(self):
        """Test column display command."""
        result = self.menu_system._cmd_column_display()
        assert isinstance(result, str)


class TestMenuSystemTitlesCommands:
    """Tests for titles commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_titles_both(self):
        """Test titles both command."""
        result = self.menu_system._cmd_titles_both()
        assert isinstance(result, str)

    def test_cmd_titles_horizontal(self):
        """Test titles horizontal command."""
        result = self.menu_system._cmd_titles_horizontal()
        assert isinstance(result, str)

    def test_cmd_titles_vertical(self):
        """Test titles vertical command."""
        result = self.menu_system._cmd_titles_vertical()
        assert isinstance(result, str)

    def test_cmd_titles_clear(self):
        """Test titles clear command."""
        result = self.menu_system._cmd_titles_clear()
        assert isinstance(result, str)


class TestMenuSystemWindowCommands:
    """Tests for window commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_window_horizontal(self):
        """Test window horizontal command."""
        result = self.menu_system._cmd_window_horizontal()
        assert isinstance(result, str)

    def test_cmd_window_vertical(self):
        """Test window vertical command."""
        result = self.menu_system._cmd_window_vertical()
        assert isinstance(result, str)

    def test_cmd_window_sync(self):
        """Test window sync command."""
        result = self.menu_system._cmd_window_sync()
        assert isinstance(result, str)

    def test_cmd_window_unsync(self):
        """Test window unsync command."""
        result = self.menu_system._cmd_window_unsync()
        assert isinstance(result, str)

    def test_cmd_window_clear(self):
        """Test window clear command."""
        result = self.menu_system._cmd_window_clear()
        assert isinstance(result, str)


class TestMenuSystemFormatCommands:
    """Tests for format commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_format_fixed(self):
        """Test format fixed command."""
        result = self.menu_system._cmd_format_fixed("2")
        assert isinstance(result, str)

    def test_cmd_format_scientific(self):
        """Test format scientific command."""
        result = self.menu_system._cmd_format_scientific("2")
        assert isinstance(result, str)

    def test_cmd_format_currency(self):
        """Test format currency command."""
        result = self.menu_system._cmd_format_currency("2")
        assert isinstance(result, str)

    def test_cmd_format_comma(self):
        """Test format comma command."""
        result = self.menu_system._cmd_format_comma("2")
        assert isinstance(result, str)

    def test_cmd_format_general(self):
        """Test format general command."""
        result = self.menu_system._cmd_format_general()
        assert isinstance(result, str)

    def test_cmd_format_plusminus(self):
        """Test format plus/minus command."""
        result = self.menu_system._cmd_format_plusminus()
        assert isinstance(result, str)

    def test_cmd_format_percent(self):
        """Test format percent command."""
        result = self.menu_system._cmd_format_percent("2")
        assert isinstance(result, str)

    def test_cmd_date_format(self):
        """Test date format command."""
        result = self.menu_system._cmd_date_format(1)
        assert isinstance(result, str)

    def test_cmd_format_text(self):
        """Test format text command."""
        result = self.menu_system._cmd_format_text()
        assert isinstance(result, str)

    def test_cmd_format_hidden(self):
        """Test format hidden command."""
        result = self.menu_system._cmd_format_hidden()
        assert isinstance(result, str)

    def test_cmd_format_reset(self):
        """Test format reset command."""
        result = self.menu_system._cmd_format_reset()
        assert isinstance(result, str)


class TestMenuSystemLabelCommands:
    """Tests for label commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_label_left(self):
        """Test label left command."""
        result = self.menu_system._cmd_label_left()
        assert isinstance(result, str)

    def test_cmd_label_right(self):
        """Test label right command."""
        result = self.menu_system._cmd_label_right()
        assert isinstance(result, str)

    def test_cmd_label_center(self):
        """Test label center command."""
        result = self.menu_system._cmd_label_center()
        assert isinstance(result, str)


class TestMenuSystemRecalcCommands:
    """Tests for recalculation commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_recalc_natural(self):
        """Test recalc natural command."""
        result = self.menu_system._cmd_recalc_natural()
        assert isinstance(result, str)

    def test_cmd_recalc_columnwise(self):
        """Test recalc columnwise command."""
        result = self.menu_system._cmd_recalc_columnwise()
        assert isinstance(result, str)

    def test_cmd_recalc_rowwise(self):
        """Test recalc rowwise command."""
        result = self.menu_system._cmd_recalc_rowwise()
        assert isinstance(result, str)

    def test_cmd_recalc_automatic(self):
        """Test recalc automatic command."""
        result = self.menu_system._cmd_recalc_automatic()
        assert isinstance(result, str)

    def test_cmd_recalc_manual(self):
        """Test recalc manual command."""
        result = self.menu_system._cmd_recalc_manual()
        assert isinstance(result, str)

    def test_cmd_recalc_iteration(self):
        """Test recalc iteration command."""
        result = self.menu_system._cmd_recalc_iteration("10")
        assert isinstance(result, str)


class TestMenuSystemDefaultCommands:
    """Tests for default commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_default_directory(self):
        """Test default directory command."""
        result = self.menu_system._cmd_default_directory("/tmp")
        assert isinstance(result, str)

    def test_cmd_default_status(self):
        """Test default status command."""
        result = self.menu_system._cmd_default_status()
        assert isinstance(result, str)

    def test_cmd_default_update(self):
        """Test default update command."""
        result = self.menu_system._cmd_default_update()
        assert isinstance(result, str)


class TestMenuSystemZeroCommands:
    """Tests for zero display commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_zero_yes(self):
        """Test zero yes command."""
        result = self.menu_system._cmd_zero_yes()
        assert isinstance(result, str)

    def test_cmd_zero_no(self):
        """Test zero no command."""
        result = self.menu_system._cmd_zero_no()
        assert isinstance(result, str)

    def test_cmd_zero_label(self):
        """Test zero label command."""
        result = self.menu_system._cmd_zero_label("N/A")
        assert isinstance(result, str)


class TestMenuSystemRangeCommands:
    """Tests for range commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_range_erase(self):
        """Test range erase command."""
        result = self.menu_system._cmd_range_erase("A1:B10")
        assert isinstance(result, str)

    def test_cmd_range_justify(self):
        """Test range justify command."""
        result = self.menu_system._cmd_range_justify("A1:B10")
        assert isinstance(result, str)

    def test_cmd_range_input(self):
        """Test range input command."""
        result = self.menu_system._cmd_range_input("A1:B10")
        assert isinstance(result, str)

    def test_cmd_range_value(self):
        """Test range value command."""
        result = self.menu_system._cmd_range_value("A1:B10")
        assert isinstance(result, str)

    def test_cmd_range_transpose(self):
        """Test range transpose command."""
        result = self.menu_system._cmd_range_transpose("A1:B10")
        assert isinstance(result, str)


class TestMenuSystemNameCommands:
    """Tests for name commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_name_create(self):
        """Test name create command."""
        result = self.menu_system._cmd_name_create("MyRange")
        assert isinstance(result, str)

    def test_cmd_name_delete(self):
        """Test name delete command."""
        result = self.menu_system._cmd_name_delete("MyRange")
        assert isinstance(result, str)

    def test_cmd_name_reset(self):
        """Test name reset command."""
        result = self.menu_system._cmd_name_reset()
        assert isinstance(result, str)

    def test_cmd_name_table(self):
        """Test name table command."""
        result = self.menu_system._cmd_name_table()
        assert isinstance(result, str)


class TestMenuSystemCopyMoveCommands:
    """Tests for copy/move commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_copy(self):
        """Test copy command."""
        result = self.menu_system._cmd_copy("A1:B10")
        assert isinstance(result, str)

    def test_cmd_move(self):
        """Test move command."""
        result = self.menu_system._cmd_move("A1:B10")
        assert isinstance(result, str)


class TestMenuSystemFileCommands:
    """Tests for file commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_file_retrieve(self):
        """Test file retrieve command."""
        result = self.menu_system._cmd_file_retrieve("test.wk1")
        assert isinstance(result, str)

    def test_cmd_file_save(self):
        """Test file save command."""
        result = self.menu_system._cmd_file_save("test.wk1")
        assert isinstance(result, str)

    def test_cmd_file_erase(self):
        """Test file erase command."""
        result = self.menu_system._cmd_file_erase("test.wk1")
        assert isinstance(result, str)

    def test_cmd_file_list(self):
        """Test file list command."""
        result = self.menu_system._cmd_file_list()
        assert isinstance(result, str)

    def test_cmd_file_directory(self):
        """Test file directory command."""
        result = self.menu_system._cmd_file_directory("/tmp")
        assert isinstance(result, str)


class TestMenuSystemGraphCommands:
    """Tests for graph commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_graph_type_line(self):
        """Test graph type line command."""
        result = self.menu_system._cmd_graph_type_line()
        assert isinstance(result, str)

    def test_cmd_graph_type_bar(self):
        """Test graph type bar command."""
        result = self.menu_system._cmd_graph_type_bar()
        assert isinstance(result, str)

    def test_cmd_graph_type_xy(self):
        """Test graph type XY command."""
        result = self.menu_system._cmd_graph_type_xy()
        assert isinstance(result, str)

    def test_cmd_graph_type_stacked(self):
        """Test graph type stacked command."""
        result = self.menu_system._cmd_graph_type_stacked()
        assert isinstance(result, str)

    def test_cmd_graph_type_pie(self):
        """Test graph type pie command."""
        result = self.menu_system._cmd_graph_type_pie()
        assert isinstance(result, str)

    def test_cmd_graph_x(self):
        """Test graph X range command."""
        result = self.menu_system._cmd_graph_x("A1:A10")
        assert isinstance(result, str)

    def test_cmd_graph_a(self):
        """Test graph A range command."""
        result = self.menu_system._cmd_graph_a("B1:B10")
        assert isinstance(result, str)

    def test_cmd_graph_reset(self):
        """Test graph reset command."""
        result = self.menu_system._cmd_graph_reset()
        assert isinstance(result, str)

    def test_cmd_graph_view(self):
        """Test graph view command."""
        result = self.menu_system._cmd_graph_view()
        assert isinstance(result, str)

    def test_cmd_graph_save(self):
        """Test graph save command."""
        result = self.menu_system._cmd_graph_save("chart.png")
        assert isinstance(result, str)


class TestMenuSystemDataCommands:
    """Tests for data commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_data_fill(self):
        """Test data fill command."""
        result = self.menu_system._cmd_data_fill("A1:A10")
        assert isinstance(result, str)

    def test_cmd_data_table1(self):
        """Test data table 1 command."""
        result = self.menu_system._cmd_data_table1("A1:A10")
        assert isinstance(result, str)

    def test_cmd_data_table_reset(self):
        """Test data table reset command."""
        result = self.menu_system._cmd_data_table_reset()
        assert isinstance(result, str)

    def test_cmd_sort_data_range(self):
        """Test sort data range command."""
        result = self.menu_system._cmd_sort_data_range("A1:C10")
        assert isinstance(result, str)

    def test_cmd_sort_primary_key(self):
        """Test sort primary key command."""
        result = self.menu_system._cmd_sort_primary_key("A1")
        assert isinstance(result, str)

    def test_cmd_sort_reset(self):
        """Test sort reset command."""
        result = self.menu_system._cmd_sort_reset()
        assert isinstance(result, str)

    def test_cmd_sort_go(self):
        """Test sort go command."""
        result = self.menu_system._cmd_sort_go()
        assert isinstance(result, str)


class TestMenuSystemQueryCommands:
    """Tests for query commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_query_input(self):
        """Test query input command."""
        result = self.menu_system._cmd_query_input("A1:C10")
        assert isinstance(result, str)

    def test_cmd_query_criterion(self):
        """Test query criterion command."""
        result = self.menu_system._cmd_query_criterion("E1:F2")
        assert isinstance(result, str)

    def test_cmd_query_output(self):
        """Test query output command."""
        result = self.menu_system._cmd_query_output("H1")
        assert isinstance(result, str)

    def test_cmd_query_find(self):
        """Test query find command."""
        result = self.menu_system._cmd_query_find()
        assert isinstance(result, str)

    def test_cmd_query_extract(self):
        """Test query extract command."""
        result = self.menu_system._cmd_query_extract()
        assert isinstance(result, str)

    def test_cmd_query_unique(self):
        """Test query unique command."""
        result = self.menu_system._cmd_query_unique()
        assert isinstance(result, str)

    def test_cmd_query_delete(self):
        """Test query delete command."""
        result = self.menu_system._cmd_query_delete()
        assert isinstance(result, str)

    def test_cmd_query_reset(self):
        """Test query reset command."""
        result = self.menu_system._cmd_query_reset()
        assert isinstance(result, str)


class TestMenuSystemSystemCommands:
    """Tests for system commands."""

    def setup_method(self):
        """Set up test fixtures."""
        self.menu_system = MenuSystem()

    def test_cmd_system(self):
        """Test system command."""
        result = self.menu_system._cmd_system()
        assert isinstance(result, str)

    def test_cmd_quit_yes(self):
        """Test quit yes command."""
        result = self.menu_system._cmd_quit_yes()
        assert isinstance(result, str)

    def test_cmd_quit_no(self):
        """Test quit no command."""
        result = self.menu_system._cmd_quit_no()
        assert isinstance(result, str)

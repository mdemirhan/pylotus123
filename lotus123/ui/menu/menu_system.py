"""Lotus 1-2-3 style hierarchical menu system.

The menu system uses single-letter shortcuts and hierarchical navigation:
- / (slash) activates the main menu
- First letter of each option selects it
- Escape returns to previous menu or cancels
- Enter confirms selection
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Callable, Any, TYPE_CHECKING

if TYPE_CHECKING:
    from ...core.spreadsheet import Spreadsheet


class MenuAction(Enum):
    """Types of menu actions."""
    SUBMENU = auto()      # Opens a submenu
    COMMAND = auto()      # Executes a command
    INPUT = auto()        # Requires text input
    RANGE = auto()        # Requires range input
    CONFIRM = auto()      # Requires yes/no confirmation


@dataclass
class MenuItem:
    """Single menu item."""
    key: str              # Single letter shortcut
    label: str            # Display label
    action: MenuAction    # Type of action
    submenu: Menu | None = None  # Submenu if action is SUBMENU
    handler: Callable[..., Any] | None = None  # Handler for commands
    help_text: str = ""   # Brief help description

    @property
    def display(self) -> str:
        """Get display string with highlighted shortcut."""
        if self.key.upper() in self.label.upper():
            idx = self.label.upper().index(self.key.upper())
            return f"{self.label[:idx]}[{self.label[idx]}]{self.label[idx+1:]}"
        return f"[{self.key}] {self.label}"


@dataclass
class Menu:
    """A menu with multiple items."""
    name: str
    items: list[MenuItem] = field(default_factory=list)
    parent: Menu | None = None

    def get_item(self, key: str) -> MenuItem | None:
        """Get item by its shortcut key."""
        key = key.upper()
        for item in self.items:
            if item.key.upper() == key:
                return item
        return None

    def get_display_line(self) -> str:
        """Get single-line display of menu items."""
        parts = []
        for item in self.items:
            # Highlight the shortcut letter
            label = item.label
            key = item.key.upper()
            if key in label.upper():
                idx = label.upper().index(key)
                parts.append(f"{label[:idx]}{label[idx]}{label[idx+1:]}")
            else:
                parts.append(f"{item.key}:{label}")
        return "  ".join(parts)

    def get_keys(self) -> str:
        """Get all valid shortcut keys."""
        return "".join(item.key.upper() for item in self.items)


class MenuState(Enum):
    """Current state of menu system."""
    INACTIVE = auto()     # Menu not active
    ACTIVE = auto()       # Menu is displayed
    AWAITING_INPUT = auto()  # Waiting for text input
    AWAITING_RANGE = auto()  # Waiting for range selection
    AWAITING_CONFIRM = auto()  # Waiting for Y/N


@dataclass
class MenuContext:
    """Context for menu operations."""
    current_menu: Menu | None = None
    menu_path: list[str] = field(default_factory=list)
    state: MenuState = MenuState.INACTIVE
    pending_action: MenuItem | None = None
    input_buffer: str = ""
    input_prompt: str = ""
    error_message: str = ""

    def get_path_string(self) -> str:
        """Get menu path as string."""
        return " > ".join(self.menu_path) if self.menu_path else ""


class MenuSystem:
    """Lotus 1-2-3 style menu system manager."""

    def __init__(self, spreadsheet: Spreadsheet = None):
        self.spreadsheet = spreadsheet
        self.context = MenuContext()
        self.main_menu = self._build_main_menu()
        self.handlers: dict[str, Callable] = {}

    def _build_main_menu(self) -> Menu:
        """Build the main menu structure."""
        main = Menu("Main")

        # Worksheet menu
        worksheet = Menu("Worksheet")
        worksheet.items = [
            MenuItem("G", "Global", MenuAction.SUBMENU,
                    submenu=self._build_worksheet_global_menu(),
                    help_text="Global worksheet settings"),
            MenuItem("I", "Insert", MenuAction.SUBMENU,
                    submenu=self._build_insert_menu(),
                    help_text="Insert rows/columns"),
            MenuItem("D", "Delete", MenuAction.SUBMENU,
                    submenu=self._build_delete_menu(),
                    help_text="Delete rows/columns"),
            MenuItem("C", "Column", MenuAction.SUBMENU,
                    submenu=self._build_column_menu(),
                    help_text="Column operations"),
            MenuItem("E", "Erase", MenuAction.COMMAND,
                    handler=self._cmd_worksheet_erase,
                    help_text="Erase worksheet"),
            MenuItem("T", "Titles", MenuAction.SUBMENU,
                    submenu=self._build_titles_menu(),
                    help_text="Freeze titles"),
            MenuItem("W", "Window", MenuAction.SUBMENU,
                    submenu=self._build_window_menu(),
                    help_text="Split window"),
            MenuItem("S", "Status", MenuAction.COMMAND,
                    handler=self._cmd_worksheet_status,
                    help_text="Display status"),
            MenuItem("P", "Page", MenuAction.COMMAND,
                    handler=self._cmd_worksheet_page,
                    help_text="Page break"),
        ]
        worksheet.parent = main
        for item in worksheet.items:
            if item.submenu:
                item.submenu.parent = worksheet

        # Range menu
        range_menu = Menu("Range")
        range_menu.items = [
            MenuItem("F", "Format", MenuAction.SUBMENU,
                    submenu=self._build_range_format_menu(),
                    help_text="Format cells"),
            MenuItem("L", "Label", MenuAction.SUBMENU,
                    submenu=self._build_label_menu(),
                    help_text="Label alignment"),
            MenuItem("E", "Erase", MenuAction.RANGE,
                    handler=self._cmd_range_erase,
                    help_text="Erase range"),
            MenuItem("N", "Name", MenuAction.SUBMENU,
                    submenu=self._build_name_menu(),
                    help_text="Named ranges"),
            MenuItem("J", "Justify", MenuAction.RANGE,
                    handler=self._cmd_range_justify,
                    help_text="Justify text"),
            MenuItem("P", "Protect", MenuAction.SUBMENU,
                    submenu=self._build_protect_menu(),
                    help_text="Cell protection"),
            MenuItem("U", "Unprotect", MenuAction.RANGE,
                    handler=self._cmd_range_unprotect,
                    help_text="Unprotect cells"),
            MenuItem("I", "Input", MenuAction.RANGE,
                    handler=self._cmd_range_input,
                    help_text="Data entry range"),
            MenuItem("V", "Value", MenuAction.RANGE,
                    handler=self._cmd_range_value,
                    help_text="Convert formulas to values"),
            MenuItem("T", "Transpose", MenuAction.RANGE,
                    handler=self._cmd_range_transpose,
                    help_text="Transpose range"),
        ]
        range_menu.parent = main
        for item in range_menu.items:
            if item.submenu:
                item.submenu.parent = range_menu

        # Copy menu
        copy_menu = Menu("Copy")
        copy_menu.items = [
            MenuItem("C", "Copy", MenuAction.RANGE,
                    handler=self._cmd_copy,
                    help_text="Copy range"),
        ]
        copy_menu.parent = main

        # Move menu
        move_menu = Menu("Move")
        move_menu.items = [
            MenuItem("M", "Move", MenuAction.RANGE,
                    handler=self._cmd_move,
                    help_text="Move range"),
        ]
        move_menu.parent = main

        # File menu
        file_menu = Menu("File")
        file_menu.items = [
            MenuItem("R", "Retrieve", MenuAction.INPUT,
                    handler=self._cmd_file_retrieve,
                    help_text="Load file"),
            MenuItem("S", "Save", MenuAction.INPUT,
                    handler=self._cmd_file_save,
                    help_text="Save file"),
            MenuItem("C", "Combine", MenuAction.SUBMENU,
                    submenu=self._build_combine_menu(),
                    help_text="Combine files"),
            MenuItem("X", "Xtract", MenuAction.SUBMENU,
                    submenu=self._build_extract_menu(),
                    help_text="Extract to file"),
            MenuItem("E", "Erase", MenuAction.INPUT,
                    handler=self._cmd_file_erase,
                    help_text="Delete file"),
            MenuItem("L", "List", MenuAction.COMMAND,
                    handler=self._cmd_file_list,
                    help_text="List files"),
            MenuItem("I", "Import", MenuAction.SUBMENU,
                    submenu=self._build_import_menu(),
                    help_text="Import text"),
            MenuItem("D", "Directory", MenuAction.INPUT,
                    handler=self._cmd_file_directory,
                    help_text="Change directory"),
        ]
        file_menu.parent = main
        for item in file_menu.items:
            if item.submenu:
                item.submenu.parent = file_menu

        # Print menu
        print_menu = Menu("Print")
        print_menu.items = [
            MenuItem("P", "Printer", MenuAction.SUBMENU,
                    submenu=self._build_printer_menu(),
                    help_text="Print to printer"),
            MenuItem("F", "File", MenuAction.INPUT,
                    handler=self._cmd_print_file,
                    help_text="Print to file"),
        ]
        print_menu.parent = main
        for item in print_menu.items:
            if item.submenu:
                item.submenu.parent = print_menu

        # Graph menu
        graph_menu = Menu("Graph")
        graph_menu.items = [
            MenuItem("T", "Type", MenuAction.SUBMENU,
                    submenu=self._build_graph_type_menu(),
                    help_text="Chart type"),
            MenuItem("X", "X", MenuAction.RANGE,
                    handler=self._cmd_graph_x,
                    help_text="X-axis range"),
            MenuItem("A", "A", MenuAction.RANGE,
                    handler=self._cmd_graph_a,
                    help_text="First data range"),
            MenuItem("B", "B", MenuAction.RANGE,
                    handler=self._cmd_graph_b,
                    help_text="Second data range"),
            MenuItem("C", "C", MenuAction.RANGE,
                    handler=self._cmd_graph_c,
                    help_text="Third data range"),
            MenuItem("D", "D", MenuAction.RANGE,
                    handler=self._cmd_graph_d,
                    help_text="Fourth data range"),
            MenuItem("E", "E", MenuAction.RANGE,
                    handler=self._cmd_graph_e,
                    help_text="Fifth data range"),
            MenuItem("F", "F", MenuAction.RANGE,
                    handler=self._cmd_graph_f,
                    help_text="Sixth data range"),
            MenuItem("R", "Reset", MenuAction.COMMAND,
                    handler=self._cmd_graph_reset,
                    help_text="Reset graph"),
            MenuItem("V", "View", MenuAction.COMMAND,
                    handler=self._cmd_graph_view,
                    help_text="View graph"),
            MenuItem("S", "Save", MenuAction.INPUT,
                    handler=self._cmd_graph_save,
                    help_text="Save graph"),
            MenuItem("O", "Options", MenuAction.SUBMENU,
                    submenu=self._build_graph_options_menu(),
                    help_text="Graph options"),
            MenuItem("N", "Name", MenuAction.SUBMENU,
                    submenu=self._build_graph_name_menu(),
                    help_text="Named graphs"),
        ]
        graph_menu.parent = main
        for item in graph_menu.items:
            if item.submenu:
                item.submenu.parent = graph_menu

        # Data menu
        data_menu = Menu("Data")
        data_menu.items = [
            MenuItem("F", "Fill", MenuAction.RANGE,
                    handler=self._cmd_data_fill,
                    help_text="Fill range"),
            MenuItem("T", "Table", MenuAction.SUBMENU,
                    submenu=self._build_data_table_menu(),
                    help_text="Data table"),
            MenuItem("S", "Sort", MenuAction.SUBMENU,
                    submenu=self._build_data_sort_menu(),
                    help_text="Sort data"),
            MenuItem("Q", "Query", MenuAction.SUBMENU,
                    submenu=self._build_data_query_menu(),
                    help_text="Database query"),
            MenuItem("D", "Distribution", MenuAction.RANGE,
                    handler=self._cmd_data_distribution,
                    help_text="Frequency distribution"),
            MenuItem("M", "Matrix", MenuAction.SUBMENU,
                    submenu=self._build_data_matrix_menu(),
                    help_text="Matrix operations"),
            MenuItem("R", "Regression", MenuAction.COMMAND,
                    handler=self._cmd_data_regression,
                    help_text="Regression analysis"),
            MenuItem("P", "Parse", MenuAction.RANGE,
                    handler=self._cmd_data_parse,
                    help_text="Parse strings"),
        ]
        data_menu.parent = main
        for item in data_menu.items:
            if item.submenu:
                item.submenu.parent = data_menu

        # System menu
        system_menu = Menu("System")
        system_menu.items = [
            MenuItem("S", "System", MenuAction.COMMAND,
                    handler=self._cmd_system,
                    help_text="OS shell"),
        ]
        system_menu.parent = main

        # Quit menu
        quit_menu = Menu("Quit")
        quit_menu.items = [
            MenuItem("Y", "Yes", MenuAction.COMMAND,
                    handler=self._cmd_quit_yes,
                    help_text="Quit and save"),
            MenuItem("N", "No", MenuAction.COMMAND,
                    handler=self._cmd_quit_no,
                    help_text="Cancel quit"),
        ]
        quit_menu.parent = main

        # Build main menu
        main.items = [
            MenuItem("W", "Worksheet", MenuAction.SUBMENU, submenu=worksheet,
                    help_text="Worksheet operations"),
            MenuItem("R", "Range", MenuAction.SUBMENU, submenu=range_menu,
                    help_text="Range operations"),
            MenuItem("C", "Copy", MenuAction.SUBMENU, submenu=copy_menu,
                    help_text="Copy cells"),
            MenuItem("M", "Move", MenuAction.SUBMENU, submenu=move_menu,
                    help_text="Move cells"),
            MenuItem("F", "File", MenuAction.SUBMENU, submenu=file_menu,
                    help_text="File operations"),
            MenuItem("P", "Print", MenuAction.SUBMENU, submenu=print_menu,
                    help_text="Print operations"),
            MenuItem("G", "Graph", MenuAction.SUBMENU, submenu=graph_menu,
                    help_text="Chart/Graph operations"),
            MenuItem("D", "Data", MenuAction.SUBMENU, submenu=data_menu,
                    help_text="Database operations"),
            MenuItem("S", "System", MenuAction.SUBMENU, submenu=system_menu,
                    help_text="System commands"),
            MenuItem("Q", "Quit", MenuAction.SUBMENU, submenu=quit_menu,
                    help_text="Exit program"),
        ]

        return main

    # Submenu builders
    def _build_worksheet_global_menu(self) -> Menu:
        """Build Worksheet Global submenu."""
        menu = Menu("Global")
        menu.items = [
            MenuItem("F", "Format", MenuAction.SUBMENU,
                    submenu=self._build_global_format_menu(),
                    help_text="Default format"),
            MenuItem("L", "Label-Prefix", MenuAction.SUBMENU,
                    submenu=self._build_label_prefix_menu(),
                    help_text="Default label prefix"),
            MenuItem("C", "Column-Width", MenuAction.INPUT,
                    handler=self._cmd_global_colwidth,
                    help_text="Default column width"),
            MenuItem("R", "Recalculation", MenuAction.SUBMENU,
                    submenu=self._build_recalc_menu(),
                    help_text="Recalculation settings"),
            MenuItem("P", "Protection", MenuAction.SUBMENU,
                    submenu=self._build_global_protection_menu(),
                    help_text="Worksheet protection"),
            MenuItem("D", "Default", MenuAction.SUBMENU,
                    submenu=self._build_default_menu(),
                    help_text="Default settings"),
            MenuItem("Z", "Zero", MenuAction.SUBMENU,
                    submenu=self._build_zero_menu(),
                    help_text="Zero display"),
        ]
        return menu

    def _build_global_format_menu(self) -> Menu:
        """Build format submenu for global settings."""
        return self._build_format_submenu()

    def _build_format_submenu(self) -> Menu:
        """Build standard format submenu."""
        menu = Menu("Format")
        menu.items = [
            MenuItem("F", "Fixed", MenuAction.INPUT,
                    handler=self._cmd_format_fixed, help_text="Fixed decimals"),
            MenuItem("S", "Scientific", MenuAction.INPUT,
                    handler=self._cmd_format_scientific, help_text="Scientific notation"),
            MenuItem("C", "Currency", MenuAction.INPUT,
                    handler=self._cmd_format_currency, help_text="Currency format"),
            MenuItem(",", ",", MenuAction.INPUT,
                    handler=self._cmd_format_comma, help_text="Comma format"),
            MenuItem("G", "General", MenuAction.COMMAND,
                    handler=self._cmd_format_general, help_text="General format"),
            MenuItem("+", "+/-", MenuAction.COMMAND,
                    handler=self._cmd_format_plusminus, help_text="Bar graph format"),
            MenuItem("P", "Percent", MenuAction.INPUT,
                    handler=self._cmd_format_percent, help_text="Percentage format"),
            MenuItem("D", "Date", MenuAction.SUBMENU,
                    submenu=self._build_date_format_menu(), help_text="Date format"),
            MenuItem("T", "Text", MenuAction.COMMAND,
                    handler=self._cmd_format_text, help_text="Text format"),
            MenuItem("H", "Hidden", MenuAction.COMMAND,
                    handler=self._cmd_format_hidden, help_text="Hidden format"),
            MenuItem("R", "Reset", MenuAction.COMMAND,
                    handler=self._cmd_format_reset, help_text="Reset to default"),
        ]
        return menu

    def _build_date_format_menu(self) -> Menu:
        """Build date format submenu."""
        menu = Menu("Date")
        menu.items = [
            MenuItem("1", "DD-MMM-YY", MenuAction.COMMAND,
                    handler=lambda: self._cmd_date_format(1), help_text="26-Dec-91"),
            MenuItem("2", "DD-MMM", MenuAction.COMMAND,
                    handler=lambda: self._cmd_date_format(2), help_text="26-Dec"),
            MenuItem("3", "MMM-YY", MenuAction.COMMAND,
                    handler=lambda: self._cmd_date_format(3), help_text="Dec-91"),
            MenuItem("4", "MM/DD/YY", MenuAction.COMMAND,
                    handler=lambda: self._cmd_date_format(4), help_text="12/26/91"),
            MenuItem("5", "MM/DD", MenuAction.COMMAND,
                    handler=lambda: self._cmd_date_format(5), help_text="12/26"),
        ]
        return menu

    def _build_label_prefix_menu(self) -> Menu:
        """Build label prefix submenu."""
        menu = Menu("Label-Prefix")
        menu.items = [
            MenuItem("L", "Left", MenuAction.COMMAND,
                    handler=self._cmd_label_left, help_text="Left align ('prefix)"),
            MenuItem("R", "Right", MenuAction.COMMAND,
                    handler=self._cmd_label_right, help_text="Right align (\"prefix)"),
            MenuItem("C", "Center", MenuAction.COMMAND,
                    handler=self._cmd_label_center, help_text="Center (^prefix)"),
        ]
        return menu

    def _build_recalc_menu(self) -> Menu:
        """Build recalculation submenu."""
        menu = Menu("Recalculation")
        menu.items = [
            MenuItem("N", "Natural", MenuAction.COMMAND,
                    handler=self._cmd_recalc_natural, help_text="Natural order"),
            MenuItem("C", "Columnwise", MenuAction.COMMAND,
                    handler=self._cmd_recalc_columnwise, help_text="Column order"),
            MenuItem("R", "Rowwise", MenuAction.COMMAND,
                    handler=self._cmd_recalc_rowwise, help_text="Row order"),
            MenuItem("A", "Automatic", MenuAction.COMMAND,
                    handler=self._cmd_recalc_automatic, help_text="Auto recalc"),
            MenuItem("M", "Manual", MenuAction.COMMAND,
                    handler=self._cmd_recalc_manual, help_text="Manual recalc"),
            MenuItem("I", "Iteration", MenuAction.INPUT,
                    handler=self._cmd_recalc_iteration, help_text="Iteration count"),
        ]
        return menu

    def _build_global_protection_menu(self) -> Menu:
        """Build global protection submenu."""
        menu = Menu("Protection")
        menu.items = [
            MenuItem("E", "Enable", MenuAction.COMMAND,
                    handler=self._cmd_protection_enable, help_text="Enable protection"),
            MenuItem("D", "Disable", MenuAction.COMMAND,
                    handler=self._cmd_protection_disable, help_text="Disable protection"),
        ]
        return menu

    def _build_default_menu(self) -> Menu:
        """Build default settings submenu."""
        menu = Menu("Default")
        menu.items = [
            MenuItem("P", "Printer", MenuAction.SUBMENU,
                    help_text="Default printer settings"),
            MenuItem("D", "Directory", MenuAction.INPUT,
                    handler=self._cmd_default_directory, help_text="Default directory"),
            MenuItem("S", "Status", MenuAction.COMMAND,
                    handler=self._cmd_default_status, help_text="Show defaults"),
            MenuItem("U", "Update", MenuAction.COMMAND,
                    handler=self._cmd_default_update, help_text="Save defaults"),
            MenuItem("O", "Other", MenuAction.SUBMENU,
                    help_text="Other settings"),
        ]
        return menu

    def _build_zero_menu(self) -> Menu:
        """Build zero display submenu."""
        menu = Menu("Zero")
        menu.items = [
            MenuItem("Y", "Yes", MenuAction.COMMAND,
                    handler=self._cmd_zero_yes, help_text="Display zeros"),
            MenuItem("N", "No", MenuAction.COMMAND,
                    handler=self._cmd_zero_no, help_text="Suppress zeros"),
            MenuItem("L", "Label", MenuAction.INPUT,
                    handler=self._cmd_zero_label, help_text="Zero label"),
        ]
        return menu

    def _build_insert_menu(self) -> Menu:
        """Build insert submenu."""
        menu = Menu("Insert")
        menu.items = [
            MenuItem("R", "Row", MenuAction.COMMAND,
                    handler=self._cmd_insert_row, help_text="Insert row"),
            MenuItem("C", "Column", MenuAction.COMMAND,
                    handler=self._cmd_insert_column, help_text="Insert column"),
        ]
        return menu

    def _build_delete_menu(self) -> Menu:
        """Build delete submenu."""
        menu = Menu("Delete")
        menu.items = [
            MenuItem("R", "Row", MenuAction.COMMAND,
                    handler=self._cmd_delete_row, help_text="Delete row"),
            MenuItem("C", "Column", MenuAction.COMMAND,
                    handler=self._cmd_delete_column, help_text="Delete column"),
        ]
        return menu

    def _build_column_menu(self) -> Menu:
        """Build column submenu."""
        menu = Menu("Column")
        menu.items = [
            MenuItem("S", "Set-Width", MenuAction.INPUT,
                    handler=self._cmd_column_setwidth, help_text="Set column width"),
            MenuItem("R", "Reset-Width", MenuAction.COMMAND,
                    handler=self._cmd_column_resetwidth, help_text="Reset to default"),
            MenuItem("H", "Hide", MenuAction.COMMAND,
                    handler=self._cmd_column_hide, help_text="Hide column"),
            MenuItem("D", "Display", MenuAction.COMMAND,
                    handler=self._cmd_column_display, help_text="Show column"),
        ]
        return menu

    def _build_titles_menu(self) -> Menu:
        """Build titles submenu."""
        menu = Menu("Titles")
        menu.items = [
            MenuItem("B", "Both", MenuAction.COMMAND,
                    handler=self._cmd_titles_both, help_text="Freeze rows and columns"),
            MenuItem("H", "Horizontal", MenuAction.COMMAND,
                    handler=self._cmd_titles_horizontal, help_text="Freeze rows"),
            MenuItem("V", "Vertical", MenuAction.COMMAND,
                    handler=self._cmd_titles_vertical, help_text="Freeze columns"),
            MenuItem("C", "Clear", MenuAction.COMMAND,
                    handler=self._cmd_titles_clear, help_text="Unfreeze all"),
        ]
        return menu

    def _build_window_menu(self) -> Menu:
        """Build window submenu."""
        menu = Menu("Window")
        menu.items = [
            MenuItem("H", "Horizontal", MenuAction.COMMAND,
                    handler=self._cmd_window_horizontal, help_text="Split horizontally"),
            MenuItem("V", "Vertical", MenuAction.COMMAND,
                    handler=self._cmd_window_vertical, help_text="Split vertically"),
            MenuItem("S", "Sync", MenuAction.COMMAND,
                    handler=self._cmd_window_sync, help_text="Sync scrolling"),
            MenuItem("U", "Unsync", MenuAction.COMMAND,
                    handler=self._cmd_window_unsync, help_text="Unsync scrolling"),
            MenuItem("C", "Clear", MenuAction.COMMAND,
                    handler=self._cmd_window_clear, help_text="Remove split"),
        ]
        return menu

    def _build_range_format_menu(self) -> Menu:
        """Build range format submenu."""
        return self._build_format_submenu()

    def _build_label_menu(self) -> Menu:
        """Build label alignment submenu."""
        menu = Menu("Label")
        menu.items = [
            MenuItem("L", "Left", MenuAction.RANGE,
                    handler=self._cmd_label_range_left, help_text="Left align"),
            MenuItem("R", "Right", MenuAction.RANGE,
                    handler=self._cmd_label_range_right, help_text="Right align"),
            MenuItem("C", "Center", MenuAction.RANGE,
                    handler=self._cmd_label_range_center, help_text="Center"),
        ]
        return menu

    def _build_name_menu(self) -> Menu:
        """Build named range submenu."""
        menu = Menu("Name")
        menu.items = [
            MenuItem("C", "Create", MenuAction.INPUT,
                    handler=self._cmd_name_create, help_text="Create name"),
            MenuItem("D", "Delete", MenuAction.INPUT,
                    handler=self._cmd_name_delete, help_text="Delete name"),
            MenuItem("L", "Labels", MenuAction.SUBMENU,
                    submenu=self._build_name_labels_menu(), help_text="Create from labels"),
            MenuItem("R", "Reset", MenuAction.COMMAND,
                    handler=self._cmd_name_reset, help_text="Delete all names"),
            MenuItem("T", "Table", MenuAction.COMMAND,
                    handler=self._cmd_name_table, help_text="List names"),
        ]
        return menu

    def _build_name_labels_menu(self) -> Menu:
        """Build name from labels submenu."""
        menu = Menu("Labels")
        menu.items = [
            MenuItem("R", "Right", MenuAction.RANGE,
                    handler=self._cmd_name_labels_right, help_text="Names from left column"),
            MenuItem("D", "Down", MenuAction.RANGE,
                    handler=self._cmd_name_labels_down, help_text="Names from top row"),
            MenuItem("L", "Left", MenuAction.RANGE,
                    handler=self._cmd_name_labels_left, help_text="Names from right column"),
            MenuItem("U", "Up", MenuAction.RANGE,
                    handler=self._cmd_name_labels_up, help_text="Names from bottom row"),
        ]
        return menu

    def _build_protect_menu(self) -> Menu:
        """Build protect submenu."""
        menu = Menu("Protect")
        menu.items = [
            MenuItem("P", "Protect", MenuAction.RANGE,
                    handler=self._cmd_protect_range, help_text="Protect cells"),
        ]
        return menu

    def _build_combine_menu(self) -> Menu:
        """Build file combine submenu."""
        menu = Menu("Combine")
        menu.items = [
            MenuItem("C", "Copy", MenuAction.INPUT,
                    handler=self._cmd_combine_copy, help_text="Copy entire file"),
            MenuItem("A", "Add", MenuAction.INPUT,
                    handler=self._cmd_combine_add, help_text="Add values"),
            MenuItem("S", "Subtract", MenuAction.INPUT,
                    handler=self._cmd_combine_subtract, help_text="Subtract values"),
        ]
        return menu

    def _build_extract_menu(self) -> Menu:
        """Build file extract submenu."""
        menu = Menu("Xtract")
        menu.items = [
            MenuItem("F", "Formulas", MenuAction.INPUT,
                    handler=self._cmd_extract_formulas, help_text="Extract with formulas"),
            MenuItem("V", "Values", MenuAction.INPUT,
                    handler=self._cmd_extract_values, help_text="Extract values only"),
        ]
        return menu

    def _build_import_menu(self) -> Menu:
        """Build import submenu."""
        menu = Menu("Import")
        menu.items = [
            MenuItem("T", "Text", MenuAction.INPUT,
                    handler=self._cmd_import_text, help_text="Import text file"),
            MenuItem("N", "Numbers", MenuAction.INPUT,
                    handler=self._cmd_import_numbers, help_text="Import numeric data"),
        ]
        return menu

    def _build_printer_menu(self) -> Menu:
        """Build printer submenu."""
        menu = Menu("Printer")
        menu.items = [
            MenuItem("R", "Range", MenuAction.RANGE,
                    handler=self._cmd_print_range, help_text="Print range"),
            MenuItem("L", "Line", MenuAction.COMMAND,
                    handler=self._cmd_print_line, help_text="Line feed"),
            MenuItem("P", "Page", MenuAction.COMMAND,
                    handler=self._cmd_print_page, help_text="Page feed"),
            MenuItem("O", "Options", MenuAction.SUBMENU,
                    submenu=self._build_print_options_menu(), help_text="Print options"),
            MenuItem("C", "Clear", MenuAction.SUBMENU,
                    submenu=self._build_print_clear_menu(), help_text="Clear settings"),
            MenuItem("A", "Align", MenuAction.COMMAND,
                    handler=self._cmd_print_align, help_text="Align paper"),
            MenuItem("G", "Go", MenuAction.COMMAND,
                    handler=self._cmd_print_go, help_text="Start printing"),
            MenuItem("Q", "Quit", MenuAction.COMMAND,
                    handler=self._cmd_print_quit, help_text="Exit print menu"),
        ]
        return menu

    def _build_print_options_menu(self) -> Menu:
        """Build print options submenu."""
        menu = Menu("Options")
        menu.items = [
            MenuItem("H", "Header", MenuAction.INPUT,
                    handler=self._cmd_print_header, help_text="Page header"),
            MenuItem("F", "Footer", MenuAction.INPUT,
                    handler=self._cmd_print_footer, help_text="Page footer"),
            MenuItem("M", "Margins", MenuAction.SUBMENU,
                    help_text="Page margins"),
            MenuItem("B", "Borders", MenuAction.SUBMENU,
                    help_text="Print borders"),
            MenuItem("S", "Setup", MenuAction.INPUT,
                    handler=self._cmd_print_setup, help_text="Printer setup"),
            MenuItem("P", "Pg-Length", MenuAction.INPUT,
                    handler=self._cmd_print_pagelength, help_text="Page length"),
            MenuItem("O", "Other", MenuAction.SUBMENU,
                    help_text="Other options"),
        ]
        return menu

    def _build_print_clear_menu(self) -> Menu:
        """Build print clear submenu."""
        menu = Menu("Clear")
        menu.items = [
            MenuItem("A", "All", MenuAction.COMMAND,
                    handler=self._cmd_print_clear_all, help_text="Clear all"),
            MenuItem("R", "Range", MenuAction.COMMAND,
                    handler=self._cmd_print_clear_range, help_text="Clear range"),
            MenuItem("B", "Borders", MenuAction.COMMAND,
                    handler=self._cmd_print_clear_borders, help_text="Clear borders"),
            MenuItem("F", "Format", MenuAction.COMMAND,
                    handler=self._cmd_print_clear_format, help_text="Clear format"),
        ]
        return menu

    def _build_graph_type_menu(self) -> Menu:
        """Build graph type submenu."""
        menu = Menu("Type")
        menu.items = [
            MenuItem("L", "Line", MenuAction.COMMAND,
                    handler=self._cmd_graph_type_line, help_text="Line chart"),
            MenuItem("B", "Bar", MenuAction.COMMAND,
                    handler=self._cmd_graph_type_bar, help_text="Bar chart"),
            MenuItem("X", "XY", MenuAction.COMMAND,
                    handler=self._cmd_graph_type_xy, help_text="XY scatter"),
            MenuItem("S", "Stacked-Bar", MenuAction.COMMAND,
                    handler=self._cmd_graph_type_stacked, help_text="Stacked bar"),
            MenuItem("P", "Pie", MenuAction.COMMAND,
                    handler=self._cmd_graph_type_pie, help_text="Pie chart"),
        ]
        return menu

    def _build_graph_options_menu(self) -> Menu:
        """Build graph options submenu."""
        menu = Menu("Options")
        menu.items = [
            MenuItem("L", "Legend", MenuAction.SUBMENU,
                    help_text="Legend settings"),
            MenuItem("F", "Format", MenuAction.SUBMENU,
                    help_text="Format settings"),
            MenuItem("T", "Titles", MenuAction.SUBMENU,
                    help_text="Chart titles"),
            MenuItem("G", "Grid", MenuAction.SUBMENU,
                    help_text="Grid lines"),
            MenuItem("S", "Scale", MenuAction.SUBMENU,
                    help_text="Scale settings"),
            MenuItem("C", "Color", MenuAction.SUBMENU,
                    help_text="Color/B&W"),
            MenuItem("D", "Data-Labels", MenuAction.SUBMENU,
                    help_text="Data labels"),
        ]
        return menu

    def _build_graph_name_menu(self) -> Menu:
        """Build graph name submenu."""
        menu = Menu("Name")
        menu.items = [
            MenuItem("U", "Use", MenuAction.INPUT,
                    handler=self._cmd_graph_name_use, help_text="Use named graph"),
            MenuItem("C", "Create", MenuAction.INPUT,
                    handler=self._cmd_graph_name_create, help_text="Create named graph"),
            MenuItem("D", "Delete", MenuAction.INPUT,
                    handler=self._cmd_graph_name_delete, help_text="Delete named graph"),
            MenuItem("R", "Reset", MenuAction.COMMAND,
                    handler=self._cmd_graph_name_reset, help_text="Delete all graphs"),
        ]
        return menu

    def _build_data_table_menu(self) -> Menu:
        """Build data table submenu."""
        menu = Menu("Table")
        menu.items = [
            MenuItem("1", "1", MenuAction.RANGE,
                    handler=self._cmd_data_table1, help_text="One-variable table"),
            MenuItem("2", "2", MenuAction.RANGE,
                    handler=self._cmd_data_table2, help_text="Two-variable table"),
            MenuItem("R", "Reset", MenuAction.COMMAND,
                    handler=self._cmd_data_table_reset, help_text="Reset table"),
        ]
        return menu

    def _build_data_sort_menu(self) -> Menu:
        """Build data sort submenu."""
        menu = Menu("Sort")
        menu.items = [
            MenuItem("D", "Data-Range", MenuAction.RANGE,
                    handler=self._cmd_sort_data_range, help_text="Data to sort"),
            MenuItem("P", "Primary-Key", MenuAction.RANGE,
                    handler=self._cmd_sort_primary_key, help_text="Primary sort key"),
            MenuItem("S", "Secondary-Key", MenuAction.RANGE,
                    handler=self._cmd_sort_secondary_key, help_text="Secondary sort key"),
            MenuItem("R", "Reset", MenuAction.COMMAND,
                    handler=self._cmd_sort_reset, help_text="Reset sort settings"),
            MenuItem("G", "Go", MenuAction.COMMAND,
                    handler=self._cmd_sort_go, help_text="Execute sort"),
        ]
        return menu

    def _build_data_query_menu(self) -> Menu:
        """Build data query submenu."""
        menu = Menu("Query")
        menu.items = [
            MenuItem("I", "Input", MenuAction.RANGE,
                    handler=self._cmd_query_input, help_text="Input range"),
            MenuItem("C", "Criterion", MenuAction.RANGE,
                    handler=self._cmd_query_criterion, help_text="Criteria range"),
            MenuItem("O", "Output", MenuAction.RANGE,
                    handler=self._cmd_query_output, help_text="Output range"),
            MenuItem("F", "Find", MenuAction.COMMAND,
                    handler=self._cmd_query_find, help_text="Find matching"),
            MenuItem("E", "Extract", MenuAction.COMMAND,
                    handler=self._cmd_query_extract, help_text="Extract records"),
            MenuItem("U", "Unique", MenuAction.COMMAND,
                    handler=self._cmd_query_unique, help_text="Unique records"),
            MenuItem("D", "Delete", MenuAction.COMMAND,
                    handler=self._cmd_query_delete, help_text="Delete matching"),
            MenuItem("R", "Reset", MenuAction.COMMAND,
                    handler=self._cmd_query_reset, help_text="Reset query settings"),
        ]
        return menu

    def _build_data_matrix_menu(self) -> Menu:
        """Build data matrix submenu."""
        menu = Menu("Matrix")
        menu.items = [
            MenuItem("I", "Invert", MenuAction.RANGE,
                    handler=self._cmd_matrix_invert, help_text="Invert matrix"),
            MenuItem("M", "Multiply", MenuAction.RANGE,
                    handler=self._cmd_matrix_multiply, help_text="Multiply matrices"),
        ]
        return menu

    # Menu navigation
    def activate(self) -> None:
        """Activate the main menu."""
        self.context.current_menu = self.main_menu
        self.context.menu_path = ["MENU"]
        self.context.state = MenuState.ACTIVE
        self.context.error_message = ""

    def deactivate(self) -> None:
        """Deactivate the menu system."""
        self.context.current_menu = None
        self.context.menu_path = []
        self.context.state = MenuState.INACTIVE
        self.context.pending_action = None
        self.context.input_buffer = ""
        self.context.error_message = ""

    def go_back(self) -> bool:
        """Go back to parent menu. Returns False if should deactivate."""
        if not self.context.current_menu:
            return False

        if self.context.state in (MenuState.AWAITING_INPUT, MenuState.AWAITING_RANGE,
                                   MenuState.AWAITING_CONFIRM):
            self.context.state = MenuState.ACTIVE
            self.context.pending_action = None
            self.context.input_buffer = ""
            return True

        parent = self.context.current_menu.parent
        if parent:
            self.context.current_menu = parent
            if self.context.menu_path:
                self.context.menu_path.pop()
            return True

        return False

    def handle_key(self, key: str) -> tuple[bool, str | None]:
        """Handle a key press.

        Returns:
            Tuple of (handled, result_message)
        """
        if self.context.state == MenuState.INACTIVE:
            return (False, None)

        # Handle input mode
        if self.context.state == MenuState.AWAITING_INPUT:
            if key == "\n" or key == "Enter":
                result = self._execute_input_action()
                return (True, result)
            elif key == "Escape":
                self.go_back()
                return (True, None)
            elif key == "Backspace":
                self.context.input_buffer = self.context.input_buffer[:-1]
                return (True, None)
            else:
                self.context.input_buffer += key
                return (True, None)

        # Handle confirmation mode
        if self.context.state == MenuState.AWAITING_CONFIRM:
            if key.upper() == "Y":
                result = self._execute_confirm_action(True)
                return (True, result)
            elif key.upper() == "N" or key == "Escape":
                self.go_back()
                return (True, None)
            return (True, None)

        # Handle normal menu navigation
        if key == "Escape":
            if not self.go_back():
                self.deactivate()
            return (True, None)

        # Look for menu item
        if self.context.current_menu:
            item = self.context.current_menu.get_item(key)
            if item:
                return self._select_item(item)

        return (True, None)  # Absorb unknown keys in menu mode

    def _select_item(self, item: MenuItem) -> tuple[bool, str | None]:
        """Select a menu item."""
        if item.action == MenuAction.SUBMENU and item.submenu:
            self.context.current_menu = item.submenu
            self.context.menu_path.append(item.label)
            return (True, None)

        elif item.action == MenuAction.COMMAND:
            result = self._execute_command(item)
            self.deactivate()
            return (True, result)

        elif item.action == MenuAction.INPUT:
            self.context.state = MenuState.AWAITING_INPUT
            self.context.pending_action = item
            self.context.input_prompt = item.help_text + ": "
            self.context.input_buffer = ""
            return (True, None)

        elif item.action == MenuAction.RANGE:
            self.context.state = MenuState.AWAITING_RANGE
            self.context.pending_action = item
            self.context.input_prompt = "Enter range: "
            self.context.input_buffer = ""
            return (True, None)

        elif item.action == MenuAction.CONFIRM:
            self.context.state = MenuState.AWAITING_CONFIRM
            self.context.pending_action = item
            return (True, None)

        return (True, None)

    def _execute_command(self, item: MenuItem) -> str | None:
        """Execute a command menu item."""
        if item.handler:
            try:
                result = item.handler()
                return result if isinstance(result, str) else None
            except Exception as e:
                return f"Error: {e}"
        return None

    def _execute_input_action(self) -> str | None:
        """Execute the pending input action."""
        item = self.context.pending_action
        if not item or not item.handler:
            self.go_back()
            return None

        try:
            result = item.handler(self.context.input_buffer)
            self.deactivate()
            return result if isinstance(result, str) else None
        except Exception as e:
            self.context.error_message = str(e)
            return f"Error: {e}"

    def _execute_confirm_action(self, confirmed: bool) -> str | None:
        """Execute the pending confirm action."""
        item = self.context.pending_action
        if not item or not item.handler:
            self.go_back()
            return None

        try:
            result = item.handler(confirmed)
            self.deactivate()
            return result if isinstance(result, str) else None
        except Exception as e:
            return f"Error: {e}"

    def get_display(self) -> str:
        """Get the current menu display string."""
        if self.context.state == MenuState.INACTIVE:
            return ""

        parts = []

        # Menu path
        path = self.context.get_path_string()
        if path:
            parts.append(path)

        # Current menu items or input prompt
        if self.context.state == MenuState.AWAITING_INPUT:
            parts.append(f"{self.context.input_prompt}{self.context.input_buffer}_")
        elif self.context.state == MenuState.AWAITING_RANGE:
            parts.append(f"{self.context.input_prompt}{self.context.input_buffer}_")
        elif self.context.state == MenuState.AWAITING_CONFIRM:
            parts.append("Are you sure? (Y/N)")
        elif self.context.current_menu:
            parts.append(self.context.current_menu.get_display_line())

        return "  ".join(parts)

    def is_active(self) -> bool:
        """Check if menu is active."""
        return self.context.state != MenuState.INACTIVE

    # Command handlers - these are stubs that can be overridden
    def _cmd_worksheet_erase(self) -> str:
        return "Worksheet erased"

    def _cmd_worksheet_status(self) -> str:
        return "Status displayed"

    def _cmd_worksheet_page(self) -> str:
        return "Page break inserted"

    def _cmd_insert_row(self) -> str:
        if self.spreadsheet:
            # Would need current row from UI
            pass
        return "Row inserted"

    def _cmd_insert_column(self) -> str:
        return "Column inserted"

    def _cmd_delete_row(self) -> str:
        return "Row deleted"

    def _cmd_delete_column(self) -> str:
        return "Column deleted"

    def _cmd_column_setwidth(self, width: str) -> str:
        try:
            w = int(width)
            return f"Column width set to {w}"
        except ValueError:
            return "Invalid width"

    def _cmd_column_resetwidth(self) -> str:
        return "Column width reset"

    def _cmd_column_hide(self) -> str:
        return "Column hidden"

    def _cmd_column_display(self) -> str:
        return "Column displayed"

    def _cmd_titles_both(self) -> str:
        return "Rows and columns frozen"

    def _cmd_titles_horizontal(self) -> str:
        return "Rows frozen"

    def _cmd_titles_vertical(self) -> str:
        return "Columns frozen"

    def _cmd_titles_clear(self) -> str:
        return "Titles cleared"

    def _cmd_window_horizontal(self) -> str:
        return "Window split horizontally"

    def _cmd_window_vertical(self) -> str:
        return "Window split vertically"

    def _cmd_window_sync(self) -> str:
        return "Windows synchronized"

    def _cmd_window_unsync(self) -> str:
        return "Windows unsynchronized"

    def _cmd_window_clear(self) -> str:
        return "Window split removed"

    def _cmd_global_colwidth(self, width: str) -> str:
        return f"Default column width set to {width}"

    def _cmd_format_fixed(self, decimals: str) -> str:
        return f"Fixed format with {decimals} decimals"

    def _cmd_format_scientific(self, decimals: str) -> str:
        return f"Scientific format with {decimals} decimals"

    def _cmd_format_currency(self, decimals: str) -> str:
        return f"Currency format with {decimals} decimals"

    def _cmd_format_comma(self, decimals: str) -> str:
        return f"Comma format with {decimals} decimals"

    def _cmd_format_general(self) -> str:
        return "General format"

    def _cmd_format_plusminus(self) -> str:
        return "+/- format"

    def _cmd_format_percent(self, decimals: str) -> str:
        return f"Percent format with {decimals} decimals"

    def _cmd_date_format(self, fmt: int) -> str:
        return f"Date format {fmt}"

    def _cmd_format_text(self) -> str:
        return "Text format"

    def _cmd_format_hidden(self) -> str:
        return "Hidden format"

    def _cmd_format_reset(self) -> str:
        return "Format reset"

    def _cmd_label_left(self) -> str:
        return "Default left alignment"

    def _cmd_label_right(self) -> str:
        return "Default right alignment"

    def _cmd_label_center(self) -> str:
        return "Default center alignment"

    def _cmd_recalc_natural(self) -> str:
        return "Natural recalculation order"

    def _cmd_recalc_columnwise(self) -> str:
        return "Column-wise recalculation"

    def _cmd_recalc_rowwise(self) -> str:
        return "Row-wise recalculation"

    def _cmd_recalc_automatic(self) -> str:
        return "Automatic recalculation"

    def _cmd_recalc_manual(self) -> str:
        return "Manual recalculation"

    def _cmd_recalc_iteration(self, count: str) -> str:
        return f"Iteration count set to {count}"

    def _cmd_protection_enable(self) -> str:
        if self.spreadsheet:
            self.spreadsheet.protection.enable_protection()
        return "Protection enabled"

    def _cmd_protection_disable(self) -> str:
        if self.spreadsheet:
            self.spreadsheet.protection.disable_protection()
        return "Protection disabled"

    def _cmd_default_directory(self, path: str) -> str:
        return f"Default directory: {path}"

    def _cmd_default_status(self) -> str:
        return "Showing defaults"

    def _cmd_default_update(self) -> str:
        return "Defaults saved"

    def _cmd_zero_yes(self) -> str:
        return "Zeros displayed"

    def _cmd_zero_no(self) -> str:
        return "Zeros suppressed"

    def _cmd_zero_label(self, label: str) -> str:
        return f"Zero label: {label}"

    def _cmd_range_erase(self, range_ref: str) -> str:
        return f"Range {range_ref} erased"

    def _cmd_range_justify(self, range_ref: str) -> str:
        return f"Range {range_ref} justified"

    def _cmd_range_unprotect(self, range_ref: str) -> str:
        return f"Range {range_ref} unprotected"

    def _cmd_range_input(self, range_ref: str) -> str:
        return f"Input range: {range_ref}"

    def _cmd_range_value(self, range_ref: str) -> str:
        return f"Formulas converted to values in {range_ref}"

    def _cmd_range_transpose(self, range_ref: str) -> str:
        return f"Range {range_ref} transposed"

    def _cmd_label_range_left(self, range_ref: str) -> str:
        return f"Left alignment for {range_ref}"

    def _cmd_label_range_right(self, range_ref: str) -> str:
        return f"Right alignment for {range_ref}"

    def _cmd_label_range_center(self, range_ref: str) -> str:
        return f"Center alignment for {range_ref}"

    def _cmd_name_create(self, name: str) -> str:
        return f"Name '{name}' created"

    def _cmd_name_delete(self, name: str) -> str:
        return f"Name '{name}' deleted"

    def _cmd_name_reset(self) -> str:
        return "All names deleted"

    def _cmd_name_table(self) -> str:
        return "Name table displayed"

    def _cmd_name_labels_right(self, range_ref: str) -> str:
        return f"Names created from {range_ref}"

    def _cmd_name_labels_down(self, range_ref: str) -> str:
        return f"Names created from {range_ref}"

    def _cmd_name_labels_left(self, range_ref: str) -> str:
        return f"Names created from {range_ref}"

    def _cmd_name_labels_up(self, range_ref: str) -> str:
        return f"Names created from {range_ref}"

    def _cmd_protect_range(self, range_ref: str) -> str:
        return f"Range {range_ref} protected"

    def _cmd_copy(self, range_ref: str) -> str:
        return f"Range {range_ref} copied"

    def _cmd_move(self, range_ref: str) -> str:
        return f"Range {range_ref} moved"

    def _cmd_file_retrieve(self, filename: str) -> str:
        return f"File {filename} loaded"

    def _cmd_file_save(self, filename: str) -> str:
        return f"File saved as {filename}"

    def _cmd_combine_copy(self, filename: str) -> str:
        return f"File {filename} combined (copy)"

    def _cmd_combine_add(self, filename: str) -> str:
        return f"File {filename} combined (add)"

    def _cmd_combine_subtract(self, filename: str) -> str:
        return f"File {filename} combined (subtract)"

    def _cmd_extract_formulas(self, filename: str) -> str:
        return f"Extracted with formulas to {filename}"

    def _cmd_extract_values(self, filename: str) -> str:
        return f"Extracted values to {filename}"

    def _cmd_file_erase(self, filename: str) -> str:
        return f"File {filename} deleted"

    def _cmd_file_list(self) -> str:
        return "File list displayed"

    def _cmd_import_text(self, filename: str) -> str:
        return f"Text imported from {filename}"

    def _cmd_import_numbers(self, filename: str) -> str:
        return f"Numbers imported from {filename}"

    def _cmd_file_directory(self, path: str) -> str:
        return f"Directory changed to {path}"

    def _cmd_print_range(self, range_ref: str) -> str:
        return f"Print range: {range_ref}"

    def _cmd_print_line(self) -> str:
        return "Line feed"

    def _cmd_print_page(self) -> str:
        return "Page feed"

    def _cmd_print_align(self) -> str:
        return "Paper aligned"

    def _cmd_print_go(self) -> str:
        return "Printing..."

    def _cmd_print_quit(self) -> str:
        return "Print menu closed"

    def _cmd_print_header(self, header: str) -> str:
        return f"Header: {header}"

    def _cmd_print_footer(self, footer: str) -> str:
        return f"Footer: {footer}"

    def _cmd_print_setup(self, setup: str) -> str:
        return f"Printer setup: {setup}"

    def _cmd_print_pagelength(self, length: str) -> str:
        return f"Page length: {length}"

    def _cmd_print_clear_all(self) -> str:
        return "All print settings cleared"

    def _cmd_print_clear_range(self) -> str:
        return "Print range cleared"

    def _cmd_print_clear_borders(self) -> str:
        return "Print borders cleared"

    def _cmd_print_clear_format(self) -> str:
        return "Print format cleared"

    def _cmd_print_file(self, filename: str) -> str:
        return f"Printed to file {filename}"

    def _cmd_graph_type_line(self) -> str:
        return "Line chart selected"

    def _cmd_graph_type_bar(self) -> str:
        return "Bar chart selected"

    def _cmd_graph_type_xy(self) -> str:
        return "XY scatter selected"

    def _cmd_graph_type_stacked(self) -> str:
        return "Stacked bar selected"

    def _cmd_graph_type_pie(self) -> str:
        return "Pie chart selected"

    def _cmd_graph_x(self, range_ref: str) -> str:
        return f"X-axis range: {range_ref}"

    def _cmd_graph_a(self, range_ref: str) -> str:
        return f"Data range A: {range_ref}"

    def _cmd_graph_b(self, range_ref: str) -> str:
        return f"Data range B: {range_ref}"

    def _cmd_graph_c(self, range_ref: str) -> str:
        return f"Data range C: {range_ref}"

    def _cmd_graph_d(self, range_ref: str) -> str:
        return f"Data range D: {range_ref}"

    def _cmd_graph_e(self, range_ref: str) -> str:
        return f"Data range E: {range_ref}"

    def _cmd_graph_f(self, range_ref: str) -> str:
        return f"Data range F: {range_ref}"

    def _cmd_graph_reset(self) -> str:
        return "Graph reset"

    def _cmd_graph_view(self) -> str:
        return "Viewing graph..."

    def _cmd_graph_save(self, filename: str) -> str:
        return f"Graph saved to {filename}"

    def _cmd_graph_name_use(self, name: str) -> str:
        return f"Using graph '{name}'"

    def _cmd_graph_name_create(self, name: str) -> str:
        return f"Graph '{name}' created"

    def _cmd_graph_name_delete(self, name: str) -> str:
        return f"Graph '{name}' deleted"

    def _cmd_graph_name_reset(self) -> str:
        return "All graphs deleted"

    def _cmd_data_fill(self, range_ref: str) -> str:
        return f"Fill range: {range_ref}"

    def _cmd_data_table1(self, range_ref: str) -> str:
        return f"Table 1 range: {range_ref}"

    def _cmd_data_table2(self, range_ref: str) -> str:
        return f"Table 2 range: {range_ref}"

    def _cmd_data_table_reset(self) -> str:
        return "Table reset"

    def _cmd_sort_data_range(self, range_ref: str) -> str:
        return f"Sort data range: {range_ref}"

    def _cmd_sort_primary_key(self, range_ref: str) -> str:
        return f"Primary key: {range_ref}"

    def _cmd_sort_secondary_key(self, range_ref: str) -> str:
        return f"Secondary key: {range_ref}"

    def _cmd_sort_reset(self) -> str:
        return "Sort settings reset"

    def _cmd_sort_go(self) -> str:
        return "Sorting..."

    def _cmd_query_input(self, range_ref: str) -> str:
        return f"Query input range: {range_ref}"

    def _cmd_query_criterion(self, range_ref: str) -> str:
        return f"Query criteria range: {range_ref}"

    def _cmd_query_output(self, range_ref: str) -> str:
        return f"Query output range: {range_ref}"

    def _cmd_query_find(self) -> str:
        return "Finding matching records..."

    def _cmd_query_extract(self) -> str:
        return "Extracting records..."

    def _cmd_query_unique(self) -> str:
        return "Extracting unique records..."

    def _cmd_query_delete(self) -> str:
        return "Deleting matching records..."

    def _cmd_query_reset(self) -> str:
        return "Query settings reset"

    def _cmd_data_distribution(self, range_ref: str) -> str:
        return f"Distribution for {range_ref}"

    def _cmd_matrix_invert(self, range_ref: str) -> str:
        return f"Matrix {range_ref} inverted"

    def _cmd_matrix_multiply(self, range_ref: str) -> str:
        return f"Matrix {range_ref} multiplied"

    def _cmd_data_regression(self) -> str:
        return "Regression analysis completed"

    def _cmd_data_parse(self, range_ref: str) -> str:
        return f"Parsing {range_ref}"

    def _cmd_system(self) -> str:
        return "Entering system shell..."

    def _cmd_quit_yes(self) -> str:
        return "QUIT"

    def _cmd_quit_no(self) -> str:
        return "Quit cancelled"

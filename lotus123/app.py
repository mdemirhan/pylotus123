"""Lotus 1-2-3 Clone - Main TUI Application."""
from __future__ import annotations
import os
import json
from pathlib import Path
from dataclasses import dataclass, asdict
from enum import Enum, auto
from textual import on, events
from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Container, Horizontal, Vertical, Center, Middle
from textual.screen import ModalScreen
from textual.widgets import Static, Input, Label, Button, DirectoryTree, Footer, ListView, ListItem
from textual.message import Message
from textual.reactive import reactive
from textual.css.query import NoMatches
from rich.text import Text
from rich.style import Style

from .spreadsheet import Spreadsheet, index_to_col, make_cell_ref, parse_cell_ref
from .core.reference import CellReference, RangeReference, adjust_formula_references
from .core.formatting import parse_format_code, format_value, FormatCode
from .utils.undo import UndoManager, CellChangeCommand, RangeChangeCommand, InsertRowCommand, DeleteRowCommand
from .charting.chart import Chart, ChartType
from .charting.renderer import TextChartRenderer


# =============================================================================
# Configuration
# =============================================================================

CONFIG_DIR = Path.home() / ".config" / "lotus123"
CONFIG_FILE = CONFIG_DIR / "config.json"


@dataclass
class AppConfig:
    """Application configuration."""
    theme: str = "LOTUS"
    default_col_width: int = 10
    recent_files: list[str] = None

    def __post_init__(self):
        if self.recent_files is None:
            self.recent_files = []

    @classmethod
    def load(cls) -> "AppConfig":
        """Load configuration from file."""
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, "r") as f:
                    data = json.load(f)
                return cls(**data)
        except Exception:
            pass
        return cls()

    def save(self) -> None:
        """Save configuration to file."""
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            with open(CONFIG_FILE, "w") as f:
                json.dump(asdict(self), f, indent=2)
        except Exception:
            pass


# =============================================================================
# Color Themes
# =============================================================================

class ThemeType(Enum):
    """Available color themes."""
    LOTUS = auto()
    TOMORROW = auto()
    MOCHA = auto()


@dataclass
class Theme:
    """Color theme definition."""
    name: str
    background: str
    foreground: str
    header_bg: str
    header_fg: str
    cell_bg: str
    cell_fg: str
    selected_bg: str
    selected_fg: str
    border: str
    menu_bg: str
    menu_fg: str
    menu_highlight: str
    status_bg: str
    status_fg: str
    input_bg: str
    input_fg: str
    accent: str


THEMES: dict[ThemeType, Theme] = {
    ThemeType.LOTUS: Theme(
        name="Lotus 1-2-3",
        background="#000080",
        foreground="#ffffff",
        header_bg="#00aaaa",
        header_fg="#000000",
        cell_bg="#000080",
        cell_fg="#ffffff",
        selected_bg="#ffffff",
        selected_fg="#000000",
        border="#00aaaa",
        menu_bg="#00aaaa",
        menu_fg="#000000",
        menu_highlight="#ffffff",
        status_bg="#000080",
        status_fg="#ffffff",
        input_bg="#000080",
        input_fg="#ffffff",
        accent="#00aaaa",
    ),
    ThemeType.TOMORROW: Theme(
        name="Tomorrow Night",
        background="#1d1f21",
        foreground="#c5c8c6",
        header_bg="#373b41",
        header_fg="#c5c8c6",
        cell_bg="#1d1f21",
        cell_fg="#c5c8c6",
        selected_bg="#81a2be",
        selected_fg="#1d1f21",
        border="#373b41",
        menu_bg="#282a2e",
        menu_fg="#c5c8c6",
        menu_highlight="#81a2be",
        status_bg="#282a2e",
        status_fg="#969896",
        input_bg="#282a2e",
        input_fg="#c5c8c6",
        accent="#81a2be",
    ),
    ThemeType.MOCHA: Theme(
        name="Mocha",
        background="#1e1e2e",
        foreground="#cdd6f4",
        header_bg="#313244",
        header_fg="#cdd6f4",
        cell_bg="#1e1e2e",
        cell_fg="#cdd6f4",
        selected_bg="#89b4fa",
        selected_fg="#1e1e2e",
        border="#45475a",
        menu_bg="#313244",
        menu_fg="#cdd6f4",
        menu_highlight="#89b4fa",
        status_bg="#181825",
        status_fg="#a6adc8",
        input_bg="#313244",
        input_fg="#cdd6f4",
        accent="#89b4fa",
    ),
}


def get_theme_type(name: str) -> ThemeType:
    """Get ThemeType from string name."""
    try:
        return ThemeType[name.upper()]
    except KeyError:
        return ThemeType.LOTUS


# =============================================================================
# Spreadsheet Grid Widget
# =============================================================================

class SpreadsheetGrid(Static, can_focus=True):
    """The main spreadsheet grid display with mouse support and range selection."""

    cursor_row = reactive(0)
    cursor_col = reactive(0)
    scroll_row = reactive(0)
    scroll_col = reactive(0)
    # Range selection (anchor point for shift+arrow selection)
    select_anchor_row = reactive(-1)
    select_anchor_col = reactive(-1)
    # Freeze titles (number of rows/cols to keep visible)
    freeze_rows = reactive(0)
    freeze_cols = reactive(0)

    class CellSelected(Message):
        def __init__(self, row: int, col: int) -> None:
            self.row = row
            self.col = col
            super().__init__()

    class CellClicked(Message):
        def __init__(self, row: int, col: int) -> None:
            self.row = row
            self.col = col
            super().__init__()

    class RangeSelected(Message):
        def __init__(self, start_row: int, start_col: int, end_row: int, end_col: int) -> None:
            self.start_row = start_row
            self.start_col = start_col
            self.end_row = end_row
            self.end_col = end_col
            super().__init__()

    def __init__(self, spreadsheet: Spreadsheet, theme: Theme, **kwargs):
        super().__init__(**kwargs)
        self.spreadsheet = spreadsheet
        self.theme = theme
        self._visible_rows = 20
        self._visible_cols = 8
        self._col_positions: list[tuple[int, int, int]] = []
        self._row_positions: list[tuple[int, int]] = []
        self.default_col_width = 10  # Default column width
        self.show_zero = True  # Whether to display zero values

    @property
    def has_selection(self) -> bool:
        """Check if there's an active range selection."""
        return self.select_anchor_row >= 0 and self.select_anchor_col >= 0

    @property
    def selection_range(self) -> tuple[int, int, int, int]:
        """Get the normalized selection range (top-left to bottom-right)."""
        if not self.has_selection:
            return (self.cursor_row, self.cursor_col, self.cursor_row, self.cursor_col)
        r1, c1 = min(self.select_anchor_row, self.cursor_row), min(self.select_anchor_col, self.cursor_col)
        r2, c2 = max(self.select_anchor_row, self.cursor_row), max(self.select_anchor_col, self.cursor_col)
        return (r1, c1, r2, c2)

    def start_selection(self) -> None:
        """Start a range selection from current cursor position."""
        self.select_anchor_row = self.cursor_row
        self.select_anchor_col = self.cursor_col

    def clear_selection(self) -> None:
        """Clear the range selection."""
        self.select_anchor_row = -1
        self.select_anchor_col = -1

    def is_in_selection(self, row: int, col: int) -> bool:
        """Check if a cell is within the current selection."""
        if not self.has_selection:
            return row == self.cursor_row and col == self.cursor_col
        r1, c1, r2, c2 = self.selection_range
        return r1 <= row <= r2 and c1 <= col <= c2

    @property
    def visible_rows(self) -> int:
        return self._visible_rows

    @property
    def visible_cols(self) -> int:
        return self._visible_cols

    def on_mount(self) -> None:
        self._calculate_visible_area()
        self.refresh_grid()

    def on_resize(self, event: events.Resize) -> None:
        self._calculate_visible_area()
        self.refresh_grid()

    def _calculate_visible_area(self) -> None:
        """Calculate how many rows/cols fit in the current size."""
        if self.size.height > 2:
            self._visible_rows = self.size.height - 2
        if self.size.width > 6:
            # Calculate visible columns based on actual column widths
            # Start with row number area (4 chars + 1 border)
            used_width = 5
            visible_cols = 0
            for c in range(self.scroll_col, self.spreadsheet.cols):
                col_width = self.spreadsheet.get_col_width(c) + 1  # +1 for border
                if used_width + col_width > self.size.width:
                    break
                used_width += col_width
                visible_cols += 1
            self._visible_cols = max(1, visible_cols)

    def watch_cursor_row(self, value: int) -> None:
        self._ensure_visible()
        self.refresh_grid()
        self.post_message(self.CellSelected(value, self.cursor_col))

    def watch_cursor_col(self, value: int) -> None:
        self._ensure_visible()
        self.refresh_grid()
        self.post_message(self.CellSelected(self.cursor_row, value))

    def watch_scroll_row(self, value: int) -> None:
        self.refresh_grid()

    def watch_scroll_col(self, value: int) -> None:
        # Recalculate visible columns when scroll position changes
        # since visible cols depends on actual column widths from scroll position
        self._calculate_visible_area()
        self.refresh_grid()

    def _ensure_visible(self) -> None:
        if self.cursor_row < self.scroll_row:
            self.scroll_row = self.cursor_row
        elif self.cursor_row >= self.scroll_row + self._visible_rows:
            self.scroll_row = self.cursor_row - self._visible_rows + 1

        if self.cursor_col < self.scroll_col:
            self.scroll_col = self.cursor_col
        elif self.cursor_col >= self.scroll_col + self._visible_cols:
            self.scroll_col = self.cursor_col - self._visible_cols + 1

    def set_theme(self, theme: Theme) -> None:
        self.theme = theme
        self.refresh_grid()

    def refresh_grid(self) -> None:
        """Redraw the grid."""
        lines = []
        t = self.theme

        header_style = Style(color=t.header_fg, bgcolor=t.header_bg, bold=True)
        cell_style = Style(color=t.cell_fg, bgcolor=t.cell_bg)
        selected_style = Style(color=t.selected_fg, bgcolor=t.selected_bg, bold=True)
        border_style = Style(color=t.border, bgcolor=t.cell_bg)

        self._col_positions = []
        self._row_positions = []

        # Header row
        header = Text()
        header.append("    ", header_style)
        header.append("\u2502", border_style)

        x_pos = 5
        for c in range(self.scroll_col, self.scroll_col + self._visible_cols):
            if c >= self.spreadsheet.cols:
                break
            col_width = self.spreadsheet.get_col_width(c)
            col_name = index_to_col(c)
            header.append(col_name.center(col_width), header_style)
            header.append("\u2502", border_style)
            self._col_positions.append((c, x_pos, x_pos + col_width))
            x_pos += col_width + 1
        lines.append(header)

        # Separator
        sep = Text()
        sep.append("\u2500" * 4, border_style)
        sep.append("\u253c", border_style)
        for c in range(self.scroll_col, self.scroll_col + self._visible_cols):
            if c >= self.spreadsheet.cols:
                break
            col_width = self.spreadsheet.get_col_width(c)
            sep.append("\u2500" * col_width, border_style)
            sep.append("\u253c", border_style)
        lines.append(sep)

        # Data rows
        for row_idx, r in enumerate(range(self.scroll_row, self.scroll_row + self._visible_rows)):
            if r >= self.spreadsheet.rows:
                break
            row_text = Text()
            row_num = str(r + 1).rjust(4)
            row_text.append(row_num, header_style)
            row_text.append("\u2502", border_style)

            self._row_positions.append((r, row_idx + 2))

            for c in range(self.scroll_col, self.scroll_col + self._visible_cols):
                if c >= self.spreadsheet.cols:
                    break
                col_width = self.spreadsheet.get_col_width(c)
                value = self.spreadsheet.get_display_value(r, c)
                # Hide zero values if show_zero is False
                if not self.show_zero and value in ("0", "0.0", "0.00"):
                    value = ""
                display = value[:col_width].ljust(col_width)

                if self.is_in_selection(r, c):
                    row_text.append(display, selected_style)
                else:
                    row_text.append(display, cell_style)
                row_text.append("\u2502", border_style)

            lines.append(row_text)

        # Build content by appending lines with explicit newlines
        content = Text()
        for i, line in enumerate(lines):
            content.append_text(line)
            if i < len(lines) - 1:
                content.append("\n")
        self.update(content)
        self.refresh()

    def on_click(self, event: events.Click) -> None:
        """Handle mouse clicks to select cells."""
        click_x = event.x
        click_y = event.y

        clicked_row = None
        for row_idx, y in self._row_positions:
            if y == click_y:
                clicked_row = row_idx
                break

        clicked_col = None
        for col_idx, start_x, end_x in self._col_positions:
            if start_x <= click_x < end_x:
                clicked_col = col_idx
                break

        if clicked_row is not None and clicked_col is not None:
            self.cursor_row = clicked_row
            self.cursor_col = clicked_col
            self.post_message(self.CellClicked(clicked_row, clicked_col))

    def move_cursor(self, dr: int, dc: int) -> None:
        new_row = max(0, min(self.spreadsheet.rows - 1, self.cursor_row + dr))
        new_col = max(0, min(self.spreadsheet.cols - 1, self.cursor_col + dc))
        self.cursor_row = new_row
        self.cursor_col = new_col

    def goto_cell(self, ref: str) -> None:
        try:
            row, col = parse_cell_ref(ref)
            self.cursor_row = max(0, min(self.spreadsheet.rows - 1, row))
            self.cursor_col = max(0, min(self.spreadsheet.cols - 1, col))
        except ValueError:
            pass


# =============================================================================
# Menu System
# =============================================================================

class LotusMenu(Static, can_focus=True):
    """Lotus-style menu bar with mouse and keyboard support."""

    class MenuItemSelected(Message):
        def __init__(self, path: str) -> None:
            self.path = path
            super().__init__()

    class MenuActivated(Message):
        """Sent when menu is activated."""
        pass

    class MenuDeactivated(Message):
        """Sent when menu is deactivated."""
        pass

    MENU_STRUCTURE = {
        "Worksheet": {"key": "W", "items": [
            ("G", "Global", [
                ("F", "Format"),
                ("L", "Label-Prefix"),
                ("C", "Column-Width"),
                ("R", "Recalculation"),
                ("P", "Protection"),
                ("Z", "Zero"),
            ]),
            ("I", "Insert"),
            ("D", "Delete"),
            ("C", "Column"),
            ("E", "Erase"),
        ]},
        "Range": {"key": "R", "items": [("F", "Format"), ("L", "Label"), ("E", "Erase"), ("N", "Name"), ("P", "Protect")]},
        "Copy": {"key": "C", "items": []},
        "Move": {"key": "M", "items": []},
        "File": {"key": "F", "items": [("R", "Retrieve"), ("S", "Save"), ("N", "New"), ("Q", "Quit")]},
        "Graph": {"key": "G", "items": [
            ("T", "Type", [("L", "Line"), ("B", "Bar"), ("X", "XY"), ("S", "Stacked"), ("P", "Pie")]),
            ("X", "X-Range"),
            ("A", "A-Range"),
            ("B", "B-Range"),
            ("C", "C-Range"),
            ("V", "View"),
            ("R", "Reset"),
            ("S", "Save"),
        ]},
        "Data": {"key": "D", "items": [("F", "Fill"), ("S", "Sort"), ("Q", "Query")]},
        "Quit": {"key": "Q", "items": [("Y", "Yes"), ("N", "No")]},
    }

    def __init__(self, theme: Theme, **kwargs):
        super().__init__(**kwargs)
        self.theme = theme
        self.active = False
        self.current_menu: str | None = None
        self.submenu_path: list[str] = []  # Track submenu navigation
        self._menu_positions: list[tuple[str, int, int]] = []

    def set_theme(self, theme: Theme) -> None:
        self.theme = theme
        self._update_display()

    def on_mount(self) -> None:
        self._update_display()

    def _get_current_items(self) -> list:
        """Get the current menu items based on navigation state."""
        if self.current_menu is None:
            return []
        items = self.MENU_STRUCTURE[self.current_menu]["items"]
        # Navigate into submenus
        for submenu_name in self.submenu_path:
            for item in items:
                if len(item) >= 3 and item[1] == submenu_name:
                    items = item[2]
                    break
        return items

    def _update_display(self) -> None:
        t = self.theme
        text = Text()
        style = Style(color=t.menu_fg, bgcolor=t.menu_bg)
        highlight = Style(color=t.menu_highlight, bgcolor=t.menu_bg, bold=True)
        selected = Style(color=t.selected_fg, bgcolor=t.selected_bg, bold=True)

        if not self.active:
            text.append(" Press ", style)
            text.append("/", highlight)
            text.append(" for menu  |  ", style)
            x_pos = 21
            self._menu_positions = []
            for name, data in self.MENU_STRUCTURE.items():
                key = data["key"]
                start_x = x_pos
                text.append(key, highlight)
                text.append(name[1:] + "  ", style)
                end_x = x_pos + len(name) + 2
                self._menu_positions.append((name, start_x, end_x))
                x_pos = end_x
        else:
            if self.current_menu is None:
                text.append(" MENU: ", selected)
                x_pos = 7
                self._menu_positions = []
                for name, data in self.MENU_STRUCTURE.items():
                    key = data["key"]
                    start_x = x_pos
                    text.append(key, highlight)
                    text.append(name[1:] + "  ", style)
                    end_x = x_pos + len(name) + 2
                    self._menu_positions.append((name, start_x, end_x))
                    x_pos = end_x
            else:
                # Build title showing navigation path
                path_str = self.current_menu
                if self.submenu_path:
                    path_str += ">" + ">".join(self.submenu_path)
                text.append(f" {path_str}: ", selected)

                items = self._get_current_items()
                for item in items:
                    key = item[0]
                    label = item[1]
                    has_submenu = len(item) >= 3
                    text.append(key, highlight)
                    text.append(label[1:], style)
                    if has_submenu:
                        text.append(">", highlight)
                    text.append("  ", style)
                text.append(" [ESC=Back]", style)

        self.update(text)
        self.refresh()

    def activate(self) -> None:
        self.active = True
        self.current_menu = None
        self.submenu_path = []
        self._update_display()
        self.focus()
        self.post_message(self.MenuActivated())

    def deactivate(self) -> None:
        self.active = False
        self.current_menu = None
        self.submenu_path = []
        self._update_display()
        self.post_message(self.MenuDeactivated())

    def on_click(self, event: events.Click) -> None:
        """Handle menu clicks."""
        was_inactive = not self.active

        for name, start_x, end_x in self._menu_positions:
            if start_x <= event.x < end_x:
                if was_inactive:
                    self.active = True
                    self.post_message(self.MenuActivated())

                items = self.MENU_STRUCTURE[name]["items"]
                if items:
                    self.current_menu = name
                    self.submenu_path = []
                    self._update_display()
                    self.focus()
                else:
                    self.post_message(self.MenuItemSelected(name))
                    self.deactivate()
                return

        # Click on menu bar area but not on a specific menu item - activate menu
        if was_inactive:
            self.activate()

    def on_key(self, event: events.Key) -> None:
        """Handle keyboard navigation."""
        if not self.active:
            return

        if event.key == "escape":
            if self.submenu_path:
                # Go back one submenu level
                self.submenu_path.pop()
                self._update_display()
            elif self.current_menu:
                self.current_menu = None
                self._update_display()
            else:
                self.deactivate()
            event.prevent_default()
            event.stop()
            return

        char = event.character.upper() if event.character else ""

        if self.current_menu is None:
            for name, data in self.MENU_STRUCTURE.items():
                if char == data["key"]:
                    items = data["items"]
                    if items:
                        self.current_menu = name
                        self.submenu_path = []
                        self._update_display()
                    else:
                        self.post_message(self.MenuItemSelected(name))
                        self.deactivate()
                    event.prevent_default()
                    event.stop()
                    return
        else:
            items = self._get_current_items()
            for item in items:
                key = item[0]
                label = item[1]
                has_submenu = len(item) >= 3
                if char == key:
                    if has_submenu:
                        # Enter submenu
                        self.submenu_path.append(label)
                        self._update_display()
                    else:
                        # Build full path for message
                        path = self.current_menu
                        if self.submenu_path:
                            path += ":" + ":".join(self.submenu_path)
                        path += ":" + label
                        self.post_message(self.MenuItemSelected(path))
                        self.deactivate()
                    event.prevent_default()
                    event.stop()
                    return


# =============================================================================
# Dialogs
# =============================================================================

class FileDialog(ModalScreen[str | None]):
    """File open/save dialog."""

    BINDINGS = [Binding("escape", "cancel", "Cancel")]

    CSS = """
    FileDialog {
        align: center middle;
    }

    #file-dialog-container {
        width: 80%;
        height: 80%;
        max-width: 100;
        max-height: 40;
        background: $surface;
        border: thick $accent;
        padding: 1 2;
    }

    #file-tree {
        height: 1fr;
        margin: 1 0;
    }

    #filename-input {
        margin: 1 0;
    }

    #dialog-buttons {
        height: 3;
        align: center middle;
    }

    #dialog-buttons Button {
        margin: 0 2;
    }
    """

    def __init__(self, mode: str = "open", initial_path: str = ".", **kwargs):
        super().__init__(**kwargs)
        self.mode = mode
        self.initial_path = initial_path

    def compose(self) -> ComposeResult:
        title = "Open File" if self.mode == "open" else "Save File"
        with Container(id="file-dialog-container"):
            yield Label(f"[bold]{title}[/bold]", id="dialog-title")
            yield DirectoryTree(self.initial_path, id="file-tree")
            yield Input(placeholder="Filename...", id="filename-input")
            with Horizontal(id="dialog-buttons"):
                yield Button("OK", id="ok-btn", variant="primary")
                yield Button("Cancel", id="cancel-btn")

    @on(DirectoryTree.FileSelected)
    def on_file_selected(self, event: DirectoryTree.FileSelected) -> None:
        """Handle file selection - fill input or open if already selected."""
        input_widget = self.query_one("#filename-input", Input)
        selected_path = str(event.path)

        # If the same file is already in the input, open it
        if input_widget.value == selected_path:
            self.dismiss(selected_path)
        else:
            # Otherwise, fill the input with the selected path
            input_widget.value = selected_path

    @on(Input.Submitted, "#filename-input")
    def on_filename_submitted(self, event: Input.Submitted) -> None:
        """Open file when Enter is pressed in the filename input."""
        filename = event.value
        if filename:
            self.dismiss(filename)

    @on(Button.Pressed, "#ok-btn")
    def on_ok(self) -> None:
        filename = self.query_one("#filename-input", Input).value
        self.dismiss(filename if filename else None)

    @on(Button.Pressed, "#cancel-btn")
    def action_cancel(self) -> None:
        self.dismiss(None)


class CommandInput(ModalScreen[str | None]):
    """Modal input for commands like Goto."""

    BINDINGS = [Binding("escape", "cancel", "Cancel")]

    CSS = """
    CommandInput {
        align: center middle;
    }

    #cmd-dialog-container {
        width: 60;
        height: 9;
        background: $surface;
        border: thick $accent;
        padding: 1 2;
    }

    #cmd-input {
        margin-top: 1;
    }
    """

    def __init__(self, prompt: str, **kwargs):
        super().__init__(**kwargs)
        self.prompt = prompt

    def compose(self) -> ComposeResult:
        with Container(id="cmd-dialog-container"):
            yield Label(self.prompt, id="cmd-prompt")
            yield Input(id="cmd-input")

    def on_mount(self) -> None:
        self.query_one("#cmd-input").focus()

    @on(Input.Submitted)
    def on_submit(self, event: Input.Submitted) -> None:
        self.dismiss(event.value)

    def action_cancel(self) -> None:
        self.dismiss(None)


class ThemeItem(ListItem):
    """A theme list item."""
    def __init__(self, theme_type: ThemeType, theme: Theme, **kwargs):
        super().__init__(**kwargs)
        self.theme_type = theme_type
        self.theme_data = theme

    def compose(self) -> ComposeResult:
        yield Label(f"  {self.theme_data.name}")


class ThemeDialog(ModalScreen[ThemeType | None]):
    """Theme selection dialog with keyboard support."""

    BINDINGS = [
        Binding("escape", "cancel", "Cancel"),
        Binding("enter", "select", "Select"),
        Binding("up", "cursor_up", "Up", show=False),
        Binding("down", "cursor_down", "Down", show=False),
        Binding("1", "select_1", "1", show=False),
        Binding("2", "select_2", "2", show=False),
        Binding("3", "select_3", "3", show=False),
    ]

    CSS = """
    ThemeDialog {
        align: center middle;
    }

    #theme-dialog-container {
        width: 50;
        height: 15;
        background: $surface;
        border: thick $accent;
        padding: 1 2;
    }

    #theme-list {
        height: 1fr;
        margin: 1 0;
        border: solid $primary;
    }

    #theme-list > ListItem {
        padding: 0 1;
    }

    #theme-list > ListItem:hover {
        background: $accent;
    }

    #theme-list > ListItem.-highlight {
        background: $accent;
    }

    #dialog-buttons {
        height: 3;
        align: center middle;
    }
    """

    def __init__(self, current: ThemeType = ThemeType.LOTUS, **kwargs):
        super().__init__(**kwargs)
        self.current = current
        self._theme_types = list(THEMES.keys())

    def compose(self) -> ComposeResult:
        with Container(id="theme-dialog-container"):
            yield Label("[bold]Select Theme (1-3 or Enter)[/bold]", id="theme-title")
            yield ListView(
                ThemeItem(ThemeType.LOTUS, THEMES[ThemeType.LOTUS], id="theme-lotus"),
                ThemeItem(ThemeType.TOMORROW, THEMES[ThemeType.TOMORROW], id="theme-tomorrow"),
                ThemeItem(ThemeType.MOCHA, THEMES[ThemeType.MOCHA], id="theme-mocha"),
                id="theme-list"
            )
            with Horizontal(id="dialog-buttons"):
                yield Button("Cancel", id="cancel-btn")

    def on_mount(self) -> None:
        list_view = self.query_one("#theme-list", ListView)
        list_view.focus()
        # Set initial selection
        idx = self._theme_types.index(self.current)
        list_view.index = idx

    def action_cancel(self) -> None:
        self.dismiss(None)

    def action_select(self) -> None:
        list_view = self.query_one("#theme-list", ListView)
        if list_view.highlighted_child:
            item = list_view.highlighted_child
            if isinstance(item, ThemeItem):
                self.dismiss(item.theme_type)

    def action_cursor_up(self) -> None:
        self.query_one("#theme-list", ListView).action_cursor_up()

    def action_cursor_down(self) -> None:
        self.query_one("#theme-list", ListView).action_cursor_down()

    def action_select_1(self) -> None:
        self.dismiss(ThemeType.LOTUS)

    def action_select_2(self) -> None:
        self.dismiss(ThemeType.TOMORROW)

    def action_select_3(self) -> None:
        self.dismiss(ThemeType.MOCHA)

    @on(ListView.Selected)
    def on_list_selected(self, event: ListView.Selected) -> None:
        if isinstance(event.item, ThemeItem):
            self.dismiss(event.item.theme_type)

    @on(Button.Pressed, "#cancel-btn")
    def on_cancel_pressed(self) -> None:
        self.dismiss(None)


# =============================================================================
# Chart View Screen
# =============================================================================

class ChartViewScreen(ModalScreen):
    """Modal screen for viewing a chart."""

    BINDINGS = [
        Binding("escape", "dismiss_screen", "Close"),
        Binding("enter", "dismiss_screen", "Close"),
        Binding("q", "dismiss_screen", "Close"),
    ]

    CSS = """
    ChartViewScreen {
        align: center middle;
    }

    #chart-container {
        width: 80%;
        height: 80%;
        border: thick $accent;
        background: $surface;
        padding: 1;
    }

    #chart-content {
        width: 100%;
        height: 100%;
    }

    #chart-footer {
        height: 1;
        dock: bottom;
        text-align: center;
    }
    """

    def __init__(self, chart_lines: list[str], **kwargs):
        super().__init__(**kwargs)
        self.chart_lines = chart_lines

    def compose(self) -> ComposeResult:
        with Container(id="chart-container"):
            yield Static("\n".join(self.chart_lines), id="chart-content")
            yield Static("Press ESC or Enter to close", id="chart-footer")

    def action_dismiss_screen(self) -> None:
        self.dismiss(None)


# =============================================================================
# Main Application
# =============================================================================

class LotusApp(App):
    """Main Lotus 1-2-3 Clone Application."""

    BINDINGS = [
        Binding("ctrl+s", "save", "Save"),
        Binding("ctrl+o", "open_file", "Open"),
        Binding("ctrl+n", "new_file", "New"),
        Binding("ctrl+g", "goto", "Goto"),
        Binding("ctrl+t", "change_theme", "Theme"),
        Binding("ctrl+q", "quit_app", "Quit"),
        Binding("ctrl+z", "undo", "Undo"),
        Binding("ctrl+y", "redo", "Redo"),
        Binding("ctrl+c", "copy", "Copy"),
        Binding("ctrl+x", "cut", "Cut"),
        Binding("ctrl+v", "paste", "Paste"),
        Binding("f2", "edit_cell", "Edit"),
        Binding("f4", "toggle_absolute", "Abs", show=False),
        Binding("f5", "goto", "Goto", show=False),
        Binding("f9", "recalculate", "Calc", show=False),
        Binding("delete", "clear_cell", "Clear", show=False),
        Binding("escape", "cancel_edit", "Cancel", show=False),
    ]

    CSS = """
    Screen {
        layout: vertical;
        background: $surface;
    }

    #menu-bar {
        height: 1;
        width: 100%;
    }

    #cell-input-container {
        dock: top;
        height: 3;
        width: 100%;
        padding: 0 1;
    }

    #cell-ref {
        width: 10;
        text-style: bold;
    }

    #cell-input {
        width: 1fr;
    }

    #grid-container {
        width: 100%;
        height: 1fr;
    }

    #grid {
        width: 100%;
        height: 100%;
    }

    #status-bar {
        dock: bottom;
        height: 1;
        width: 100%;
    }

    Footer {
        dock: bottom;
    }
    """

    def __init__(self, initial_file: str | None = None):
        super().__init__()
        self._initial_file = initial_file
        self.config = AppConfig.load()
        self.spreadsheet = Spreadsheet()
        self.current_theme_type = get_theme_type(self.config.theme)
        self.color_theme = THEMES[self.current_theme_type]
        self.editing = False
        self._cell_clipboard: tuple[int, int, str] | None = None
        self._range_clipboard: list[list[str]] | None = None  # For range copy/paste
        self._clipboard_is_cut = False  # Track if last copy was a cut
        self._menu_active = False
        self.undo_manager = UndoManager(max_history=100)
        self._recalc_mode = "auto"  # "auto" or "manual"
        self.chart = Chart()  # Current chart configuration
        self._chart_renderer = TextChartRenderer(self.spreadsheet)
        # Global worksheet settings
        self._global_format_code = "G"  # Default format for new cells
        self._global_label_prefix = "'"  # Default label alignment (left)
        self._global_col_width = 10  # Default column width
        self._global_protection = False  # Worksheet protection enabled
        self._global_zero_display = True  # Show zero values
        self._dirty = False  # Track unsaved changes

    @property
    def _has_modal(self) -> bool:
        """Check if a modal screen is currently open."""
        return len(self.screen_stack) > 1

    def _mark_dirty(self) -> None:
        """Mark the spreadsheet as having unsaved changes."""
        self._dirty = True
        self._update_title()

    def _generate_css(self) -> str:
        """Generate CSS based on current theme."""
        t = self.color_theme
        return f"""
        Screen {{
            background: {t.background};
        }}

        #menu-bar {{
            background: {t.menu_bg};
            color: {t.menu_fg};
        }}

        #cell-input-container {{
            background: {t.background};
        }}

        #cell-ref {{
            color: {t.accent};
        }}

        #cell-input {{
            background: {t.input_bg};
            color: {t.input_fg};
            border: solid {t.accent};
        }}

        #grid {{
            background: {t.cell_bg};
        }}

        #status-bar {{
            background: {t.status_bg};
            color: {t.status_fg};
        }}

        Footer {{
            background: {t.menu_bg};
            color: {t.menu_fg};
        }}

        /* Dialog styling */
        ModalScreen {{
            background: {t.background} 80%;
        }}

        #file-dialog-container, #cmd-dialog-container, #theme-dialog-container {{
            background: {t.background};
            border: thick {t.accent};
        }}

        #dialog-title, #cmd-prompt, #theme-title {{
            color: {t.accent};
        }}

        #theme-list {{
            background: {t.cell_bg};
            border: solid {t.border};
        }}

        #theme-list > ListItem {{
            color: {t.cell_fg};
            background: {t.cell_bg};
        }}

        #theme-list > ListItem:hover {{
            background: {t.header_bg};
            color: {t.header_fg};
        }}

        #theme-list > ListItem.-highlight {{
            background: {t.selected_bg};
            color: {t.selected_fg};
        }}
        """

    def compose(self) -> ComposeResult:
        yield LotusMenu(self.color_theme, id="menu-bar")
        with Horizontal(id="cell-input-container"):
            yield Static("A1:", id="cell-ref")
            yield Input(id="cell-input", placeholder="Enter value or formula...")
        with Container(id="grid-container"):
            yield SpreadsheetGrid(self.spreadsheet, self.color_theme, id="grid")
        yield Static(" READY", id="status-bar")
        yield Footer()

    def on_mount(self) -> None:
        self._update_title()
        self.sub_title = f"Theme: {self.color_theme.name}"
        self._apply_theme()
        self._update_status()
        self.query_one("#grid", SpreadsheetGrid).focus()

        # Load initial file if provided via command line
        if self._initial_file:
            self._load_initial_file()

    def _update_title(self) -> None:
        """Update the window title with filename and dirty indicator."""
        filename = self.spreadsheet.filename or "Untitled"
        if "/" in filename or "\\" in filename:
            filename = Path(filename).name
        dirty_indicator = " *" if self._dirty else ""
        self.title = f"Lotus 1-2-3 Clone - {filename}{dirty_indicator}"

    def _apply_theme(self) -> None:
        """Apply the current theme to all widgets."""
        t = self.color_theme
        self.stylesheet.add_source(self._generate_css(), "theme")

        try:
            menu_bar = self.query_one("#menu-bar", LotusMenu)
            menu_bar.set_theme(t)

            self.query_one("#cell-ref").styles.color = t.accent
            self.query_one("#cell-input").styles.background = t.input_bg
            self.query_one("#cell-input").styles.color = t.input_fg
            self.query_one("#cell-input").styles.border = ("solid", t.accent)

            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.set_theme(t)

            self.query_one("#status-bar").styles.background = t.status_bg
            self.query_one("#status-bar").styles.color = t.status_fg

            self.query_one("#cell-input-container").styles.background = t.background
        except NoMatches:
            pass

    @on(SpreadsheetGrid.CellSelected)
    def on_cell_selected(self, event: SpreadsheetGrid.CellSelected) -> None:
        self._update_status()

    @on(SpreadsheetGrid.CellClicked)
    def on_cell_clicked(self, event: SpreadsheetGrid.CellClicked) -> None:
        self._update_status()
        if not self._menu_active:
            self.query_one("#grid", SpreadsheetGrid).focus()

    @on(LotusMenu.MenuItemSelected)
    def on_menu_item_selected(self, event: LotusMenu.MenuItemSelected) -> None:
        self._handle_menu(event.path)
        self._menu_active = False
        self.query_one("#grid", SpreadsheetGrid).focus()

    @on(LotusMenu.MenuActivated)
    def on_menu_activated(self, event: LotusMenu.MenuActivated) -> None:
        self._menu_active = True

    @on(LotusMenu.MenuDeactivated)
    def on_menu_deactivated(self, event: LotusMenu.MenuDeactivated) -> None:
        self._menu_active = False
        self.query_one("#grid", SpreadsheetGrid).focus()

    def _update_status(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        ref = make_cell_ref(grid.cursor_row, grid.cursor_col)
        cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)
        value = self.spreadsheet.get_display_value(grid.cursor_row, grid.cursor_col)

        self.query_one("#cell-ref", Static).update(f"{ref}:")

        if cell.is_formula:
            status = f" {ref}: {cell.raw_value} = {value}"
        elif value:
            status = f" {ref}: {value}"
        else:
            status = f" {ref}: READY"
        self.query_one("#status-bar", Static).update(status)

        if not self.editing:
            self.query_one("#cell-input", Input).value = cell.raw_value

    def action_move_up(self) -> None:
        if not self.editing and not self._menu_active and not self._has_modal:
            self.query_one("#grid", SpreadsheetGrid).move_cursor(-1, 0)

    def action_move_down(self) -> None:
        if not self.editing and not self._menu_active and not self._has_modal:
            self.query_one("#grid", SpreadsheetGrid).move_cursor(1, 0)

    def action_move_left(self) -> None:
        if not self.editing and not self._menu_active and not self._has_modal:
            self.query_one("#grid", SpreadsheetGrid).move_cursor(0, -1)

    def action_move_right(self) -> None:
        if not self.editing and not self._menu_active and not self._has_modal:
            self.query_one("#grid", SpreadsheetGrid).move_cursor(0, 1)

    def action_page_up(self) -> None:
        if not self.editing and not self._menu_active and not self._has_modal:
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.move_cursor(-grid.visible_rows, 0)

    def action_page_down(self) -> None:
        if not self.editing and not self._menu_active and not self._has_modal:
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.move_cursor(grid.visible_rows, 0)

    def action_go_home(self) -> None:
        if not self.editing and not self._menu_active and not self._has_modal:
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.cursor_col = 0

    def action_go_end(self) -> None:
        if not self.editing and not self._menu_active and not self._has_modal:
            grid = self.query_one("#grid", SpreadsheetGrid)
            for c in range(self.spreadsheet.cols - 1, -1, -1):
                cell = self.spreadsheet.get_cell_if_exists(grid.cursor_row, c)
                if cell and cell.raw_value:
                    grid.cursor_col = c
                    return
            grid.cursor_col = 0

    def action_edit_cell(self) -> None:
        if not self._menu_active:
            self.editing = True
            cell_input = self.query_one("#cell-input", Input)
            cell_input.focus()

    def action_cancel_edit(self) -> None:
        if self.editing:
            self.editing = False
            self._update_status()
            self.query_one("#grid", SpreadsheetGrid).focus()
        elif self._menu_active:
            menu = self.query_one("#menu-bar", LotusMenu)
            menu.deactivate()

    @on(Input.Submitted, "#cell-input")
    def on_cell_input_submitted(self, event: Input.Submitted) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        # Create undo command
        cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)
        old_value = cell.raw_value
        cmd = CellChangeCommand(
            spreadsheet=self.spreadsheet,
            row=grid.cursor_row,
            col=grid.cursor_col,
            new_value=event.value,
            old_value=old_value,
        )
        self.undo_manager.execute(cmd)
        self._mark_dirty()
        self.editing = False
        grid.refresh_grid()
        grid.move_cursor(1, 0)
        grid.focus()
        self._update_status()

    def action_clear_cell(self) -> None:
        if not self.editing and not self._menu_active:
            grid = self.query_one("#grid", SpreadsheetGrid)
            if grid.has_selection:
                # Clear selected range
                r1, c1, r2, c2 = grid.selection_range
                changes = []
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        cell = self.spreadsheet.get_cell(r, c)
                        if cell.raw_value:
                            changes.append((r, c, "", cell.raw_value))
                if changes:
                    cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                    self.undo_manager.execute(cmd)
                grid.clear_selection()
            else:
                # Clear single cell
                cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)
                if cell.raw_value:
                    cmd = CellChangeCommand(
                        spreadsheet=self.spreadsheet,
                        row=grid.cursor_row,
                        col=grid.cursor_col,
                        new_value="",
                        old_value=cell.raw_value,
                    )
                    self.undo_manager.execute(cmd)
            grid.refresh_grid()
            self._update_status()

    def action_undo(self) -> None:
        """Undo the last action."""
        if not self.editing and not self._menu_active:
            cmd = self.undo_manager.undo()
            if cmd:
                grid = self.query_one("#grid", SpreadsheetGrid)
                grid.refresh_grid()
                self._update_status()
                self._mark_dirty()
                self.notify(f"Undo: {cmd.description}")
            else:
                self.notify("Nothing to undo")

    def action_redo(self) -> None:
        """Redo the last undone action."""
        if not self.editing and not self._menu_active:
            cmd = self.undo_manager.redo()
            if cmd:
                grid = self.query_one("#grid", SpreadsheetGrid)
                grid.refresh_grid()
                self._update_status()
                self._mark_dirty()
                self.notify(f"Redo: {cmd.description}")
            else:
                self.notify("Nothing to redo")

    def action_copy(self) -> None:
        """Copy selected cell(s) to clipboard."""
        if not self.editing and not self._menu_active:
            self._copy_cells()

    def action_cut(self) -> None:
        """Cut selected cell(s) to clipboard."""
        if not self.editing and not self._menu_active:
            self._cut_cells()

    def action_paste(self) -> None:
        """Paste clipboard to current position."""
        if not self.editing and not self._menu_active:
            self._paste_cells()

    def action_toggle_absolute(self) -> None:
        """Toggle between relative and absolute references in formula (F4)."""
        if self.editing:
            cell_input = self.query_one("#cell-input", Input)
            value = cell_input.value
            if value.startswith("=") or value.startswith("@"):
                # Simple toggle - cycle through reference types
                # This is a simplified version; full implementation would
                # track cursor position and toggle the reference at cursor
                import re
                def toggle_ref(m):
                    ref = m.group(0)
                    if ref.startswith("$") and "$" in ref[1:]:
                        # $A$1 -> A$1
                        return ref[1:].replace("$", "", 1)
                    elif "$" in ref:
                        # A$1 or $A1 -> A1
                        return ref.replace("$", "")
                    else:
                        # A1 -> $A$1
                        col = re.match(r"([A-Za-z]+)", ref).group(1)
                        row = ref[len(col):]
                        return f"${col}${row}"
                new_value = re.sub(r'\$?[A-Za-z]+\$?\d+', toggle_ref, value)
                cell_input.value = new_value

    def action_recalculate(self) -> None:
        """Force recalculation of all formulas (F9)."""
        self.spreadsheet.recalculate()
        grid = self.query_one("#grid", SpreadsheetGrid)
        grid.refresh_grid()
        self._update_status()
        self.notify("Recalculated")

    def action_show_menu(self) -> None:
        if not self.editing and not self._has_modal:
            menu = self.query_one("#menu-bar", LotusMenu)
            menu.activate()

    def _handle_menu(self, result: str | None) -> None:
        if not result:
            return

        if result == "Goto" or result == "Worksheet:Goto":
            self.action_goto()
        elif result == "Copy":
            self._menu_copy()
        elif result == "Move":
            self._menu_move()
        elif result == "File:New":
            self.action_new_file()
        elif result == "File:Retrieve":
            self.action_open_file()
        elif result == "File:Save":
            self.action_save()
        elif result == "File:Quit" or result == "Quit:Yes":
            self.action_quit_app()
        elif result == "Quit:No":
            pass
        elif result == "Range:Erase":
            self.action_clear_cell()
        elif result == "Range:Format":
            self._range_format()
        elif result == "Range:Label":
            self._range_label()
        elif result == "Range:Name":
            self._range_name()
        elif result == "Range:Protect":
            self._range_protect()
        elif result == "Data:Fill":
            self._data_fill()
        elif result == "Data:Sort":
            self._data_sort()
        elif result == "Data:Query":
            self._data_query()
        elif result == "Worksheet:Insert":
            self._insert_row()
        elif result == "Worksheet:Delete":
            self._delete_row()
        elif result == "Worksheet:Column":
            self._set_column_width()
        elif result == "Worksheet:Erase":
            self._worksheet_erase()
        # Worksheet Global submenu
        elif result == "Worksheet:Global:Format":
            self._global_format()
        elif result == "Worksheet:Global:Label-Prefix":
            self._global_label_prefix()
        elif result == "Worksheet:Global:Column-Width":
            self._global_column_width()
        elif result == "Worksheet:Global:Recalculation":
            self._global_recalculation()
        elif result == "Worksheet:Global:Protection":
            self._global_protection()
        elif result == "Worksheet:Global:Zero":
            self._global_zero()
        # Graph menu handlers
        elif result == "Graph:Type:Line":
            self._set_chart_type(ChartType.LINE)
        elif result == "Graph:Type:Bar":
            self._set_chart_type(ChartType.BAR)
        elif result == "Graph:Type:XY":
            self._set_chart_type(ChartType.XY_SCATTER)
        elif result == "Graph:Type:Stacked":
            self._set_chart_type(ChartType.STACKED_BAR)
        elif result == "Graph:Type:Pie":
            self._set_chart_type(ChartType.PIE)
        elif result == "Graph:X-Range":
            self._set_chart_x_range()
        elif result == "Graph:A-Range":
            self._set_chart_a_range()
        elif result == "Graph:B-Range":
            self._set_chart_b_range()
        elif result == "Graph:Titles":
            self._set_chart_titles()
        elif result == "Graph:View":
            self._view_chart()
        elif result == "Graph:Reset":
            self._reset_chart()

    def action_goto(self) -> None:
        self.push_screen(CommandInput("Go to cell (e.g., A1):"), self._do_goto)

    def _do_goto(self, result: str | None) -> None:
        if result:
            self.query_one("#grid", SpreadsheetGrid).goto_cell(result.upper())
            self._update_status()

    def _set_column_width(self) -> None:
        self.push_screen(CommandInput("Column width (3-50):"), self._do_set_width)

    def _do_set_width(self, result: str | None) -> None:
        if result:
            try:
                width = int(result)
                grid = self.query_one("#grid", SpreadsheetGrid)
                self.spreadsheet.set_col_width(grid.cursor_col, width)
                grid.refresh_grid()
            except ValueError:
                pass

    def action_new_file(self) -> None:
        self.spreadsheet.clear()
        self.spreadsheet.filename = ""
        self.undo_manager.clear()
        self._dirty = False
        grid = self.query_one("#grid", SpreadsheetGrid)
        grid.cursor_row = 0
        grid.cursor_col = 0
        grid.scroll_row = 0
        grid.scroll_col = 0
        grid.refresh_grid()
        self._update_status()
        self._update_title()
        self.notify("New spreadsheet created")

    def action_open_file(self) -> None:
        self.push_screen(FileDialog(mode="open"), self._do_open)

    def _load_initial_file(self) -> None:
        """Load the initial file specified via command line."""
        try:
            filepath = Path(self._initial_file)
            if filepath.exists():
                self.spreadsheet.load(str(filepath))
                self.undo_manager.clear()
                self._dirty = False
                grid = self.query_one("#grid", SpreadsheetGrid)
                grid.refresh_grid()
                self._update_status()
                self._update_title()
                # Add to recent files
                filepath_str = str(filepath)
                if filepath_str not in self.config.recent_files:
                    self.config.recent_files.insert(0, filepath_str)
                    self.config.recent_files = self.config.recent_files[:10]
                    self.config.save()
                self.notify(f"Loaded: {filepath}")
            else:
                self.notify(f"File not found: {self._initial_file}", severity="error")
        except Exception as e:
            self.notify(f"Error loading file: {e}", severity="error")

    def _do_open(self, result: str | None) -> None:
        if result:
            try:
                self.spreadsheet.load(result)
                self.undo_manager.clear()
                self._dirty = False
                grid = self.query_one("#grid", SpreadsheetGrid)
                # Force scroll position reset by setting to non-zero first
                # This ensures the reactive watchers fire even if already at 0
                grid.scroll_row = 1
                grid.scroll_col = 1
                grid.scroll_row = 0
                grid.scroll_col = 0
                grid.cursor_row = 0
                grid.cursor_col = 0
                grid._calculate_visible_area()
                grid.refresh_grid()
                self._update_status()
                self._update_title()
                # Add to recent files
                if result not in self.config.recent_files:
                    self.config.recent_files.insert(0, result)
                    self.config.recent_files = self.config.recent_files[:10]
                    self.config.save()
                self.notify(f"Loaded: {result}")
            except Exception as e:
                self.notify(f"Error: {e}", severity="error")

    def action_save(self) -> None:
        if self.spreadsheet.filename:
            self.spreadsheet.save(self.spreadsheet.filename)
            self._dirty = False
            self._update_title()
            self.notify(f"Saved: {self.spreadsheet.filename}")
        else:
            self._save_as()

    def _save_as(self) -> None:
        self.push_screen(FileDialog(mode="save"), self._do_save)

    def _do_save(self, result: str | None) -> None:
        if result:
            try:
                if not result.endswith('.json'):
                    result += '.json'
                self.spreadsheet.save(result)
                self._dirty = False
                self._update_title()
                self.notify(f"Saved: {result}")
            except Exception as e:
                self.notify(f"Error: {e}", severity="error")

    def action_change_theme(self) -> None:
        self.push_screen(ThemeDialog(self.current_theme_type), self._do_change_theme)

    def _do_change_theme(self, result: ThemeType | None) -> None:
        if result:
            self.current_theme_type = result
            self.color_theme = THEMES[result]
            self.sub_title = f"Theme: {self.color_theme.name}"
            self._apply_theme()
            # Save preference
            self.config.theme = result.name
            self.config.save()
            self.notify(f"Theme changed to {self.color_theme.name}")

    def action_quit_app(self) -> None:
        if self._dirty:
            self.push_screen(
                CommandInput("Save changes before quitting? (Y/N/C=Cancel):"),
                self._do_quit_confirm
            )
        else:
            self.config.save()
            self.exit()

    def _do_quit_confirm(self, result: str | None) -> None:
        if not result:
            return  # Cancel
        response = result.strip().upper()
        if response.startswith("Y"):
            # Save then quit
            if self.spreadsheet.filename:
                self.spreadsheet.save(self.spreadsheet.filename)
                self.config.save()
                self.exit()
            else:
                # No filename - show save dialog
                self.push_screen(FileDialog(mode="save"), self._do_save_and_quit)
        elif response.startswith("N"):
            # Quit without saving
            self.config.save()
            self.exit()
        # else Cancel - do nothing

    def _do_save_and_quit(self, result: str | None) -> None:
        if result:
            self.spreadsheet.save(result)
        self.config.save()
        self.exit()

    def _copy_cell(self) -> None:
        """Legacy single cell copy."""
        self._copy_cells()

    def _cut_cell(self) -> None:
        """Legacy single cell cut."""
        self._cut_cells()

    def _paste_cell(self) -> None:
        """Legacy single cell paste."""
        self._paste_cells()

    def _menu_copy(self) -> None:
        """Lotus-style Copy: copy from current selection to a destination."""
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        source_range = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
        self._pending_source_range = (r1, c1, r2, c2)

        self.push_screen(
            CommandInput(f"Copy {source_range} TO (e.g., D1):"),
            self._do_menu_copy
        )

    def _do_menu_copy(self, result: str | None) -> None:
        if not result:
            return
        try:
            dest_row, dest_col = parse_cell_ref(result.upper())
            r1, c1, r2, c2 = self._pending_source_range

            changes = []
            for r_offset in range(r2 - r1 + 1):
                for c_offset in range(c2 - c1 + 1):
                    src_row, src_col = r1 + r_offset, c1 + c_offset
                    target_row, target_col = dest_row + r_offset, dest_col + c_offset

                    if target_row >= self.spreadsheet.rows or target_col >= self.spreadsheet.cols:
                        continue

                    src_cell = self.spreadsheet.get_cell(src_row, src_col)
                    target_cell = self.spreadsheet.get_cell(target_row, target_col)
                    old_value = target_cell.raw_value
                    new_value = src_cell.raw_value

                    # Adjust formula references
                    if new_value and (new_value.startswith("=") or new_value.startswith("@")):
                        row_delta = target_row - src_row
                        col_delta = target_col - src_col
                        new_value = new_value[0] + adjust_formula_references(
                            new_value[1:], row_delta, col_delta
                        )

                    if new_value != old_value:
                        changes.append((target_row, target_col, new_value, old_value))

            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid = self.query_one("#grid", SpreadsheetGrid)
                grid.refresh_grid()
                self._update_status()
                self.notify(f"Copied {len(changes)} cell(s)")

        except ValueError as e:
            self.notify(f"Invalid destination: {e}", severity="error")

    def _menu_move(self) -> None:
        """Lotus-style Move: move from current selection to a destination."""
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        source_range = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
        self._pending_source_range = (r1, c1, r2, c2)

        self.push_screen(
            CommandInput(f"Move {source_range} TO (e.g., D1):"),
            self._do_menu_move
        )

    def _do_menu_move(self, result: str | None) -> None:
        if not result:
            return
        try:
            dest_row, dest_col = parse_cell_ref(result.upper())
            r1, c1, r2, c2 = self._pending_source_range

            changes = []

            # First, copy data to destination
            for r_offset in range(r2 - r1 + 1):
                for c_offset in range(c2 - c1 + 1):
                    src_row, src_col = r1 + r_offset, c1 + c_offset
                    target_row, target_col = dest_row + r_offset, dest_col + c_offset

                    if target_row >= self.spreadsheet.rows or target_col >= self.spreadsheet.cols:
                        continue

                    src_cell = self.spreadsheet.get_cell(src_row, src_col)
                    target_cell = self.spreadsheet.get_cell(target_row, target_col)
                    old_value = target_cell.raw_value
                    new_value = src_cell.raw_value

                    # Adjust formula references
                    if new_value and (new_value.startswith("=") or new_value.startswith("@")):
                        row_delta = target_row - src_row
                        col_delta = target_col - src_col
                        new_value = new_value[0] + adjust_formula_references(
                            new_value[1:], row_delta, col_delta
                        )

                    if new_value != old_value:
                        changes.append((target_row, target_col, new_value, old_value))

            # Then, clear source cells (if not overlapping with destination)
            for r_offset in range(r2 - r1 + 1):
                for c_offset in range(c2 - c1 + 1):
                    src_row, src_col = r1 + r_offset, c1 + c_offset
                    target_row, target_col = dest_row + r_offset, dest_col + c_offset

                    # Don't clear if source and dest overlap
                    if src_row == target_row and src_col == target_col:
                        continue

                    src_cell = self.spreadsheet.get_cell(src_row, src_col)
                    if src_cell.raw_value:
                        changes.append((src_row, src_col, "", src_cell.raw_value))

            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid = self.query_one("#grid", SpreadsheetGrid)
                grid.clear_selection()
                grid.cursor_row = dest_row
                grid.cursor_col = dest_col
                grid.refresh_grid()
                self._update_status()
                self.notify(f"Moved cells to {make_cell_ref(dest_row, dest_col)}")

        except ValueError as e:
            self.notify(f"Invalid destination: {e}", severity="error")

    def _copy_cells(self) -> None:
        """Copy selected cell(s) to clipboard with range support."""
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range

        # Copy as 2D array for range paste
        self._range_clipboard = []
        self._clipboard_origin = (r1, c1)
        for r in range(r1, r2 + 1):
            row_data = []
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                row_data.append(cell.raw_value)
            self._range_clipboard.append(row_data)

        self._clipboard_is_cut = False

        # Also maintain legacy single-cell clipboard
        cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)
        self._cell_clipboard = (grid.cursor_row, grid.cursor_col, cell.raw_value)

        cells_count = (r2 - r1 + 1) * (c2 - c1 + 1)
        self.notify(f"Copied {cells_count} cell(s)")

    def _cut_cells(self) -> None:
        """Cut selected cell(s) to clipboard."""
        self._copy_cells()
        self._clipboard_is_cut = True
        self.notify("Cut to clipboard")

    def _paste_cells(self) -> None:
        """Paste clipboard to current position with formula adjustment."""
        if not self._range_clipboard:
            if self._cell_clipboard:
                # Legacy single-cell paste
                grid = self.query_one("#grid", SpreadsheetGrid)
                cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)
                old_value = cell.raw_value
                new_value = self._cell_clipboard[2]

                # Adjust formula references
                if new_value.startswith("=") or new_value.startswith("@"):
                    row_delta = grid.cursor_row - self._cell_clipboard[0]
                    col_delta = grid.cursor_col - self._cell_clipboard[1]
                    new_value = new_value[0] + adjust_formula_references(
                        new_value[1:], row_delta, col_delta
                    )

                cmd = CellChangeCommand(
                    spreadsheet=self.spreadsheet,
                    row=grid.cursor_row,
                    col=grid.cursor_col,
                    new_value=new_value,
                    old_value=old_value,
                )
                self.undo_manager.execute(cmd)
                grid.refresh_grid()
                self._update_status()
                self.notify("Pasted")
            return

        grid = self.query_one("#grid", SpreadsheetGrid)
        dest_row, dest_col = grid.cursor_row, grid.cursor_col
        src_row, src_col = getattr(self, '_clipboard_origin', (0, 0))

        changes = []
        for r_offset, row_data in enumerate(self._range_clipboard):
            for c_offset, value in enumerate(row_data):
                target_row = dest_row + r_offset
                target_col = dest_col + c_offset

                if target_row >= self.spreadsheet.rows or target_col >= self.spreadsheet.cols:
                    continue

                cell = self.spreadsheet.get_cell(target_row, target_col)
                old_value = cell.raw_value
                new_value = value

                # Adjust formula references for relative copy
                if new_value and (new_value.startswith("=") or new_value.startswith("@")):
                    row_delta = target_row - (src_row + r_offset)
                    col_delta = target_col - (src_col + c_offset)
                    new_value = new_value[0] + adjust_formula_references(
                        new_value[1:], row_delta, col_delta
                    )

                if new_value != old_value:
                    changes.append((target_row, target_col, new_value, old_value))

        if changes:
            cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
            self.undo_manager.execute(cmd)

        # If it was a cut, clear the source cells
        if self._clipboard_is_cut:
            clear_changes = []
            for r_offset, row_data in enumerate(self._range_clipboard):
                for c_offset, value in enumerate(row_data):
                    if value:
                        clear_changes.append((src_row + r_offset, src_col + c_offset, "", value))
            if clear_changes:
                clear_cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=clear_changes)
                self.undo_manager.execute(clear_cmd)
            self._clipboard_is_cut = False

        grid.refresh_grid()
        self._update_status()
        cells_count = len(self._range_clipboard) * len(self._range_clipboard[0]) if self._range_clipboard else 0
        self.notify(f"Pasted {cells_count} cell(s)")

    def _insert_row(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        cmd = InsertRowCommand(spreadsheet=self.spreadsheet, row=grid.cursor_row)
        self.undo_manager.execute(cmd)
        grid.refresh_grid()
        self.notify(f"Row {grid.cursor_row + 1} inserted")

    def _delete_row(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        cmd = DeleteRowCommand(spreadsheet=self.spreadsheet, row=grid.cursor_row)
        self.undo_manager.execute(cmd)
        grid.refresh_grid()
        self._update_status()
        self.notify(f"Row {grid.cursor_row + 1} deleted")

    # =========================================================================
    # Chart/Graph Methods
    # =========================================================================

    def _set_chart_type(self, chart_type: ChartType) -> None:
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

    def _set_chart_x_range(self) -> None:
        """Set the X-axis data range."""
        grid = self.query_one("#grid", SpreadsheetGrid)
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self.chart.set_x_range(range_str)
            self.notify(f"X-Range set to {range_str}")
        else:
            self.push_screen(CommandInput("X-Range (e.g., A1:A10):"), self._do_set_x_range)

    def _do_set_x_range(self, result: str | None) -> None:
        if result:
            self.chart.set_x_range(result.upper())
            self.notify(f"X-Range set to {result.upper()}")

    def _set_chart_a_range(self) -> None:
        """Set the A data series range."""
        grid = self.query_one("#grid", SpreadsheetGrid)
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self._add_or_update_series(0, "A", range_str)
        else:
            self.push_screen(CommandInput("A-Range (e.g., B1:B10):"), self._do_set_a_range)

    def _do_set_a_range(self, result: str | None) -> None:
        if result:
            self._add_or_update_series(0, "A", result.upper())

    def _set_chart_b_range(self) -> None:
        """Set the B data series range."""
        grid = self.query_one("#grid", SpreadsheetGrid)
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self._add_or_update_series(1, "B", range_str)
        else:
            self.push_screen(CommandInput("B-Range (e.g., C1:C10):"), self._do_set_b_range)

    def _do_set_b_range(self, result: str | None) -> None:
        if result:
            self._add_or_update_series(1, "B", result.upper())

    def _add_or_update_series(self, index: int, name: str, range_str: str) -> None:
        """Add or update a data series."""
        while len(self.chart.series) <= index:
            self.chart.add_series(f"Series {len(self.chart.series) + 1}")
        self.chart.series[index].name = name
        self.chart.series[index].data_range = range_str
        self.notify(f"{name}-Range set to {range_str}")

    def _set_chart_titles(self) -> None:
        """Set chart title."""
        self.push_screen(CommandInput("Chart title:"), self._do_set_title)

    def _do_set_title(self, result: str | None) -> None:
        if result:
            self.chart.set_title(result)
            self.notify(f"Chart title set to '{result}'")

    def _view_chart(self) -> None:
        """View the current chart."""
        if not self.chart.series:
            self.notify("No data series defined. Use A-Range to set data.")
            return

        # Render the chart
        self._chart_renderer.spreadsheet = self.spreadsheet
        chart_lines = self._chart_renderer.render(self.chart, width=70, height=20)

        # Show in modal
        self.push_screen(ChartViewScreen(chart_lines))

    def _reset_chart(self) -> None:
        """Reset chart to default settings."""
        self.chart.reset()
        self.notify("Chart reset")

    # =========================================================================
    # Range Menu Methods
    # =========================================================================

    def _range_format(self) -> None:
        """Set number format for selected range."""
        self.push_screen(
            CommandInput("Format (F=Fixed, S=Scientific, C=Currency, P=Percent, G=General):"),
            self._do_range_format
        )

    def _do_range_format(self, result: str | None) -> None:
        if not result:
            return
        format_char = result.upper()[0] if result else "G"
        format_map = {
            "F": "F2",  # Fixed with 2 decimals
            "S": "S",   # Scientific
            "C": "C2",  # Currency with 2 decimals
            "P": "P2",  # Percent with 2 decimals
            "G": "G",   # General
            ",": ",2",  # Comma format
        }
        format_code = format_map.get(format_char, "G")

        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range

        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                cell.format_str = format_code

        grid.refresh_grid()
        self._update_status()
        self.notify(f"Format set to {format_code}")

    def _range_label(self) -> None:
        """Set label alignment for selected range."""
        self.push_screen(
            CommandInput("Label alignment (L=Left, R=Right, C=Center):"),
            self._do_range_label
        )

    def _do_range_label(self, result: str | None) -> None:
        if not result:
            return
        align_char = result.upper()[0] if result else "L"
        # Lotus 1-2-3 uses prefix characters: ' for left, " for right, ^ for center
        prefix_map = {"L": "'", "R": '"', "C": "^"}
        prefix = prefix_map.get(align_char, "'")

        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range

        changes = []
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                old_value = cell.raw_value
                # Only apply to non-formula, non-numeric cells
                if old_value and not cell.is_formula:
                    # Remove existing prefix if any
                    display = cell.display_value
                    new_value = prefix + display
                    if new_value != old_value:
                        changes.append((r, c, new_value, old_value))

        if changes:
            cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
            self.undo_manager.execute(cmd)
            grid.refresh_grid()
            self._update_status()

        align_names = {"L": "Left", "R": "Right", "C": "Center"}
        self.notify(f"Label alignment set to {align_names.get(align_char, 'Left')}")

    def _range_name(self) -> None:
        """Create a named range."""
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
        self._pending_range = range_str
        self.push_screen(
            CommandInput(f"Name for range {range_str}:"),
            self._do_range_name
        )

    def _do_range_name(self, result: str | None) -> None:
        if not result:
            return
        name = result.strip().upper()
        if not name:
            return
        # Store named range (simplified - would need proper named range storage)
        if not hasattr(self.spreadsheet, '_named_ranges'):
            self.spreadsheet._named_ranges = {}
        self.spreadsheet._named_ranges[name] = self._pending_range
        self.notify(f"Named range '{name}' created for {self._pending_range}")

    def _range_protect(self) -> None:
        """Toggle protection for selected range."""
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range

        # Toggle protection (simplified - cells don't have protection attribute yet)
        protected_count = 0
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                # Add protection attribute if not present
                if not hasattr(cell, '_protected'):
                    cell._protected = False
                cell._protected = not cell._protected
                if cell._protected:
                    protected_count += 1

        total_cells = (r2 - r1 + 1) * (c2 - c1 + 1)
        if protected_count > 0:
            self.notify(f"Protected {protected_count} cell(s)")
        else:
            self.notify(f"Unprotected {total_cells} cell(s)")

    # =========================================================================
    # Data Menu Methods
    # =========================================================================

    def _data_fill(self) -> None:
        """Fill a range with values."""
        grid = self.query_one("#grid", SpreadsheetGrid)
        if not grid.has_selection:
            self.notify("Select a range first")
            return
        self.push_screen(
            CommandInput("Fill with (start,step,stop) or value:"),
            self._do_data_fill
        )

    def _do_data_fill(self, result: str | None) -> None:
        if not result:
            return

        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range

        changes = []
        try:
            # Check if it's a sequence (start,step,stop) or (start,step)
            if "," in result:
                parts = [p.strip() for p in result.split(",")]
                start = float(parts[0])
                step = float(parts[1]) if len(parts) > 1 else 1
                # Fill cells with sequence
                val = start
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        cell = self.spreadsheet.get_cell(r, c)
                        old_value = cell.raw_value
                        new_value = str(int(val) if val == int(val) else val)
                        changes.append((r, c, new_value, old_value))
                        val += step
            else:
                # Fill with single value
                fill_value = result
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        cell = self.spreadsheet.get_cell(r, c)
                        old_value = cell.raw_value
                        changes.append((r, c, fill_value, old_value))

            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid.refresh_grid()
                self._update_status()
                self.notify(f"Filled {len(changes)} cell(s)")

        except ValueError as e:
            self.notify(f"Invalid fill value: {e}", severity="error")

    def _data_sort(self) -> None:
        """Sort data in selected range."""
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range

        # Show which columns are in selection
        from .spreadsheet import index_to_col
        first_col = index_to_col(c1)
        last_col = index_to_col(c2)
        col_range = first_col if c1 == c2 else f"{first_col}-{last_col}"

        self.push_screen(
            CommandInput(f"Sort column [{col_range}] (add D for descending, e.g., 'A' or 'AD'):"),
            self._do_data_sort
        )

    def _do_data_sort(self, result: str | None) -> None:
        if not result:
            return

        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range

        try:
            # Parse sort specification - e.g., "A", "B", "AD" (A descending), "BD"
            result = result.strip().upper()
            reverse = result.endswith("D")
            sort_col_letter = result.rstrip("D").rstrip("A") or result[0]

            # Calculate sort column index
            from .spreadsheet import col_to_index, index_to_col
            sort_col_abs = col_to_index(sort_col_letter)

            # If column letter is outside selection, treat as relative (A=first col, B=second, etc.)
            if sort_col_abs < c1 or sort_col_abs > c2:
                sort_col_idx = ord(sort_col_letter) - ord("A")
                sort_col_abs = c1 + sort_col_idx

            if sort_col_abs < c1 or sort_col_abs > c2:
                self.notify(f"Sort column must be within selection ({index_to_col(c1)}-{index_to_col(c2)})", severity="error")
                return

            # Collect ALL row data (entire rows within selection)
            rows_data = []
            for r in range(r1, r2 + 1):
                row_values = []
                for c in range(c1, c2 + 1):
                    cell = self.spreadsheet.get_cell(r, c)
                    row_values.append(cell.raw_value)

                # Get sort key from the sort column
                sort_val = self.spreadsheet.get_value(r, sort_col_abs)
                # Handle mixed types for sorting
                if sort_val == "" or sort_val is None:
                    sort_key = (2, "")  # Empty values sort last
                elif isinstance(sort_val, (int, float)):
                    sort_key = (0, sort_val)  # Numbers first
                else:
                    sort_key = (1, str(sort_val).lower())  # Strings after numbers

                rows_data.append((sort_key, row_values))

            # Sort rows by sort key
            rows_data.sort(key=lambda x: x[0], reverse=reverse)

            # Apply sorted data - ALL columns move together
            changes = []
            for row_idx, (_, row_values) in enumerate(rows_data):
                target_row = r1 + row_idx
                for col_idx, value in enumerate(row_values):
                    target_col = c1 + col_idx
                    cell = self.spreadsheet.get_cell(target_row, target_col)
                    old_value = cell.raw_value
                    if value != old_value:
                        changes.append((target_row, target_col, value, old_value))

            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid.refresh_grid()
                self._update_status()
                order_name = "descending" if reverse else "ascending"
                self.notify(f"Sorted {len(rows_data)} rows by column {sort_col_letter} ({order_name})")
            else:
                self.notify("Data already sorted")

        except Exception as e:
            self.notify(f"Sort error: {e}", severity="error")

    def _data_query(self) -> None:
        """Database query operations."""
        # Simplified - just show a message about query functionality
        self.notify("Data Query: Select criteria range, then input range. Use @D functions for queries.")

    # =========================================================================
    # Worksheet Global Methods
    # =========================================================================

    def _global_format(self) -> None:
        """Set default number format for new cells."""
        self.push_screen(
            CommandInput(f"Default format (F=Fixed, S=Scientific, C=Currency, P=Percent, G=General) [{self._global_format_code}]:"),
            self._do_global_format
        )

    def _do_global_format(self, result: str | None) -> None:
        if not result:
            return
        format_char = result.upper()[0] if result else "G"
        format_map = {"F": "F2", "S": "S", "C": "C2", "P": "P2", "G": "G", ",": ",2"}
        self._global_format_code = format_map.get(format_char, "G")
        self.notify(f"Default format set to {self._global_format_code}")

    def _global_label_prefix(self) -> None:
        """Set default label alignment prefix."""
        self.push_screen(
            CommandInput("Default label alignment (L=Left, R=Right, C=Center):"),
            self._do_global_label_prefix
        )

    def _do_global_label_prefix(self, result: str | None) -> None:
        if not result:
            return
        align_char = result.upper()[0] if result else "L"
        prefix_map = {"L": "'", "R": '"', "C": "^"}
        self._global_label_prefix = prefix_map.get(align_char, "'")
        align_names = {"'": "Left", '"': "Right", "^": "Center"}
        self.notify(f"Default label alignment set to {align_names.get(self._global_label_prefix, 'Left')}")

    def _global_column_width(self) -> None:
        """Set default column width for all columns."""
        self.push_screen(
            CommandInput(f"Default column width (3-50) [{self._global_col_width}]:"),
            self._do_global_column_width
        )

    def _do_global_column_width(self, result: str | None) -> None:
        if not result:
            return
        try:
            width = int(result)
            width = max(3, min(50, width))
            self._global_col_width = width
            # Apply to all columns that don't have custom widths
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.default_col_width = width
            grid.refresh_grid()
            self.notify(f"Default column width set to {width}")
        except ValueError:
            self.notify("Invalid width", severity="error")

    def _global_recalculation(self) -> None:
        """Toggle automatic/manual recalculation."""
        if self._recalc_mode == "auto":
            self._recalc_mode = "manual"
            self.notify("Recalculation: Manual (press F9 to recalculate)")
        else:
            self._recalc_mode = "auto"
            self.spreadsheet._invalidate_cache()
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.refresh_grid()
            self.notify("Recalculation: Automatic")

    def _global_protection(self) -> None:
        """Toggle worksheet protection."""
        self._global_protection = not self._global_protection
        if self._global_protection:
            self.notify("Worksheet protection ENABLED - protected cells cannot be edited")
        else:
            self.notify("Worksheet protection DISABLED")

    def _global_zero(self) -> None:
        """Toggle display of zero values."""
        self._global_zero_display = not self._global_zero_display
        grid = self.query_one("#grid", SpreadsheetGrid)
        grid.show_zero = self._global_zero_display
        grid.refresh_grid()
        if self._global_zero_display:
            self.notify("Zero values: Displayed")
        else:
            self.notify("Zero values: Hidden (blank)")

    def _worksheet_erase(self) -> None:
        """Erase entire worksheet."""
        self.push_screen(
            CommandInput("Erase entire worksheet? (Y/N):"),
            self._do_worksheet_erase
        )

    def _do_worksheet_erase(self, result: str | None) -> None:
        if result and result.upper().startswith("Y"):
            self.spreadsheet.clear()
            self.undo_manager.clear()
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.refresh_grid()
            self._update_status()
            self.notify("Worksheet erased")

    def on_key(self, event: events.Key) -> None:
        """Handle key presses for navigation and direct cell input."""
        # Don't intercept if modal is open - let modal handle keys
        if self._has_modal:
            return

        # Handle / key for menu
        if event.key == "slash" or event.character == "/":
            if not self.editing:
                self.action_show_menu()
                event.prevent_default()
                event.stop()
                return

        # Don't intercept navigation if editing or menu is active
        if self.editing or self._menu_active:
            return

        # Arrow key navigation (with Shift for range selection)
        grid = self.query_one("#grid", SpreadsheetGrid)

        # Check for Shift+Arrow for range selection
        if event.key == "shift+up":
            if not grid.has_selection:
                grid.start_selection()
            grid.move_cursor(-1, 0)
            event.prevent_default()
            return
        elif event.key == "shift+down":
            if not grid.has_selection:
                grid.start_selection()
            grid.move_cursor(1, 0)
            event.prevent_default()
            return
        elif event.key == "shift+left":
            if not grid.has_selection:
                grid.start_selection()
            grid.move_cursor(0, -1)
            event.prevent_default()
            return
        elif event.key == "shift+right":
            if not grid.has_selection:
                grid.start_selection()
            grid.move_cursor(0, 1)
            event.prevent_default()
            return

        # Regular arrow keys (clear selection)
        if event.key == "up":
            grid.clear_selection()
            grid.move_cursor(-1, 0)
            event.prevent_default()
            return
        elif event.key == "down":
            grid.clear_selection()
            grid.move_cursor(1, 0)
            event.prevent_default()
            return
        elif event.key == "left":
            grid.clear_selection()
            grid.move_cursor(0, -1)
            event.prevent_default()
            return
        elif event.key == "right":
            grid.clear_selection()
            grid.move_cursor(0, 1)
            event.prevent_default()
            return
        elif event.key == "pageup":
            grid.move_cursor(-grid.visible_rows, 0)
            event.prevent_default()
            return
        elif event.key == "pagedown":
            grid.move_cursor(grid.visible_rows, 0)
            event.prevent_default()
            return
        elif event.key == "ctrl+d":
            # Half page down (Vim-style)
            grid.move_cursor(grid.visible_rows // 2, 0)
            event.prevent_default()
            return
        elif event.key == "ctrl+u":
            # Half page up (Vim-style)
            grid.move_cursor(-(grid.visible_rows // 2), 0)
            event.prevent_default()
            return
        elif event.key == "home":
            grid.cursor_col = 0
            event.prevent_default()
            return
        elif event.key == "end":
            for c in range(self.spreadsheet.cols - 1, -1, -1):
                cell = self.spreadsheet.get_cell_if_exists(grid.cursor_row, c)
                if cell and cell.raw_value:
                    grid.cursor_col = c
                    break
            else:
                grid.cursor_col = 0
            event.prevent_default()
            return
        elif event.key == "enter":
            self.action_edit_cell()
            event.prevent_default()
            return
        elif event.key in ("delete", "backspace"):
            self.action_clear_cell()
            event.prevent_default()
            return

        # Start editing on printable character (except /)
        if event.character and event.character.isprintable() and event.character != '/':
            cell_input = self.query_one("#cell-input", Input)
            # Disable select-on-focus to prevent text selection when focusing
            cell_input.select_on_focus = False
            # Clear existing value and focus
            cell_input.value = ""
            cell_input.focus()
            self.editing = True
            # Insert the character directly using the Input's internal method
            cell_input.insert_text_at_cursor(event.character)
            event.prevent_default()


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Lotus 1-2-3 Clone - A terminal-based spreadsheet application"
    )
    parser.add_argument(
        "file",
        nargs="?",
        help="Spreadsheet file to open on startup"
    )
    args = parser.parse_args()

    app = LotusApp(initial_file=args.file)
    app.run()


if __name__ == "__main__":
    main()

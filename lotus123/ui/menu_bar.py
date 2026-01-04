"""Lotus-style menu bar widget.

Provides the horizontal menu bar with keyboard and mouse support,
featuring single-letter shortcuts and hierarchical navigation.
"""

from typing import Any

from rich.style import Style
from rich.text import Text
from textual import events
from textual.message import Message
from textual.widgets import Static

from .themes import Theme

# Menu item can be a 2-tuple (key, name) or 3-tuple (key, name, subitems)
MenuItem = tuple[str, str] | tuple[str, str, list[Any]]
MenuData = dict[str, str | list[MenuItem]]


class LotusMenu(Static, can_focus=True):
    """Lotus-style menu bar with mouse and keyboard support."""

    class MenuItemSelected(Message):
        """Sent when a menu item is selected."""

        def __init__(self, path: str) -> None:
            self.path = path
            super().__init__()

    class MenuActivated(Message):
        """Sent when menu is activated."""

        pass

    class MenuDeactivated(Message):
        """Sent when menu is deactivated."""

        pass

    MENU_STRUCTURE: dict[str, MenuData] = {
        "Worksheet": {
            "key": "W",
            "items": [
                (
                    "G",
                    "Global",
                    [
                        ("F", "Format"),
                        ("L", "Label-Prefix"),
                        ("C", "Column-Width"),
                        ("R", "Recalculation"),
                        ("Z", "Zero"),
                    ],
                ),
                ("I", "Insert", [("R", "Rows"), ("C", "Columns")]),
                ("D", "Delete", [("R", "Rows"), ("C", "Columns")]),
                ("C", "Column"),
                ("E", "Erase"),
            ],
        },
        "Range": {
            "key": "R",
            "items": [
                ("F", "Format"),
                ("L", "Label"),
                ("E", "Erase"),
                ("N", "Name"),
            ],
        },
        "Copy": {"key": "C", "items": []},
        "Move": {"key": "M", "items": []},
        "File": {
            "key": "F",
            "items": [
                ("R", "Retrieve"),
                ("S", "Save"),
                ("A", "Save As"),
                ("N", "New"),
                (
                    "I",
                    "Import",
                    [
                        ("C", "CSV"),
                        ("T", "TSV"),
                        ("W", "WK1"),
                        ("X", "XLSX"),
                    ],
                ),
                (
                    "X",
                    "Export",
                    [
                        ("C", "CSV"),
                        ("T", "TSV"),
                        ("W", "WK1"),
                        ("X", "XLSX"),
                    ],
                ),
                ("Q", "Quit"),
            ],
        },
        "Graph": {
            "key": "G",
            "items": [
                (
                    "T",
                    "Type",
                    [
                        ("L", "Line"),
                        ("B", "Bar"),
                        ("X", "XY"),
                        ("S", "Stacked"),
                        ("P", "Pie"),
                    ],
                ),
                ("X", "X-Range"),
                ("A", "A-Range"),
                ("B", "B-Range"),
                ("C", "C-Range"),
                ("D", "D-Range"),
                ("E", "E-Range"),
                ("F", "F-Range"),
                ("V", "View"),
                ("R", "Reset"),
                ("S", "Save"),
                ("L", "Load"),
            ],
        },
        "Data": {
            "key": "D",
            "items": [
                ("F", "Fill"),
                ("S", "Sort"),
                (
                    "Q",
                    "Query",
                    [
                        ("I", "Input"),
                        ("C", "Criteria"),
                        ("O", "Output"),
                        ("F", "Find"),
                        ("X", "Extract"),
                        ("U", "Unique"),
                        ("D", "Delete"),
                        ("R", "Reset"),
                    ],
                ),
            ],
        },
        "System": {
            "key": "S",
            "items": [
                ("T", "Theme"),
            ],
        },
        "Quit": {"key": "Q", "items": [("Y", "Yes"), ("N", "No")]},
    }

    def __init__(self, theme: Theme, **kwargs: Any) -> None:
        super().__init__(**kwargs)
        self.theme = theme
        self.active = False
        self.current_menu: str | None = None
        self.submenu_path: list[str] = []
        self._menu_positions: list[tuple[str, int, int]] = []

    def set_theme(self, theme: Theme) -> None:
        """Update the menu's theme."""
        self.theme = theme
        self.styles.background = theme.menu_bg
        self._update_display()

    def on_mount(self) -> None:
        self._update_display()

    def _get_current_items(self) -> list[Any]:
        """Get the current menu items based on navigation state."""
        if self.current_menu is None:
            return []
        menu_items = self.MENU_STRUCTURE[self.current_menu]["items"]
        items: list[Any] = list(menu_items) if isinstance(menu_items, list) else []
        # Navigate into submenus
        for submenu_name in self.submenu_path:
            for item in items:
                if len(item) >= 3 and item[1] == submenu_name:
                    items = list(item[2])
                    break
        return items

    def _update_display(self) -> None:
        """Update the menu bar display."""

        def append_menu(x_pos: int) -> None:
            self._menu_positions = []
            for name, data in self.MENU_STRUCTURE.items():
                start_x = x_pos
                text.append(str(data["key"]), highlight)
                text.append(name[1:] + "  ", style)
                end_x = x_pos + len(name) + 2
                self._menu_positions.append((name, start_x, end_x))
                x_pos = end_x

        t = self.theme
        text = Text()
        style = Style(color=t.menu_fg, bgcolor=t.menu_bg)
        highlight = Style(color=t.menu_highlight, bgcolor=t.menu_bg, bold=True)
        selected = Style(color=t.selected_fg, bgcolor=t.selected_bg, bold=True)

        if not self.active:
            text.append(" Press ", style)
            text.append("/", highlight)
            text.append(" for menu  |  ", style)
            append_menu(21)
        else:
            if self.current_menu is None:
                text.append(" MENU: ", selected)
                append_menu(7)
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
                    # Find the key in the label (case-insensitive) and highlight it
                    key_idx = label.upper().find(key.upper())
                    if key_idx == 0:
                        # Key is first character
                        text.append(key, highlight)
                        text.append(label[1:], style)
                    elif key_idx > 0:
                        # Key is in the middle of the label
                        text.append(label[:key_idx], style)
                        text.append(label[key_idx], highlight)
                        text.append(label[key_idx + 1 :], style)
                    else:
                        # Key not found in label, show key in brackets
                        text.append(f"[{key}]", highlight)
                        text.append(label, style)
                    if has_submenu:
                        text.append(">", highlight)
                    text.append("  ", style)
                text.append(" [ESC=Back]", style)

        self.update(text)
        self.refresh()

    def activate(self) -> None:
        """Activate the menu system."""
        self.active = True
        self.current_menu = None
        self.submenu_path = []
        self._update_display()
        self.focus()
        self.post_message(self.MenuActivated())

    def deactivate(self) -> None:
        """Deactivate the menu system."""
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

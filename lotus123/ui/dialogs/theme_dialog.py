"""Theme selection dialog.

Provides a modal dialog for selecting the application color theme
with keyboard and mouse support.
"""

from __future__ import annotations

from typing import Any

from textual import on
from textual.app import ComposeResult
from textual.binding import Binding
from textual.containers import Container, Horizontal
from textual.screen import ModalScreen
from textual.widgets import Button, Label, ListItem, ListView

from ..themes import THEMES, Theme, ThemeType


class ThemeItem(ListItem):
    """A theme list item."""

    def __init__(self, theme_type: ThemeType, theme: Theme, **kwargs: Any) -> None:
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
        padding: 1 2;
    }

    #theme-list {
        height: 1fr;
        margin: 1 0;
    }

    #theme-list > ListItem {
        padding: 0 1;
    }

    #dialog-buttons {
        height: 3;
        align: center middle;
    }

    #dialog-buttons Button {
        margin: 0 2;
    }
    """

    def __init__(self, current: ThemeType = ThemeType.LOTUS, **kwargs: Any) -> None:
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
                id="theme-list",
            )
            with Horizontal(id="dialog-buttons"):
                yield Button("OK", id="ok-btn", variant="primary")
                yield Button("Cancel", id="cancel-btn")

    def on_mount(self) -> None:
        # Get theme fresh from the app's current setting
        theme_type = self.app.current_theme_type  # type: ignore[attr-defined]
        theme = THEMES[theme_type]

        container = self.query_one("#theme-dialog-container")
        container.styles.background = theme.cell_bg
        container.styles.border = ("thick", theme.accent)

        list_view = self.query_one("#theme-list", ListView)
        list_view.styles.border = ("solid", theme.accent)
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

    @on(Button.Pressed, "#ok-btn")
    def on_ok_pressed(self) -> None:
        self.action_select()

    @on(Button.Pressed, "#cancel-btn")
    def on_cancel_pressed(self) -> None:
        self.dismiss(None)

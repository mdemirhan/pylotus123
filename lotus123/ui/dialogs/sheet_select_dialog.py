"""Sheet selection dialog for multi-sheet XLSX import.

Provides a modal dialog for selecting which sheet to import from an
Excel workbook with multiple sheets.
"""

from typing import Any

from textual import on
from textual.app import ComposeResult
from textual.binding import Binding
from textual.containers import Container, Horizontal
from textual.screen import ModalScreen
from textual.widgets import Button, Label, ListItem, ListView


class SheetItem(ListItem):
    """A sheet list item."""

    def __init__(self, sheet_name: str, index: int, **kwargs: Any) -> None:
        super().__init__(**kwargs)
        self.sheet_name = sheet_name
        self.index = index

    def compose(self) -> ComposeResult:
        yield Label(f"  {self.index + 1}. {self.sheet_name}")


class SheetSelectDialog(ModalScreen[str | None]):
    """Sheet selection dialog for multi-sheet XLSX files.

    Returns the selected sheet name or None if cancelled.
    """

    BINDINGS = [
        Binding("escape", "cancel", "Cancel"),
        Binding("enter", "select", "Select"),
        Binding("up", "cursor_up", "Up", show=False),
        Binding("down", "cursor_down", "Down", show=False),
        Binding("1", "select_1", "1", show=False),
        Binding("2", "select_2", "2", show=False),
        Binding("3", "select_3", "3", show=False),
        Binding("4", "select_4", "4", show=False),
        Binding("5", "select_5", "5", show=False),
        Binding("6", "select_6", "6", show=False),
        Binding("7", "select_7", "7", show=False),
        Binding("8", "select_8", "8", show=False),
        Binding("9", "select_9", "9", show=False),
    ]

    CSS = """
    SheetSelectDialog {
        align: center middle;
    }

    #sheet-dialog-container {
        width: 60;
        height: auto;
        max-height: 20;
        padding: 1 2;
        background: $surface;
        border: thick $primary;
    }

    #sheet-title {
        text-align: center;
        padding-bottom: 1;
    }

    #sheet-list {
        height: auto;
        max-height: 12;
        margin: 1 0;
        border: solid $primary;
    }

    #sheet-list > ListItem {
        padding: 0 1;
    }

    #sheet-list > ListItem:hover {
        background: $accent;
    }

    #dialog-buttons {
        height: 3;
        align: center middle;
    }

    #dialog-buttons Button {
        margin: 0 2;
    }
    """

    def __init__(self, sheet_names: list[str], **kwargs: Any) -> None:
        """Initialize dialog with sheet names.

        Args:
            sheet_names: List of sheet names from the workbook
        """
        super().__init__(**kwargs)
        self.sheet_names = sheet_names

    def compose(self) -> ComposeResult:
        with Container(id="sheet-dialog-container"):
            yield Label(
                f"[bold]Select Sheet to Import ({len(self.sheet_names)} available)[/bold]",
                id="sheet-title",
            )
            yield ListView(
                *[SheetItem(name, i, id=f"sheet-{i}") for i, name in enumerate(self.sheet_names)],
                id="sheet-list",
            )
            with Horizontal(id="dialog-buttons"):
                yield Button("Import", id="ok-btn", variant="primary")
                yield Button("Cancel", id="cancel-btn")

    def on_mount(self) -> None:
        """Focus the list view on mount."""
        list_view = self.query_one("#sheet-list", ListView)
        list_view.focus()
        # Select first sheet by default
        list_view.index = 0

    def action_cancel(self) -> None:
        """Cancel the dialog."""
        self.dismiss(None)

    def action_select(self) -> None:
        """Select the currently highlighted sheet."""
        list_view = self.query_one("#sheet-list", ListView)
        if list_view.highlighted_child:
            item = list_view.highlighted_child
            if isinstance(item, SheetItem):
                self.dismiss(item.sheet_name)

    def action_cursor_up(self) -> None:
        """Move cursor up."""
        self.query_one("#sheet-list", ListView).action_cursor_up()

    def action_cursor_down(self) -> None:
        """Move cursor down."""
        self.query_one("#sheet-list", ListView).action_cursor_down()

    def _select_by_index(self, index: int) -> None:
        """Select sheet by 1-based index."""
        if 0 <= index < len(self.sheet_names):
            self.dismiss(self.sheet_names[index])

    def action_select_1(self) -> None:
        self._select_by_index(0)

    def action_select_2(self) -> None:
        self._select_by_index(1)

    def action_select_3(self) -> None:
        self._select_by_index(2)

    def action_select_4(self) -> None:
        self._select_by_index(3)

    def action_select_5(self) -> None:
        self._select_by_index(4)

    def action_select_6(self) -> None:
        self._select_by_index(5)

    def action_select_7(self) -> None:
        self._select_by_index(6)

    def action_select_8(self) -> None:
        self._select_by_index(7)

    def action_select_9(self) -> None:
        self._select_by_index(8)

    @on(ListView.Selected)
    def on_list_selected(self, event: ListView.Selected) -> None:
        """Handle double-click/enter on list item."""
        if isinstance(event.item, SheetItem):
            self.dismiss(event.item.sheet_name)

    @on(Button.Pressed, "#ok-btn")
    def on_ok_pressed(self) -> None:
        """Handle OK button press."""
        self.action_select()

    @on(Button.Pressed, "#cancel-btn")
    def on_cancel_pressed(self) -> None:
        """Handle Cancel button press."""
        self.dismiss(None)

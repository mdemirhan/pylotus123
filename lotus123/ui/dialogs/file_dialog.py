"""File open/save dialog.

Provides a modal dialog for browsing and selecting files,
with directory tree navigation and filename input.
"""
from __future__ import annotations

from typing import Any

from textual import on
from textual.binding import Binding
from textual.containers import Container, Horizontal
from textual.screen import ModalScreen
from textual.widgets import Button, DirectoryTree, Input, Label
from textual.app import ComposeResult


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

    def __init__(self, mode: str = "open", initial_path: str = ".", **kwargs: Any) -> None:
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

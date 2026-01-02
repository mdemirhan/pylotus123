"""File open/save dialog.

Provides a modal dialog for browsing and selecting files,
with directory tree navigation and filename input.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable

from textual import on
from textual.app import ComposeResult
from textual.binding import Binding
from textual.containers import Container, Horizontal
from textual.screen import ModalScreen
from textual.widgets import Button, Checkbox, DirectoryTree, Input, Label


class FilteredDirectoryTree(DirectoryTree):
    """DirectoryTree that filters files by extension."""

    def __init__(
        self,
        path: str | Path,
        extensions: list[str] | None = None,
        **kwargs: Any,
    ) -> None:
        super().__init__(path, **kwargs)
        # Normalize extensions to lowercase with leading dot
        self._extensions: set[str] | None = None
        if extensions:
            self._extensions = {
                ext.lower() if ext.startswith(".") else f".{ext.lower()}"
                for ext in extensions
            }
        self._show_all = False

    def set_show_all(self, show_all: bool) -> None:
        """Toggle showing all files."""
        self._show_all = show_all
        self.reload()

    def filter_paths(self, paths: Iterable[Path]) -> Iterable[Path]:
        """Filter paths to only show matching extensions."""
        if self._show_all or not self._extensions:
            return paths

        result = []
        for path in paths:
            if path.is_dir():
                result.append(path)
            elif path.suffix.lower() in self._extensions:
                result.append(path)
        return result


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
        padding: 1 2;
    }

    #file-tree {
        height: 1fr;
        margin: 1 0;
    }

    #filter-row {
        height: auto;
        margin: 0 0;
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

    def __init__(
        self,
        mode: str = "open",
        initial_path: str = ".",
        title: str | None = None,
        file_extensions: list[str] | None = None,
        **kwargs: Any,
    ) -> None:
        super().__init__(**kwargs)
        self.mode = mode
        self.initial_path = initial_path
        self._custom_title = title
        self._file_extensions = file_extensions

    def compose(self) -> ComposeResult:
        if self._custom_title:
            title = self._custom_title
        else:
            title = "Open File" if self.mode == "open" else "Save File"
        with Container(id="file-dialog-container"):
            yield Label(f"[bold]{title}[/bold]", id="dialog-title")
            yield FilteredDirectoryTree(
                self.initial_path,
                extensions=self._file_extensions,
                id="file-tree",
            )
            if self._file_extensions:
                with Horizontal(id="filter-row"):
                    yield Checkbox("Show all files", id="show-all-checkbox")
            yield Input(placeholder="Filename...", id="filename-input")
            with Horizontal(id="dialog-buttons"):
                yield Button("OK", id="ok-btn", variant="primary")
                yield Button("Cancel", id="cancel-btn")

    def on_mount(self) -> None:
        # Get theme fresh from the app's current setting
        from ..themes import THEMES
        theme_type = self.app.current_theme_type  # type: ignore[attr-defined]
        theme = THEMES[theme_type]

        container = self.query_one("#file-dialog-container")
        container.styles.background = theme.cell_bg
        container.styles.border = ("thick", theme.accent)

    @on(Checkbox.Changed, "#show-all-checkbox")
    def on_show_all_changed(self, event: Checkbox.Changed) -> None:
        """Toggle showing all files in the directory tree."""
        tree = self.query_one("#file-tree", FilteredDirectoryTree)
        tree.set_show_all(event.value)

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

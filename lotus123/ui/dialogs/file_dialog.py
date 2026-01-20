"""File open/save dialog.

Provides a modal dialog for browsing and selecting files,
with directory tree navigation and filename input.
"""

from pathlib import Path
from typing import Any, Iterable, override

from textual import on
from textual.app import ComposeResult
from textual.binding import Binding
from textual.containers import Container, Horizontal
from textual.screen import ModalScreen
from textual.widgets import Button, Checkbox, DirectoryTree, Input, Label, Static


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
                ext.lower() if ext.startswith(".") else f".{ext.lower()}" for ext in extensions
            }
        self._show_all = False

    def set_show_all(self, show_all: bool) -> None:
        """Toggle showing all files."""
        self._show_all = show_all
        self.reload()

    @override
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

    #path-row {
        height: 3;
        margin: 1 0 0 0;
        align: left middle;
    }

    #go-up-btn {
        min-width: 10;
    }

    #current-path {
        margin-left: 1;
        height: 3;
        content-align: left middle;
    }

    .file-tree {
        height: 1fr;
        margin: 0 0 1 0;
    }

    #filter-row {
        height: 3;
        margin: 0 0;
        align: left middle;
    }

    #extension-display {
        margin-left: 1;
        height: 3;
        content-align: left middle;
        color: $text-muted;
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
        self._current_path = Path(initial_path).resolve()
        self._tree_counter = 0  # For generating unique tree IDs

    @override
    def compose(self) -> ComposeResult:
        if self._custom_title:
            title = self._custom_title
        else:
            title = "Open File" if self.mode == "open" else "Save File"

        # For save mode with extensions, prepare initial filename with extension
        initial_filename = ""
        if self.mode == "save" and self._file_extensions:
            # Use first extension as default
            ext = self._file_extensions[0]
            initial_filename = ext if ext.startswith(".") else f".{ext}"

        with Container(id="file-dialog-container"):
            yield Label(f"[bold]{title}[/bold]", id="dialog-title")
            with Horizontal(id="path-row"):
                yield Button("Go Up", id="go-up-btn")
                yield Static(str(self._current_path), id="current-path")
            tree = FilteredDirectoryTree(
                self._current_path,
                extensions=self._file_extensions,
                classes="file-tree",
            )
            yield tree
            if self._file_extensions:
                # Format extensions for display (e.g., "*.json, *.csv")
                ext_display = ", ".join(
                    f"*{ext}" if ext.startswith(".") else f"*.{ext}"
                    for ext in self._file_extensions
                )
                with Horizontal(id="filter-row"):
                    yield Checkbox("Show all files", id="show-all-checkbox")
                    yield Static(f"({ext_display})", id="extension-display")
            yield Input(value=initial_filename, placeholder="Filename...", id="filename-input")
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

        # Set focus based on mode
        if self.mode == "open":
            # Focus on file tree for open dialog
            tree = self.query_one(".file-tree", FilteredDirectoryTree)
            tree.focus()
        else:
            # Focus on input for save dialog
            input_widget = self.query_one("#filename-input", Input)
            input_widget.focus()
            # Set cursor position after focus using timer
            if input_widget.value:
                from textual.widgets._input import Selection

                def set_cursor() -> None:
                    input_widget.selection = Selection(0, 0)
                    input_widget.cursor_position = 0

                # Use timer to ensure input is ready after focus
                self.set_timer(0.05, set_cursor)

    @on(Checkbox.Changed, "#show-all-checkbox")
    def on_show_all_changed(self, event: Checkbox.Changed) -> None:
        """Toggle showing all files in the directory tree."""
        tree = self.query_one(".file-tree", FilteredDirectoryTree)
        tree.set_show_all(event.value)

        # Update extension display
        ext_label = self.query_one("#extension-display", Static)
        if event.value:
            ext_label.update("(*)")
        else:
            if self._file_extensions:
                ext_display = ", ".join(
                    f"*{ext}" if ext.startswith(".") else f"*.{ext}"
                    for ext in self._file_extensions
                )
                ext_label.update(f"({ext_display})")
            else:
                ext_label.update("(*)")

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

    @on(Button.Pressed, "#go-up-btn")
    def on_go_up(self) -> None:
        """Navigate to the parent directory."""
        parent = self._current_path.parent
        if parent == self._current_path:
            # Already at root
            return
        self._current_path = parent
        self._replace_tree()

    def _replace_tree(self) -> None:
        """Replace the directory tree with a new one at the current path."""
        old_tree = self.query_one(".file-tree", FilteredDirectoryTree)
        container = self.query_one("#file-dialog-container", Container)

        # Update the path display
        path_label = self.query_one("#current-path", Static)
        path_label.update(str(self._current_path))

        # Get current show_all state if checkbox exists
        show_all = False
        try:
            checkbox = self.query_one("#show-all-checkbox", Checkbox)
            show_all = checkbox.value
        except Exception:
            pass

        # Remove old tree first
        old_tree.remove()

        # Generate unique ID for the new tree
        self._tree_counter += 1
        new_tree = FilteredDirectoryTree(
            self._current_path,
            extensions=self._file_extensions,
            id=f"file-tree-{self._tree_counter}",
            classes="file-tree",
        )
        new_tree.set_show_all(show_all)

        # Mount after the path-row
        path_row = self.query_one("#path-row", Horizontal)
        container.mount(new_tree, after=path_row)

"""Command input dialog.

Provides a modal dialog for entering commands like Goto cell reference,
column width values, and other text inputs.
"""

from __future__ import annotations

from typing import Any

from textual import on
from textual.app import ComposeResult
from textual.binding import Binding
from textual.containers import Container
from textual.screen import ModalScreen
from textual.widgets import Input, Label
from textual.widgets._input import Selection


class CommandInput(ModalScreen[str | None]):
    """Modal input for commands like Goto."""

    BINDINGS = [Binding("escape", "cancel", "Cancel")]

    CSS = """
    CommandInput {
        align: center middle;
    }

    #cmd-dialog-container {
        width: 80;
        height: auto;
        min-height: 7;
        padding: 1 2;
    }

    #cmd-prompt {
        width: 100%;
    }

    #cmd-input {
        margin-top: 1;
    }
    """

    def __init__(self, prompt: str, default: str = "", **kwargs: Any) -> None:
        super().__init__(**kwargs)
        self.prompt = prompt
        self.default = default

    def compose(self) -> ComposeResult:
        with Container(id="cmd-dialog-container"):
            yield Label(self.prompt, id="cmd-prompt")
            yield Input(value=self.default, id="cmd-input")

    def on_mount(self) -> None:
        # Get theme fresh from the app's current setting
        from ..themes import THEMES

        theme_type = self.app.current_theme_type  # type: ignore[attr-defined]
        theme = THEMES[theme_type]

        container = self.query_one("#cmd-dialog-container")
        container.styles.background = theme.cell_bg
        container.styles.border = ("thick", theme.accent)
        input_widget = self.query_one("#cmd-input", Input)
        input_widget.focus()
        # Select all text so typing replaces it
        if self.default:
            input_widget.selection = Selection(0, len(self.default))

    @on(Input.Submitted)
    def on_submit(self, event: Input.Submitted) -> None:
        self.dismiss(event.value)

    def action_cancel(self) -> None:
        self.dismiss(None)

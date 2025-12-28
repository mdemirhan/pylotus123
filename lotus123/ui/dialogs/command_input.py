"""Command input dialog.

Provides a modal dialog for entering commands like Goto cell reference,
column width values, and other text inputs.
"""
from __future__ import annotations

from typing import Any

from textual import on
from textual.binding import Binding
from textual.containers import Container
from textual.screen import ModalScreen
from textual.widgets import Input, Label
from textual.app import ComposeResult


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

    def __init__(self, prompt: str, **kwargs: Any) -> None:
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

"""File operation handler methods for LotusApp."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from ..ui import THEMES, CommandInput, FileDialog, ThemeDialog, ThemeType
from .base import BaseHandler

if TYPE_CHECKING:
    from .base import AppProtocol


class FileHandler(BaseHandler):
    """Handler for file operations (new, open, save, quit, theme)."""

    def __init__(self, app: "AppProtocol") -> None:
        super().__init__(app)

    def new_file(self) -> None:
        """Create a new empty spreadsheet."""
        self.spreadsheet.clear()
        self.spreadsheet.filename = ""
        self.undo_manager.clear()
        self._app._dirty = False
        grid = self.get_grid()
        grid.cursor_row = 0
        grid.cursor_col = 0
        grid.scroll_row = 0
        grid.scroll_col = 0
        grid.refresh_grid()
        self._app._update_status()
        self._app._update_title()
        self.notify("New spreadsheet created")

    def open_file(self) -> None:
        """Show the file open dialog."""
        self._app.push_screen(FileDialog(mode="open"), self._do_open)

    def load_initial_file(self, filepath: str) -> None:
        """Load a file specified at startup."""
        try:
            path = Path(filepath)
            if path.exists():
                self.spreadsheet.load(str(path))
                self.undo_manager.clear()
                self._app._dirty = False
                grid = self.get_grid()
                grid.refresh_grid()
                self._app._update_status()
                self._app._update_title()
                self._app.config.add_recent_file(str(path))
                self._app.config.save()
                self.notify(f"Loaded: {path}")
            else:
                self.notify(f"File not found: {filepath}", severity="error")
        except Exception as e:
            self.notify(f"Error loading file: {e}", severity="error")

    def _do_open(self, result: str | None) -> None:
        if result:
            try:
                self.spreadsheet.load(result)
                self.undo_manager.clear()
                self._app._dirty = False
                grid = self.get_grid()
                grid.scroll_row = 1
                grid.scroll_col = 1
                grid.scroll_row = 0
                grid.scroll_col = 0
                grid.cursor_row = 0
                grid.cursor_col = 0
                grid.recalculate_visible_area()
                grid.refresh_grid()
                self._app._update_status()
                self._app._update_title()
                self._app.config.add_recent_file(result)
                self._app.config.save()
                self.notify(f"Loaded: {result}")
            except Exception as e:
                self.notify(f"Error: {e}", severity="error")

    def save(self) -> None:
        """Save the current spreadsheet."""
        if self.spreadsheet.filename:
            self.spreadsheet.save(self.spreadsheet.filename)
            self._app._dirty = False
            self._app._update_title()
            self.notify(f"Saved: {self.spreadsheet.filename}")
        else:
            self.save_as()

    def save_as(self) -> None:
        """Show the save-as dialog."""
        self._app.push_screen(FileDialog(mode="save"), self._do_save)

    def _do_save(self, result: str | None) -> None:
        if result:
            try:
                if not result.endswith(".json"):
                    result += ".json"
                self.spreadsheet.save(result)
                self._app._dirty = False
                self._app._update_title()
                self.notify(f"Saved: {result}")
            except Exception as e:
                self.notify(f"Error: {e}", severity="error")

    def change_theme(self) -> None:
        """Show the theme selection dialog."""
        self._app.push_screen(
            ThemeDialog(self._app.current_theme_type), self._do_change_theme
        )

    def _do_change_theme(self, result: ThemeType | None) -> None:
        if result:
            self._app.current_theme_type = result
            self._app.color_theme = THEMES[result]
            self._app.sub_title = f"Theme: {self._app.color_theme.name}"
            self._app._apply_theme()
            self._app.config.theme = result.name
            self._app.config.save()
            self.notify(f"Theme changed to {self._app.color_theme.name}")

    def quit_app(self) -> None:
        """Quit the application, prompting to save if dirty."""
        if self._app._dirty:
            self._app.push_screen(
                CommandInput("Save changes before quitting? (Y/N/C=Cancel):"),
                self._do_quit_confirm,
            )
        else:
            self._app.config.save()
            self._app.exit()

    def _do_quit_confirm(self, result: str | None) -> None:
        if not result:
            return
        response = result.strip().upper()
        if response.startswith("Y"):
            if self.spreadsheet.filename:
                self.spreadsheet.save(self.spreadsheet.filename)
                self._app.config.save()
                self._app.exit()
            else:
                self._app.push_screen(FileDialog(mode="save"), self._do_save_and_quit)
        elif response.startswith("N"):
            self._app.config.save()
            self._app.exit()

    def _do_save_and_quit(self, result: str | None) -> None:
        if result:
            self.spreadsheet.save(result)
        self._app.config.save()
        self._app.exit()

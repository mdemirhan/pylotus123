"""Navigation handler methods for LotusApp."""

from ..ui import CommandInput
from .base import AppProtocol, BaseHandler


class NavigationHandler(BaseHandler):
    """Handler for navigation operations."""

    def __init__(self, app: AppProtocol) -> None:
        super().__init__(app)

    def scroll_half_down(self) -> None:
        """Scroll down by half a page (Ctrl+D)."""
        grid = self.get_grid()
        grid.move_cursor(grid.visible_rows // 2, 0)

    def scroll_half_up(self) -> None:
        """Scroll up by half a page (Ctrl+U)."""
        grid = self.get_grid()
        grid.move_cursor(-(grid.visible_rows // 2), 0)

    def page_up(self) -> None:
        """Move up by one page."""
        grid = self.get_grid()
        grid.move_cursor(-grid.visible_rows, 0)

    def page_down(self) -> None:
        """Move down by one page."""
        grid = self.get_grid()
        grid.move_cursor(grid.visible_rows, 0)

    def go_home(self) -> None:
        """Go to beginning of current row."""
        grid = self.get_grid()
        grid.cursor_col = 0
        grid.scroll_col = 0
        grid.refresh_grid()
        self.update_status()

    def go_end(self) -> None:
        """Go to end of current row (last used cell)."""
        grid = self.get_grid()
        # Find last used column in current row
        last_col = 0
        for col in range(self.spreadsheet.cols):
            cell = self.spreadsheet.get_cell_if_exists(grid.cursor_row, col)
            if cell and cell.raw_value:
                last_col = col
        grid.cursor_col = last_col
        grid.refresh_grid()
        self.update_status()

    def go_start(self) -> None:
        """Go to cell A1 (Ctrl+Home)."""
        grid = self.get_grid()
        grid.cursor_row = 0
        grid.cursor_col = 0
        grid.scroll_row = 0
        grid.scroll_col = 0
        grid.refresh_grid()
        self.update_status()

    def go_last(self) -> None:
        """Go to last used cell (Ctrl+End)."""
        grid = self.get_grid()
        # Find last used cell
        last_row, last_col = 0, 0
        for row in range(self.spreadsheet.rows):
            for col in range(self.spreadsheet.cols):
                cell = self.spreadsheet.get_cell_if_exists(row, col)
                if cell and cell.raw_value:
                    last_row = max(last_row, row)
                    last_col = max(last_col, col)
        grid.cursor_row = last_row
        grid.cursor_col = last_col
        grid.refresh_grid()
        self.update_status()

    def goto(self) -> None:
        """Show the goto dialog."""
        self._app.push_screen(CommandInput("Go to cell (e.g., A1):"), self._do_goto)

    def _do_goto(self, result: str | None) -> None:
        if result:
            self.get_grid().goto_cell(result.upper())
            self.update_status()

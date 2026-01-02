"""Lotus 1-2-3 Clone - Main TUI Application."""

from __future__ import annotations

from pathlib import Path

from textual import events, on
from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Container, Horizontal
from textual.css.query import NoMatches
from textual.widgets import Footer, Input, Static

from .charting import Chart, ChartType
from .core import Spreadsheet, make_cell_ref

# Handler classes
from .handlers import (
    ChartHandler,
    ClipboardHandler,
    DataHandler,
    FileHandler,
    ImportExportHandler,
    NavigationHandler,
    QueryHandler,
    RangeHandler,
    WorksheetHandler,
)

# UI components
from .ui import (
    THEMES,
    AppConfig,
    LotusMenu,
    Mode,
    SpreadsheetGrid,
    StatusBarWidget,
    get_theme_type,
)
from .utils.undo import (
    CellChangeCommand,
    RangeChangeCommand,
    UndoManager,
)


class LotusApp(App[None]):
    """Main Lotus 1-2-3 Clone Application."""

    BINDINGS = [
        # File operations
        Binding("ctrl+s", "save", "Save"),
        Binding("ctrl+o", "open_file", "Open"),
        Binding("ctrl+n", "new_file", "New"),
        # Navigation
        Binding("ctrl+g", "goto", "Goto"),
        Binding("ctrl+d", "scroll_half_down", "PgDn"),
        Binding("ctrl+u", "scroll_half_up", "PgUp"),
        # Edit operations
        Binding("ctrl+z", "undo", "Undo"),
        Binding("ctrl+y", "redo", "Redo"),
        Binding("ctrl+c", "copy", "Copy"),
        Binding("ctrl+x", "cut", "Cut"),
        Binding("ctrl+v", "paste", "Paste"),
        # UI
        Binding("ctrl+t", "change_theme", "Theme"),
        Binding("ctrl+q", "quit_app", "Quit"),
        # Function keys
        Binding("f2", "edit_cell", "Edit"),
        Binding("f4", "toggle_absolute", "Abs", show=False),
        Binding("f5", "goto", "Goto", show=False),
        Binding("f9", "recalculate", "Calc", show=False),
        # Navigation (hidden from footer)
        Binding("pageup", "page_up", "PgUp", show=False),
        Binding("pagedown", "page_down", "PgDn", show=False),
        Binding("home", "go_home", "Home", show=False),
        Binding("end", "go_end", "End", show=False),
        Binding("ctrl+home", "go_start", "Start", show=False),
        Binding("ctrl+end", "go_last", "Last", show=False),
        # Other
        Binding("delete", "clear_cell", "Clear", show=False),
        Binding("escape", "cancel_edit", "Cancel", show=False),
    ]

    CSS = """
    Screen {
        layout: vertical;
        background: $surface;
    }

    #menu-bar {
        height: 1;
        width: 100%;
    }

    #cell-input-container {
        dock: top;
        height: 3;
        width: 100%;
        padding: 0 1;
    }

    #cell-ref {
        width: 10;
        text-style: bold;
    }

    #cell-input {
        width: 1fr;
    }

    #grid-container {
        width: 100%;
        height: 1fr;
    }

    #grid {
        width: 100%;
        height: 100%;
    }

    #status-bar {
        height: 1;
        width: 100%;
    }

    Footer {
        height: auto;
    }

    ToastRack {
        margin-bottom: 2;
    }
    """

    def __init__(self, initial_file: str | None = None):
        super().__init__()
        self._initial_file = initial_file
        self.config = AppConfig.load()
        self.spreadsheet = Spreadsheet()
        self.current_theme_type = get_theme_type(self.config.theme)
        self.color_theme = THEMES[self.current_theme_type]
        self.editing = False
        self._menu_active = False
        self.undo_manager = UndoManager(max_history=100)
        self.recalc_mode = "auto"
        self.chart = Chart()
        # Global worksheet settings (public - shared across handlers)
        self.global_format_code = "G"
        self.global_label_prefix = "'"
        self.global_col_width = 10
        self.global_zero_display = True
        self._dirty = False

        # Initialize handlers with explicit dependency injection
        # Note: type: ignore needed because Textual's overloaded methods
        # can't be fully matched by Protocol signatures
        self._chart_handler = ChartHandler(self)  # type: ignore[arg-type]
        self._clipboard_handler = ClipboardHandler(self)  # type: ignore[arg-type]
        self._data_handler = DataHandler(self)  # type: ignore[arg-type]
        self._file_handler = FileHandler(self)  # type: ignore[arg-type]
        self._import_export_handler = ImportExportHandler(self)  # type: ignore[arg-type]
        self._navigation_handler = NavigationHandler(self)  # type: ignore[arg-type]
        self._query_handler = QueryHandler(self)  # type: ignore[arg-type]
        self._range_handler = RangeHandler(self)  # type: ignore[arg-type]
        self._worksheet_handler = WorksheetHandler(self)  # type: ignore[arg-type]

    @property
    def _has_modal(self) -> bool:
        """Check if a modal screen is currently open."""
        return len(self.screen_stack) > 1

    def _mark_dirty(self) -> None:
        """Mark the spreadsheet as having unsaved changes."""
        self._dirty = True
        self._update_title()

    def _generate_css(self) -> str:
        """Generate CSS based on current theme."""
        t = self.color_theme
        return f"""
        Screen {{
            background: {t.background};
        }}

        #menu-bar {{
            background: {t.menu_bg};
            color: {t.menu_fg};
        }}

        #cell-input-container {{
            background: {t.background};
        }}

        #cell-ref {{
            color: {t.accent};
        }}

        #cell-input {{
            background: {t.input_bg};
            color: {t.input_fg};
            border: solid {t.accent};
        }}

        #grid {{
            background: {t.cell_bg};
        }}

        #status-bar {{
            background: {t.status_bg};
            color: {t.status_fg};
        }}

        Footer {{
            background: {t.menu_bg};
            color: {t.menu_fg};
        }}

        /* Dialog styling */
        ModalScreen {{
            background: {t.background} 80%;
        }}

        #file-dialog-container, #cmd-dialog-container, #theme-dialog-container {{
            background: {t.background};
            border: thick {t.accent};
        }}

        #dialog-title, #cmd-prompt, #theme-title {{
            color: {t.accent};
        }}

        #theme-list {{
            background: {t.cell_bg};
            border: solid {t.border};
        }}

        #theme-list > ListItem {{
            color: {t.cell_fg};
            background: {t.cell_bg};
        }}

        #theme-list > ListItem:hover {{
            background: {t.header_bg};
            color: {t.header_fg};
        }}

        #theme-list > ListItem.-highlight {{
            background: {t.selected_bg};
            color: {t.selected_fg};
        }}
        """

    def compose(self) -> ComposeResult:
        yield LotusMenu(self.color_theme, id="menu-bar")
        with Horizontal(id="cell-input-container"):
            yield Static("A1:", id="cell-ref")
            yield Input(id="cell-input", placeholder="Enter value or formula...")
        with Container(id="grid-container"):
            yield SpreadsheetGrid(self.spreadsheet, self.color_theme, id="grid")
        yield StatusBarWidget(self.spreadsheet, id="status-bar")
        yield Footer()

    def on_mount(self) -> None:
        self._update_title()
        self.sub_title = f"Theme: {self.color_theme.name}"
        self._apply_theme()
        self._update_status()
        self.query_one("#grid", SpreadsheetGrid).focus()

        if self._initial_file:
            self._file_handler.load_initial_file(self._initial_file)

    def _update_title(self) -> None:
        """Update the window title with filename and dirty indicator."""
        filename = self.spreadsheet.filename or "Untitled"
        if "/" in filename or "\\" in filename:
            filename = Path(filename).name
        dirty_indicator = " *" if self._dirty else ""
        self.title = f"Lotus 1-2-3 Clone - {filename}{dirty_indicator}"

    def _apply_theme(self) -> None:
        """Apply the current theme to all widgets."""
        from textual.widgets import Footer

        t = THEMES[self.current_theme_type]
        self.color_theme = t
        self.stylesheet.add_source(self._generate_css())

        try:
            menu_bar = self.query_one("#menu-bar", LotusMenu)
            menu_bar.set_theme(t)

            self.query_one("#cell-ref").styles.color = t.accent
            self.query_one("#cell-input").styles.background = t.input_bg
            self.query_one("#cell-input").styles.color = t.input_fg
            self.query_one("#cell-input").styles.border = ("solid", t.accent)

            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.set_theme(t)

            self.query_one("#grid-container").styles.background = t.cell_bg

            self.query_one("#status-bar").styles.background = t.status_bg
            self.query_one("#status-bar").styles.color = t.status_fg

            self.query_one("#cell-input-container").styles.background = t.background

            # Update Footer (bottom bar with keyboard shortcuts)
            footer = self.query_one(Footer)
            footer.styles.background = t.menu_bg
            footer.styles.color = t.menu_fg
        except NoMatches:
            pass

    @on(SpreadsheetGrid.CellSelected)
    def on_cell_selected(self, event: SpreadsheetGrid.CellSelected) -> None:
        self._update_status()

    @on(SpreadsheetGrid.CellClicked)
    def on_cell_clicked(self, event: SpreadsheetGrid.CellClicked) -> None:
        self._update_status()
        if not self._menu_active:
            self.query_one("#grid", SpreadsheetGrid).focus()

    @on(LotusMenu.MenuItemSelected)
    def on_menu_item_selected(self, event: LotusMenu.MenuItemSelected) -> None:
        self._menu_active = False
        self._handle_menu(event.path)
        self.query_one("#grid", SpreadsheetGrid).focus()

    @on(LotusMenu.MenuActivated)
    def on_menu_activated(self, event: LotusMenu.MenuActivated) -> None:
        self._menu_active = True
        self.query_one("#status-bar", StatusBarWidget).set_mode(Mode.MENU)

    @on(LotusMenu.MenuDeactivated)
    def on_menu_deactivated(self, event: LotusMenu.MenuDeactivated) -> None:
        self._menu_active = False
        self.query_one("#status-bar", StatusBarWidget).set_mode(Mode.READY)
        self.query_one("#grid", SpreadsheetGrid).focus()

    def _update_status(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        ref = make_cell_ref(grid.cursor_row, grid.cursor_col)
        cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)

        self.query_one("#cell-ref", Static).update(f"{ref}:")

        status_bar = self.query_one("#status-bar", StatusBarWidget)
        status_bar.update_cell(grid.cursor_row, grid.cursor_col)
        status_bar.update_from_spreadsheet()
        status_bar.set_modified(self._dirty)  # Must be after update_from_spreadsheet

        if not self.editing:
            self.query_one("#cell-input", Input).value = cell.raw_value

    def action_edit_cell(self) -> None:
        if not self._menu_active:
            self.editing = True
            self.query_one("#status-bar", StatusBarWidget).set_mode(Mode.EDIT)
            cell_input = self.query_one("#cell-input", Input)
            cell_input.focus()

    def action_cancel_edit(self) -> None:
        if self.editing:
            self.editing = False
            self.query_one("#status-bar", StatusBarWidget).set_mode(Mode.READY)
            self._update_status()
            self.query_one("#grid", SpreadsheetGrid).focus()
        elif self._menu_active:
            menu = self.query_one("#menu-bar", LotusMenu)
            menu.deactivate()

    @on(Input.Submitted, "#cell-input")
    def on_cell_input_submitted(self, event: Input.Submitted) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)
        old_value = cell.raw_value
        cmd = CellChangeCommand(
            spreadsheet=self.spreadsheet,
            row=grid.cursor_row,
            col=grid.cursor_col,
            new_value=event.value,
            old_value=old_value,
        )
        self.undo_manager.execute(cmd)
        self._mark_dirty()
        self.editing = False
        self.query_one("#status-bar", StatusBarWidget).set_mode(Mode.READY)
        grid.refresh_grid()
        grid.move_cursor(1, 0)
        grid.focus()
        self._update_status()

    def action_clear_cell(self) -> None:
        if not self.editing and not self._menu_active:
            grid = self.query_one("#grid", SpreadsheetGrid)
            if grid.has_selection:
                r1, c1, r2, c2 = grid.selection_range
                changes = []
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        cell = self.spreadsheet.get_cell(r, c)
                        if cell.raw_value:
                            changes.append((r, c, "", cell.raw_value))
                if changes:
                    cmd = RangeChangeCommand(
                        spreadsheet=self.spreadsheet, changes=changes
                    )
                    self.undo_manager.execute(cmd)
                    self._mark_dirty()
                grid.clear_selection()
            else:
                cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)
                if cell.raw_value:
                    cell_cmd = CellChangeCommand(
                        spreadsheet=self.spreadsheet,
                        row=grid.cursor_row,
                        col=grid.cursor_col,
                        new_value="",
                        old_value=cell.raw_value,
                    )
                    self.undo_manager.execute(cell_cmd)
                    self._mark_dirty()
            grid.refresh_grid()
            self._update_status()

    def action_undo(self) -> None:
        if not self.editing and not self._menu_active:
            cmd = self.undo_manager.undo()
            if cmd:
                grid = self.query_one("#grid", SpreadsheetGrid)
                grid.refresh_grid()
                self._update_status()
                self._mark_dirty()
                self.notify(f"Undo: {cmd.description}")
            else:
                self.notify("Nothing to undo")

    def action_redo(self) -> None:
        if not self.editing and not self._menu_active:
            cmd = self.undo_manager.redo()
            if cmd:
                grid = self.query_one("#grid", SpreadsheetGrid)
                grid.refresh_grid()
                self._update_status()
                self._mark_dirty()
                self.notify(f"Redo: {cmd.description}")
            else:
                self.notify("Nothing to redo")

    def action_copy(self) -> None:
        if not self.editing and not self._menu_active:
            self._clipboard_handler.copy_cells()

    def action_cut(self) -> None:
        if not self.editing and not self._menu_active:
            self._clipboard_handler.cut_cells()

    def action_paste(self) -> None:
        if not self.editing and not self._menu_active:
            self._clipboard_handler.paste_cells()

    def action_toggle_absolute(self) -> None:
        if self.editing:
            cell_input = self.query_one("#cell-input", Input)
            value = cell_input.value
            if value.startswith("=") or value.startswith("@"):
                import re

                def toggle_ref(m: re.Match[str]) -> str:
                    ref = m.group(0)
                    if ref.startswith("$") and "$" in ref[1:]:
                        return ref[1:].replace("$", "", 1)
                    elif "$" in ref:
                        return ref.replace("$", "")
                    else:
                        match = re.match(r"([A-Za-z]+)", ref)
                        if match:
                            col = match.group(1)
                            row = ref[len(col) :]
                            return f"${col}${row}"
                        return ref

                new_value = re.sub(r"\$?[A-Za-z]+\$?\d+", toggle_ref, value)
                cell_input.value = new_value

    def action_recalculate(self) -> None:
        self.spreadsheet.invalidate_cache()
        grid = self.query_one("#grid", SpreadsheetGrid)
        grid.refresh_grid()
        self._update_status()
        self.notify("Recalculated")

    def action_show_menu(self) -> None:
        if not self.editing and not self._has_modal:
            menu = self.query_one("#menu-bar", LotusMenu)
            menu.activate()

    # Navigation actions - delegate to handler
    def action_scroll_half_down(self) -> None:
        if not self.editing and not self._menu_active:
            self._navigation_handler.scroll_half_down()

    def action_scroll_half_up(self) -> None:
        if not self.editing and not self._menu_active:
            self._navigation_handler.scroll_half_up()

    def action_page_up(self) -> None:
        if not self.editing and not self._menu_active:
            self._navigation_handler.page_up()

    def action_page_down(self) -> None:
        if not self.editing and not self._menu_active:
            self._navigation_handler.page_down()

    def action_go_home(self) -> None:
        if not self.editing and not self._menu_active:
            self._navigation_handler.go_home()

    def action_go_end(self) -> None:
        if not self.editing and not self._menu_active:
            self._navigation_handler.go_end()

    def action_go_start(self) -> None:
        if not self.editing and not self._menu_active:
            self._navigation_handler.go_start()

    def action_go_last(self) -> None:
        if not self.editing and not self._menu_active:
            self._navigation_handler.go_last()

    def action_goto(self) -> None:
        self._navigation_handler.goto()

    # File actions - delegate to handler
    def action_new_file(self) -> None:
        self._file_handler.new_file()

    def action_open_file(self) -> None:
        self._file_handler.open_file()

    def action_save(self) -> None:
        self._file_handler.save()

    def action_change_theme(self) -> None:
        self._file_handler.change_theme()

    def action_quit_app(self) -> None:
        self._file_handler.quit_app()

    def _handle_menu(self, result: str | None) -> None:
        if not result:
            return

        # Navigation
        if result == "Goto" or result == "Worksheet:Goto":
            self._navigation_handler.goto()
        # Clipboard
        elif result == "Copy":
            self._clipboard_handler.menu_copy()
        elif result == "Move":
            self._clipboard_handler.menu_move()
        # File
        elif result == "File:New":
            self._file_handler.new_file()
        elif result == "File:Retrieve":
            self._file_handler.open_file()
        elif result == "File:Save":
            self._file_handler.save()
        elif result == "File:Save As":
            self._file_handler.save_as()
        elif result == "File:Quit" or result == "Quit:Yes":
            self._file_handler.quit_app()
        elif result == "Quit:No":
            pass
        # Range
        elif result == "Range:Erase":
            self.action_clear_cell()
        elif result == "Range:Format":
            self._range_handler.range_format()
        elif result == "Range:Label":
            self._range_handler.range_label()
        elif result == "Range:Name":
            self._range_handler.range_name()
        # Data
        elif result == "Data:Fill":
            self._data_handler.data_fill()
        elif result == "Data:Sort":
            self._data_handler.data_sort()
        # Query
        elif result == "Data:Query:Input":
            self._query_handler.set_input()
        elif result == "Data:Query:Criteria":
            self._query_handler.set_criteria()
        elif result == "Data:Query:Output":
            self._query_handler.set_output()
        elif result == "Data:Query:Find":
            self._query_handler.find()
        elif result == "Data:Query:Extract":
            self._query_handler.extract()
        elif result == "Data:Query:Unique":
            self._query_handler.unique()
        elif result == "Data:Query:Delete":
            self._query_handler.delete()
        elif result == "Data:Query:Reset":
            self._query_handler.reset()
        # Worksheet
        elif result == "Worksheet:Insert:Rows":
            self._worksheet_handler.insert_rows()
        elif result == "Worksheet:Insert:Columns":
            self._worksheet_handler.insert_columns()
        elif result == "Worksheet:Delete:Rows":
            self._worksheet_handler.delete_rows()
        elif result == "Worksheet:Delete:Columns":
            self._worksheet_handler.delete_columns()
        elif result == "Worksheet:Column":
            self._worksheet_handler.set_column_width()
        elif result == "Worksheet:Erase":
            self._worksheet_handler.worksheet_erase()
        elif result == "Worksheet:Global:Format":
            self._worksheet_handler.global_format()
        elif result == "Worksheet:Global:Label-Prefix":
            self._worksheet_handler.global_label_prefix()
        elif result == "Worksheet:Global:Column-Width":
            self._worksheet_handler.global_column_width()
        elif result == "Worksheet:Global:Recalculation":
            self._worksheet_handler.global_recalculation()
        elif result == "Worksheet:Global:Zero":
            self._worksheet_handler.global_zero()
        # Graph/Chart
        elif result == "Graph:Type:Line":
            self._chart_handler.set_chart_type(ChartType.LINE)
        elif result == "Graph:Type:Bar":
            self._chart_handler.set_chart_type(ChartType.BAR)
        elif result == "Graph:Type:XY":
            self._chart_handler.set_chart_type(ChartType.XY_SCATTER)
        elif result == "Graph:Type:Stacked":
            self._chart_handler.set_chart_type(ChartType.STACKED_BAR)
        elif result == "Graph:Type:Pie":
            self._chart_handler.set_chart_type(ChartType.PIE)
        elif result == "Graph:X-Range":
            self._chart_handler.set_x_range()
        elif result == "Graph:A-Range":
            self._chart_handler.set_a_range()
        elif result == "Graph:B-Range":
            self._chart_handler.set_b_range()
        elif result == "Graph:C-Range":
            self._chart_handler.set_c_range()
        elif result == "Graph:D-Range":
            self._chart_handler.set_d_range()
        elif result == "Graph:E-Range":
            self._chart_handler.set_e_range()
        elif result == "Graph:F-Range":
            self._chart_handler.set_f_range()
        elif result == "Graph:View":
            self._chart_handler.view_chart()
        elif result == "Graph:Reset":
            self._chart_handler.reset_chart()
        elif result == "Graph:Save":
            self._chart_handler.save_chart()
        elif result == "Graph:Load":
            self._chart_handler.load_chart()
        # Import/Export
        elif result == "File:Import:CSV":
            self._import_export_handler.import_csv()
        elif result == "File:Import:TSV":
            self._import_export_handler.import_tsv()
        elif result == "File:Import:WK1":
            self._import_export_handler.import_wk1()
        elif result == "File:Import:XLSX":
            self._import_export_handler.import_xlsx()
        elif result == "File:Export:CSV":
            self._import_export_handler.export_csv()
        elif result == "File:Export:TSV":
            self._import_export_handler.export_tsv()
        elif result == "File:Export:WK1":
            self._import_export_handler.export_wk1()
        elif result == "File:Export:XLSX":
            self._import_export_handler.export_xlsx()
        # System menu
        elif result == "System:Theme":
            self._file_handler.change_theme()

    def on_key(self, event: events.Key) -> None:
        """Handle key presses for navigation and direct cell input."""
        if self._has_modal:
            return

        if event.key == "slash" or event.character == "/":
            if not self.editing:
                self.action_show_menu()
                event.prevent_default()
                event.stop()
                return

        if self.editing or self._menu_active:
            return

        grid = self.query_one("#grid", SpreadsheetGrid)

        # Shift+Arrow for range selection
        if event.key == "shift+up":
            if not grid.has_selection:
                grid.start_selection()
            grid.move_cursor(-1, 0)
            event.prevent_default()
            return
        elif event.key == "shift+down":
            if not grid.has_selection:
                grid.start_selection()
            grid.move_cursor(1, 0)
            event.prevent_default()
            return
        elif event.key == "shift+left":
            if not grid.has_selection:
                grid.start_selection()
            grid.move_cursor(0, -1)
            event.prevent_default()
            return
        elif event.key == "shift+right":
            if not grid.has_selection:
                grid.start_selection()
            grid.move_cursor(0, 1)
            event.prevent_default()
            return

        # Regular arrow keys
        if event.key == "up":
            grid.clear_selection()
            grid.move_cursor(-1, 0)
            event.prevent_default()
            return
        elif event.key == "down":
            grid.clear_selection()
            grid.move_cursor(1, 0)
            event.prevent_default()
            return
        elif event.key == "left":
            grid.clear_selection()
            grid.move_cursor(0, -1)
            event.prevent_default()
            return
        elif event.key == "right":
            grid.clear_selection()
            grid.move_cursor(0, 1)
            event.prevent_default()
            return
        elif event.key == "enter":
            self.action_edit_cell()
            event.prevent_default()
            return
        elif event.key in ("delete", "backspace"):
            self.action_clear_cell()
            event.prevent_default()
            return

        # Start editing on printable character
        if event.character and event.character.isprintable() and event.character != "/":
            cell_input = self.query_one("#cell-input", Input)
            cell_input.select_on_focus = False
            cell_input.value = ""
            cell_input.focus()
            self.editing = True
            if event.character in "0123456789.+-@#(":
                self.query_one("#status-bar", StatusBarWidget).set_mode(Mode.VALUE)
            elif event.character == "=":
                self.query_one("#status-bar", StatusBarWidget).set_mode(Mode.VALUE)
            else:
                self.query_one("#status-bar", StatusBarWidget).set_mode(Mode.LABEL)
            cell_input.insert_text_at_cursor(event.character)
            event.prevent_default()


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(
        description="Lotus 1-2-3 Clone - A terminal-based spreadsheet application"
    )
    parser.add_argument("file", nargs="?", help="Spreadsheet file to open on startup")
    args = parser.parse_args()
    app = LotusApp(initial_file=args.file)
    app.run()


if __name__ == "__main__":
    main()

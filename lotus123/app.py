"""Lotus 1-2-3 Clone - Main TUI Application."""
from __future__ import annotations

from pathlib import Path

from textual import on, events
from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Container, Horizontal
from textual.widgets import Static, Input, Footer
from textual.css.query import NoMatches

from .core.spreadsheet import Spreadsheet
from .core.reference import index_to_col, make_cell_ref, parse_cell_ref, col_to_index
from .core.reference import adjust_formula_references
from .utils.undo import (
    UndoManager, CellChangeCommand, RangeChangeCommand,
    InsertRowCommand, DeleteRowCommand
)
from .charting.chart import Chart, ChartType
from .charting.renderer import TextChartRenderer

# UI components
from .ui import (
    Theme, ThemeType, THEMES, get_theme_type,
    AppConfig,
    SpreadsheetGrid,
    LotusMenu,
    StatusBarWidget, Mode,
    FileDialog, CommandInput, ThemeDialog, ChartViewScreen,
)


class LotusApp(App):
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
        self._cell_clipboard: tuple[int, int, str] | None = None
        self._range_clipboard: list[list[str]] | None = None
        self._clipboard_is_cut = False
        self._menu_active = False
        self.undo_manager = UndoManager(max_history=100)
        self._recalc_mode = "auto"
        self.chart = Chart()
        self._chart_renderer = TextChartRenderer(self.spreadsheet)
        # Global worksheet settings
        self._global_format_code = "G"
        self._global_label_prefix = "'"
        self._global_col_width = 10
        self._global_protection = False
        self._global_zero_display = True
        self._dirty = False

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
            self._load_initial_file()

    def _update_title(self) -> None:
        """Update the window title with filename and dirty indicator."""
        filename = self.spreadsheet.filename or "Untitled"
        if "/" in filename or "\\" in filename:
            filename = Path(filename).name
        dirty_indicator = " *" if self._dirty else ""
        self.title = f"Lotus 1-2-3 Clone - {filename}{dirty_indicator}"

    def _apply_theme(self) -> None:
        """Apply the current theme to all widgets."""
        t = self.color_theme
        self.stylesheet.add_source(self._generate_css(), "theme")

        try:
            menu_bar = self.query_one("#menu-bar", LotusMenu)
            menu_bar.set_theme(t)

            self.query_one("#cell-ref").styles.color = t.accent
            self.query_one("#cell-input").styles.background = t.input_bg
            self.query_one("#cell-input").styles.color = t.input_fg
            self.query_one("#cell-input").styles.border = ("solid", t.accent)

            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.set_theme(t)

            self.query_one("#status-bar").styles.background = t.status_bg
            self.query_one("#status-bar").styles.color = t.status_fg

            self.query_one("#cell-input-container").styles.background = t.background
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
        self._handle_menu(event.path)
        self._menu_active = False
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

        # Update status bar widget
        status_bar = self.query_one("#status-bar", StatusBarWidget)
        status_bar.update_cell(grid.cursor_row, grid.cursor_col)
        status_bar.set_modified(self._dirty)
        status_bar.update_from_spreadsheet()

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
                    cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                    self.undo_manager.execute(cmd)
                grid.clear_selection()
            else:
                cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)
                if cell.raw_value:
                    cmd = CellChangeCommand(
                        spreadsheet=self.spreadsheet,
                        row=grid.cursor_row,
                        col=grid.cursor_col,
                        new_value="",
                        old_value=cell.raw_value,
                    )
                    self.undo_manager.execute(cmd)
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
            self._copy_cells()

    def action_cut(self) -> None:
        if not self.editing and not self._menu_active:
            self._cut_cells()

    def action_paste(self) -> None:
        if not self.editing and not self._menu_active:
            self._paste_cells()

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
                        col = re.match(r"([A-Za-z]+)", ref).group(1)
                        row = ref[len(col):]
                        return f"${col}${row}"
                new_value = re.sub(r'\$?[A-Za-z]+\$?\d+', toggle_ref, value)
                cell_input.value = new_value

    def action_recalculate(self) -> None:
        self.spreadsheet._invalidate_cache()
        grid = self.query_one("#grid", SpreadsheetGrid)
        grid.refresh_grid()
        self._update_status()
        self.notify("Recalculated")

    # Navigation actions (for key bindings)
    def action_scroll_half_down(self) -> None:
        """Scroll down by half a page (Ctrl+D)."""
        if not self.editing and not self._menu_active:
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.move_cursor(grid.visible_rows // 2, 0)

    def action_scroll_half_up(self) -> None:
        """Scroll up by half a page (Ctrl+U)."""
        if not self.editing and not self._menu_active:
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.move_cursor(-(grid.visible_rows // 2), 0)

    def action_page_up(self) -> None:
        """Move up by one page."""
        if not self.editing and not self._menu_active:
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.move_cursor(-grid.visible_rows, 0)

    def action_page_down(self) -> None:
        """Move down by one page."""
        if not self.editing and not self._menu_active:
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.move_cursor(grid.visible_rows, 0)

    def action_go_home(self) -> None:
        """Go to beginning of current row."""
        if not self.editing and not self._menu_active:
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.cursor_col = 0
            grid.scroll_col = 0
            grid.refresh_grid()
            self._update_status()

    def action_go_end(self) -> None:
        """Go to end of current row (last used cell)."""
        if not self.editing and not self._menu_active:
            grid = self.query_one("#grid", SpreadsheetGrid)
            # Find last used column in current row
            last_col = 0
            for col in range(self.spreadsheet.cols):
                cell = self.spreadsheet.get_cell_if_exists(grid.cursor_row, col)
                if cell and cell.raw_value:
                    last_col = col
            grid.cursor_col = last_col
            grid.refresh_grid()
            self._update_status()

    def action_go_start(self) -> None:
        """Go to cell A1 (Ctrl+Home)."""
        if not self.editing and not self._menu_active:
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.cursor_row = 0
            grid.cursor_col = 0
            grid.scroll_row = 0
            grid.scroll_col = 0
            grid.refresh_grid()
            self._update_status()

    def action_go_last(self) -> None:
        """Go to last used cell (Ctrl+End)."""
        if not self.editing and not self._menu_active:
            grid = self.query_one("#grid", SpreadsheetGrid)
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
            self._update_status()

    def action_show_menu(self) -> None:
        if not self.editing and not self._has_modal:
            menu = self.query_one("#menu-bar", LotusMenu)
            menu.activate()

    def _handle_menu(self, result: str | None) -> None:
        if not result:
            return

        if result == "Goto" or result == "Worksheet:Goto":
            self.action_goto()
        elif result == "Copy":
            self._menu_copy()
        elif result == "Move":
            self._menu_move()
        elif result == "File:New":
            self.action_new_file()
        elif result == "File:Retrieve":
            self.action_open_file()
        elif result == "File:Save":
            self.action_save()
        elif result == "File:Quit" or result == "Quit:Yes":
            self.action_quit_app()
        elif result == "Quit:No":
            pass
        elif result == "Range:Erase":
            self.action_clear_cell()
        elif result == "Range:Format":
            self._range_format()
        elif result == "Range:Label":
            self._range_label()
        elif result == "Range:Name":
            self._range_name()
        elif result == "Range:Protect":
            self._range_protect()
        elif result == "Data:Fill":
            self._data_fill()
        elif result == "Data:Sort":
            self._data_sort()
        elif result == "Data:Query":
            self._data_query()
        elif result == "Worksheet:Insert":
            self._insert_row()
        elif result == "Worksheet:Delete":
            self._delete_row()
        elif result == "Worksheet:Column":
            self._set_column_width()
        elif result == "Worksheet:Erase":
            self._worksheet_erase()
        elif result == "Worksheet:Global:Format":
            self._global_format()
        elif result == "Worksheet:Global:Label-Prefix":
            self._global_label_prefix_action()
        elif result == "Worksheet:Global:Column-Width":
            self._global_column_width()
        elif result == "Worksheet:Global:Recalculation":
            self._global_recalculation()
        elif result == "Worksheet:Global:Protection":
            self._global_protection_action()
        elif result == "Worksheet:Global:Zero":
            self._global_zero()
        elif result == "Graph:Type:Line":
            self._set_chart_type(ChartType.LINE)
        elif result == "Graph:Type:Bar":
            self._set_chart_type(ChartType.BAR)
        elif result == "Graph:Type:XY":
            self._set_chart_type(ChartType.XY_SCATTER)
        elif result == "Graph:Type:Stacked":
            self._set_chart_type(ChartType.STACKED_BAR)
        elif result == "Graph:Type:Pie":
            self._set_chart_type(ChartType.PIE)
        elif result == "Graph:X-Range":
            self._set_chart_x_range()
        elif result == "Graph:A-Range":
            self._set_chart_a_range()
        elif result == "Graph:B-Range":
            self._set_chart_b_range()
        elif result == "Graph:View":
            self._view_chart()
        elif result == "Graph:Reset":
            self._reset_chart()

    def action_goto(self) -> None:
        self.push_screen(CommandInput("Go to cell (e.g., A1):"), self._do_goto)

    def _do_goto(self, result: str | None) -> None:
        if result:
            self.query_one("#grid", SpreadsheetGrid).goto_cell(result.upper())
            self._update_status()

    def _set_column_width(self) -> None:
        self.push_screen(CommandInput("Column width (3-50):"), self._do_set_width)

    def _do_set_width(self, result: str | None) -> None:
        if result:
            try:
                width = int(result)
                grid = self.query_one("#grid", SpreadsheetGrid)
                self.spreadsheet.set_col_width(grid.cursor_col, width)
                grid.refresh_grid()
            except ValueError:
                pass

    def action_new_file(self) -> None:
        self.spreadsheet.clear()
        self.spreadsheet.filename = ""
        self.undo_manager.clear()
        self._dirty = False
        grid = self.query_one("#grid", SpreadsheetGrid)
        grid.cursor_row = 0
        grid.cursor_col = 0
        grid.scroll_row = 0
        grid.scroll_col = 0
        grid.refresh_grid()
        self._update_status()
        self._update_title()
        self.notify("New spreadsheet created")

    def action_open_file(self) -> None:
        self.push_screen(FileDialog(mode="open"), self._do_open)

    def _load_initial_file(self) -> None:
        try:
            filepath = Path(self._initial_file)
            if filepath.exists():
                self.spreadsheet.load(str(filepath))
                self.undo_manager.clear()
                self._dirty = False
                grid = self.query_one("#grid", SpreadsheetGrid)
                grid.refresh_grid()
                self._update_status()
                self._update_title()
                self.config.add_recent_file(str(filepath))
                self.config.save()
                self.notify(f"Loaded: {filepath}")
            else:
                self.notify(f"File not found: {self._initial_file}", severity="error")
        except Exception as e:
            self.notify(f"Error loading file: {e}", severity="error")

    def _do_open(self, result: str | None) -> None:
        if result:
            try:
                self.spreadsheet.load(result)
                self.undo_manager.clear()
                self._dirty = False
                grid = self.query_one("#grid", SpreadsheetGrid)
                grid.scroll_row = 1
                grid.scroll_col = 1
                grid.scroll_row = 0
                grid.scroll_col = 0
                grid.cursor_row = 0
                grid.cursor_col = 0
                grid._calculate_visible_area()
                grid.refresh_grid()
                self._update_status()
                self._update_title()
                self.config.add_recent_file(result)
                self.config.save()
                self.notify(f"Loaded: {result}")
            except Exception as e:
                self.notify(f"Error: {e}", severity="error")

    def action_save(self) -> None:
        if self.spreadsheet.filename:
            self.spreadsheet.save(self.spreadsheet.filename)
            self._dirty = False
            self._update_title()
            self.notify(f"Saved: {self.spreadsheet.filename}")
        else:
            self._save_as()

    def _save_as(self) -> None:
        self.push_screen(FileDialog(mode="save"), self._do_save)

    def _do_save(self, result: str | None) -> None:
        if result:
            try:
                if not result.endswith('.json'):
                    result += '.json'
                self.spreadsheet.save(result)
                self._dirty = False
                self._update_title()
                self.notify(f"Saved: {result}")
            except Exception as e:
                self.notify(f"Error: {e}", severity="error")

    def action_change_theme(self) -> None:
        self.push_screen(ThemeDialog(self.current_theme_type), self._do_change_theme)

    def _do_change_theme(self, result: ThemeType | None) -> None:
        if result:
            self.current_theme_type = result
            self.color_theme = THEMES[result]
            self.sub_title = f"Theme: {self.color_theme.name}"
            self._apply_theme()
            self.config.theme = result.name
            self.config.save()
            self.notify(f"Theme changed to {self.color_theme.name}")

    def action_quit_app(self) -> None:
        if self._dirty:
            self.push_screen(
                CommandInput("Save changes before quitting? (Y/N/C=Cancel):"),
                self._do_quit_confirm
            )
        else:
            self.config.save()
            self.exit()

    def _do_quit_confirm(self, result: str | None) -> None:
        if not result:
            return
        response = result.strip().upper()
        if response.startswith("Y"):
            if self.spreadsheet.filename:
                self.spreadsheet.save(self.spreadsheet.filename)
                self.config.save()
                self.exit()
            else:
                self.push_screen(FileDialog(mode="save"), self._do_save_and_quit)
        elif response.startswith("N"):
            self.config.save()
            self.exit()

    def _do_save_and_quit(self, result: str | None) -> None:
        if result:
            self.spreadsheet.save(result)
        self.config.save()
        self.exit()

    def _menu_copy(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        source_range = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
        self._pending_source_range = (r1, c1, r2, c2)
        self.push_screen(
            CommandInput(f"Copy {source_range} TO (e.g., D1):"),
            self._do_menu_copy
        )

    def _do_menu_copy(self, result: str | None) -> None:
        if not result:
            return
        try:
            dest_row, dest_col = parse_cell_ref(result.upper())
            r1, c1, r2, c2 = self._pending_source_range
            changes = []
            for r_offset in range(r2 - r1 + 1):
                for c_offset in range(c2 - c1 + 1):
                    src_row, src_col = r1 + r_offset, c1 + c_offset
                    target_row, target_col = dest_row + r_offset, dest_col + c_offset
                    if target_row >= self.spreadsheet.rows or target_col >= self.spreadsheet.cols:
                        continue
                    src_cell = self.spreadsheet.get_cell(src_row, src_col)
                    target_cell = self.spreadsheet.get_cell(target_row, target_col)
                    old_value = target_cell.raw_value
                    new_value = src_cell.raw_value
                    if new_value and (new_value.startswith("=") or new_value.startswith("@")):
                        row_delta = target_row - src_row
                        col_delta = target_col - src_col
                        new_value = new_value[0] + adjust_formula_references(
                            new_value[1:], row_delta, col_delta
                        )
                    if new_value != old_value:
                        changes.append((target_row, target_col, new_value, old_value))
            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid = self.query_one("#grid", SpreadsheetGrid)
                grid.refresh_grid()
                self._update_status()
                self.notify(f"Copied {len(changes)} cell(s)")
        except ValueError as e:
            self.notify(f"Invalid destination: {e}", severity="error")

    def _menu_move(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        source_range = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
        self._pending_source_range = (r1, c1, r2, c2)
        self.push_screen(
            CommandInput(f"Move {source_range} TO (e.g., D1):"),
            self._do_menu_move
        )

    def _do_menu_move(self, result: str | None) -> None:
        if not result:
            return
        try:
            dest_row, dest_col = parse_cell_ref(result.upper())
            r1, c1, r2, c2 = self._pending_source_range
            changes = []
            for r_offset in range(r2 - r1 + 1):
                for c_offset in range(c2 - c1 + 1):
                    src_row, src_col = r1 + r_offset, c1 + c_offset
                    target_row, target_col = dest_row + r_offset, dest_col + c_offset
                    if target_row >= self.spreadsheet.rows or target_col >= self.spreadsheet.cols:
                        continue
                    src_cell = self.spreadsheet.get_cell(src_row, src_col)
                    target_cell = self.spreadsheet.get_cell(target_row, target_col)
                    old_value = target_cell.raw_value
                    new_value = src_cell.raw_value
                    if new_value and (new_value.startswith("=") or new_value.startswith("@")):
                        row_delta = target_row - src_row
                        col_delta = target_col - src_col
                        new_value = new_value[0] + adjust_formula_references(
                            new_value[1:], row_delta, col_delta
                        )
                    if new_value != old_value:
                        changes.append((target_row, target_col, new_value, old_value))
            for r_offset in range(r2 - r1 + 1):
                for c_offset in range(c2 - c1 + 1):
                    src_row, src_col = r1 + r_offset, c1 + c_offset
                    target_row, target_col = dest_row + r_offset, dest_col + c_offset
                    if src_row == target_row and src_col == target_col:
                        continue
                    src_cell = self.spreadsheet.get_cell(src_row, src_col)
                    if src_cell.raw_value:
                        changes.append((src_row, src_col, "", src_cell.raw_value))
            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid = self.query_one("#grid", SpreadsheetGrid)
                grid.clear_selection()
                grid.cursor_row = dest_row
                grid.cursor_col = dest_col
                grid.refresh_grid()
                self._update_status()
                self.notify(f"Moved cells to {make_cell_ref(dest_row, dest_col)}")
        except ValueError as e:
            self.notify(f"Invalid destination: {e}", severity="error")

    def _copy_cells(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        self._range_clipboard = []
        self._clipboard_origin = (r1, c1)
        for r in range(r1, r2 + 1):
            row_data = []
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                row_data.append(cell.raw_value)
            self._range_clipboard.append(row_data)
        self._clipboard_is_cut = False
        cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)
        self._cell_clipboard = (grid.cursor_row, grid.cursor_col, cell.raw_value)
        cells_count = (r2 - r1 + 1) * (c2 - c1 + 1)
        self.notify(f"Copied {cells_count} cell(s)")

    def _cut_cells(self) -> None:
        self._copy_cells()
        self._clipboard_is_cut = True
        self.notify("Cut to clipboard")

    def _paste_cells(self) -> None:
        if not self._range_clipboard:
            if self._cell_clipboard:
                grid = self.query_one("#grid", SpreadsheetGrid)
                cell = self.spreadsheet.get_cell(grid.cursor_row, grid.cursor_col)
                old_value = cell.raw_value
                new_value = self._cell_clipboard[2]
                if new_value.startswith("=") or new_value.startswith("@"):
                    row_delta = grid.cursor_row - self._cell_clipboard[0]
                    col_delta = grid.cursor_col - self._cell_clipboard[1]
                    new_value = new_value[0] + adjust_formula_references(
                        new_value[1:], row_delta, col_delta
                    )
                cmd = CellChangeCommand(
                    spreadsheet=self.spreadsheet,
                    row=grid.cursor_row,
                    col=grid.cursor_col,
                    new_value=new_value,
                    old_value=old_value,
                )
                self.undo_manager.execute(cmd)
                grid.refresh_grid()
                self._update_status()
                self.notify("Pasted")
            return

        grid = self.query_one("#grid", SpreadsheetGrid)
        dest_row, dest_col = grid.cursor_row, grid.cursor_col
        src_row, src_col = getattr(self, '_clipboard_origin', (0, 0))
        changes = []
        for r_offset, row_data in enumerate(self._range_clipboard):
            for c_offset, value in enumerate(row_data):
                target_row = dest_row + r_offset
                target_col = dest_col + c_offset
                if target_row >= self.spreadsheet.rows or target_col >= self.spreadsheet.cols:
                    continue
                cell = self.spreadsheet.get_cell(target_row, target_col)
                old_value = cell.raw_value
                new_value = value
                if new_value and (new_value.startswith("=") or new_value.startswith("@")):
                    row_delta = target_row - (src_row + r_offset)
                    col_delta = target_col - (src_col + c_offset)
                    new_value = new_value[0] + adjust_formula_references(
                        new_value[1:], row_delta, col_delta
                    )
                if new_value != old_value:
                    changes.append((target_row, target_col, new_value, old_value))
        if changes:
            cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
            self.undo_manager.execute(cmd)
        if self._clipboard_is_cut:
            clear_changes = []
            for r_offset, row_data in enumerate(self._range_clipboard):
                for c_offset, value in enumerate(row_data):
                    if value:
                        clear_changes.append((src_row + r_offset, src_col + c_offset, "", value))
            if clear_changes:
                clear_cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=clear_changes)
                self.undo_manager.execute(clear_cmd)
            self._clipboard_is_cut = False
        grid.refresh_grid()
        self._update_status()
        cells_count = len(self._range_clipboard) * len(self._range_clipboard[0]) if self._range_clipboard else 0
        self.notify(f"Pasted {cells_count} cell(s)")

    def _insert_row(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        cmd = InsertRowCommand(spreadsheet=self.spreadsheet, row=grid.cursor_row)
        self.undo_manager.execute(cmd)
        grid.refresh_grid()
        self.notify(f"Row {grid.cursor_row + 1} inserted")

    def _delete_row(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        cmd = DeleteRowCommand(spreadsheet=self.spreadsheet, row=grid.cursor_row)
        self.undo_manager.execute(cmd)
        grid.refresh_grid()
        self._update_status()
        self.notify(f"Row {grid.cursor_row + 1} deleted")

    # Chart methods
    def _set_chart_type(self, chart_type: ChartType) -> None:
        self.chart.set_type(chart_type)
        type_names = {
            ChartType.LINE: "Line", ChartType.BAR: "Bar",
            ChartType.XY_SCATTER: "XY Scatter", ChartType.STACKED_BAR: "Stacked Bar",
            ChartType.PIE: "Pie",
        }
        self.notify(f"Chart type set to {type_names.get(chart_type, 'Unknown')}")

    def _set_chart_x_range(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self.chart.set_x_range(range_str)
            self.notify(f"X-Range set to {range_str}")
        else:
            self.push_screen(CommandInput("X-Range (e.g., A1:A10):"), self._do_set_x_range)

    def _do_set_x_range(self, result: str | None) -> None:
        if result:
            self.chart.set_x_range(result.upper())
            self.notify(f"X-Range set to {result.upper()}")

    def _set_chart_a_range(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self._add_or_update_series(0, "A", range_str)
        else:
            self.push_screen(CommandInput("A-Range (e.g., B1:B10):"), self._do_set_a_range)

    def _do_set_a_range(self, result: str | None) -> None:
        if result:
            self._add_or_update_series(0, "A", result.upper())

    def _set_chart_b_range(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        if grid.has_selection:
            r1, c1, r2, c2 = grid.selection_range
            range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
            self._add_or_update_series(1, "B", range_str)
        else:
            self.push_screen(CommandInput("B-Range (e.g., C1:C10):"), self._do_set_b_range)

    def _do_set_b_range(self, result: str | None) -> None:
        if result:
            self._add_or_update_series(1, "B", result.upper())

    def _add_or_update_series(self, index: int, name: str, range_str: str) -> None:
        while len(self.chart.series) <= index:
            self.chart.add_series(f"Series {len(self.chart.series) + 1}")
        self.chart.series[index].name = name
        self.chart.series[index].data_range = range_str
        self.notify(f"{name}-Range set to {range_str}")

    def _view_chart(self) -> None:
        if not self.chart.series:
            self.notify("No data series defined. Use A-Range to set data.")
            return
        self._chart_renderer.spreadsheet = self.spreadsheet
        chart_lines = self._chart_renderer.render(self.chart, width=70, height=20)
        self.push_screen(ChartViewScreen(chart_lines))

    def _reset_chart(self) -> None:
        self.chart.reset()
        self.notify("Chart reset")

    # Range menu methods
    def _range_format(self) -> None:
        self.push_screen(
            CommandInput("Format (F=Fixed, S=Scientific, C=Currency, P=Percent, G=General):"),
            self._do_range_format
        )

    def _do_range_format(self, result: str | None) -> None:
        if not result:
            return
        format_char = result.upper()[0] if result else "G"
        format_map = {"F": "F2", "S": "S", "C": "C2", "P": "P2", "G": "G", ",": ",2"}
        format_code = format_map.get(format_char, "G")
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                cell.format_code = format_code
        grid.refresh_grid()
        self._update_status()
        self.notify(f"Format set to {format_code}")

    def _range_label(self) -> None:
        self.push_screen(
            CommandInput("Label alignment (L=Left, R=Right, C=Center):"),
            self._do_range_label
        )

    def _do_range_label(self, result: str | None) -> None:
        if not result:
            return
        align_char = result.upper()[0] if result else "L"
        prefix_map = {"L": "'", "R": '"', "C": "^"}
        prefix = prefix_map.get(align_char, "'")
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        changes = []
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                old_value = cell.raw_value
                if old_value and not cell.is_formula:
                    display = cell.display_value
                    new_value = prefix + display
                    if new_value != old_value:
                        changes.append((r, c, new_value, old_value))
        if changes:
            cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
            self.undo_manager.execute(cmd)
            grid.refresh_grid()
            self._update_status()
        align_names = {"L": "Left", "R": "Right", "C": "Center"}
        self.notify(f"Label alignment set to {align_names.get(align_char, 'Left')}")

    def _range_name(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        range_str = f"{make_cell_ref(r1, c1)}:{make_cell_ref(r2, c2)}"
        self._pending_range = range_str
        self.push_screen(CommandInput(f"Name for range {range_str}:"), self._do_range_name)

    def _do_range_name(self, result: str | None) -> None:
        if not result:
            return
        name = result.strip().upper()
        if not name:
            return
        if not hasattr(self.spreadsheet, '_named_ranges'):
            self.spreadsheet._named_ranges = {}
        self.spreadsheet._named_ranges[name] = self._pending_range
        self.notify(f"Named range '{name}' created for {self._pending_range}")

    def _range_protect(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        protected_count = 0
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = self.spreadsheet.get_cell(r, c)
                if not hasattr(cell, '_protected'):
                    cell._protected = False
                cell._protected = not cell._protected
                if cell._protected:
                    protected_count += 1
        total_cells = (r2 - r1 + 1) * (c2 - c1 + 1)
        if protected_count > 0:
            self.notify(f"Protected {protected_count} cell(s)")
        else:
            self.notify(f"Unprotected {total_cells} cell(s)")

    # Data menu methods
    def _data_fill(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        if not grid.has_selection:
            self.notify("Select a range first")
            return
        self.push_screen(
            CommandInput("Fill with (start,step,stop) or value:"),
            self._do_data_fill
        )

    def _do_data_fill(self, result: str | None) -> None:
        if not result:
            return
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        changes = []
        try:
            if "," in result:
                parts = [p.strip() for p in result.split(",")]
                start = float(parts[0])
                step = float(parts[1]) if len(parts) > 1 else 1
                val = start
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        cell = self.spreadsheet.get_cell(r, c)
                        old_value = cell.raw_value
                        new_value = str(int(val) if val == int(val) else val)
                        changes.append((r, c, new_value, old_value))
                        val += step
            else:
                fill_value = result
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        cell = self.spreadsheet.get_cell(r, c)
                        old_value = cell.raw_value
                        changes.append((r, c, fill_value, old_value))
            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid.refresh_grid()
                self._update_status()
                self.notify(f"Filled {len(changes)} cell(s)")
        except ValueError as e:
            self.notify(f"Invalid fill value: {e}", severity="error")

    def _data_sort(self) -> None:
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        first_col = index_to_col(c1)
        last_col = index_to_col(c2)
        col_range = first_col if c1 == c2 else f"{first_col}-{last_col}"
        self.push_screen(
            CommandInput(f"Sort column [{col_range}] (add D for descending, e.g., 'A' or 'AD'):"),
            self._do_data_sort
        )

    def _do_data_sort(self, result: str | None) -> None:
        if not result:
            return
        grid = self.query_one("#grid", SpreadsheetGrid)
        r1, c1, r2, c2 = grid.selection_range
        try:
            result = result.strip().upper()
            reverse = result.endswith("D")
            sort_col_letter = result.rstrip("D").rstrip("A") or result[0]
            sort_col_abs = col_to_index(sort_col_letter)
            if sort_col_abs < c1 or sort_col_abs > c2:
                sort_col_idx = ord(sort_col_letter) - ord("A")
                sort_col_abs = c1 + sort_col_idx
            if sort_col_abs < c1 or sort_col_abs > c2:
                self.notify(f"Sort column must be within selection ({index_to_col(c1)}-{index_to_col(c2)})", severity="error")
                return
            rows_data = []
            for r in range(r1, r2 + 1):
                row_values = []
                for c in range(c1, c2 + 1):
                    cell = self.spreadsheet.get_cell(r, c)
                    row_values.append(cell.raw_value)
                sort_val = self.spreadsheet.get_value(r, sort_col_abs)
                if sort_val == "" or sort_val is None:
                    sort_key = (2, "")
                elif isinstance(sort_val, (int, float)):
                    sort_key = (0, sort_val)
                else:
                    sort_key = (1, str(sort_val).lower())
                rows_data.append((sort_key, row_values))
            rows_data.sort(key=lambda x: x[0], reverse=reverse)
            changes = []
            for row_idx, (_, row_values) in enumerate(rows_data):
                target_row = r1 + row_idx
                for col_idx, value in enumerate(row_values):
                    target_col = c1 + col_idx
                    cell = self.spreadsheet.get_cell(target_row, target_col)
                    old_value = cell.raw_value
                    if value != old_value:
                        changes.append((target_row, target_col, value, old_value))
            if changes:
                cmd = RangeChangeCommand(spreadsheet=self.spreadsheet, changes=changes)
                self.undo_manager.execute(cmd)
                grid.refresh_grid()
                self._update_status()
                order_name = "descending" if reverse else "ascending"
                self.notify(f"Sorted {len(rows_data)} rows by column {sort_col_letter} ({order_name})")
            else:
                self.notify("Data already sorted")
        except Exception as e:
            self.notify(f"Sort error: {e}", severity="error")

    def _data_query(self) -> None:
        self.notify("Data Query: Select criteria range, then input range. Use @D functions for queries.")

    # Worksheet global methods
    def _global_format(self) -> None:
        self.push_screen(
            CommandInput(f"Default format (F=Fixed, S=Scientific, C=Currency, P=Percent, G=General) [{self._global_format_code}]:"),
            self._do_global_format
        )

    def _do_global_format(self, result: str | None) -> None:
        if not result:
            return
        format_char = result.upper()[0] if result else "G"
        format_map = {"F": "F2", "S": "S", "C": "C2", "P": "P2", "G": "G", ",": ",2"}
        self._global_format_code = format_map.get(format_char, "G")
        self.notify(f"Default format set to {self._global_format_code}")

    def _global_label_prefix_action(self) -> None:
        self.push_screen(
            CommandInput("Default label alignment (L=Left, R=Right, C=Center):"),
            self._do_global_label_prefix
        )

    def _do_global_label_prefix(self, result: str | None) -> None:
        if not result:
            return
        align_char = result.upper()[0] if result else "L"
        prefix_map = {"L": "'", "R": '"', "C": "^"}
        self._global_label_prefix = prefix_map.get(align_char, "'")
        align_names = {"'": "Left", '"': "Right", "^": "Center"}
        self.notify(f"Default label alignment set to {align_names.get(self._global_label_prefix, 'Left')}")

    def _global_column_width(self) -> None:
        self.push_screen(
            CommandInput(f"Default column width (3-50) [{self._global_col_width}]:"),
            self._do_global_column_width
        )

    def _do_global_column_width(self, result: str | None) -> None:
        if not result:
            return
        try:
            width = int(result)
            width = max(3, min(50, width))
            self._global_col_width = width
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.default_col_width = width
            grid.refresh_grid()
            self.notify(f"Default column width set to {width}")
        except ValueError:
            self.notify("Invalid width", severity="error")

    def _global_recalculation(self) -> None:
        if self._recalc_mode == "auto":
            self._recalc_mode = "manual"
            self.notify("Recalculation: Manual (press F9 to recalculate)")
        else:
            self._recalc_mode = "auto"
            self.spreadsheet._invalidate_cache()
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.refresh_grid()
            self.notify("Recalculation: Automatic")

    def _global_protection_action(self) -> None:
        self._global_protection = not self._global_protection
        if self._global_protection:
            self.notify("Worksheet protection ENABLED - protected cells cannot be edited")
        else:
            self.notify("Worksheet protection DISABLED")

    def _global_zero(self) -> None:
        self._global_zero_display = not self._global_zero_display
        grid = self.query_one("#grid", SpreadsheetGrid)
        grid.show_zero = self._global_zero_display
        grid.refresh_grid()
        if self._global_zero_display:
            self.notify("Zero values: Displayed")
        else:
            self.notify("Zero values: Hidden (blank)")

    def _worksheet_erase(self) -> None:
        self.push_screen(
            CommandInput("Erase entire worksheet? (Y/N):"),
            self._do_worksheet_erase
        )

    def _do_worksheet_erase(self, result: str | None) -> None:
        if result and result.upper().startswith("Y"):
            self.spreadsheet.clear()
            self.undo_manager.clear()
            grid = self.query_one("#grid", SpreadsheetGrid)
            grid.refresh_grid()
            self._update_status()
            self.notify("Worksheet erased")

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
        # pageup, pagedown, ctrl+d, ctrl+u, home, end are handled by BINDINGS
        elif event.key == "enter":
            self.action_edit_cell()
            event.prevent_default()
            return
        elif event.key in ("delete", "backspace"):
            self.action_clear_cell()
            event.prevent_default()
            return

        # Start editing on printable character
        if event.character and event.character.isprintable() and event.character != '/':
            cell_input = self.query_one("#cell-input", Input)
            cell_input.select_on_focus = False
            cell_input.value = ""
            cell_input.focus()
            self.editing = True
            # Determine mode based on first character (Lotus 1-2-3 style)
            if event.character in '0123456789.+-@#(':
                self.query_one("#status-bar", StatusBarWidget).set_mode(Mode.VALUE)
            elif event.character == '=':
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

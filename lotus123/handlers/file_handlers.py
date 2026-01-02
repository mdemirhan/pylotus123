"""File operation handler methods for LotusApp."""

from __future__ import annotations

import json
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
        self._pending_save_path: str = ""
        self._pending_xlsx_import_path: str = ""

    def _sync_global_settings_to_spreadsheet(self) -> None:
        """Sync app global settings to spreadsheet before save."""
        self.spreadsheet.global_settings["format_code"] = self._app.global_format_code
        self.spreadsheet.global_settings["label_prefix"] = self._app.global_label_prefix
        self.spreadsheet.global_settings["default_col_width"] = self._app.global_col_width
        self.spreadsheet.global_settings["zero_display"] = self._app.global_zero_display

    def _sync_global_settings_from_spreadsheet(self) -> None:
        """Sync global settings from spreadsheet to app after load."""
        grid = self.get_grid()

        # Global settings
        gs = self.spreadsheet.global_settings
        self._app.global_format_code = gs.get("format_code", "G")
        self._app.global_label_prefix = gs.get("label_prefix", "'")
        self._app.global_col_width = gs.get("default_col_width", 10)
        self._app.global_zero_display = gs.get("zero_display", True)

        # Apply zero display to grid
        grid.show_zero = self._app.global_zero_display
        grid.default_col_width = self._app.global_col_width

    def new_file(self) -> None:
        """Create a new empty spreadsheet."""
        self.spreadsheet.clear()
        self.spreadsheet.filename = ""
        self.undo_manager.clear()
        self._app._dirty = False

        # Reset global settings to defaults
        self._app.global_format_code = "G"
        self._app.global_label_prefix = "'"
        self._app.global_col_width = 10
        self._app.global_zero_display = True

        grid = self.get_grid()
        grid.cursor_row = 0
        grid.cursor_col = 0
        grid.scroll_row = 0
        grid.scroll_col = 0
        grid.show_zero = True
        grid.default_col_width = 10
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
                # Restore global settings from loaded file
                self._sync_global_settings_from_spreadsheet()
                grid = self.get_grid()
                grid.refresh_grid()
                self._app._update_status()
                self._app._update_title()
                self._app.config.add_recent_file(str(path))
                self._app.config.save()
                self.notify(f"Loaded: {path}")
            else:
                self.notify(f"File not found: {filepath}", severity="error")
        except FileNotFoundError:
            self.notify(f"File not found: {filepath}", severity="error")
        except PermissionError:
            self.notify(f"Permission denied: {filepath}", severity="error")
        except json.JSONDecodeError:
            self.notify(f"Invalid file format: {filepath}", severity="error")
        except (OSError, IOError) as e:
            self.notify(f"Error reading file: {e}", severity="error")

    def _do_open(self, result: str | None) -> None:
        if not result:
            return

        # Detect file type by extension
        path = Path(result)
        ext = path.suffix.lower()

        # Handle non-JSON files as imports
        if ext in (".csv", ".tsv", ".wk1", ".xlsx"):
            self._import_non_json_file(result, ext)
            return

        # Handle JSON (native) format
        try:
            self.spreadsheet.load(result)
            self.undo_manager.clear()
            self._app._dirty = False
            # Restore global settings from loaded file
            self._sync_global_settings_from_spreadsheet()
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
        except FileNotFoundError:
            self.notify(f"File not found: {result}", severity="error")
        except PermissionError:
            self.notify(f"Permission denied: {result}", severity="error")
        except json.JSONDecodeError:
            self.notify(f"Invalid file format: {result}", severity="error")
        except (OSError, IOError) as e:
            self.notify(f"Error reading file: {e}", severity="error")

    def _import_non_json_file(self, filepath: str, ext: str) -> None:
        """Import a non-JSON file format (CSV, TSV, WK1, XLSX).

        This is called when the user opens a non-JSON file.
        The file is imported, not opened - save will require Save As.
        """
        path = Path(filepath)
        filename = path.name

        try:
            if ext == ".csv":
                from ..io import ImportOptions, TextImporter
                from ..io.text_import import ImportFormat

                self.spreadsheet.clear()
                importer = TextImporter(self.spreadsheet)
                options = ImportOptions(format=ImportFormat.CSV)
                importer.import_file(filepath, options)
                self._finalize_import(f"Imported {filename}. Use 'Save As' to save as Lotus JSON.")

            elif ext == ".tsv":
                from ..io import ImportOptions, TextImporter
                from ..io.text_import import ImportFormat

                self.spreadsheet.clear()
                importer = TextImporter(self.spreadsheet)
                options = ImportOptions(format=ImportFormat.TSV)
                importer.import_file(filepath, options)
                self._finalize_import(f"Imported {filename}. Use 'Save As' to save as Lotus JSON.")

            elif ext == ".wk1":
                from ..io.wk1 import Wk1Reader

                reader = Wk1Reader(self.spreadsheet)
                reader.load(filepath)
                self._finalize_import(f"Imported {filename}. Use 'Save As' to save as Lotus JSON.")

            elif ext == ".xlsx":
                # For XLSX, we may need to show sheet selection dialog
                from ..io.xlsx import get_xlsx_sheet_names

                self._pending_xlsx_import_path = filepath
                sheet_names = get_xlsx_sheet_names(filepath)

                if len(sheet_names) > 1:
                    # Multiple sheets - show selection dialog
                    from ..ui.dialogs import SheetSelectDialog

                    self._app.push_screen(
                        SheetSelectDialog(sheet_names),
                        self._do_import_xlsx_from_open,
                    )
                else:
                    # Single sheet - import directly
                    self._perform_xlsx_import_from_open(filepath, sheet_names[0] if sheet_names else None)

        except FileNotFoundError:
            self.notify(f"File not found: {filepath}", severity="error")
        except PermissionError:
            self.notify(f"Permission denied: {filepath}", severity="error")
        except ImportError as e:
            self.notify(str(e), severity="error")
        except Exception as e:
            self.notify(f"Error importing {filename}: {e}", severity="error")

    def _do_import_xlsx_from_open(self, sheet_name: str | None) -> None:
        """Handle sheet selection for XLSX import from open dialog."""
        if not sheet_name:
            self.notify("Import cancelled", severity="warning")
            return
        self._perform_xlsx_import_from_open(self._pending_xlsx_import_path, sheet_name)

    def _perform_xlsx_import_from_open(self, filepath: str, sheet_name: str | None) -> None:
        """Perform XLSX import from open dialog."""
        try:
            from ..io.xlsx import XlsxReader

            reader = XlsxReader(self.spreadsheet)
            warnings = reader.load(filepath, sheet_name)

            filename = Path(filepath).name
            msg = f"Imported {filename}. Use 'Save As' to save as Lotus JSON."
            if warnings.has_warnings():
                msg += f" {warnings.to_message()}"
                self._finalize_import(msg, severity="warning")
            else:
                self._finalize_import(msg)

        except Exception as e:
            self.notify(f"Error importing XLSX: {e}", severity="error")

    def _finalize_import(self, message: str, severity: str = "information") -> None:
        """Finalize an import operation and refresh the UI."""
        self.spreadsheet.filename = ""  # Not a native file
        self._app._dirty = True
        self.undo_manager.clear()

        grid = self.get_grid()
        grid.cursor_row = 0
        grid.cursor_col = 0
        grid.scroll_row = 0
        grid.scroll_col = 0
        grid.recalculate_visible_area()
        grid.refresh_grid()
        self._app._update_status()
        self._app._update_title()

        self.notify(message, severity=severity)

    def save(self) -> None:
        """Save the current spreadsheet."""
        if self.spreadsheet.filename:
            self._sync_global_settings_to_spreadsheet()
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
            if not result.endswith(".json"):
                result += ".json"
            # Check if file exists
            if Path(result).exists():
                self._pending_save_path = result
                self._app.push_screen(
                    CommandInput(f"File '{result}' exists. Overwrite? (Y/N):"),
                    self._do_save_confirm,
                )
            else:
                self._perform_save(result)

    def _do_save_confirm(self, result: str | None) -> None:
        if result and result.strip().upper().startswith("Y"):
            self._perform_save(self._pending_save_path)
        else:
            self.notify("Save cancelled", severity="warning")

    def _perform_save(self, filepath: str) -> None:
        try:
            self._sync_global_settings_to_spreadsheet()
            self.spreadsheet.save(filepath)
            self._app._dirty = False
            self._app._update_title()
            self.notify(f"Saved: {filepath}")
        except PermissionError:
            self.notify(f"Permission denied: {filepath}", severity="error")
        except (OSError, IOError) as e:
            self.notify(f"Error saving file: {e}", severity="error")

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
                self._sync_global_settings_to_spreadsheet()
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
            if not result.endswith(".json"):
                result += ".json"
            if Path(result).exists():
                self._pending_save_path = result
                self._app.push_screen(
                    CommandInput(f"File '{result}' exists. Overwrite? (Y/N):"),
                    self._do_save_and_quit_confirm,
                )
            else:
                self._sync_global_settings_to_spreadsheet()
                self.spreadsheet.save(result)
                self._app.config.save()
                self._app.exit()
        else:
            self._app.config.save()
            self._app.exit()

    def _do_save_and_quit_confirm(self, result: str | None) -> None:
        if result and result.strip().upper().startswith("Y"):
            self._sync_global_settings_to_spreadsheet()
            self.spreadsheet.save(self._pending_save_path)
            self._app.config.save()
            self._app.exit()

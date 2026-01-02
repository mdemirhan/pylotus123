"""File operation handler methods for LotusApp."""

from __future__ import annotations

import json
from pathlib import Path
from typing import TYPE_CHECKING, Callable, Literal

from ..ui import THEMES, CommandInput, FileDialog, ThemeDialog, ThemeType
from .base import BaseHandler

if TYPE_CHECKING:
    from .base import AppProtocol


class FileHandler(BaseHandler):
    """Handler for file operations (new, open, save, quit, theme)."""

    # Extensions for importable formats (non-native)
    IMPORT_EXTENSIONS = (".csv", ".tsv", ".wk1", ".wks", ".xlsx", ".xls")
    # All supported extensions for open dialog
    OPEN_EXTENSIONS = [".json"] + list(IMPORT_EXTENSIONS)

    def __init__(self, app: "AppProtocol") -> None:
        super().__init__(app)
        self._pending_xlsx_import_path: str = ""
        self._pending_open_path: str = ""  # File path selected in open dialog
        self._pending_action: Callable[[], None] | None = None  # Callback for after save confirmation

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

    def _confirm_save_if_dirty(self, prompt: str, on_proceed: Callable[[], None]) -> None:
        """Prompt to save if dirty, then call on_proceed.

        Args:
            prompt: The confirmation prompt message
            on_proceed: Callback to execute after saving or declining to save
        """
        if self.is_dirty:
            self._pending_action = on_proceed
            self._app.push_screen(
                CommandInput(prompt),
                self._do_save_confirm,
            )
        else:
            on_proceed()

    def _do_save_confirm(self, result: str | None) -> None:
        """Handle Y/N/C response for save confirmation."""
        if not result or not self._pending_action:
            self._pending_action = None
            return
        response = result.strip().upper()
        if response.startswith("Y"):
            if self.spreadsheet.filename:
                self._sync_global_settings_to_spreadsheet()
                self.spreadsheet.save(self.spreadsheet.filename)
                self._pending_action()
                self._pending_action = None
            else:
                self._app.push_screen(
                    FileDialog(mode="save", file_extensions=[".json"]),
                    self._do_save_then_action,
                )
        elif response.startswith("N"):
            self._pending_action()
            self._pending_action = None
        else:
            # Cancel (C) or anything else: do nothing
            self._pending_action = None

    def _do_save_then_action(self, result: str | None) -> None:
        """Save to chosen file then execute pending action."""
        if result and self._pending_action:
            self._sync_global_settings_to_spreadsheet()
            self.spreadsheet.save(result)
            self._app.config.add_recent_file(result)
            self._app.config.save()
            self._pending_action()
        self._pending_action = None

    def new_file(self) -> None:
        """Create a new empty spreadsheet, prompting to save if dirty."""
        self._confirm_save_if_dirty(
            "Save changes before creating new file? (Y/N/C=Cancel):",
            self._do_new_file,
        )

    def _do_new_file(self) -> None:
        """Actually create a new empty spreadsheet."""
        self.spreadsheet.clear()
        self.spreadsheet.filename = ""
        self.undo_manager.clear()
        self.is_dirty = False

        # Reset global settings to defaults
        self._app.global_format_code = "G"
        self._app.global_label_prefix = "'"
        self._app.global_col_width = 10
        self._app.global_zero_display = True

        grid = self.reset_view()
        grid.show_zero = True
        grid.default_col_width = 10
        self.notify("New spreadsheet created")

    def open_file(self) -> None:
        """Show the file open dialog."""
        self._app.push_screen(
            FileDialog(mode="open", file_extensions=self.OPEN_EXTENSIONS),
            self._do_open,
        )

    def load_initial_file(self, filepath: str) -> None:
        """Load a file specified at startup.

        Supports JSON (native), XLSX, CSV, TSV, and WK1 formats.
        """
        if not Path(filepath).exists():
            self.notify(f"File not found: {filepath}", severity="error")
            return
        self._load_file(filepath)

    def _do_open(self, result: str | None) -> None:
        """Handle file open dialog result."""
        if not result:
            return
        self._pending_open_path = result
        self._confirm_save_if_dirty(
            "Save changes before opening another file? (Y/N/C=Cancel):",
            self._do_open_after_save,
        )

    def _do_open_after_save(self) -> None:
        """Load the pending file after save confirmation."""
        if self._pending_open_path:
            self._load_file(self._pending_open_path)
            self._pending_open_path = ""

    def _load_file(self, filepath: str) -> None:
        """Load a file by path, detecting format by extension."""
        path = Path(filepath)
        ext = path.suffix.lower()

        # Handle non-JSON files as imports
        if ext in self.IMPORT_EXTENSIONS:
            self._import_non_json_file(filepath, ext)
            return

        # Handle JSON (native) format
        try:
            self.spreadsheet.load(filepath)
            self.undo_manager.clear()
            self.is_dirty = False
            self._sync_global_settings_from_spreadsheet()
            self.reset_view()
            self._app.config.add_recent_file(filepath)
            self._app.config.save()
            self.notify(f"Loaded: {filepath}")
        except FileNotFoundError:
            self.notify(f"File not found: {filepath}", severity="error")
        except PermissionError:
            self.notify(f"Permission denied: {filepath}", severity="error")
        except (json.JSONDecodeError, UnicodeDecodeError):
            self.notify(f"Invalid file format: {filepath}", severity="error")
        except (OSError, IOError) as e:
            self.notify(f"Error reading file: {e}", severity="error")

    def _import_non_json_file(self, filepath: str, ext: str) -> None:
        """Import a non-JSON file format (CSV, TSV, WK1/WKS, XLSX/XLS).

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

            elif ext in (".wk1", ".wks"):
                from ..io.wk1 import Wk1Reader

                reader = Wk1Reader(self.spreadsheet)
                reader.load(filepath)
                self._finalize_import(f"Imported {filename}. Use 'Save As' to save as Lotus JSON.")

            elif ext in (".xlsx", ".xls"):
                # For XLSX/XLS, we may need to show sheet selection dialog
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

    def _finalize_import(
        self, message: str, severity: Literal["information", "warning", "error"] = "information"
    ) -> None:
        """Finalize an import operation and refresh the UI."""
        self.spreadsheet.filename = ""  # Not a native file
        self.is_dirty = True
        self.undo_manager.clear()
        self.reset_view()

        self.notify(message, severity=severity)

    def save(self) -> None:
        """Save the current spreadsheet."""
        if self.spreadsheet.filename:
            self._sync_global_settings_to_spreadsheet()
            self.spreadsheet.save(self.spreadsheet.filename)
            self.is_dirty = False
            self.update_title()
            self.notify(f"Saved: {self.spreadsheet.filename}")
        else:
            self.save_as()

    def save_as(self) -> None:
        """Show the save-as dialog."""
        self._app.push_screen(
            FileDialog(mode="save", title="Save As", file_extensions=[".json"]),
            self._do_save,
        )

    def _do_save(self, result: str | None) -> None:
        if result:
            if not result.endswith(".json"):
                result += ".json"
            self.confirm_overwrite(result, self._perform_save)

    def _perform_save(self, filepath: str) -> None:
        try:
            self._sync_global_settings_to_spreadsheet()
            self.spreadsheet.save(filepath)
            self.is_dirty = False
            self.update_title()
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
            self.apply_theme()
            self._app.config.theme = result.name
            self._app.config.save()
            self.notify(f"Theme changed to {self._app.color_theme.name}")

    def quit_app(self) -> None:
        """Quit the application, prompting to save if dirty."""
        self._confirm_save_if_dirty(
            "Save changes before quitting? (Y/N/C=Cancel):",
            self._do_quit,
        )

    def _do_quit(self) -> None:
        """Actually quit the application."""
        self._app.config.save()
        self._app.exit()

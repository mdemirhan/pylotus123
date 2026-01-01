"""Import/Export handler for CSV, TSV, and WK1 file formats."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from ..ui import CommandInput, FileDialog
from .base import BaseHandler

if TYPE_CHECKING:
    from .base import AppProtocol


class ImportExportHandler(BaseHandler):
    """Handler for import/export operations (CSV, TSV, WK1)."""

    def __init__(self, app: "AppProtocol") -> None:
        super().__init__(app)
        self._pending_export_format: str = ""
        self._pending_export_path: str = ""

    # ===== CSV Operations =====

    def import_csv(self) -> None:
        """Show file dialog to import a CSV file."""
        self._app.push_screen(
            FileDialog(mode="open", title="Import CSV"),
            self._do_import_csv,
        )

    def _do_import_csv(self, result: str | None) -> None:
        """Handle CSV import dialog result."""
        if not result:
            return

        try:
            from ..io import ImportOptions, TextImporter
            from ..io.text_import import ImportFormat

            # Clear and import
            self.spreadsheet.clear()
            importer = TextImporter(self.spreadsheet)
            options = ImportOptions(format=ImportFormat.CSV)
            row_count = importer.import_file(result, options)

            # Update state
            self.spreadsheet.filename = ""  # Not a native file
            self._app._dirty = True
            self.undo_manager.clear()

            # Refresh UI
            grid = self.get_grid()
            grid.cursor_row = 0
            grid.cursor_col = 0
            grid.scroll_row = 0
            grid.scroll_col = 0
            grid.refresh_grid()
            self._app._update_status()
            self._app._update_title()

            self.notify(f"Imported {row_count} rows from CSV")

        except FileNotFoundError:
            self.notify(f"File not found: {result}", severity="error")
        except PermissionError:
            self.notify(f"Permission denied: {result}", severity="error")
        except Exception as e:
            self.notify(f"Error importing CSV: {e}", severity="error")

    def export_csv(self) -> None:
        """Show file dialog to export to CSV."""
        self._pending_export_format = "csv"
        self._app.push_screen(
            FileDialog(mode="save", title="Export CSV"),
            self._do_export_csv,
        )

    def _do_export_csv(self, result: str | None) -> None:
        """Handle CSV export dialog result."""
        if not result:
            return

        # Ensure .csv extension
        if not result.lower().endswith(".csv"):
            result += ".csv"

        # Check if file exists
        if Path(result).exists():
            self._pending_export_path = result
            self._app.push_screen(
                CommandInput(f"File '{result}' exists. Overwrite? (Y/N):"),
                self._do_export_csv_confirm,
            )
        else:
            self._perform_export_csv(result)

    def _do_export_csv_confirm(self, result: str | None) -> None:
        """Handle CSV overwrite confirmation."""
        if result and result.strip().upper().startswith("Y"):
            self._perform_export_csv(self._pending_export_path)
        else:
            self.notify("Export cancelled", severity="warning")

    def _perform_export_csv(self, filepath: str) -> None:
        """Actually perform the CSV export."""
        try:
            from ..io import ExportOptions, TextExporter
            from ..io.text_export import ExportFormat

            exporter = TextExporter(self.spreadsheet)
            options = ExportOptions(
                format=ExportFormat.CSV,
                use_formulas=False,  # Export calculated values
            )
            row_count = exporter.export_file(filepath, options)

            self.notify(f"Exported {row_count} rows to {Path(filepath).name}")

        except PermissionError:
            self.notify(f"Permission denied: {filepath}", severity="error")
        except Exception as e:
            self.notify(f"Error exporting CSV: {e}", severity="error")

    # ===== TSV Operations =====

    def import_tsv(self) -> None:
        """Show file dialog to import a TSV file."""
        self._app.push_screen(
            FileDialog(mode="open", title="Import TSV"),
            self._do_import_tsv,
        )

    def _do_import_tsv(self, result: str | None) -> None:
        """Handle TSV import dialog result."""
        if not result:
            return

        try:
            from ..io import ImportOptions, TextImporter
            from ..io.text_import import ImportFormat

            # Clear and import
            self.spreadsheet.clear()
            importer = TextImporter(self.spreadsheet)
            options = ImportOptions(format=ImportFormat.TSV)
            row_count = importer.import_file(result, options)

            # Update state
            self.spreadsheet.filename = ""
            self._app._dirty = True
            self.undo_manager.clear()

            # Refresh UI
            grid = self.get_grid()
            grid.cursor_row = 0
            grid.cursor_col = 0
            grid.scroll_row = 0
            grid.scroll_col = 0
            grid.refresh_grid()
            self._app._update_status()
            self._app._update_title()

            self.notify(f"Imported {row_count} rows from TSV")

        except FileNotFoundError:
            self.notify(f"File not found: {result}", severity="error")
        except PermissionError:
            self.notify(f"Permission denied: {result}", severity="error")
        except Exception as e:
            self.notify(f"Error importing TSV: {e}", severity="error")

    def export_tsv(self) -> None:
        """Show file dialog to export to TSV."""
        self._pending_export_format = "tsv"
        self._app.push_screen(
            FileDialog(mode="save", title="Export TSV"),
            self._do_export_tsv,
        )

    def _do_export_tsv(self, result: str | None) -> None:
        """Handle TSV export dialog result."""
        if not result:
            return

        # Ensure .tsv extension
        if not result.lower().endswith(".tsv"):
            result += ".tsv"

        # Check if file exists
        if Path(result).exists():
            self._pending_export_path = result
            self._app.push_screen(
                CommandInput(f"File '{result}' exists. Overwrite? (Y/N):"),
                self._do_export_tsv_confirm,
            )
        else:
            self._perform_export_tsv(result)

    def _do_export_tsv_confirm(self, result: str | None) -> None:
        """Handle TSV overwrite confirmation."""
        if result and result.strip().upper().startswith("Y"):
            self._perform_export_tsv(self._pending_export_path)
        else:
            self.notify("Export cancelled", severity="warning")

    def _perform_export_tsv(self, filepath: str) -> None:
        """Actually perform the TSV export."""
        try:
            from ..io import ExportOptions, TextExporter
            from ..io.text_export import ExportFormat

            exporter = TextExporter(self.spreadsheet)
            options = ExportOptions(
                format=ExportFormat.TSV,
                use_formulas=False,
            )
            row_count = exporter.export_file(filepath, options)

            self.notify(f"Exported {row_count} rows to {Path(filepath).name}")

        except PermissionError:
            self.notify(f"Permission denied: {filepath}", severity="error")
        except Exception as e:
            self.notify(f"Error exporting TSV: {e}", severity="error")

    # ===== WK1 Operations =====

    def import_wk1(self) -> None:
        """Show file dialog to import a Lotus WK1 file."""
        self._app.push_screen(
            FileDialog(mode="open", title="Import Lotus 1-2-3"),
            self._do_import_wk1,
        )

    def _do_import_wk1(self, result: str | None) -> None:
        """Handle WK1 import dialog result."""
        if not result:
            return

        try:
            from ..io.wk1 import Wk1Reader

            reader = Wk1Reader(self.spreadsheet)
            reader.load(result)

            # Update state - keep filename empty since it's not native format
            self.spreadsheet.filename = ""
            self._app._dirty = True
            self.undo_manager.clear()

            # Refresh UI
            grid = self.get_grid()
            grid.cursor_row = 0
            grid.cursor_col = 0
            grid.scroll_row = 0
            grid.scroll_col = 0
            grid.refresh_grid()
            self._app._update_status()
            self._app._update_title()

            self.notify(f"Imported from {Path(result).name}")

        except FileNotFoundError:
            self.notify(f"File not found: {result}", severity="error")
        except PermissionError:
            self.notify(f"Permission denied: {result}", severity="error")
        except ValueError as e:
            self.notify(f"Invalid WK1 file: {e}", severity="error")
        except Exception as e:
            self.notify(f"Error importing WK1: {e}", severity="error")

    def export_wk1(self) -> None:
        """Show file dialog to export to Lotus WK1 format."""
        self._pending_export_format = "wk1"
        self._app.push_screen(
            FileDialog(mode="save", title="Export Lotus 1-2-3"),
            self._do_export_wk1,
        )

    def _do_export_wk1(self, result: str | None) -> None:
        """Handle WK1 export dialog result."""
        if not result:
            return

        # Ensure .wk1 extension
        if not result.lower().endswith(".wk1"):
            result += ".wk1"

        # Check if file exists
        if Path(result).exists():
            self._pending_export_path = result
            self._app.push_screen(
                CommandInput(f"File '{result}' exists. Overwrite? (Y/N):"),
                self._do_export_wk1_confirm,
            )
        else:
            self._perform_export_wk1(result)

    def _do_export_wk1_confirm(self, result: str | None) -> None:
        """Handle WK1 overwrite confirmation."""
        if result and result.strip().upper().startswith("Y"):
            self._perform_export_wk1(self._pending_export_path)
        else:
            self.notify("Export cancelled", severity="warning")

    def _perform_export_wk1(self, filepath: str) -> None:
        """Actually perform the WK1 export."""
        try:
            from ..io.wk1 import Wk1Writer

            writer = Wk1Writer(self.spreadsheet)
            writer.save(filepath)

            self.notify(f"Exported to {Path(filepath).name}")

        except PermissionError:
            self.notify(f"Permission denied: {filepath}", severity="error")
        except Exception as e:
            self.notify(f"Error exporting WK1: {e}", severity="error")

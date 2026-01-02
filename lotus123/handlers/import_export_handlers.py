"""Import/Export handler for CSV, TSV, WK1, and XLSX file formats."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from ..ui import CommandInput, FileDialog
from ..ui.dialogs import SheetSelectDialog
from .base import BaseHandler

if TYPE_CHECKING:
    from .base import AppProtocol


class ImportExportHandler(BaseHandler):
    """Handler for import/export operations (CSV, TSV, WK1, XLSX)."""

    def __init__(self, app: "AppProtocol") -> None:
        super().__init__(app)
        self._pending_export_format: str = ""
        self._pending_export_path: str = ""
        self._pending_xlsx_path: str = ""

    # ===== CSV Operations =====

    def import_csv(self) -> None:
        """Show file dialog to import a CSV file."""
        self._app.push_screen(
            FileDialog(mode="open", title="Import CSV", file_extensions=[".csv"]),
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
            FileDialog(mode="open", title="Import TSV", file_extensions=[".tsv"]),
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
            FileDialog(mode="open", title="Import Lotus 1-2-3", file_extensions=[".wk1", ".wks"]),
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

    # ===== XLSX Operations =====

    def import_xlsx(self) -> None:
        """Show file dialog to import an Excel XLSX file."""
        self._app.push_screen(
            FileDialog(mode="open", title="Import Excel XLSX", file_extensions=[".xlsx", ".xls"]),
            self._do_import_xlsx_file,
        )

    def _do_import_xlsx_file(self, result: str | None) -> None:
        """Handle XLSX file selection dialog result."""
        if not result:
            return

        self._pending_xlsx_path = result

        try:
            from ..io.xlsx import get_xlsx_sheet_names

            sheet_names = get_xlsx_sheet_names(result)

            if len(sheet_names) > 1:
                # Multiple sheets - show selection dialog
                self._app.push_screen(
                    SheetSelectDialog(sheet_names),
                    self._do_import_xlsx_sheet,
                )
            else:
                # Single sheet - import directly
                self._perform_import_xlsx(result, sheet_names[0] if sheet_names else None)

        except ImportError as e:
            self.notify(str(e), severity="error")
        except FileNotFoundError:
            self.notify(f"File not found: {result}", severity="error")
        except PermissionError:
            self.notify(f"Permission denied: {result}", severity="error")
        except Exception as e:
            self.notify(f"Error reading XLSX: {e}", severity="error")

    def _do_import_xlsx_sheet(self, sheet_name: str | None) -> None:
        """Handle sheet selection dialog result."""
        if not sheet_name:
            self.notify("Import cancelled", severity="warning")
            return

        self._perform_import_xlsx(self._pending_xlsx_path, sheet_name)

    def _perform_import_xlsx(self, filepath: str, sheet_name: str | None) -> None:
        """Actually perform the XLSX import."""
        try:
            from ..io.xlsx import XlsxReader

            reader = XlsxReader(self.spreadsheet)
            warnings = reader.load(filepath, sheet_name)

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

            # Show result with warnings
            if warnings.has_warnings():
                self.notify(
                    f"Imported from XLSX. {warnings.to_message()}",
                    severity="warning",
                )
            else:
                self.notify(f"Imported from {Path(filepath).name}")

        except ImportError as e:
            self.notify(str(e), severity="error")
        except FileNotFoundError:
            self.notify(f"File not found: {filepath}", severity="error")
        except PermissionError:
            self.notify(f"Permission denied: {filepath}", severity="error")
        except Exception as e:
            self.notify(f"Error importing XLSX: {e}", severity="error")

    def export_xlsx(self) -> None:
        """Show file dialog to export to Excel XLSX format."""
        self._pending_export_format = "xlsx"
        self._app.push_screen(
            FileDialog(mode="save", title="Export Excel XLSX"),
            self._do_export_xlsx,
        )

    def _do_export_xlsx(self, result: str | None) -> None:
        """Handle XLSX export dialog result."""
        if not result:
            return

        # Ensure .xlsx extension
        if not result.lower().endswith(".xlsx"):
            result += ".xlsx"

        # Check if file exists
        if Path(result).exists():
            self._pending_export_path = result
            self._app.push_screen(
                CommandInput(f"File '{result}' exists. Overwrite? (Y/N):"),
                self._do_export_xlsx_confirm,
            )
        else:
            self._perform_export_xlsx(result)

    def _do_export_xlsx_confirm(self, result: str | None) -> None:
        """Handle XLSX overwrite confirmation."""
        if result and result.strip().upper().startswith("Y"):
            self._perform_export_xlsx(self._pending_export_path)
        else:
            self.notify("Export cancelled", severity="warning")

    def _perform_export_xlsx(self, filepath: str) -> None:
        """Actually perform the XLSX export."""
        try:
            from ..io.xlsx import XlsxWriter

            writer = XlsxWriter(self.spreadsheet)
            writer.save(filepath)

            self.notify(f"Exported to {Path(filepath).name}")

        except ImportError as e:
            self.notify(str(e), severity="error")
        except PermissionError:
            self.notify(f"Permission denied: {filepath}", severity="error")
        except Exception as e:
            self.notify(f"Error exporting XLSX: {e}", severity="error")

"""Import/Export handler for CSV, TSV, WK1, and XLSX file formats."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Callable

from ..io.text_import import ImportFormat
from ..io.text_export import ExportFormat
from ..ui import FileDialog
from ..ui.dialogs import SheetSelectDialog
from .base import BaseHandler

if TYPE_CHECKING:
    from .base import AppProtocol


class ImportExportHandler(BaseHandler):
    """Handler for import/export operations (CSV, TSV, WK1, XLSX)."""

    def __init__(self, app: "AppProtocol") -> None:
        super().__init__(app)
        self._pending_export_format: str = ""
        self._pending_xlsx_path: str = ""

    # ===== Common Import/Export Helpers =====

    def _import_text_file(
        self, filepath: str, format_type: ImportFormat, format_name: str
    ) -> None:
        """Common text file import logic for CSV/TSV."""
        try:
            from ..io import ImportOptions, TextImporter

            self.spreadsheet.clear()
            importer = TextImporter(self.spreadsheet)
            options = ImportOptions(format=format_type)
            row_count = importer.import_file(filepath, options)

            self.spreadsheet.filename = ""
            self.is_dirty = True
            self.undo_manager.clear()
            self.reset_view()

            self.notify(f"Imported {row_count} rows from {format_name}")

        except FileNotFoundError:
            self.notify(f"File not found: {filepath}", severity="error")
        except PermissionError:
            self.notify(f"Permission denied: {filepath}", severity="error")
        except Exception as e:
            self.notify(f"Error importing {format_name}: {e}", severity="error")

    def _export_text_file(
        self, filepath: str, format_type: ExportFormat, format_name: str
    ) -> None:
        """Common text file export logic for CSV/TSV."""
        try:
            from ..io import ExportOptions, TextExporter

            exporter = TextExporter(self.spreadsheet)
            options = ExportOptions(format=format_type, use_formulas=False)
            row_count = exporter.export_file(filepath, options)

            self.notify(f"Exported {row_count} rows to {Path(filepath).name}")

        except PermissionError:
            self.notify(f"Permission denied: {filepath}", severity="error")
        except Exception as e:
            self.notify(f"Error exporting {format_name}: {e}", severity="error")

    def _handle_export_dialog(
        self,
        result: str | None,
        extension: str,
        perform_export: Callable[[str], None],
    ) -> None:
        """Common export dialog result handling."""
        if not result:
            return

        if not result.lower().endswith(extension):
            result += extension

        self.confirm_overwrite(result, perform_export, cancel_message="Export cancelled")

    def _refresh_after_binary_import(self, filename: str) -> None:
        """Common UI refresh after WK1/XLSX import."""
        self.spreadsheet.filename = ""
        self.is_dirty = True
        self.undo_manager.clear()

        grid = self.get_grid()
        grid.cursor_row = 0
        grid.cursor_col = 0
        grid.scroll_row = 0
        grid.scroll_col = 0
        grid.refresh_grid()
        self.update_status()
        self.update_title()

        self.notify(f"Imported from {Path(filename).name}")

    # ===== CSV Operations =====

    def import_csv(self) -> None:
        """Show file dialog to import a CSV file."""
        self._app.push_screen(
            FileDialog(mode="open", title="Import CSV", file_extensions=[".csv"]),
            self._do_import_csv,
        )

    def _do_import_csv(self, result: str | None) -> None:
        """Handle CSV import dialog result."""
        if result:
            self._import_text_file(result, ImportFormat.CSV, "CSV")

    def export_csv(self) -> None:
        """Show file dialog to export to CSV."""
        self._app.push_screen(
            FileDialog(mode="save", title="Export CSV", file_extensions=[".csv"]),
            self._do_export_csv,
        )

    def _do_export_csv(self, result: str | None) -> None:
        """Handle CSV export dialog result."""
        self._handle_export_dialog(result, ".csv", self._perform_export_csv)

    def _perform_export_csv(self, filepath: str) -> None:
        """Actually perform the CSV export."""
        self._export_text_file(filepath, ExportFormat.CSV, "CSV")

    # ===== TSV Operations =====

    def import_tsv(self) -> None:
        """Show file dialog to import a TSV file."""
        self._app.push_screen(
            FileDialog(mode="open", title="Import TSV", file_extensions=[".tsv"]),
            self._do_import_tsv,
        )

    def _do_import_tsv(self, result: str | None) -> None:
        """Handle TSV import dialog result."""
        if result:
            self._import_text_file(result, ImportFormat.TSV, "TSV")

    def export_tsv(self) -> None:
        """Show file dialog to export to TSV."""
        self._app.push_screen(
            FileDialog(mode="save", title="Export TSV", file_extensions=[".tsv"]),
            self._do_export_tsv,
        )

    def _do_export_tsv(self, result: str | None) -> None:
        """Handle TSV export dialog result."""
        self._handle_export_dialog(result, ".tsv", self._perform_export_tsv)

    def _perform_export_tsv(self, filepath: str) -> None:
        """Actually perform the TSV export."""
        self._export_text_file(filepath, ExportFormat.TSV, "TSV")

    # ===== WK1 Operations =====

    def import_wk1(self) -> None:
        """Show file dialog to import a Lotus WK1 file."""
        self._app.push_screen(
            FileDialog(
                mode="open", title="Import Lotus 1-2-3", file_extensions=[".wk1", ".wks"]
            ),
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
            self._refresh_after_binary_import(result)

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
        self._app.push_screen(
            FileDialog(mode="save", title="Export Lotus 1-2-3", file_extensions=[".wk1"]),
            self._do_export_wk1,
        )

    def _do_export_wk1(self, result: str | None) -> None:
        """Handle WK1 export dialog result."""
        self._handle_export_dialog(result, ".wk1", self._perform_export_wk1)

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
            FileDialog(
                mode="open", title="Import Excel XLSX", file_extensions=[".xlsx", ".xls"]
            ),
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
                self._app.push_screen(
                    SheetSelectDialog(sheet_names),
                    self._do_import_xlsx_sheet,
                )
            else:
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

            self.spreadsheet.filename = ""
            self.is_dirty = True
            self.undo_manager.clear()

            grid = self.get_grid()
            grid.cursor_row = 0
            grid.cursor_col = 0
            grid.scroll_row = 0
            grid.scroll_col = 0
            grid.refresh_grid()
            self.update_status()
            self.update_title()

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
        self._app.push_screen(
            FileDialog(mode="save", title="Export Excel XLSX", file_extensions=[".xlsx"]),
            self._do_export_xlsx,
        )

    def _do_export_xlsx(self, result: str | None) -> None:
        """Handle XLSX export dialog result."""
        self._handle_export_dialog(result, ".xlsx", self._perform_export_xlsx)

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

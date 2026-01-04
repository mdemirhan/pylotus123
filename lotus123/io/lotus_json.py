"""JSON serialization for Spreadsheet."""

import json
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet

# Constants for default values to save space
DEFAULT_COL_WIDTH = 10
DEFAULT_ROW_HEIGHT = 1
MAX_ROWS = 65536
MAX_COLS = 256


class LotusJsonSerializer:
    """Handles saving and loading spreadsheets to/from JSON."""

    @staticmethod
    def save(spreadsheet: Spreadsheet, filename: str) -> None:
        """Save spreadsheet to JSON file."""
        data = {
            "version": 2,
            "rows": spreadsheet.rows,
            "cols": spreadsheet.cols,
            "col_widths": spreadsheet.get_all_col_widths(),
            "row_heights": spreadsheet.get_all_row_heights(),
            "cells": {
                f"{r},{c}": cell.to_dict()
                for r, c, cell in spreadsheet.iter_cells()
                if not cell.is_empty
            },
            "named_ranges": spreadsheet.named_ranges.to_dict(),
            "frozen_rows": spreadsheet.frozen_rows,
            "frozen_cols": spreadsheet.frozen_cols,
            "global_settings": spreadsheet.global_settings,
        }
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

    @staticmethod
    def load(spreadsheet: Spreadsheet, filename: str) -> None:
        """Load spreadsheet from JSON file."""
        with open(filename, "r", encoding="utf-8") as f:
            data = json.load(f)

        # Validate dimensions to prevent loading excessively large files
        loaded_rows = data.get("rows", MAX_ROWS)
        loaded_cols = data.get("cols", MAX_COLS)
        if loaded_rows > MAX_ROWS or loaded_rows < 1:
            raise ValueError(f"Invalid row count: {loaded_rows} (must be 1-{MAX_ROWS})")
        if loaded_cols > MAX_COLS or loaded_cols < 1:
            raise ValueError(f"Invalid column count: {loaded_cols} (must be 1-{MAX_COLS})")

        spreadsheet.clear()

        spreadsheet.rows = loaded_rows
        spreadsheet.cols = loaded_cols
        spreadsheet.set_all_col_widths({int(k): v for k, v in data.get("col_widths", {}).items()})
        spreadsheet.set_all_row_heights({int(k): v for k, v in data.get("row_heights", {}).items()})

        for key, cell_data in data.get("cells", {}).items():
            r, c = map(int, key.split(","))
            spreadsheet.set_cell_data(r, c, cell_data)

        if "named_ranges" in data:
            spreadsheet.named_ranges.from_dict(data["named_ranges"])

        spreadsheet.frozen_rows = data.get("frozen_rows", 0)
        spreadsheet.frozen_cols = data.get("frozen_cols", 0)

        # Load global settings (with defaults for backward compatibility)
        if "global_settings" in data:
            spreadsheet.global_settings.update(data["global_settings"])

        # Rebuild dependency graph if engine is attached
        spreadsheet.rebuild_dependency_graph()

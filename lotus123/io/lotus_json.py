"""JSON serialization for Spreadsheet."""

from __future__ import annotations

import json
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet
    from ..core.cell import Cell

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
            "col_widths": spreadsheet._col_widths,
            "row_heights": spreadsheet._row_heights,
            "cells": {
                f"{r},{c}": cell.to_dict()
                for (r, c), cell in spreadsheet._cells.items()
                if not cell.is_empty
            },
            "named_ranges": spreadsheet.named_ranges.to_dict(),
            "protection": spreadsheet.protection.to_dict(),
            "frozen_rows": spreadsheet.frozen_rows,
            "frozen_cols": spreadsheet.frozen_cols,
        }
        with open(filename, "w") as f:
            json.dump(data, f, indent=2)

    @staticmethod
    def load(spreadsheet: Spreadsheet, filename: str) -> None:
        """Load spreadsheet from JSON file."""
        from ..core.cell import Cell

        with open(filename, "r") as f:
            data = json.load(f)

        spreadsheet.clear()

        spreadsheet.rows = data.get("rows", MAX_ROWS)
        spreadsheet.cols = data.get("cols", MAX_COLS)
        spreadsheet._col_widths = {int(k): v for k, v in data.get("col_widths", {}).items()}
        spreadsheet._row_heights = {int(k): v for k, v in data.get("row_heights", {}).items()}

        for key, cell_data in data.get("cells", {}).items():
            r, c = map(int, key.split(","))
            spreadsheet._cells[(r, c)] = Cell.from_dict(cell_data)

        if "named_ranges" in data:
            spreadsheet.named_ranges.from_dict(data["named_ranges"])

        if "protection" in data:
            spreadsheet.protection.from_dict(data["protection"])

        spreadsheet.frozen_rows = data.get("frozen_rows", 0)
        spreadsheet.frozen_cols = data.get("frozen_cols", 0)

        # Rebuild dependency graph if engine is attached
        if spreadsheet._recalc_engine:
            spreadsheet._recalc_engine._rebuild_dependency_graph()

"""Lotus 1-2-3 Clone - A full-featured terminal spreadsheet application.

This package provides:
- Core spreadsheet data model with 256 columns x 65,536 rows
- Complete formula engine with 100+ functions
- Text alignment and numeric formatting
- Named ranges and cell protection
- Undo/redo support
- Database operations (sort, query, fill)
- Charting capabilities
- Text file import/export
- Lotus-style menu system
"""

# Core data model
# App
from .app import LotusApp

# Charting
from .charting import Chart, ChartRenderer, ChartType, TextChartRenderer
from .core import (
    Cell,
    CellReference,
    CellType,
    FormatCode,
    NamedRangeManager,
    ProtectionManager,
    RangeReference,
    TextAlignment,
    col_to_index,
    format_value,
    index_to_col,
    make_cell_ref,
    parse_cell_ref,
)
from .core.spreadsheet import Spreadsheet

# Data operations
from .data import CriteriaParser, DatabaseOperations, FillOperations, FillType, SortOrder

# Formula engine
from .formula import (
    FormulaEvaluator,
    FormulaParser,
    FunctionRegistry,
    RecalcEngine,
    RecalcMode,
    RecalcOrder,
)

# File I/O
from .io import ExportOptions, ImportOptions, TextExporter, TextImporter

# UI
from .ui import (
    FrozenTitles,
    Menu,
    MenuAction,
    MenuItem,
    MenuState,
    MenuSystem,
    Mode,
    ModeIndicator,
    SplitType,
    StatusBar,
    TitleFreezeType,
    ViewPort,
    WindowManager,
    WindowSplit,
)

# Utilities
from .utils import Clipboard, ClipboardMode, Command, UndoManager

__version__ = "1.0.0"

__all__ = [
    # Core
    "Cell",
    "CellType",
    "TextAlignment",
    "Spreadsheet",
    "CellReference",
    "RangeReference",
    "parse_cell_ref",
    "make_cell_ref",
    "col_to_index",
    "index_to_col",
    "FormatCode",
    "format_value",
    "NamedRangeManager",
    "ProtectionManager",
    # Formula
    "FormulaParser",
    "FormulaEvaluator",
    "RecalcEngine",
    "RecalcMode",
    "RecalcOrder",
    "FunctionRegistry",
    # Utils
    "UndoManager",
    "Command",
    "Clipboard",
    "ClipboardMode",
    # Data
    "DatabaseOperations",
    "SortOrder",
    "CriteriaParser",
    "FillOperations",
    "FillType",
    # Charting
    "Chart",
    "ChartType",
    "ChartRenderer",
    "TextChartRenderer",
    # I/O
    "TextImporter",
    "TextExporter",
    "ImportOptions",
    "ExportOptions",
    # UI
    "StatusBar",
    "ModeIndicator",
    "Mode",
    "MenuSystem",
    "Menu",
    "MenuItem",
    "MenuAction",
    "MenuState",
    "WindowManager",
    "ViewPort",
    "FrozenTitles",
    "WindowSplit",
    "SplitType",
    "TitleFreezeType",
    # App
    "LotusApp",
]

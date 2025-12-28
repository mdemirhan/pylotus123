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

# For backwards compatibility, import from both old and new locations
# Core data model (new modular structure)
from .core import (
    Cell,
    CellType,
    TextAlignment,
    CellReference,
    RangeReference,
    parse_cell_ref,
    make_cell_ref,
    col_to_index,
    index_to_col,
    FormatCode,
    format_value,
    NamedRangeManager,
    ProtectionManager,
)
from .core.spreadsheet import Spreadsheet

# Formula engine (new modular structure)
from .formula import (
    FormulaParser,
    FormulaEvaluator,
    RecalcEngine,
    RecalcMode,
    RecalcOrder,
    FunctionRegistry,
)

# Utilities
from .utils import (
    UndoManager,
    Command,
    Clipboard,
    ClipboardMode,
)

# Data operations
from .data import (
    DatabaseOperations,
    SortOrder,
    CriteriaParser,
    FillOperations,
    FillType,
)

# Charting
from .charting import (
    Chart,
    ChartType,
    ChartRenderer,
    TextChartRenderer,
)

# File I/O
from .io import (
    TextImporter,
    TextExporter,
    ImportOptions,
    ExportOptions,
)

# UI
from .ui import (
    StatusBar,
    ModeIndicator,
    Mode,
    MenuSystem,
    Menu,
    MenuItem,
    MenuAction,
    MenuState,
    WindowManager,
    ViewPort,
    FrozenTitles,
    WindowSplit,
    SplitType,
    TitleFreezeType,
)

# App (still uses old app.py, will be updated separately)
from .app import LotusApp

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

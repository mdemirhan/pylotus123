# Lotus 1-2-3 Clone - Architecture Design

## Overview

This document describes the modular architecture of the Lotus 1-2-3 clone, a fully-functional
terminal-based spreadsheet application built with Python and the Textual TUI framework.

## Directory Structure

```
lotus123/
├── __init__.py              # Package exports
├── app.py                   # Main TUI application entry point
│
├── core/                    # Core data model
│   ├── __init__.py
│   ├── cell.py              # Cell class with data types, formats, alignment
│   ├── spreadsheet.py       # Main spreadsheet managing the grid
│   ├── reference.py         # Cell/range reference handling (A1, $A$1, ranges)
│   ├── formatting.py        # Numeric, date, time format definitions
│   ├── named_ranges.py      # Named range management
│   └── protection.py        # Cell and worksheet protection
│
├── formula/                 # Formula engine
│   ├── __init__.py
│   ├── parser.py            # Formula tokenizer and parser
│   ├── evaluator.py         # Expression evaluation engine
│   ├── recalc.py            # Recalculation engine with dependency tracking
│   └── functions/           # Function implementations (180+ functions)
│       ├── __init__.py      # Function registry
│       ├── math.py          # @SUM, @ABS, @SQRT, @EXP, @LN, @LOG, trig (27 functions)
│       ├── statistical.py   # @AVG, @STD, @VAR, @COUNT, @MIN, @MAX (32 functions)
│       ├── string.py        # @LEFT, @RIGHT, @MID, @LEN, @TRIM, @UPPER (29 functions)
│       ├── logical.py       # @IF, @AND, @OR, @NOT, @TRUE, @FALSE (24 functions)
│       ├── lookup.py        # @VLOOKUP, @HLOOKUP, @INDEX, @MATCH (14 functions)
│       ├── datetime.py      # @DATE, @NOW, @TODAY, @YEAR, @MONTH (17 functions)
│       ├── financial.py     # @PMT, @PV, @FV, @NPV, @IRR, @RATE (14 functions)
│       ├── database.py      # @DSUM, @DAVG, @DCOUNT, @DMAX, @DMIN (13 functions)
│       └── info.py          # @ISNUMBER, @ISSTRING, @TYPE, @CELL (10 functions)
│
├── ui/                      # User interface components
│   ├── __init__.py
│   ├── grid.py              # SpreadsheetGrid widget with selection
│   ├── menu_bar.py          # LotusMenu widget with hierarchical menus
│   ├── status_bar.py        # Status bar with mode indicators
│   ├── themes.py            # Theme definitions and management
│   ├── config.py            # Application configuration
│   ├── window.py            # Window splitting and frozen titles
│   ├── menu/                # Lotus-style menu system
│   │   ├── __init__.py
│   │   └── menu_system.py   # Complete menu system with all commands
│   └── dialogs/             # Modal dialog screens
│       ├── __init__.py
│       ├── file_dialog.py   # File open/save
│       ├── command_input.py # Generic command/formula input
│       ├── theme_dialog.py  # Theme selection dialog
│       └── chart_view.py    # Chart display screen
│
├── data/                    # Data operations
│   ├── __init__.py
│   ├── database.py          # Sort, Query, Extract operations
│   ├── criteria.py          # Criteria parsing (wildcards, comparisons)
│   └── fill.py              # Fill operations (sequences, patterns)
│
├── io/                      # File I/O operations
│   ├── __init__.py
│   ├── lotus_json.py        # JSON save/load (enhanced)
│   ├── text_import.py       # Import structured/unstructured text
│   └── text_export.py       # Export to text formats
│
├── charting/                # Graphics/charting module
│   ├── __init__.py
│   ├── chart.py             # Chart data model (ChartType, ChartSeries, etc.)
│   ├── renderer.py          # Base chart renderer and TextChartRenderer
│   └── renderers/           # Type-specific chart renderers
│       ├── __init__.py      # Renderer registry
│       ├── line.py          # Line chart renderer
│       ├── bar.py           # Bar chart renderer
│       ├── pie.py           # Pie chart renderer
│       └── scatter.py       # Scatter plot renderer
│
└── utils/                   # Utility modules
    ├── __init__.py
    ├── undo.py              # Undo/redo manager with command pattern
    └── clipboard.py         # Clipboard management for cells/ranges
```

## Core Design Principles

### 1. Separation of Concerns
- **Data Model (core/)**: Pure data representation, no UI dependencies
- **Formula Engine (formula/)**: Parsing and evaluation, uses core but no UI
- **UI Components (ui/)**: Presentation layer, uses core and formula
- **Data Operations (data/)**: Business logic for data manipulation

### 2. Extensibility Patterns

#### Function Registry Pattern (formula/functions/)
```python
# Each function module registers functions with a decorator
@register_function("SUM", "AVG", "COUNT")
def aggregate_function(name: str, values: list) -> Any:
    ...
```

#### Command Pattern (utils/undo.py)
```python
class Command(Protocol):
    def execute(self) -> None: ...
    def undo(self) -> None: ...

class UndoManager:
    def execute(self, command: Command) -> None: ...
    def undo(self) -> None: ...
    def redo(self) -> None: ...
```

#### Plugin-Ready Menu System (ui/menu/)
```python
class MenuItem:
    key: str           # Single keystroke
    label: str         # Display label
    action: Callable   # Handler function
    submenu: list[MenuItem] | None
```

### 3. Data Types & Cell Model

```python
class CellType(Enum):
    EMPTY = auto()
    NUMBER = auto()
    TEXT = auto()
    FORMULA = auto()
    DATE = auto()
    TIME = auto()
    ERROR = auto()

class TextAlignment(Enum):
    DEFAULT = auto()    # Based on content type
    LEFT = auto()       # Apostrophe prefix (')
    RIGHT = auto()      # Quotation mark prefix (")
    CENTER = auto()     # Caret prefix (^)
    REPEAT = auto()     # Backslash prefix (\)

class Cell:
    raw_value: str
    cell_type: CellType
    alignment: TextAlignment
    format_code: str
    is_protected: bool
    # ... computed properties for display
```

### 4. Reference System

```python
class CellReference:
    col: int
    row: int
    col_absolute: bool  # $A vs A
    row_absolute: bool  # $1 vs 1

    def adjust(self, row_delta: int, col_delta: int) -> CellReference:
        """Adjust relative references during copy/paste."""
        ...

class RangeReference:
    start: CellReference
    end: CellReference
```

### 5. Formatting System

```python
class FormatCode(Enum):
    GENERAL = "G"       # Automatic
    FIXED = "F"         # Fixed decimal (0-15 places)
    SCIENTIFIC = "S"    # Scientific notation
    CURRENCY = "C"      # Currency with symbol
    COMMA = ","         # Thousands separator
    PERCENT = "P"       # Percentage
    DATE = "D"          # Date formats (D1-D5)
    TIME = "T"          # Time formats (T1-T4)
    TEXT = "T"          # Display formula as text
    HIDDEN = "H"        # Suppress display
    PLUSMINUS = "+"     # Horizontal bar graph
```

### 6. Recalculation Engine

```python
class RecalcMode(Enum):
    AUTOMATIC = auto()
    MANUAL = auto()

class RecalcOrder(Enum):
    NATURAL = auto()      # Dependency-based
    COLUMN_WISE = auto()  # Left to right, top to bottom
    ROW_WISE = auto()     # Top to bottom, left to right

class RecalcEngine:
    mode: RecalcMode
    order: RecalcOrder

    def build_dependency_graph(self) -> dict[CellRef, set[CellRef]]: ...
    def recalculate(self, changed_cells: set[CellRef] | None = None) -> None: ...
    def detect_circular(self) -> list[CellRef]: ...
```

## Feature Implementation Map

| Feature | Module | Status |
|---------|--------|--------|
| 256×65536 grid | core/spreadsheet.py | ✓ Complete |
| Cell data types | core/cell.py | ✓ Complete |
| Text alignment prefixes | core/cell.py | ✓ Complete |
| Numeric formats | core/formatting.py | ✓ Complete |
| Date/time formats | core/formatting.py | ✓ Complete |
| Absolute references | core/reference.py | ✓ Complete |
| Formula parsing | formula/parser.py | ✓ Complete |
| Math functions (27) | formula/functions/math.py | ✓ Complete |
| Statistical functions (32) | formula/functions/statistical.py | ✓ Complete |
| String functions (29) | formula/functions/string.py | ✓ Complete |
| Logical functions (24) | formula/functions/logical.py | ✓ Complete |
| Lookup functions (14) | formula/functions/lookup.py | ✓ Complete |
| Date/time functions (17) | formula/functions/datetime.py | ✓ Complete |
| Financial functions (14) | formula/functions/financial.py | ✓ Complete |
| Database functions (13) | formula/functions/database.py | ✓ Complete |
| Info functions (10) | formula/functions/info.py | ✓ Complete |
| Recalc engine | formula/recalc.py | ✓ Complete |
| Menu system | ui/menu/ | ✓ Complete |
| Worksheet commands | ui/menu/menu_system.py | ✓ Complete |
| Range commands | ui/menu/menu_system.py | ✓ Complete |
| Named ranges | core/named_ranges.py | ✓ Complete |
| Database operations | data/database.py | ✓ Complete |
| Undo/redo | utils/undo.py | ✓ Complete |
| Window splitting | ui/window.py | ✓ Complete |
| Frozen titles | ui/window.py | ✓ Complete |
| Cell protection | core/protection.py | ✓ Complete |
| Charting (5 types) | charting/ | ✓ Complete |
| Text import/export | io/ | ✓ Complete |
| Status indicators | ui/status_bar.py | ✓ Complete |
| Theme support | ui/themes.py | ✓ Complete |
| Clipboard operations | utils/clipboard.py | ✓ Complete |
| Fill operations | data/fill.py | ✓ Complete |
| Criteria parsing | data/criteria.py | ✓ Complete |

## Testing Strategy

- Unit tests for each module in `tests/` mirroring source structure
- Integration tests for cross-module functionality
- UI tests using Textual's testing framework
- Property-based tests for formula evaluation

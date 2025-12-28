# Lotus 1-2-3 Clone - Architecture Design

## Overview

This document describes the modular architecture designed to support all Lotus 1-2-3 features
while maintaining extensibility for future enhancements.

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
│   └── functions/           # Function implementations (organized by category)
│       ├── __init__.py      # Function registry
│       ├── math.py          # @SUM, @ABS, @SQRT, @EXP, @LN, @LOG, trig, etc.
│       ├── statistical.py   # @AVG, @STD, @VAR, @COUNT, @MIN, @MAX
│       ├── string.py        # @LEFT, @RIGHT, @MID, @LEN, @TRIM, @UPPER, etc.
│       ├── logical.py       # @IF, @AND, @OR, @NOT, @TRUE, @FALSE, @ISERR, etc.
│       ├── lookup.py        # @VLOOKUP, @HLOOKUP, @INDEX, @CHOOSE, @CELL, etc.
│       ├── datetime.py      # @DATE, @NOW, @TODAY, @YEAR, @MONTH, @DAY, etc.
│       └── info.py          # @ISNUMBER, @ISSTRING, @ISNA, @ISERR, @COLS, @ROWS
│
├── ui/                      # User interface components
│   ├── __init__.py
│   ├── grid.py              # SpreadsheetGrid widget
│   ├── menu/                # Lotus-style menu system
│   │   ├── __init__.py
│   │   ├── main_menu.py     # Main slash menu
│   │   ├── file_menu.py     # /File operations
│   │   ├── worksheet_menu.py # /Worksheet commands
│   │   ├── range_menu.py    # /Range commands
│   │   ├── data_menu.py     # /Data commands
│   │   └── graph_menu.py    # /Graph commands
│   ├── dialogs/             # Modal dialog screens
│   │   ├── __init__.py
│   │   ├── file_dialog.py   # File open/save
│   │   ├── input_dialog.py  # Generic command input
│   │   ├── format_dialog.py # Format selection
│   │   └── range_dialog.py  # Range input
│   ├── status_bar.py        # Status bar with mode indicators
│   └── chart_view.py        # Chart display widget
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
│   ├── chart.py             # Base chart data model
│   ├── line_chart.py        # Line graph renderer
│   ├── bar_chart.py         # Bar/stacked bar chart
│   ├── pie_chart.py         # Pie chart
│   └── scatter_chart.py     # XY scatter plot
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
| 256×65536 grid | core/spreadsheet.py | Pending |
| Cell data types | core/cell.py | Pending |
| Text alignment prefixes | core/cell.py | Pending |
| Numeric formats | core/formatting.py | Pending |
| Date/time formats | core/formatting.py | Pending |
| Absolute references | core/reference.py | Pending |
| Formula parsing | formula/parser.py | Exists (enhance) |
| Math functions | formula/functions/math.py | Partial |
| Statistical functions | formula/functions/statistical.py | Pending |
| String functions | formula/functions/string.py | Partial |
| Logical functions | formula/functions/logical.py | Partial |
| Lookup functions | formula/functions/lookup.py | Pending |
| Date/time functions | formula/functions/datetime.py | Pending |
| Recalc engine | formula/recalc.py | Pending |
| Menu system | ui/menu/ | Basic (expand) |
| Worksheet commands | ui/menu/worksheet_menu.py | Pending |
| Range commands | ui/menu/range_menu.py | Pending |
| Named ranges | core/named_ranges.py | Pending |
| Database operations | data/database.py | Pending |
| Undo/redo | utils/undo.py | Pending |
| Window splitting | ui/grid.py | Pending |
| Frozen titles | ui/grid.py | Pending |
| Cell protection | core/protection.py | Pending |
| Charting | charting/ | Pending |
| Text import/export | io/ | Pending |
| Status indicators | ui/status_bar.py | Pending |

## Migration Strategy

1. **Phase 1**: Create new package structure alongside existing code
2. **Phase 2**: Migrate core functionality to new modules
3. **Phase 3**: Update imports in app.py to use new structure
4. **Phase 4**: Add new features in modular fashion
5. **Phase 5**: Comprehensive testing at each phase

## Testing Strategy

- Unit tests for each module in `tests/` mirroring source structure
- Integration tests for cross-module functionality
- UI tests using Textual's testing framework
- Property-based tests for formula evaluation

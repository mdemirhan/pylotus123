# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Lotus 1-2-3 clone - a fully-functional terminal-based spreadsheet application built with Python and the Textual TUI framework. Features a 256×65,536 cell grid, 180+ formula functions, Lotus-style menu system, charting, and full undo/redo support.

## Commands

```bash
# Install dependencies
uv sync

# Run the application
uv run python main.py

# Run all tests
uv run pytest

# Run a single test file
uv run pytest tests/test_formula.py

# Run a specific test
uv run pytest tests/test_formula.py::test_function_name -v
```

## Architecture

The codebase follows a modular architecture with clear separation of concerns:

- **lotus123/app.py** - Main TUI application using Textual framework
- **lotus123/spreadsheet.py** - Core spreadsheet grid data model
- **lotus123/core/** - Core data model (cell, reference, formatting, named ranges, protection)
- **lotus123/formula/** - Formula engine with parser, evaluator, recalc engine, and 180+ functions
- **lotus123/ui/** - UI components (grid, menus, dialogs, status bar, themes, window management)
- **lotus123/data/** - Data operations (sort, query, fill, criteria parsing)
- **lotus123/io/** - File I/O (text import/export in CSV/TSV/delimited formats)
- **lotus123/charting/** - Chart data model and text-based renderers (line, bar, pie, scatter)
- **lotus123/utils/** - Utilities (undo/redo with command pattern, clipboard)

### Key Patterns

**Function Registry** (`formula/functions/`): Functions are registered via decorators for extensibility. Organized into 9 categories: math, statistical, string, logical, lookup, datetime, financial, database, info.

**Command Pattern** (`utils/undo.py`): All undoable operations implement execute/undo protocol with CellChangeCommand, RangeChangeCommand, InsertRowCommand, DeleteRowCommand.

**Menu System** (`ui/menu/menu_system.py`): Lotus-style `/` menu with hierarchical MenuItem structures supporting File, Worksheet, Range, Data, Graph, and Tools menus.

## Formula System

Supports 180+ functions across 9 categories:
- **Math** (27): SUM, ABS, SQRT, EXP, LN, LOG, trig functions, etc.
- **Statistical** (32): AVG, COUNT, MIN, MAX, MEDIAN, STD, VAR, etc.
- **String** (29): LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, etc.
- **Logical** (24): IF, AND, OR, NOT, ISERR, ISNUMBER, etc.
- **Lookup** (14): VLOOKUP, HLOOKUP, INDEX, MATCH, etc.
- **Date/Time** (17): DATE, NOW, TODAY, YEAR, MONTH, DAY, etc.
- **Financial** (14): PMT, PV, FV, NPV, IRR, RATE, etc.
- **Database** (13): DSUM, DAVG, DCOUNT, DMAX, DMIN, etc.
- **Info** (10): TYPE, CELL, ISNUMBER, ISSTRING, etc.

Both `=SUM(...)` and Lotus-style `@SUM(...)` syntax are supported.

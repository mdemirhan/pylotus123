# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Lotus 1-2-3 clone - a terminal-based spreadsheet application built with Python and Textual TUI framework.

## Commands

```bash
# Install dependencies
uv sync

# Run the application
uv run lotus123
# or
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
- **lotus123/core/** - Core data model (cell, reference, formatting)
- **lotus123/formula/** - Modular formula engine with function registry pattern
- **lotus123/ui/** - UI components including menu system and dialogs
- **lotus123/data/** - Data operations (sort, query, fill)
- **lotus123/io/** - File I/O (JSON save/load, text import/export)
- **lotus123/charting/** - Chart rendering modules
- **lotus123/utils/** - Utilities (undo/redo, clipboard)

### Key Patterns

**Function Registry** (`formula/functions/`): Functions are registered via decorators for extensibility.

**Command Pattern** (`utils/undo.py`): All undoable operations implement execute/undo protocol.

**Menu System** (`ui/menu/`): Lotus-style `/` menu with nested MenuItem structures.

## Formula System

Supports 40+ functions across categories: math, statistical, string, logical, lookup, datetime. Both `=SUM(...)` and Lotus-style `@SUM(...)` syntax are supported.

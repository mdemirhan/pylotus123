# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Lotus 1-2-3 Clone - A terminal-based spreadsheet application built with Python and Textual TUI framework. Features 256x65536 grid, 180+ formula functions, charting, and support for multiple file formats (JSON, XLSX, CSV, WK1).

## Commands

```bash
# Install dependencies
uv sync

# Run the application
uv run python main.py
uv run python main.py samples/sales_dashboard.json  # Open a file

# Testing
uv run pytest                                        # Run all tests
uv run pytest tests/core/test_spreadsheet.py        # Run specific file
uv run pytest tests/core/test_spreadsheet.py::TestCellAlignment -v  # Run specific test

# Linting and formatting
uv run ruff check                                    # Lint
uv run ruff format                                   # Format

# Type checking
uv run basedpyright
```

## Development Workflow

After making code changes, always run the type checker (`uv run basedpyright`) and address any issues before considering the work complete.

## Architecture

The codebase follows separation of concerns with these main layers:

- **core/**: Pure data model (Cell, Spreadsheet, CellReference, formatting) - no UI dependencies
- **formula/**: Formula parsing and evaluation engine with function registry pattern
- **handlers/**: Application logic using composition pattern via `AppProtocol` interface
- **ui/**: Textual-based presentation layer (grid, menus, dialogs, themes)
- **data/**: Business logic for sort, query, fill operations
- **io/**: File format handlers (JSON, XLSX, CSV/TSV, WK1)
- **charting/**: Text-based chart renderers
- **utils/**: Undo/redo (command pattern), clipboard management

### Key Patterns

**Handler Composition**: Handlers extend `BaseHandler` and receive the app via `AppProtocol` dependency injection. This provides type-safe contracts and testability.

**Function Registry**: Formula functions use a decorator-based registration pattern in `formula/functions/`.

**Command Pattern**: Undo/redo uses the command pattern in `utils/undo.py`.

### Cell Model

Cells support types (EMPTY, NUMBER, TEXT, FORMULA, DATE, TIME, ERROR) and Lotus-style alignment prefixes:
- `'` = Left, `"` = Right, `^` = Center, `\` = Repeat/Fill

### Formula Syntax

Both `=SUM(A1:A10)` and Lotus-style `@SUM(A1:A10)` are supported.

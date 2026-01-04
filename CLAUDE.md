# Lotus 1-2-3 Clone

Terminal spreadsheet app (Python/Textual). 256x65536 grid, 180+ formulas, charting, multi-format I/O.

## Commands

```bash
uv sync                      # Install dependencies
uv run python main.py [file] # Run app
uv run pytest                # Test (excludes slow tests)
uv run pytest -m slow        # Run slow UI tests only
uv run pytest -m ''          # Run all tests including slow
uv run ruff check lotus123/ && uv run ruff format lotus123/  # Lint & format
uv run basedpyright lotus123/  # Type check
```

After code changes, run `uv run ruff format lotus123/`, `uv run ruff check lotus123/ --fix` and `uv run basedpyright lotus123/`,
fixing any issues before considering work complete. Do NOT run ruff or basedpyright on the tests/ folder.

**Testing strategy**: Run fast tests (`uv run pytest`) by default. Only run slow UI tests (`uv run pytest -m ''`) for:
- Extensive/major changes
- Changes to UI components (`lotus123/ui/`, `lotus123/handlers/`, `lotus123/app.py`)

## Architecture

- **core/**: Data model (Cell, Spreadsheet) - no UI deps
- **formula/**: Parser + function registry pattern
- **handlers/**: App logic via `AppProtocol` DI
- **ui/**: Textual presentation
- **data/**: Sort, query, fill
- **io/**: JSON, XLSX, CSV, WK1
- **charting/**: Text charts
- **utils/**: Undo (command pattern), clipboard

## Key Details

- Handlers extend `BaseHandler`, receive app via `AppProtocol`
- Formula functions: decorator registration in `formula/functions/`
- Cell types: EMPTY, NUMBER, TEXT, FORMULA, DATE, TIME, ERROR
- Alignment prefixes: `'`=Left, `"`=Right, `^`=Center, `\`=Fill
- Formula syntax: `=SUM()` or Lotus-style `@SUM()`

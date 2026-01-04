# Lotus 1-2-3 Clone

Terminal spreadsheet app (Python/Textual). 256x65536 grid, 180+ formulas, charting, multi-format I/O.

## Commands

```bash
uv sync                      # Install dependencies
uv run python main.py [file] # Run app
uv run pytest                # Test
uv run ruff check && uv run ruff format  # Lint & format
uv run basedpyright          # Type check
```

After code changes, run `uv run ruff format` and `uv run basedpyright`, fixing any issues before considering work complete.

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

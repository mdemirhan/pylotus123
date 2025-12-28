# Lotus 1-2-3 Clone

A modern terminal-based spreadsheet application inspired by the classic Lotus 1-2-3.

## Features

- Classic blue/cyan color scheme reminiscent of the original
- Full keyboard navigation
- Formula support with 40+ functions
- Cell references (A1 notation) and ranges (A1:B10)
- Lotus-style `/` menu system
- Save/load spreadsheets to JSON
- Row/column insert/delete
- Copy/cut/paste cells

## Installation

```bash
uv sync
```

## Usage

```bash
uv run lotus123
# or
uv run python main.py
```

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| Arrow keys | Navigate cells |
| Enter/F2 | Edit cell |
| Delete | Clear cell |
| Escape | Cancel edit |
| / | Open menu |
| Ctrl+S | Save |
| Ctrl+O | Open |
| Ctrl+N | New |
| Ctrl+G | Goto cell |
| Ctrl+Q | Quit |
| Page Up/Down | Scroll |
| Home/End | First/last column |

## Formula Functions

### Math
`SUM`, `AVG`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `ABS`, `INT`, `ROUND`, `SQRT`, `POWER`, `MOD`, `SIN`, `COS`, `TAN`, `LOG`, `LN`, `EXP`, `PI`

### String
`LEN`, `LEFT`, `RIGHT`, `MID`, `UPPER`, `LOWER`, `TRIM`, `CONCATENATE`, `CONCAT`, `VALUE`, `TEXT`

### Logical
`IF`, `AND`, `OR`, `NOT`

### Date/Time
`NOW`, `TODAY`

## Examples

```
=A1+B1          # Add two cells
=SUM(A1:A10)    # Sum a range
=@SUM(A1:A10)   # Lotus-style @ prefix
=IF(A1>10,"Yes","No")  # Conditional
=ROUND(A1/B1,2) # Division rounded to 2 decimals
```

## Running Tests

```bash
uv run pytest
```

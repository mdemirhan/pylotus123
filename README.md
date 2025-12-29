# Lotus 1-2-3 Clone

A modern terminal-based spreadsheet application inspired by the classic Lotus 1-2-3.

## Features

- 256 columns (A-IV) x 65,536 rows grid
- 180+ formula functions across 9 categories
- Cell references (A1 notation) with absolute/relative support ($A$1, A1, $A1, A$1)
- Range references (A1:B10)
- Lotus-style `/` menu system with hierarchical menus
- Multiple themes (classic blue, modern dark, light, high contrast)
- Full keyboard navigation
- Row/column insert/delete with width/height adjustment
- Copy/cut/paste with clipboard management
- Undo/redo support
- Named ranges
- Cell protection
- Data operations (sort, query, fill)
- Text import/export (CSV, TSV, custom delimited)
- Charting (line, bar, pie, scatter plots)
- Save/load spreadsheets to JSON

## Installation

```bash
uv sync
```

## Usage

```bash
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
| Ctrl+G/F5 | Goto cell |
| Ctrl+Q | Quit |
| Ctrl+Z | Undo |
| Ctrl+Y | Redo |
| Ctrl+C | Copy |
| Ctrl+X | Cut |
| Ctrl+V | Paste |
| Ctrl+T | Change theme |
| Ctrl+D | Scroll down |
| Ctrl+U | Scroll up |
| F3 | Find |
| F4 | Toggle absolute reference |
| F9 | Recalculate |
| Page Up/Down | Scroll |
| Home/End | First/last column |

## Formula Functions

### Math (27 functions)
`SUM`, `ABS`, `INT`, `ROUND`, `MOD`, `SQRT`, `POWER`, `SIGN`, `TRUNC`, `CEILING`, `FLOOR`, `FACT`, `GCD`, `LCM`, `EXP`, `LN`, `LOG`, `SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `DEGREES`, `RADIANS`, `PI`, `RAND`

### Statistical (32 functions)
`AVG`, `AVERAGE`, `COUNT`, `COUNTA`, `COUNTBLANK`, `MIN`, `MAX`, `MEDIAN`, `STD`, `STDEV`, `VAR`, `MODE`, `LARGE`, `SMALL`, `RANK`, `PERCENTILE`, `QUARTILE`, `GEOMEAN`, `HARMEAN`, `PRODUCT`, `SUMPRODUCT`, `SUMSQ`, `COMBIN`, `PERMUT`

### String (29 functions)
`LEFT`, `RIGHT`, `MID`, `LENGTH`, `LEN`, `FIND`, `SEARCH`, `REPLACE`, `SUBSTITUTE`, `UPPER`, `LOWER`, `PROPER`, `TRIM`, `CLEAN`, `VALUE`, `STRING`, `CHAR`, `CODE`, `CONCAT`, `CONCATENATE`, `EXACT`, `REPEAT`, `REPT`, `TEXT`

### Logical (24 functions)
`IF`, `TRUE`, `FALSE`, `AND`, `OR`, `NOT`, `XOR`, `ISERR`, `ISERROR`, `ISNA`, `ISNUMBER`, `ISSTRING`, `ISTEXT`, `ISBLANK`, `ISLOGICAL`, `ISEVEN`, `ISODD`, `NA`, `ERR`, `IFERROR`, `IFNA`, `SWITCH`, `CHOOSE`

### Lookup (14 functions)
`VLOOKUP`, `HLOOKUP`, `INDEX`, `LOOKUP`, `MATCH`, `ADDRESS`, `INDIRECT`, `COLUMNS`, `COLS`, `COLUMN`, `ROWS`, `ROW`, `OFFSET`, `TRANSPOSE`

### Date/Time (17 functions)
`DATE`, `DATEVALUE`, `DAY`, `MONTH`, `YEAR`, `TIME`, `TIMEVALUE`, `HOUR`, `MINUTE`, `SECOND`, `NOW`, `TODAY`, `WEEKDAY`, `DAYS`, `EDATE`, `EOMONTH`, `YEARFRAC`

### Financial (14 functions)
`PMT`, `PV`, `FV`, `NPV`, `IRR`, `RATE`, `NPER`, `CTERM`, `SLN`, `SYD`, `DDB`, `IPMT`, `PPMT`, `TERM`

### Database (13 functions)
`DSUM`, `DAVG`, `DCOUNT`, `DCOUNTA`, `DGET`, `DMAX`, `DMIN`, `DSTD`, `DSTDEV`, `DVAR`

### Info (10 functions)
`TYPE`, `CELL`, `ISNUMBER`, `ISSTRING`, `ISERR`, `ISNA`, `INFO`, `ISFORMULA`

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

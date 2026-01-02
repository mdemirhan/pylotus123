"""Lotus WK1 format handler.

WK1 is a binary format used by Lotus 1-2-3 Release 2.
File structure: sequence of variable-length records.
Each record: 2-byte opcode (little-endian) + 2-byte length + data.

References:
- http://fileformats.archiveteam.org/wiki/Lotus_1-2-3
- https://docs.fileformat.com/spreadsheet/123/
- Lotus Worksheet File Format Specification (WSFF)
"""

from __future__ import annotations

import math
import struct
from typing import TYPE_CHECKING, BinaryIO

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


# WK1 Record Opcodes (per Lotus 1-2-3 Release 2 specification)
BOF = 0x0000      # Beginning of file
EOF = 0x0001      # End of file
CALCMODE = 0x0002  # Calculation mode (1 byte: 0=manual, 0xFF=automatic)
CALCORDER = 0x0003  # Calculation order (1 byte: 0=natural, 1=column, 2=row)
SPLIT = 0x0004    # Split window type
SYNC = 0x0005     # Split window sync
RANGE = 0x0006    # Active worksheet range
WINDOW1 = 0x0007  # Window 1 record
COLW1 = 0x0008    # Column width, window 1
NAME = 0x000B     # Named range
BLANK = 0x000C    # Blank cell (format only, no value)
INTEGER = 0x000D  # Integer number cell (16-bit signed)
NUMBER = 0x000E   # Floating point number cell (IEEE 754 double)
LABEL = 0x000F    # Label (string) cell
FORMULA = 0x0010  # Formula cell
DEFAULTFORMAT = 0x0050  # Default cell format (Symphony)
HIDCOL = 0x0064   # Hidden column

# Version codes
VERSION_WK1 = 0x0406  # Lotus 1-2-3 Release 2 (WK1)
VERSION_WKS = 0x0404  # Lotus 1-2-3 Release 1A (WKS)

# =============================================================================
# Format Byte Encoding (per Lotus WSFF specification)
# =============================================================================
# The format byte in cell records encodes:
# - Bits 0-3: Format type
# - Bits 4-6: Decimal places (0-7) or date/time variant
# - Bit 7: Protection flag (1 = protected)

# Format type values (bits 0-3)
FMT_FIXED = 0x00       # Fixed decimal
FMT_SCIENTIFIC = 0x01  # Scientific notation
FMT_CURRENCY = 0x02    # Currency
FMT_PERCENT = 0x03     # Percent
FMT_COMMA = 0x04       # Comma (thousands separator)
FMT_UNUSED5 = 0x05     # Unused
FMT_UNUSED6 = 0x06     # Unused
FMT_DATETIME = 0x07    # Date/time (variant in bits 4-6)
# 0x08-0x0E are unused
FMT_DEFAULT = 0x0F     # Default/General format

# Date/time variants when format type is FMT_DATETIME (bits 4-6)
DT_D1 = 0  # DD-MMM-YY
DT_D2 = 1  # DD-MMM
DT_D3 = 2  # MMM-YY
DT_D4 = 3  # MM/DD/YY (Long international)
DT_D5 = 4  # MM/DD
DT_T1 = 5  # HH:MM:SS AM/PM
DT_T2 = 6  # HH:MM AM/PM
DT_T3 = 7  # HH:MM:SS (24-hour)

# =============================================================================
# Formula Bytecode Opcodes (from Lotus WSFF specification)
# =============================================================================

# Data type opcodes (followed by data bytes)
OP_CONSTANT = 0x00      # 8-byte IEEE double follows
OP_VARIABLE = 0x01      # 4-byte cell reference (col 2, row 2)
OP_RANGE = 0x02         # 8-byte range (start col/row, end col/row)
OP_RETURN = 0x03        # End of formula
OP_PAREN = 0x04         # Parentheses marker (cosmetic)
OP_INTEGER = 0x05       # 2-byte signed integer follows

# Operator opcodes (single byte, binary/unary operations)
OP_UNARY_MINUS = 0x08   # Negation
OP_ADD = 0x09           # Addition
OP_SUB = 0x0A           # Subtraction
OP_MUL = 0x0B           # Multiplication
OP_DIV = 0x0C           # Division
OP_EXP = 0x0D           # Exponentiation
OP_EQ = 0x0E            # Equal
OP_NE = 0x0F            # Not equal
OP_LE = 0x10            # Less than or equal
OP_GE = 0x11            # Greater than or equal
OP_LT = 0x12            # Less than
OP_GT = 0x13            # Greater than
OP_AND = 0x14           # Logical AND
OP_OR = 0x15            # Logical OR
OP_NOT = 0x16           # Logical NOT
OP_UNARY_PLUS = 0x17    # Unary plus (ignored)

# Single-argument function opcodes
FN_NA = 0x1F            # @NA
FN_ERR = 0x20           # @ERR
FN_ABS = 0x21           # @ABS(x)
FN_INT = 0x22           # @INT(x)
FN_SQRT = 0x23          # @SQRT(x)
FN_LOG = 0x24           # @LOG(x)
FN_LN = 0x25            # @LN(x)
FN_PI = 0x26            # @PI
FN_SIN = 0x27           # @SIN(x)
FN_COS = 0x28           # @COS(x)
FN_TAN = 0x29           # @TAN(x)
FN_ATAN2 = 0x2A         # @ATAN2(x,y)
FN_ATAN = 0x2B          # @ATAN(x)
FN_ASIN = 0x2C          # @ASIN(x)
FN_ACOS = 0x2D          # @ACOS(x)
FN_EXP = 0x2E           # @EXP(x)
FN_MOD = 0x2F           # @MOD(x,y)
FN_CHOOSE = 0x30        # @CHOOSE(x,v0,v1,...vN)
FN_ISNA = 0x31          # @ISNA(x)
FN_ISERR = 0x32         # @ISERR(x)
FN_FALSE = 0x33         # @FALSE
FN_TRUE = 0x34          # @TRUE
FN_RAND = 0x35          # @RAND
FN_DATE = 0x36          # @DATE(y,m,d)
FN_TODAY = 0x37         # @TODAY
FN_PMT = 0x38           # @PMT(princ,int,term)
FN_PV = 0x39            # @PV(pmt,int,term)
FN_FV = 0x3A            # @FV(pmt,int,term)
FN_IF = 0x3B            # @IF(cond,then,else)
FN_DAY = 0x3C           # @DAY(date)
FN_MONTH = 0x3D         # @MONTH(date)
FN_YEAR = 0x3E          # @YEAR(date)
FN_ROUND = 0x3F         # @ROUND(x,n)
FN_TIME = 0x40          # @TIME(h,m,s)
FN_HOUR = 0x41          # @HOUR(time)
FN_MINUTE = 0x42        # @MINUTE(time)
FN_SECOND = 0x43        # @SECOND(time)
FN_ISN = 0x44           # @ISNUMBER(x)
FN_ISS = 0x45           # @ISSTRING(x)
FN_LENGTH = 0x46        # @LENGTH(s)
FN_VALUE = 0x47         # @VALUE(s)
FN_STRING = 0x48        # @STRING(x,n)
FN_MID = 0x49           # @MID(s,start,n)
FN_CHAR = 0x4A          # @CHAR(n)
FN_CODE = 0x4B          # @CODE(s)
FN_FIND = 0x4C          # @FIND(search,s,start)
FN_DATEVALUE = 0x4D     # @DATEVALUE(s)
FN_TIMEVALUE = 0x4E     # @TIMEVALUE(s)
FN_CELLPOINTER = 0x4F   # @CELLPOINTER(attr)

# Multi-argument function opcodes (followed by 1-byte argument count)
FN_SUM = 0x50           # @SUM(list)
FN_AVG = 0x51           # @AVG(list)
FN_COUNT = 0x52         # @COUNT(list)
FN_MIN = 0x53           # @MIN(list)
FN_MAX = 0x54           # @MAX(list)
FN_VLOOKUP = 0x55       # @VLOOKUP(x,range,col)
FN_NPV = 0x56           # @NPV(int,range)
FN_VAR = 0x57           # @VAR(list)
FN_STD = 0x58           # @STD(list)
FN_IRR = 0x59           # @IRR(guess,range)
FN_HLOOKUP = 0x5A       # @HLOOKUP(x,range,row)
FN_DSUM = 0x5B          # @DSUM(db,col,crit)
FN_DAVG = 0x5C          # @DAVG(db,col,crit)
FN_DCOUNT = 0x5D        # @DCOUNT(db,col,crit)
FN_DMIN = 0x5E          # @DMIN(db,col,crit)
FN_DMAX = 0x5F          # @DMAX(db,col,crit)
FN_DVAR = 0x60          # @DVAR(db,col,crit)
FN_DSTD = 0x61          # @DSTD(db,col,crit)
FN_INDEX = 0x62         # @INDEX(range,col,row)
FN_COLS = 0x63          # @COLS(range)
FN_ROWS = 0x64          # @ROWS(range)
FN_REPEAT = 0x65        # @REPEAT(s,n)
FN_UPPER = 0x66         # @UPPER(s)
FN_LOWER = 0x67         # @LOWER(s)
FN_LEFT = 0x68          # @LEFT(s,n)
FN_RIGHT = 0x69         # @RIGHT(s,n)
FN_REPLACE = 0x6A       # @REPLACE(s,start,n,new)
FN_PROPER = 0x6B        # @PROPER(s)
FN_CELL = 0x6C          # @CELL(attr,range)
FN_TRIM = 0x6D          # @TRIM(s)
FN_CLEAN = 0x6E         # @CLEAN(s)
FN_S = 0x6F             # @S(range)
FN_N = 0x70             # @N(range)
FN_EXACT = 0x71         # @EXACT(s1,s2)
FN_CALL = 0x72          # @CALL - external call
FN_INDIRECT = 0x73      # @@ or @INDIRECT

# Function name mapping for decompilation
FUNCTION_NAMES: dict[int, tuple[str, int]] = {
    # (name, arg_count) where -1 means variable args
    FN_NA: ("NA", 0),
    FN_ERR: ("ERR", 0),
    FN_ABS: ("ABS", 1),
    FN_INT: ("INT", 1),
    FN_SQRT: ("SQRT", 1),
    FN_LOG: ("LOG", 1),
    FN_LN: ("LN", 1),
    FN_PI: ("PI", 0),
    FN_SIN: ("SIN", 1),
    FN_COS: ("COS", 1),
    FN_TAN: ("TAN", 1),
    FN_ATAN2: ("ATAN2", 2),
    FN_ATAN: ("ATAN", 1),
    FN_ASIN: ("ASIN", 1),
    FN_ACOS: ("ACOS", 1),
    FN_EXP: ("EXP", 1),
    FN_MOD: ("MOD", 2),
    FN_CHOOSE: ("CHOOSE", -1),
    FN_ISNA: ("ISNA", 1),
    FN_ISERR: ("ISERR", 1),
    FN_FALSE: ("FALSE", 0),
    FN_TRUE: ("TRUE", 0),
    FN_RAND: ("RAND", 0),
    FN_DATE: ("DATE", 3),
    FN_TODAY: ("TODAY", 0),
    FN_PMT: ("PMT", 3),
    FN_PV: ("PV", 3),
    FN_FV: ("FV", 3),
    FN_IF: ("IF", 3),
    FN_DAY: ("DAY", 1),
    FN_MONTH: ("MONTH", 1),
    FN_YEAR: ("YEAR", 1),
    FN_ROUND: ("ROUND", 2),
    FN_TIME: ("TIME", 3),
    FN_HOUR: ("HOUR", 1),
    FN_MINUTE: ("MINUTE", 1),
    FN_SECOND: ("SECOND", 1),
    FN_ISN: ("ISNUMBER", 1),
    FN_ISS: ("ISSTRING", 1),
    FN_LENGTH: ("LENGTH", 1),
    FN_VALUE: ("VALUE", 1),
    FN_STRING: ("STRING", 2),
    FN_MID: ("MID", 3),
    FN_CHAR: ("CHAR", 1),
    FN_CODE: ("CODE", 1),
    FN_FIND: ("FIND", 3),
    FN_DATEVALUE: ("DATEVALUE", 1),
    FN_TIMEVALUE: ("TIMEVALUE", 1),
    FN_CELLPOINTER: ("CELLPOINTER", 1),
    FN_SUM: ("SUM", -1),
    FN_AVG: ("AVG", -1),
    FN_COUNT: ("COUNT", -1),
    FN_MIN: ("MIN", -1),
    FN_MAX: ("MAX", -1),
    FN_VLOOKUP: ("VLOOKUP", 3),
    FN_NPV: ("NPV", 2),
    FN_VAR: ("VAR", -1),
    FN_STD: ("STD", -1),
    FN_IRR: ("IRR", 2),
    FN_HLOOKUP: ("HLOOKUP", 3),
    FN_DSUM: ("DSUM", 3),
    FN_DAVG: ("DAVG", 3),
    FN_DCOUNT: ("DCOUNT", 3),
    FN_DMIN: ("DMIN", 3),
    FN_DMAX: ("DMAX", 3),
    FN_DVAR: ("DVAR", 3),
    FN_DSTD: ("DSTD", 3),
    FN_INDEX: ("INDEX", 3),
    FN_COLS: ("COLS", 1),
    FN_ROWS: ("ROWS", 1),
    FN_REPEAT: ("REPEAT", 2),
    FN_UPPER: ("UPPER", 1),
    FN_LOWER: ("LOWER", 1),
    FN_LEFT: ("LEFT", 2),
    FN_RIGHT: ("RIGHT", 2),
    FN_REPLACE: ("REPLACE", 4),
    FN_PROPER: ("PROPER", 1),
    FN_CELL: ("CELL", 2),
    FN_TRIM: ("TRIM", 1),
    FN_CLEAN: ("CLEAN", 1),
    FN_S: ("S", 1),
    FN_N: ("N", 1),
    FN_EXACT: ("EXACT", 2),
    FN_INDIRECT: ("INDIRECT", 1),
}

# Reverse mapping: function name to opcode
FUNCTION_OPCODES: dict[str, int] = {name: op for op, (name, _) in FUNCTION_NAMES.items()}

# Binary operator precedence (higher = binds tighter)
BINARY_PRECEDENCE: dict[int, int] = {
    OP_OR: 1,
    OP_AND: 1,
    OP_EQ: 3,
    OP_NE: 3,
    OP_LE: 3,
    OP_GE: 3,
    OP_LT: 3,
    OP_GT: 3,
    OP_ADD: 4,
    OP_SUB: 4,
    OP_MUL: 5,
    OP_DIV: 5,
    OP_EXP: 7,
}

# Operator symbols for decompilation
BINARY_OPERATORS: dict[int, str] = {
    OP_ADD: "+",
    OP_SUB: "-",
    OP_MUL: "*",
    OP_DIV: "/",
    OP_EXP: "^",
    OP_EQ: "=",
    OP_NE: "<>",
    OP_LE: "<=",
    OP_GE: ">=",
    OP_LT: "<",
    OP_GT: ">",
    OP_AND: "#AND#",
    OP_OR: "#OR#",
}


# =============================================================================
# Format Byte Encoding/Decoding
# =============================================================================

def encode_format_byte(format_code: str) -> int:
    """Encode app format code to WK1 format byte.

    Args:
        format_code: App format code (e.g., "F2", "C0", "D1", "G")

    Returns:
        WK1 format byte (0-255)
    """
    if not format_code or format_code == "G":
        return FMT_DEFAULT  # General/Default

    format_code = format_code.upper()
    fmt_char = format_code[0]
    decimals = 2  # default

    if len(format_code) > 1:
        try:
            decimals = int(format_code[1:])
            decimals = max(0, min(7, decimals))  # WK1 only supports 0-7
        except ValueError:
            pass

    if fmt_char == "F":
        return FMT_FIXED | (decimals << 4)
    elif fmt_char == "S":
        return FMT_SCIENTIFIC | (decimals << 4)
    elif fmt_char == "C":
        return FMT_CURRENCY | (decimals << 4)
    elif fmt_char == "P":
        return FMT_PERCENT | (decimals << 4)
    elif fmt_char == ",":
        return FMT_COMMA | (decimals << 4)
    elif fmt_char == "D":
        # Date formats D1-D5
        variant = decimals - 1 if 1 <= decimals <= 5 else 0
        return FMT_DATETIME | (variant << 4)
    elif fmt_char == "T":
        # Time formats T1-T4 map to variants 5-7 (T4 uses T3)
        if decimals == 1:
            variant = DT_T1
        elif decimals == 2:
            variant = DT_T2
        elif decimals in (3, 4):
            variant = DT_T3
        else:
            variant = DT_T1
        return FMT_DATETIME | (variant << 4)
    elif fmt_char == "H":
        # Hidden - use special value (we'll use 0xFF but strip protection bit)
        return 0x7F  # All bits except protection
    elif fmt_char == "+":
        # Plus/minus bar graph - not directly supported in WK1, use General
        return FMT_DEFAULT

    return FMT_DEFAULT


def decode_format_byte(fmt_byte: int) -> str:
    """Decode WK1 format byte to app format code.

    Args:
        fmt_byte: WK1 format byte

    Returns:
        App format code (e.g., "F2", "C0", "D1", "G")
    """
    # Extract format type (bits 0-3) and decimals/variant (bits 4-6)
    fmt_type = fmt_byte & 0x0F
    decimals = (fmt_byte >> 4) & 0x07
    # protection = (fmt_byte >> 7) & 0x01  # Bit 7 - not used currently

    if fmt_type == FMT_DEFAULT:
        return "G"
    elif fmt_type == FMT_FIXED:
        return f"F{decimals}"
    elif fmt_type == FMT_SCIENTIFIC:
        return f"S{decimals}"
    elif fmt_type == FMT_CURRENCY:
        return f"C{decimals}"
    elif fmt_type == FMT_PERCENT:
        return f"P{decimals}"
    elif fmt_type == FMT_COMMA:
        return f",{decimals}"
    elif fmt_type == FMT_DATETIME:
        # Date/time formats based on variant
        if decimals == DT_D1:
            return "D1"
        elif decimals == DT_D2:
            return "D2"
        elif decimals == DT_D3:
            return "D3"
        elif decimals == DT_D4:
            return "D4"
        elif decimals == DT_D5:
            return "D5"
        elif decimals == DT_T1:
            return "T1"
        elif decimals == DT_T2:
            return "T2"
        elif decimals == DT_T3:
            return "T3"
        else:
            return "D1"  # Default to D1

    return "G"  # Default to General


def _col_to_letter(col: int) -> str:
    """Convert 0-based column index to letter(s)."""
    result = ""
    col += 1  # Make 1-based
    while col > 0:
        col -= 1
        result = chr(ord('A') + (col % 26)) + result
        col //= 26
    return result


def _letter_to_col(letters: str) -> int:
    """Convert column letter(s) to 0-based index."""
    result = 0
    for char in letters.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1


class FormulaDecompiler:
    """Decompile WK1 formula bytecode to text formula."""

    def __init__(self, bytecode: bytes, formula_row: int = 0, formula_col: int = 0) -> None:
        self.bytecode = bytecode
        self.pos = 0
        self.stack: list[tuple[str, int]] = []  # (expression, precedence)
        self.formula_row = formula_row
        self.formula_col = formula_col

    def decompile(self) -> str:
        """Decompile bytecode to formula string."""
        while self.pos < len(self.bytecode):
            opcode = self.bytecode[self.pos]
            self.pos += 1

            if opcode == OP_RETURN:
                break
            elif opcode == OP_CONSTANT:
                self._read_constant()
            elif opcode == OP_INTEGER:
                self._read_integer()
            elif opcode == OP_VARIABLE:
                self._read_variable()
            elif opcode == OP_RANGE:
                self._read_range()
            elif opcode == OP_PAREN:
                # Parentheses marker - wrap top of stack
                if self.stack:
                    expr, _ = self.stack.pop()
                    self.stack.append((f"({expr})", 99))
            elif opcode == OP_UNARY_MINUS:
                if self.stack:
                    expr, prec = self.stack.pop()
                    if prec < 6:
                        expr = f"({expr})"
                    self.stack.append((f"-{expr}", 6))
            elif opcode == OP_UNARY_PLUS:
                # Unary plus is ignored
                pass
            elif opcode == OP_NOT:
                if self.stack:
                    expr, _ = self.stack.pop()
                    self.stack.append((f"#NOT#{expr}", 2))
            elif opcode in BINARY_OPERATORS:
                self._binary_op(opcode)
            elif opcode in FUNCTION_NAMES:
                self._function(opcode)
            elif opcode >= 0x50:
                # Multi-arg function with arg count
                self._multi_arg_function(opcode)

        if self.stack:
            return "=" + self.stack[-1][0]
        return "=0"

    def _read_constant(self) -> None:
        """Read 8-byte IEEE double constant."""
        if self.pos + 8 <= len(self.bytecode):
            value = struct.unpack("<d", self.bytecode[self.pos:self.pos + 8])[0]
            self.pos += 8
            if value == int(value):
                self.stack.append((str(int(value)), 99))
            else:
                self.stack.append((str(value), 99))

    def _read_integer(self) -> None:
        """Read 2-byte signed integer constant."""
        if self.pos + 2 <= len(self.bytecode):
            value = struct.unpack("<h", self.bytecode[self.pos:self.pos + 2])[0]
            self.pos += 2
            self.stack.append((str(value), 99))

    def _read_variable(self) -> None:
        """Read 4-byte cell reference.

        WK1 cell reference format:
        - Column (2 bytes): bit 15 = relative flag, bits 0-7 = column or offset
        - Row (2 bytes): bit 15 = relative flag, bits 0-13 = row or offset

        For relative references, offset is signed (2's complement in relevant bits).
        """
        if self.pos + 4 <= len(self.bytecode):
            col_word = struct.unpack("<H", self.bytecode[self.pos:self.pos + 2])[0]
            row_word = struct.unpack("<H", self.bytecode[self.pos + 2:self.pos + 4])[0]
            self.pos += 4

            # Check if column is relative (bit 15 set)
            col_relative = (col_word & 0x8000) != 0
            # Column value is in bits 0-7 (8-bit value)
            col_val = col_word & 0xFF
            if col_relative:
                # Signed 8-bit offset
                if col_val >= 128:
                    col_val -= 256
                col = self.formula_col + col_val
            else:
                col = col_val

            # Check if row is relative (bit 15 set)
            row_relative = (row_word & 0x8000) != 0
            # Row value is in bits 0-13 (14-bit value)
            row_val = row_word & 0x3FFF
            if row_relative:
                # Signed 14-bit offset
                if row_val >= 0x2000:  # 8192 = 2^13
                    row_val -= 0x4000  # 16384 = 2^14
                row = self.formula_row + row_val
            else:
                row = row_val

            cell_ref = f"{_col_to_letter(col)}{row + 1}"
            self.stack.append((cell_ref, 99))

    def _read_range(self) -> None:
        """Read 8-byte range reference.

        WK1 range reference format is same as cell reference, but for both corners.
        """
        if self.pos + 8 <= len(self.bytecode):
            start_col_word = struct.unpack("<H", self.bytecode[self.pos:self.pos + 2])[0]
            start_row_word = struct.unpack("<H", self.bytecode[self.pos + 2:self.pos + 4])[0]
            end_col_word = struct.unpack("<H", self.bytecode[self.pos + 4:self.pos + 6])[0]
            end_row_word = struct.unpack("<H", self.bytecode[self.pos + 6:self.pos + 8])[0]
            self.pos += 8

            # Resolve start cell
            start_col = self._resolve_col(start_col_word)
            start_row = self._resolve_row(start_row_word)

            # Resolve end cell
            end_col = self._resolve_col(end_col_word)
            end_row = self._resolve_row(end_row_word)

            range_ref = (
                f"{_col_to_letter(start_col)}{start_row + 1}:"
                f"{_col_to_letter(end_col)}{end_row + 1}"
            )
            self.stack.append((range_ref, 99))

    def _resolve_col(self, col_word: int) -> int:
        """Resolve column word to absolute column index."""
        col_relative = (col_word & 0x8000) != 0
        col_val = col_word & 0xFF
        if col_relative:
            if col_val >= 128:
                col_val -= 256
            return self.formula_col + col_val
        return col_val

    def _resolve_row(self, row_word: int) -> int:
        """Resolve row word to absolute row index."""
        row_relative = (row_word & 0x8000) != 0
        row_val = row_word & 0x3FFF
        if row_relative:
            if row_val >= 0x2000:
                row_val -= 0x4000
            return self.formula_row + row_val
        return row_val

    def _binary_op(self, opcode: int) -> None:
        """Handle binary operator."""
        if len(self.stack) >= 2:
            right, right_prec = self.stack.pop()
            left, left_prec = self.stack.pop()
            op_prec = BINARY_PRECEDENCE.get(opcode, 0)
            symbol = BINARY_OPERATORS[opcode]

            # Add parentheses if needed
            if left_prec < op_prec:
                left = f"({left})"
            if right_prec < op_prec or (right_prec == op_prec and opcode in (OP_SUB, OP_DIV)):
                right = f"({right})"

            self.stack.append((f"{left}{symbol}{right}", op_prec))

    def _function(self, opcode: int) -> None:
        """Handle function call."""
        if opcode not in FUNCTION_NAMES:
            return

        name, arg_count = FUNCTION_NAMES[opcode]

        if arg_count == 0:
            # No-argument function - always include () for parser compatibility
            self.stack.append((f"@{name}()", 99))
        elif arg_count == -1:
            # Variable argument function - read arg count
            if self.pos < len(self.bytecode):
                actual_count = self.bytecode[self.pos]
                self.pos += 1
                args: list[str] = []
                for _ in range(actual_count):
                    if self.stack:
                        args.insert(0, self.stack.pop()[0])
                self.stack.append((f"@{name}({','.join(args)})", 99))
        else:
            # Fixed argument function
            args = []
            for _ in range(arg_count):
                if self.stack:
                    args.insert(0, self.stack.pop()[0])
            self.stack.append((f"@{name}({','.join(args)})", 99))

    def _multi_arg_function(self, opcode: int) -> None:
        """Handle multi-argument function (opcode >= 0x50)."""
        if opcode in FUNCTION_NAMES:
            name, expected_args = FUNCTION_NAMES[opcode]
            if expected_args == -1:
                # Variable args - read count from bytecode
                if self.pos < len(self.bytecode):
                    arg_count = self.bytecode[self.pos]
                    self.pos += 1
                    args: list[str] = []
                    for _ in range(arg_count):
                        if self.stack:
                            args.insert(0, self.stack.pop()[0])
                    self.stack.append((f"@{name}({','.join(args)})", 99))
            else:
                self._function(opcode)


class FormulaCompiler:
    """Compile text formula to WK1 bytecode."""

    def __init__(self, formula: str, formula_row: int = 0, formula_col: int = 0) -> None:
        # Remove leading = or @
        self.formula = formula.lstrip("=@").strip()
        self.pos = 0
        self.bytecode = bytearray()
        self.formula_row = formula_row
        self.formula_col = formula_col

    def compile(self) -> bytes:
        """Compile formula to bytecode."""
        try:
            self._parse_expression()
            self.bytecode.append(OP_RETURN)
            return bytes(self.bytecode)
        except Exception:
            # On any error, return empty formula (just return)
            return bytes([OP_RETURN])

    def _peek(self) -> str:
        """Peek at current character."""
        self._skip_whitespace()
        if self.pos < len(self.formula):
            return self.formula[self.pos]
        return ""

    def _advance(self) -> str:
        """Consume and return current character."""
        self._skip_whitespace()
        if self.pos < len(self.formula):
            char = self.formula[self.pos]
            self.pos += 1
            return char
        return ""

    def _skip_whitespace(self) -> None:
        """Skip whitespace."""
        while self.pos < len(self.formula) and self.formula[self.pos].isspace():
            self.pos += 1

    def _parse_expression(self, min_prec: int = 0) -> None:
        """Parse expression with operator precedence."""
        self._parse_unary()

        while True:
            self._skip_whitespace()
            op = self._peek_operator()
            if op is None:
                break

            op_prec = BINARY_PRECEDENCE.get(op, 0)
            if op_prec < min_prec:
                break

            self._consume_operator(op)
            self._parse_expression(op_prec + 1)
            self.bytecode.append(op)

    def _peek_operator(self) -> int | None:
        """Check for binary operator at current position."""
        if self.pos >= len(self.formula):
            return None

        remaining = self.formula[self.pos:]

        # Multi-char operators first
        if remaining.startswith("<>"):
            return OP_NE
        if remaining.startswith("<="):
            return OP_LE
        if remaining.startswith(">="):
            return OP_GE
        if remaining.upper().startswith("#AND#"):
            return OP_AND
        if remaining.upper().startswith("#OR#"):
            return OP_OR

        # Single char operators
        char = remaining[0]
        if char == "+":
            return OP_ADD
        if char == "-":
            return OP_SUB
        if char == "*":
            return OP_MUL
        if char == "/":
            return OP_DIV
        if char == "^":
            return OP_EXP
        if char == "=":
            return OP_EQ
        if char == "<":
            return OP_LT
        if char == ">":
            return OP_GT

        return None

    def _consume_operator(self, op: int) -> None:
        """Consume operator from input."""
        if op == OP_NE:
            self.pos += 2
        elif op == OP_LE:
            self.pos += 2
        elif op == OP_GE:
            self.pos += 2
        elif op == OP_AND:
            self.pos += 5
        elif op == OP_OR:
            self.pos += 4
        else:
            self.pos += 1

    def _parse_unary(self) -> None:
        """Parse unary expression."""
        self._skip_whitespace()
        char = self._peek()

        if char == "-":
            self._advance()
            self._parse_unary()
            self.bytecode.append(OP_UNARY_MINUS)
        elif char == "+":
            self._advance()
            self._parse_unary()
            # Unary plus is a no-op
        elif char.upper() == "#" and self.formula[self.pos:].upper().startswith("#NOT#"):
            self.pos += 5
            self._parse_unary()
            self.bytecode.append(OP_NOT)
        else:
            self._parse_primary()

    def _parse_primary(self) -> None:
        """Parse primary expression (number, cell ref, function, parentheses)."""
        self._skip_whitespace()
        char = self._peek()

        if char == "(":
            self._advance()  # consume (
            self._parse_expression()
            self.bytecode.append(OP_PAREN)
            if self._peek() == ")":
                self._advance()  # consume )
        elif char == "@":
            self._parse_function()
        elif char.isdigit() or char == ".":
            self._parse_number()
        elif char.isalpha() or char == "$":
            # Could be cell reference, range, or function without @
            # $ prefix indicates absolute reference (e.g., $A$1)
            self._parse_cell_or_function()
        elif char == '"':
            # String literal - not directly supported, skip
            self._advance()
            while self.pos < len(self.formula) and self.formula[self.pos] != '"':
                self.pos += 1
            if self.pos < len(self.formula):
                self._advance()
            # Push 0 as placeholder
            self.bytecode.append(OP_INTEGER)
            self.bytecode.extend(struct.pack("<h", 0))

    def _parse_number(self) -> None:
        """Parse numeric literal."""
        start = self.pos
        has_dot = False

        while self.pos < len(self.formula):
            char = self.formula[self.pos]
            if char.isdigit():
                self.pos += 1
            elif char == "." and not has_dot:
                has_dot = True
                self.pos += 1
            elif char.upper() == "E":
                # Scientific notation
                self.pos += 1
                if self.pos < len(self.formula) and self.formula[self.pos] in "+-":
                    self.pos += 1
                while self.pos < len(self.formula) and self.formula[self.pos].isdigit():
                    self.pos += 1
                break
            else:
                break

        num_str = self.formula[start:self.pos]
        try:
            value = float(num_str)
            if value == int(value) and -32768 <= value <= 32767:
                # Use integer encoding
                self.bytecode.append(OP_INTEGER)
                self.bytecode.extend(struct.pack("<h", int(value)))
            else:
                # Use double encoding
                self.bytecode.append(OP_CONSTANT)
                self.bytecode.extend(struct.pack("<d", value))
        except ValueError:
            # Default to 0
            self.bytecode.append(OP_INTEGER)
            self.bytecode.extend(struct.pack("<h", 0))

    def _parse_cell_or_function(self) -> None:
        """Parse cell reference, range, or function name."""
        start = self.pos

        # Read identifier
        while self.pos < len(self.formula):
            char = self.formula[self.pos]
            if char.isalnum() or char == "$":
                self.pos += 1
            else:
                break

        ident = self.formula[start:self.pos]

        # Check if it's a function (followed by parenthesis)
        self._skip_whitespace()
        if self._peek() == "(":
            # It's a function - self.pos already points at "("
            self._parse_function_by_name(ident.upper())
        elif self._peek() == ":" or ":" in ident:
            # It's a range
            self.pos = start  # Reset
            self._parse_range()
        else:
            # It's a cell reference
            self._parse_cell_ref(ident)

    def _parse_cell_ref(self, ref: str) -> None:
        """Parse cell reference like A1, $A$1.

        References with $ are absolute; without $ are relative.
        """
        # Check for absolute column ($A) and row ($1)
        col_absolute = ref.startswith("$")
        # Find if there's a $ before the row number
        row_absolute = False
        in_col = True
        for i, char in enumerate(ref):
            if char == "$":
                if in_col and i > 0:
                    # $ not at start means row absolute
                    row_absolute = True
                continue
            if char.isdigit():
                # Check if there's a $ right before the digits
                if i > 0 and ref[i - 1] == "$":
                    row_absolute = True
                in_col = False

        # Remove $ signs for parsing
        clean_ref = ref.replace("$", "")

        # Extract column letters and row number
        col_str = ""
        row_str = ""
        for char in clean_ref:
            if char.isalpha():
                col_str += char
            elif char.isdigit():
                row_str += char

        if col_str and row_str:
            target_col = _letter_to_col(col_str)
            target_row = int(row_str) - 1  # Convert to 0-based

            # Encode column
            if col_absolute:
                col_word = target_col  # No relative flag
            else:
                # Relative: compute offset and set bit 15
                col_offset = target_col - self.formula_col
                # Encode as signed 8-bit in bits 0-7
                if col_offset < 0:
                    col_offset += 256
                col_word = 0x8000 | (col_offset & 0xFF)

            # Encode row
            if row_absolute:
                row_word = target_row  # No relative flag
            else:
                # Relative: compute offset and set bit 15
                row_offset = target_row - self.formula_row
                # Encode as signed 14-bit in bits 0-13
                if row_offset < 0:
                    row_offset += 0x4000
                row_word = 0x8000 | (row_offset & 0x3FFF)

            self.bytecode.append(OP_VARIABLE)
            self.bytecode.extend(struct.pack("<HH", col_word, row_word))
        else:
            # Invalid reference, push 0
            self.bytecode.append(OP_INTEGER)
            self.bytecode.extend(struct.pack("<h", 0))

    def _parse_range(self) -> None:
        """Parse range reference like A1:B10, $A$1:$B$10."""
        start = self.pos

        # Read until we have the full range
        while self.pos < len(self.formula):
            char = self.formula[self.pos]
            if char.isalnum() or char in "$:":
                self.pos += 1
            else:
                break

        range_str = self.formula[start:self.pos]

        if ":" in range_str:
            parts = range_str.split(":")
            if len(parts) == 2:
                start_ref = parts[0]
                end_ref = parts[1]

                start_word = self._encode_cell_ref(start_ref)
                end_word = self._encode_cell_ref(end_ref)

                if start_word and end_word:
                    self.bytecode.append(OP_RANGE)
                    self.bytecode.extend(struct.pack("<HHHH",
                        start_word[0], start_word[1], end_word[0], end_word[1]))
                    return

        # Fallback - invalid range
        self.bytecode.append(OP_INTEGER)
        self.bytecode.extend(struct.pack("<h", 0))

    def _encode_cell_ref(self, ref: str) -> tuple[int, int] | None:
        """Encode a cell reference to (col_word, row_word) tuple.

        Returns None if the reference is invalid.
        """
        # Check for absolute column ($A) and row ($1)
        col_absolute = ref.startswith("$")
        row_absolute = False
        in_col = True
        for i, char in enumerate(ref):
            if char == "$":
                if in_col and i > 0:
                    row_absolute = True
                continue
            if char.isdigit():
                if i > 0 and ref[i - 1] == "$":
                    row_absolute = True
                in_col = False

        # Remove $ signs for parsing
        clean_ref = ref.replace("$", "")

        # Extract column letters and row number
        col_str = ""
        row_str = ""
        for char in clean_ref:
            if char.isalpha():
                col_str += char
            elif char.isdigit():
                row_str += char

        if not col_str or not row_str:
            return None

        target_col = _letter_to_col(col_str)
        target_row = int(row_str) - 1

        # Encode column
        if col_absolute:
            col_word = target_col
        else:
            col_offset = target_col - self.formula_col
            if col_offset < 0:
                col_offset += 256
            col_word = 0x8000 | (col_offset & 0xFF)

        # Encode row
        if row_absolute:
            row_word = target_row
        else:
            row_offset = target_row - self.formula_row
            if row_offset < 0:
                row_offset += 0x4000
            row_word = 0x8000 | (row_offset & 0x3FFF)

        return (col_word, row_word)

    def _parse_function(self) -> None:
        """Parse @FUNCTION(...) style function."""
        self._advance()  # consume @

        # Read function name
        start = self.pos
        while self.pos < len(self.formula):
            char = self.formula[self.pos]
            if char.isalnum():
                self.pos += 1
            else:
                break

        func_name = self.formula[start:self.pos].upper()
        self._parse_function_by_name(func_name)

    def _parse_function_by_name(self, func_name: str) -> None:
        """Parse function by name."""
        if func_name not in FUNCTION_OPCODES:
            # Unknown function - push 0
            self._skip_to_matching_paren()
            self.bytecode.append(OP_INTEGER)
            self.bytecode.extend(struct.pack("<h", 0))
            return

        opcode = FUNCTION_OPCODES[func_name]
        _, expected_args = FUNCTION_NAMES[opcode]

        # Parse arguments
        args: list[bytearray] = []
        self._skip_whitespace()

        if self._peek() == "(":
            self._advance()  # consume (

            while True:
                self._skip_whitespace()
                if self._peek() == ")":
                    self._advance()
                    break
                if self._peek() == "":
                    break

                # Parse one argument (stop at comma or closing paren)
                paren_depth = 0
                arg_start_pos = self.pos

                # Find end of argument
                while self.pos < len(self.formula):
                    char = self.formula[self.pos]
                    if char == "(":
                        paren_depth += 1
                        self.pos += 1
                    elif char == ")":
                        if paren_depth == 0:
                            break
                        paren_depth -= 1
                        self.pos += 1
                    elif char == "," and paren_depth == 0:
                        break
                    else:
                        self.pos += 1

                # Compile the argument (pass formula position for relative refs)
                arg_str = self.formula[arg_start_pos:self.pos]
                if arg_str.strip():
                    arg_comp = FormulaCompiler(arg_str, self.formula_row, self.formula_col)
                    arg_comp._parse_expression()
                    args.append(arg_comp.bytecode)

                if self._peek() == ",":
                    self._advance()

        # Write arguments to bytecode
        for arg in args:
            self.bytecode.extend(arg)

        # Write function opcode
        self.bytecode.append(opcode)

        # For variable-arg functions, write arg count
        if expected_args == -1:
            self.bytecode.append(len(args))

    def _skip_to_matching_paren(self) -> None:
        """Skip to matching closing parenthesis."""
        if self._peek() == "(":
            self._advance()
            depth = 1
            while depth > 0 and self.pos < len(self.formula):
                char = self._advance()
                if char == "(":
                    depth += 1
                elif char == ")":
                    depth -= 1


def decompile_formula(bytecode: bytes, formula_row: int = 0, formula_col: int = 0) -> str:
    """Decompile WK1 formula bytecode to text formula.

    Args:
        bytecode: Formula bytecode from WK1 file
        formula_row: Row index of the cell containing the formula (0-based)
        formula_col: Column index of the cell containing the formula (0-based)

    Returns:
        Text formula string (e.g., "=A1+B1")
    """
    return FormulaDecompiler(bytecode, formula_row, formula_col).decompile()


def compile_formula(formula: str, formula_row: int = 0, formula_col: int = 0) -> bytes:
    """Compile text formula to WK1 bytecode.

    Args:
        formula: Text formula (e.g., "=A1+B1" or "=@SUM(A1:A10)")
        formula_row: Row index of the cell containing the formula (0-based)
        formula_col: Column index of the cell containing the formula (0-based)

    Returns:
        Formula bytecode for WK1 file
    """
    return FormulaCompiler(formula, formula_row, formula_col).compile()


class Wk1Reader:
    """Read Lotus WK1 files."""

    def __init__(self, spreadsheet: Spreadsheet) -> None:
        """Initialize reader with target spreadsheet.

        Args:
            spreadsheet: Spreadsheet to load data into
        """
        self.spreadsheet = spreadsheet

    def load(self, filepath: str) -> None:
        """Load WK1 file into spreadsheet.

        Args:
            filepath: Path to WK1 file

        Raises:
            ValueError: If file format is invalid
            FileNotFoundError: If file doesn't exist
        """
        self.spreadsheet.clear()

        with open(filepath, "rb") as f:
            self._read_file(f)

        self.spreadsheet.filename = filepath
        self.spreadsheet.modified = False
        self.spreadsheet.rebuild_dependency_graph()

    def _read_file(self, f: BinaryIO) -> None:
        """Read records from WK1 file.

        Args:
            f: Binary file handle
        """
        # Read BOF record
        opcode, length = self._read_record_header(f)
        if opcode != BOF:
            raise ValueError("Invalid WK1 file: missing BOF record")

        data = f.read(length)
        if len(data) < 2:
            raise ValueError("Invalid WK1 file: truncated BOF record")

        version = struct.unpack("<H", data[:2])[0]
        if version not in (VERSION_WK1, VERSION_WKS):
            raise ValueError(f"Unsupported WK1 version: {version:#06x}")

        # Read records until EOF
        while True:
            opcode, length = self._read_record_header(f)
            if opcode == EOF:
                break
            if opcode is None:
                break  # End of file

            data = f.read(length)
            if len(data) < length:
                break  # Truncated file

            self._process_record(opcode, data)

    def _read_record_header(self, f: BinaryIO) -> tuple[int | None, int]:
        """Read 4-byte record header.

        Args:
            f: Binary file handle

        Returns:
            Tuple of (opcode, length), or (None, 0) if EOF
        """
        header = f.read(4)
        if len(header) < 4:
            return None, 0
        opcode, length = struct.unpack("<HH", header)
        return opcode, length

    def _process_record(self, opcode: int, data: bytes) -> None:
        """Process a single record.

        Args:
            opcode: Record type
            data: Record data
        """
        if opcode == LABEL:
            self._read_label(data)
        elif opcode == INTEGER:
            self._read_integer(data)
        elif opcode == NUMBER:
            self._read_number(data)
        elif opcode == FORMULA:
            self._read_formula(data)
        elif opcode == COLW1:
            self._read_column_width(data)
        elif opcode == NAME:
            self._read_named_range(data)
        elif opcode == CALCMODE:
            self._read_calcmode(data)
        elif opcode == CALCORDER:
            self._read_calcorder(data)
        elif opcode == BLANK:
            self._read_blank(data)
        # Ignore other records (RANGE, WINDOW1, etc.)

    def _read_label(self, data: bytes) -> None:
        """Read string cell.

        Format: format(1) + col(2) + row(2) + string(null-terminated)
        """
        if len(data) < 6:
            return

        fmt_byte = data[0]
        col = struct.unpack("<H", data[1:3])[0]
        row = struct.unpack("<H", data[3:5])[0]

        # String is null-terminated, may have prefix character
        string_data = data[5:]
        null_pos = string_data.find(0)
        if null_pos != -1:
            string_data = string_data[:null_pos]

        # Decode using CP437 (DOS codepage)
        try:
            text = string_data.decode("cp437", errors="replace")
        except Exception:
            text = string_data.decode("latin-1", errors="replace")

        # Keep Lotus label prefix if present for alignment
        # ' = left, " = right, ^ = center, \ = repeat
        self.spreadsheet.set_cell(row, col, text)

        # Apply format code if not default
        format_code = decode_format_byte(fmt_byte)
        if format_code != "G":
            cell = self.spreadsheet.get_cell(row, col)
            cell.format_code = format_code

    def _read_integer(self, data: bytes) -> None:
        """Read integer cell.

        Format: format(1) + col(2) + row(2) + value(2)
        """
        if len(data) < 7:
            return

        fmt_byte = data[0]
        col = struct.unpack("<H", data[1:3])[0]
        row = struct.unpack("<H", data[3:5])[0]
        value = struct.unpack("<h", data[5:7])[0]  # Signed 16-bit

        self.spreadsheet.set_cell(row, col, str(value))

        # Apply format code if not default
        format_code = decode_format_byte(fmt_byte)
        if format_code != "G":
            cell = self.spreadsheet.get_cell(row, col)
            cell.format_code = format_code

    def _read_number(self, data: bytes) -> None:
        """Read floating-point cell.

        Format: format(1) + col(2) + row(2) + value(8)
        """
        if len(data) < 13:
            return

        fmt_byte = data[0]
        col = struct.unpack("<H", data[1:3])[0]
        row = struct.unpack("<H", data[3:5])[0]
        value = struct.unpack("<d", data[5:13])[0]  # IEEE 754 double

        # Format nicely - avoid unnecessary decimals
        # Handle IEEE 754 special values (infinity, NaN) that can't be converted to int
        if math.isfinite(value) and value == int(value):
            self.spreadsheet.set_cell(row, col, str(int(value)))
        else:
            self.spreadsheet.set_cell(row, col, str(value))

        # Apply format code if not default
        format_code = decode_format_byte(fmt_byte)
        if format_code != "G":
            cell = self.spreadsheet.get_cell(row, col)
            cell.format_code = format_code

    def _read_formula(self, data: bytes) -> None:
        """Read formula cell.

        Format: format(1) + col(2) + row(2) + value(8) + length(2) + formula_bytes
        """
        if len(data) < 15:
            return

        fmt_byte = data[0]
        col = struct.unpack("<H", data[1:3])[0]
        row = struct.unpack("<H", data[3:5])[0]
        cached_value = struct.unpack("<d", data[5:13])[0]  # Cached result
        formula_len = struct.unpack("<H", data[13:15])[0]

        if len(data) >= 15 + formula_len:
            bytecode = data[15:15 + formula_len]
            try:
                # Pass the formula's cell location to resolve relative references
                formula_text = decompile_formula(bytecode, row, col)
                self.spreadsheet.set_cell(row, col, formula_text)
            except Exception:
                # On decompilation failure, use cached value
                if math.isfinite(cached_value) and cached_value == int(cached_value):
                    self.spreadsheet.set_cell(row, col, str(int(cached_value)))
                else:
                    self.spreadsheet.set_cell(row, col, str(cached_value))
        else:
            # Truncated formula, use cached value
            if math.isfinite(cached_value) and cached_value == int(cached_value):
                self.spreadsheet.set_cell(row, col, str(int(cached_value)))
            else:
                self.spreadsheet.set_cell(row, col, str(cached_value))

        # Apply format code if not default
        format_code = decode_format_byte(fmt_byte)
        if format_code != "G":
            cell = self.spreadsheet.get_cell(row, col)
            cell.format_code = format_code

    def _read_column_width(self, data: bytes) -> None:
        """Read column width.

        Format: col(2) + width(1)
        """
        if len(data) < 3:
            return

        col = struct.unpack("<H", data[0:2])[0]
        width = data[2]

        # Default width in WK1 is 9
        if width != 9:
            self.spreadsheet.set_col_width(col, width)

    def _read_named_range(self, data: bytes) -> None:
        """Read named range.

        Format: name(16 bytes, null-padded) + start_col(2) + start_row(2) + end_col(2) + end_row(2)
        """
        if len(data) < 24:
            return

        # Extract null-terminated name (up to 16 bytes)
        name_bytes = data[:16]
        null_pos = name_bytes.find(0)
        if null_pos != -1:
            name_bytes = name_bytes[:null_pos]

        try:
            name = name_bytes.decode("cp437", errors="replace").strip()
        except Exception:
            name = name_bytes.decode("latin-1", errors="replace").strip()

        if not name:
            return

        # Read range coordinates
        start_col = struct.unpack("<H", data[16:18])[0]
        start_row = struct.unpack("<H", data[18:20])[0]
        end_col = struct.unpack("<H", data[20:22])[0]
        end_row = struct.unpack("<H", data[22:24])[0]

        # Build reference string
        if start_col == end_col and start_row == end_row:
            ref_str = f"{_col_to_letter(start_col)}{start_row + 1}"
        else:
            ref_str = (
                f"{_col_to_letter(start_col)}{start_row + 1}:"
                f"{_col_to_letter(end_col)}{end_row + 1}"
            )

        try:
            self.spreadsheet.named_ranges.add_from_string(name, ref_str)
        except ValueError:
            pass  # Invalid name, skip

    def _read_calcmode(self, data: bytes) -> None:
        """Read calculation mode.

        Format: mode(1) - 0=manual, 0xFF=automatic
        """
        if len(data) < 1:
            return

        from ..formula.recalc import RecalcMode

        mode_byte = data[0]
        if self.spreadsheet._recalc_engine:
            if mode_byte == 0:
                self.spreadsheet._recalc_engine.mode = RecalcMode.MANUAL
            else:
                self.spreadsheet._recalc_engine.mode = RecalcMode.AUTOMATIC

    def _read_calcorder(self, data: bytes) -> None:
        """Read calculation order.

        Format: order(1) - 0=natural, 1=column-wise, 2=row-wise
        """
        if len(data) < 1:
            return

        from ..formula.recalc import RecalcOrder

        order_byte = data[0]
        if self.spreadsheet._recalc_engine:
            if order_byte == 0:
                self.spreadsheet._recalc_engine.order = RecalcOrder.NATURAL
            elif order_byte == 1:
                self.spreadsheet._recalc_engine.order = RecalcOrder.COLUMN_WISE
            elif order_byte == 2:
                self.spreadsheet._recalc_engine.order = RecalcOrder.ROW_WISE

    def _read_blank(self, data: bytes) -> None:
        """Read blank cell (format only, no value).

        Format: format(1) + col(2) + row(2)
        """
        if len(data) < 5:
            return

        fmt_byte = data[0]
        col = struct.unpack("<H", data[1:3])[0]
        row = struct.unpack("<H", data[3:5])[0]

        # Decode format and apply if not default
        format_code = decode_format_byte(fmt_byte)
        if format_code != "G":
            # Create or get cell and set format
            cell = self.spreadsheet.get_cell(row, col)
            cell.format_code = format_code


class Wk1Writer:
    """Write Lotus WK1 files."""

    def __init__(self, spreadsheet: Spreadsheet) -> None:
        """Initialize writer with source spreadsheet.

        Args:
            spreadsheet: Spreadsheet to save
        """
        self.spreadsheet = spreadsheet

    def save(self, filepath: str) -> None:
        """Save spreadsheet to WK1 format.

        Args:
            filepath: Path to save file

        Note:
            - Formulas are compiled to WK1 bytecode
            - Only basic cell types are supported
            - Maximum dimensions: 256 cols x 8192 rows
        """
        with open(filepath, "wb") as f:
            self._write_file(f)

    def _write_file(self, f: BinaryIO) -> None:
        """Write WK1 file.

        Args:
            f: Binary file handle
        """
        # Write BOF (Beginning of File)
        self._write_record(f, BOF, struct.pack("<H", VERSION_WK1))

        # Write calculation mode and order
        self._write_calcmode(f)
        self._write_calcorder(f)

        # Write RANGE (used area)
        used = self.spreadsheet.get_used_range()
        if used:
            (min_row, min_col), (max_row, max_col) = used
            # WK1 RANGE format: start_col, start_row, end_col, end_row (all 2 bytes)
            range_data = struct.pack("<HHHH", min_col, min_row, max_col, max_row)
            self._write_record(f, RANGE, range_data)

        # Write column widths
        for col, width in sorted(self.spreadsheet._col_widths.items()):
            colw_data = struct.pack("<HB", col, width)
            self._write_record(f, COLW1, colw_data)

        # Write named ranges
        self._write_named_ranges(f)

        # Write cells
        for (row, col), cell in sorted(self.spreadsheet._cells.items()):
            if cell.is_empty:
                # Write blank cell if it has a non-default format
                if cell.format_code != "G":
                    self._write_blank(f, row, col, cell.format_code)
                continue

            raw_value = cell.raw_value
            format_code = cell.format_code

            # Check if it's a formula
            if raw_value.startswith("=") or raw_value.startswith("@"):
                self._write_formula(f, row, col, raw_value, format_code)
            else:
                # Get calculated value for non-formulas
                value = self.spreadsheet.get_value(row, col)

                if value is None or value == "":
                    continue
                elif isinstance(value, str):
                    # Check for IEEE 754 special float strings
                    value_lower = value.lower()
                    if value_lower == "inf":
                        self._write_number(f, row, col, float("inf"), format_code)
                    elif value_lower == "-inf":
                        self._write_number(f, row, col, float("-inf"), format_code)
                    elif value_lower == "nan":
                        self._write_number(f, row, col, float("nan"), format_code)
                    else:
                        self._write_label(f, row, col, value, format_code)
                elif isinstance(value, bool):
                    # Booleans as integers
                    self._write_integer(f, row, col, 1 if value else 0, format_code)
                elif isinstance(value, int) and -32768 <= value <= 32767:
                    self._write_integer(f, row, col, value, format_code)
                elif isinstance(value, (int, float)):
                    self._write_number(f, row, col, float(value), format_code)
                else:
                    # Fallback: convert to string
                    self._write_label(f, row, col, str(value), format_code)

        # Write EOF (End of File)
        self._write_record(f, EOF, b"")

    def _write_record(self, f: BinaryIO, opcode: int, data: bytes) -> None:
        """Write a record.

        Args:
            f: Binary file handle
            opcode: Record type
            data: Record data
        """
        header = struct.pack("<HH", opcode, len(data))
        f.write(header)
        f.write(data)

    def _write_calcmode(self, f: BinaryIO) -> None:
        """Write calculation mode record.

        Args:
            f: Binary file handle
        """
        from ..formula.recalc import RecalcMode

        mode_byte = 0xFF  # Default: automatic
        if self.spreadsheet._recalc_engine:
            if self.spreadsheet._recalc_engine.mode == RecalcMode.MANUAL:
                mode_byte = 0x00

        self._write_record(f, CALCMODE, struct.pack("<B", mode_byte))

    def _write_calcorder(self, f: BinaryIO) -> None:
        """Write calculation order record.

        Args:
            f: Binary file handle
        """
        from ..formula.recalc import RecalcOrder

        order_byte = 0x00  # Default: natural
        if self.spreadsheet._recalc_engine:
            order = self.spreadsheet._recalc_engine.order
            if order == RecalcOrder.COLUMN_WISE:
                order_byte = 0x01
            elif order == RecalcOrder.ROW_WISE:
                order_byte = 0x02

        self._write_record(f, CALCORDER, struct.pack("<B", order_byte))

    def _write_named_ranges(self, f: BinaryIO) -> None:
        """Write all named range records.

        Args:
            f: Binary file handle
        """
        from ..core.reference import CellReference, RangeReference

        for named in self.spreadsheet.named_ranges.list_all():
            # Encode name (max 15 chars + null terminator = 16 bytes)
            name = named.name[:15]
            try:
                name_bytes = name.encode("cp437", errors="replace")
            except Exception:
                name_bytes = name.encode("latin-1", errors="replace")
            # Pad to 16 bytes with null
            name_bytes = name_bytes.ljust(16, b"\x00")

            # Get range coordinates
            ref = named.reference
            if isinstance(ref, CellReference):
                start_col = end_col = ref.col
                start_row = end_row = ref.row
            elif isinstance(ref, RangeReference):
                start_col = ref.start.col
                start_row = ref.start.row
                end_col = ref.end.col
                end_row = ref.end.row
            else:
                continue

            # Format: name(16) + start_col(2) + start_row(2) + end_col(2) + end_row(2)
            data = name_bytes + struct.pack("<HHHH", start_col, start_row, end_col, end_row)
            self._write_record(f, NAME, data)

    def _write_blank(
        self, f: BinaryIO, row: int, col: int, format_code: str = "G"
    ) -> None:
        """Write blank cell (format only, no value).

        Args:
            f: Binary file handle
            row: Row index (0-based)
            col: Column index (0-based)
            format_code: Cell format code
        """
        fmt_byte = encode_format_byte(format_code)
        data = struct.pack("<BHH", fmt_byte, col, row)
        self._write_record(f, BLANK, data)

    def _write_label(
        self, f: BinaryIO, row: int, col: int, text: str, format_code: str = "G"
    ) -> None:
        """Write string cell.

        Args:
            f: Binary file handle
            row: Row index (0-based)
            col: Column index (0-based)
            text: Cell text
            format_code: Cell format code
        """
        fmt_byte = encode_format_byte(format_code)

        # Check if text already has alignment prefix
        if text and text[0] in "'\"^\\":
            prefix = ""
        else:
            prefix = "'"  # Default left-aligned

        # Encode to CP437 with null terminator
        try:
            encoded = (prefix + text).encode("cp437", errors="replace") + b"\x00"
        except Exception:
            encoded = (prefix + text).encode("latin-1", errors="replace") + b"\x00"

        # Format: format_byte(1) + col(2) + row(2) + string
        data = struct.pack("<BHH", fmt_byte, col, row) + encoded
        self._write_record(f, LABEL, data)

    def _write_integer(
        self, f: BinaryIO, row: int, col: int, value: int, format_code: str = "G"
    ) -> None:
        """Write integer cell.

        Args:
            f: Binary file handle
            row: Row index (0-based)
            col: Column index (0-based)
            value: Integer value (-32768 to 32767)
            format_code: Cell format code
        """
        fmt_byte = encode_format_byte(format_code)
        # Format: format_byte(1) + col(2) + row(2) + value(2)
        data = struct.pack("<BHHh", fmt_byte, col, row, value)
        self._write_record(f, INTEGER, data)

    def _write_number(
        self, f: BinaryIO, row: int, col: int, value: float, format_code: str = "G"
    ) -> None:
        """Write floating-point cell.

        Args:
            f: Binary file handle
            row: Row index (0-based)
            col: Column index (0-based)
            value: Float value
            format_code: Cell format code
        """
        fmt_byte = encode_format_byte(format_code)
        # Format: format_byte(1) + col(2) + row(2) + value(8)
        data = struct.pack("<BHHd", fmt_byte, col, row, value)
        self._write_record(f, NUMBER, data)

    def _write_formula(
        self, f: BinaryIO, row: int, col: int, formula: str, format_code: str = "G"
    ) -> None:
        """Write formula cell.

        Args:
            f: Binary file handle
            row: Row index (0-based)
            col: Column index (0-based)
            formula: Formula text (e.g., "=A1+B1")
            format_code: Cell format code
        """
        fmt_byte = encode_format_byte(format_code)

        # Get cached value
        cached_value = self.spreadsheet.get_value(row, col)
        if cached_value is None:
            cached_value = 0.0
        elif isinstance(cached_value, bool):
            cached_value = 1.0 if cached_value else 0.0
        elif isinstance(cached_value, str):
            # Strings in formulas become 0
            cached_value = 0.0
        else:
            cached_value = float(cached_value)

        # Compile formula to bytecode (pass cell location for relative refs)
        bytecode = compile_formula(formula, row, col)

        # Format: format_byte(1) + col(2) + row(2) + value(8) + length(2) + bytecode
        data = struct.pack("<BHHdH", fmt_byte, col, row, cached_value, len(bytecode))
        data += bytecode
        self._write_record(f, FORMULA, data)

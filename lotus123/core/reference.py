"""Cell and range reference handling with absolute/relative reference support."""

import re
from dataclasses import dataclass
from typing import Iterator, override

from .errors import FormulaError

# Pattern for cell references: optional $ before column and/or row
CELL_REF_PATTERN = re.compile(r"^(\$?)([A-Za-z]+)(\$?)(\d+)$")

# Pattern for range references: cell:cell
RANGE_PATTERN = re.compile(r"^(\$?[A-Za-z]+\$?\d+):(\$?[A-Za-z]+\$?\d+)$")


def col_to_index(col: str) -> int:
    """Convert column letter(s) to 0-based index.

    Examples:
        A -> 0, Z -> 25, AA -> 26, IV -> 255
    """
    result = 0
    for char in col.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result - 1


def index_to_col(index: int) -> str:
    """Convert 0-based index to column letter(s).

    Examples:
        0 -> A, 25 -> Z, 26 -> AA, 255 -> IV
    """
    result = ""
    index += 1
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def parse_cell_ref(ref: str) -> tuple[int, int]:
    """Parse cell reference to (row, col) 0-based indices.

    Handles both regular (A1) and absolute ($A$1, $A1, A$1) references.

    Args:
        ref: Cell reference string like 'A1' or '$A$1'

    Returns:
        Tuple of (row_index, col_index) as 0-based integers

    Raises:
        ValueError: If reference format is invalid
    """
    match = CELL_REF_PATTERN.match(ref.strip())
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    _, col_str, _, row_str = match.groups()
    return int(row_str) - 1, col_to_index(col_str)


def parse_range_ref(range_ref: str) -> tuple[tuple[int, int], tuple[int, int]]:
    """Parse range reference to start and end coordinates.

    Args:
        range_ref: Range string like 'A1:B10' or '$A$1:$B$10'

    Returns:
        Tuple of ((start_row, start_col), (end_row, end_col))

    Raises:
        ValueError: If range format is invalid
    """
    match = RANGE_PATTERN.match(range_ref.strip())
    if not match:
        raise ValueError(f"Invalid range reference: {range_ref}")
    start_ref, end_ref = match.groups()
    start = parse_cell_ref(start_ref)
    end = parse_cell_ref(end_ref)
    return start, end


def make_cell_ref(
    row: int, col: int, col_absolute: bool = False, row_absolute: bool = False
) -> str:
    """Convert (row, col) 0-based indices to cell reference.

    Args:
        row: 0-based row index
        col: 0-based column index
        col_absolute: Whether to prefix column with $
        row_absolute: Whether to prefix row with $

    Returns:
        Cell reference string like 'A1' or '$A$1'
    """
    col_prefix = "$" if col_absolute else ""
    row_prefix = "$" if row_absolute else ""
    return f"{col_prefix}{index_to_col(col)}{row_prefix}{row + 1}"


@dataclass
class CellReference:
    """Represents a cell reference with absolute/relative indicators.

    Attributes:
        row: 0-based row index
        col: 0-based column index
        col_absolute: Whether column is absolute ($A vs A)
        row_absolute: Whether row is absolute ($1 vs 1)
    """

    row: int
    col: int
    col_absolute: bool = False
    row_absolute: bool = False

    @classmethod
    def parse(cls, ref: str) -> CellReference:
        """Parse a cell reference string.

        Args:
            ref: Reference string like 'A1', '$A1', 'A$1', or '$A$1'

        Returns:
            CellReference instance
        """
        match = CELL_REF_PATTERN.match(ref.strip())
        if not match:
            raise ValueError(f"Invalid cell reference: {ref}")
        col_abs, col_str, row_abs, row_str = match.groups()
        return cls(
            row=int(row_str) - 1,
            col=col_to_index(col_str),
            col_absolute=bool(col_abs),
            row_absolute=bool(row_abs),
        )

    def to_string(self) -> str:
        """Convert to reference string."""
        return make_cell_ref(self.row, self.col, self.col_absolute, self.row_absolute)

    def adjust(
        self, row_delta: int, col_delta: int, max_row: int = 8191, max_col: int = 255
    ) -> CellReference:
        """Create adjusted reference for copy/paste operations.

        Relative references are adjusted by the deltas.
        Absolute references remain unchanged.

        Args:
            row_delta: Rows to offset (for relative refs)
            col_delta: Columns to offset (for relative refs)
            max_row: Maximum valid row index
            max_col: Maximum valid column index

        Returns:
            New CellReference with adjusted coordinates
        """
        new_row = self.row if self.row_absolute else min(max(0, self.row + row_delta), max_row)
        new_col = self.col if self.col_absolute else min(max(0, self.col + col_delta), max_col)
        return CellReference(new_row, new_col, self.col_absolute, self.row_absolute)

    @override
    def __str__(self) -> str:
        return self.to_string()

    @override
    def __hash__(self) -> int:
        return hash((self.row, self.col))

    @override
    def __eq__(self, other: object) -> bool:
        if isinstance(other, CellReference):
            return self.row == other.row and self.col == other.col
        return False


@dataclass
class RangeReference:
    """Represents a range of cells.

    Attributes:
        start: Starting cell reference (top-left)
        end: Ending cell reference (bottom-right)
    """

    start: CellReference
    end: CellReference

    @classmethod
    def parse(cls, range_ref: str) -> RangeReference:
        """Parse a range reference string.

        Args:
            range_ref: Range string like 'A1:B10' or '$A$1:$B$10'

        Returns:
            RangeReference instance
        """
        match = RANGE_PATTERN.match(range_ref.strip())
        if not match:
            raise ValueError(f"Invalid range reference: {range_ref}")
        start_ref, end_ref = match.groups()
        return cls(start=CellReference.parse(start_ref), end=CellReference.parse(end_ref))

    def to_string(self) -> str:
        """Convert to range string."""
        return f"{self.start.to_string()}:{self.end.to_string()}"

    @property
    def normalized(self) -> RangeReference:
        """Return range with start as top-left and end as bottom-right."""
        min_row = min(self.start.row, self.end.row)
        max_row = max(self.start.row, self.end.row)
        min_col = min(self.start.col, self.end.col)
        max_col = max(self.start.col, self.end.col)
        return RangeReference(
            CellReference(min_row, min_col, self.start.col_absolute, self.start.row_absolute),
            CellReference(max_row, max_col, self.end.col_absolute, self.end.row_absolute),
        )

    @property
    def row_count(self) -> int:
        """Number of rows in range."""
        norm = self.normalized
        return norm.end.row - norm.start.row + 1

    @property
    def col_count(self) -> int:
        """Number of columns in range."""
        norm = self.normalized
        return norm.end.col - norm.start.col + 1

    def iter_cells(self) -> Iterator[tuple[int, int]]:
        """Iterate over all cell coordinates in the range.

        Yields:
            Tuples of (row, col) for each cell in the range
        """
        norm = self.normalized
        for row in range(norm.start.row, norm.end.row + 1):
            for col in range(norm.start.col, norm.end.col + 1):
                yield row, col

    def adjust(
        self, row_delta: int, col_delta: int, max_row: int = 8191, max_col: int = 255
    ) -> RangeReference:
        """Create adjusted range reference for copy/paste operations."""
        return RangeReference(
            self.start.adjust(row_delta, col_delta, max_row, max_col),
            self.end.adjust(row_delta, col_delta, max_row, max_col),
        )

    def contains(self, row: int, col: int) -> bool:
        """Check if a cell is within this range."""
        norm = self.normalized
        return norm.start.row <= row <= norm.end.row and norm.start.col <= col <= norm.end.col

    @override
    def __str__(self) -> str:
        return self.to_string()


def adjust_formula_references(
    formula: str, row_delta: int, col_delta: int, max_row: int = 8191, max_col: int = 255
) -> str:
    """Adjust cell references in a formula for copy/paste.

    Relative references are adjusted, absolute references ($) are kept.
    Strings and named ranges are preserved.

    Args:
        formula: Formula string (without leading =)
        row_delta: Rows to offset
        col_delta: Columns to offset
        max_row: Maximum valid row
        max_col: Maximum valid column

    Returns:
        Formula with adjusted references
    """
    from ..formula.tokenizer import Tokenizer, TokenType

    tokenizer = Tokenizer()
    tokens = tokenizer.tokenize(formula)

    result = []
    last_pos = 0

    for token in tokens:
        # Append skipped content (whitespace, etc.)
        if token.position > last_pos:
            result.append(formula[last_pos : token.position])

        # Determine replacement text
        text = token.raw_text

        if token.type == TokenType.CELL:
            try:
                ref = CellReference.parse(token.raw_text)
                text = ref.adjust(row_delta, col_delta, max_row, max_col).to_string()
            except ValueError:
                # Not a valid cell ref (e.g. named range or function looking like cell)
                pass

        elif token.type == TokenType.RANGE:
            # Tokenizer might classify named ranges as RANGE if we passed a spreadsheet,
            # but here we pass None. So this usually won't trigger unless we handle ranges explicitly.
            # However, our Tokenizer emits RANGE only if named range is resolved.
            # Standard A1:B2 is parsed as CELL, COLON, CELL by our tokenizer?
            # Let's check Tokenizer.tokenize loop.
            # No, Tokenizer splits on special chars. A1:B2 -> A1 (CELL), : (COLON), B2 (CELL).
            # So range adjustment happens automatically via cell adjustment!
            pass

        elif token.type == TokenType.EOF:
            text = ""

        result.append(text)
        last_pos = token.position + len(token.raw_text)

    # Append any trailing content
    if last_pos < len(formula):
        result.append(formula[last_pos:])

    return "".join(result)


def adjust_for_structural_change(
    formula: str,
    axis: str,
    boundary: int,
    shift: int,
    max_row: int = 8191,
    max_col: int = 255,
) -> str:
    """Adjust references based on row/column insertion/deletion.

    Args:
        formula: Formula string
        axis: 'row' or 'col'
        boundary: Index where shift occurs
        shift: Amount to shift (+1 or -1)
        max_row: Max row index
        max_col: Max col index

    Returns:
        Adjusted formula string
    """
    from ..formula.tokenizer import Tokenizer, TokenType

    tokenizer = Tokenizer()
    tokens = tokenizer.tokenize(formula)

    result = []
    last_pos = 0

    for token in tokens:
        # Append skipped content
        if token.position > last_pos:
            result.append(formula[last_pos : token.position])

        text = token.raw_text

        if token.type == TokenType.CELL:
            try:
                ref = CellReference.parse(token.raw_text)

                if axis == "row":
                    if shift < 0 and ref.row == boundary:
                        text = FormulaError.REF
                    elif ref.row >= boundary:
                        new_row = ref.row + shift
                        if 0 <= new_row <= max_row:
                            ref.row = new_row
                            text = ref.to_string()
                        else:
                            text = FormulaError.REF
                elif axis == "col":
                    if shift < 0 and ref.col == boundary:
                        text = FormulaError.REF
                    elif ref.col >= boundary:
                        new_col = ref.col + shift
                        if 0 <= new_col <= max_col:
                            ref.col = new_col
                            text = ref.to_string()
                        else:
                            text = FormulaError.REF

            except ValueError:
                pass

        elif token.type == TokenType.EOF:
            text = ""

        result.append(text)
        last_pos = token.position + len(token.raw_text)

    if last_pos < len(formula):
        result.append(formula[last_pos:])

    return "".join(result)

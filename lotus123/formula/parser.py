"""Formula parser with tokenization and expression building."""

import math
import operator
from typing import Any

from ..core.errors import FormulaError
from ..core.spreadsheet_protocol import SpreadsheetProtocol
from .context import EvaluationContext
from .functions import REGISTRY
from .tokenizer import Token, Tokenizer, TokenType


class FormulaParser:
    """Parser and evaluator for spreadsheet formulas.

    Supports:
    - Arithmetic: +, -, *, /, ^, %
    - Comparison: =, <>, <, >, <=, >=
    - Cell references: A1, $A$1, $A1, A$1
    - Ranges: A1:B10
    - Functions: @SUM(), SUM(), etc.
    - Named ranges
    """

    OPERATORS = {
        "+": (1, operator.add),
        "-": (1, operator.sub),
        "*": (2, operator.mul),
        "/": (2, operator.truediv),
        "^": (3, operator.pow),
        "%": (2, operator.mod),
    }

    COMPARISONS = {
        "=": operator.eq,
        "==": operator.eq,
        "<>": operator.ne,
        "!=": operator.ne,
        "<": operator.lt,
        ">": operator.gt,
        "<=": operator.le,
        ">=": operator.ge,
    }

    def __init__(
        self, spreadsheet: SpreadsheetProtocol, context: EvaluationContext | None = None
    ) -> None:
        self.spreadsheet = spreadsheet
        # Use singleton registry to avoid expensive re-initialization
        self.functions = REGISTRY
        self.tokenizer = Tokenizer(spreadsheet)
        self._tokens: list[Token] = []
        self._pos: int = 0
        self.context = context

    def evaluate(self, formula: str) -> Any:
        """Evaluate a formula string and return the result.

        Args:
            formula: Formula string (without leading =)

        Returns:
            Computed value or error string
        """
        formula = formula.strip()
        if not formula:
            return ""

        try:
            self._tokens = self.tokenizer.tokenize(formula)
            self._pos = 0
            result = self._parse_expression()

            # Handle NaN as error
            if isinstance(result, float) and math.isnan(result):
                return FormulaError.NUM

            return result
        except ZeroDivisionError:
            return FormulaError.DIV_ZERO
        except RecursionError:
            return FormulaError.REF
        except (ValueError, TypeError, KeyError, IndexError, AttributeError, OverflowError):
            return FormulaError.ERR

    def _current(self) -> Token:
        """Get current token."""
        if self._pos < len(self._tokens):
            return self._tokens[self._pos]
        return Token(TokenType.EOF, None)

    def _advance(self) -> Token:
        """Move to next token and return current."""
        token = self._current()
        self._pos += 1
        return token

    def _is_error(self, value: Any) -> bool:
        """Check if a value is a known error string."""
        return FormulaError.is_error(value)

    def _parse_expression(self) -> Any:
        """Parse a full expression."""
        return self._parse_comparison()

    def _parse_comparison(self) -> Any:
        """Parse comparison operators."""
        left = self._parse_arithmetic()

        # Propagate errors immediately
        if self._is_error(left):
            return left

        while self._current().type == TokenType.COMPARISON:
            op_str = self._advance().value
            right = self._parse_arithmetic()

            # Propagate errors from right operand
            if self._is_error(right):
                return right

            cmp_fn = self.COMPARISONS.get(op_str)
            if cmp_fn:
                left = cmp_fn(left, right)

        return left

    def _parse_arithmetic(self, min_prec: int = 0) -> Any:
        """Parse arithmetic with operator precedence."""
        left = self._parse_atom()

        # Propagate errors immediately
        if self._is_error(left):
            return left

        while True:
            token = self._current()
            if token.type != TokenType.OPERATOR:
                break

            op_str = token.value
            if op_str not in self.OPERATORS:
                break

            prec, op_fn = self.OPERATORS[op_str]
            if prec < min_prec:
                break

            self._advance()
            right = self._parse_arithmetic(prec + 1)

            # Propagate errors from right operand
            if self._is_error(right):
                return right

            try:
                left = op_fn(left, right)
            except ZeroDivisionError:
                left = FormulaError.DIV_ZERO
            except (ValueError, TypeError, OverflowError):
                left = FormulaError.ERR

            # Propagate errors from operation result
            if self._is_error(left):
                return left

        return left

    def _parse_atom(self) -> Any:
        """Parse atomic value."""
        token = self._current()

        # Unary minus
        if token.type == TokenType.OPERATOR and token.value == "-":
            self._advance()
            val = self._parse_atom()
            if self._is_error(val):
                return val
            return -val if isinstance(val, (int, float)) else val

        # Unary plus
        if token.type == TokenType.OPERATOR and token.value == "+":
            self._advance()
            val = self._parse_atom()
            if self._is_error(val):
                return val
            return val

        # Number
        if token.type == TokenType.NUMBER:
            self._advance()
            return token.value

        # String
        if token.type == TokenType.STRING:
            self._advance()
            return token.value

        # Cell reference (possibly part of range)
        if token.type == TokenType.CELL:
            self._advance()
            cell_ref = token.value

            # Check for range
            if self._current().type == TokenType.COLON:
                self._advance()
                if self._current().type == TokenType.CELL:
                    end_ref = self._advance().value
                    return self._get_range_values(cell_ref, end_ref)

            return self._get_cell_value(cell_ref)

        # Range (already parsed as single token)
        if token.type == TokenType.RANGE:
            self._advance()
            parts = token.value.split(":")
            if len(parts) == 2:
                return self._get_range_values(parts[0], parts[1])
            return FormulaError.REF

        # Function call
        if token.type == TokenType.FUNCTION:
            self._advance()
            return self._parse_function(token.value)

        # Parenthesized expression
        if token.type == TokenType.LPAREN:
            self._advance()
            result = self._parse_expression()
            if self._current().type == TokenType.RPAREN:
                self._advance()
            return result

        # EOF - empty expression
        if token.type == TokenType.EOF:
            return ""

        # Unknown/unexpected token - malformed formula
        self._advance()
        return FormulaError.ERR

    def _parse_function(self, name: str) -> Any:
        """Parse and evaluate a function call."""
        args = []

        # Consume opening paren
        if self._current().type == TokenType.LPAREN:
            self._advance()

        # Parse arguments
        while True:
            if self._current().type == TokenType.RPAREN:
                self._advance()
                break
            if self._current().type == TokenType.EOF:
                break
            if self._current().type == TokenType.COMMA:
                self._advance()
                continue

            arg = self._parse_expression()
            args.append(arg)

        # Look up and call function
        fn = self.functions.get(name)
        if not fn:
            return FormulaError.NAME

        try:
            return fn(*args)
        except (ValueError, TypeError, ZeroDivisionError, OverflowError, IndexError):
            return FormulaError.ERR

    def _get_cell_value(self, ref: str) -> Any:
        """Get value of a cell reference."""
        try:
            # Strip $ signs for lookup
            clean_ref = ref.replace("$", "")
            return self.spreadsheet.get_value_by_ref(clean_ref, context=self.context)
        except ValueError:
            return FormulaError.REF

    def _get_range_values(self, start_ref: str, end_ref: str) -> list[Any]:
        """Get 2D list of values from a range.

        Returns a list of lists (rows) to preserve range shape information
        needed by functions like VLOOKUP, HLOOKUP, INDEX, ROWS, COLS.
        Functions that need flat values flatten internally.
        """
        try:
            start = start_ref.replace("$", "")
            end = end_ref.replace("$", "")
            return self.spreadsheet.get_range(start, end, context=self.context)
        except ValueError:
            return [[FormulaError.REF]]

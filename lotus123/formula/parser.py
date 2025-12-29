"""Formula parser with tokenization and expression building."""

from __future__ import annotations

import operator
import re
from dataclasses import dataclass
from enum import Enum, auto
from typing import TYPE_CHECKING, Any

from .functions import FunctionRegistry

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


class TokenType(Enum):
    """Types of tokens in a formula."""

    NUMBER = auto()
    STRING = auto()
    CELL = auto()
    RANGE = auto()
    FUNCTION = auto()
    OPERATOR = auto()
    COMPARISON = auto()
    LPAREN = auto()
    RPAREN = auto()
    COMMA = auto()
    COLON = auto()
    EOF = auto()


@dataclass
class Token:
    """A single token from the formula."""

    type: TokenType
    value: Any
    position: int = 0


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

    # Regex patterns
    CELL_PATTERN = re.compile(r"\$?[A-Za-z]+\$?\d+")
    NUMBER_PATTERN = re.compile(r"\d+\.?\d*([eE][+-]?\d+)?")

    def __init__(self, spreadsheet: Spreadsheet) -> None:
        self.spreadsheet = spreadsheet
        self.functions = FunctionRegistry()
        self._tokens: list[Token] = []
        self._pos: int = 0

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
            self._tokens = self._tokenize(formula)
            self._pos = 0
            result = self._parse_expression()

            # Handle NaN as error
            if isinstance(result, float) and (result != result):  # NaN check
                return "#NUM!"

            return result
        except ZeroDivisionError:
            return "#DIV/0!"
        except Exception:
            return "#ERR!"

    def _tokenize(self, formula: str) -> list[Token]:
        """Convert formula string to tokens."""
        tokens = []
        i = 0

        while i < len(formula):
            ch = formula[i]

            # Skip whitespace
            if ch.isspace():
                i += 1
                continue

            # String literal
            if ch == '"':
                j = i + 1
                while j < len(formula) and formula[j] != '"':
                    if formula[j] == "\\" and j + 1 < len(formula):
                        j += 2
                    else:
                        j += 1
                value = formula[i + 1 : j].replace('\\"', '"')
                tokens.append(Token(TokenType.STRING, value, i))
                i = j + 1
                continue

            # Multi-char comparison operators
            if i + 1 < len(formula):
                two = formula[i : i + 2]
                if two in ("<>", "<=", ">=", "!=", "=="):
                    tokens.append(Token(TokenType.COMPARISON, two, i))
                    i += 2
                    continue

            # Single char comparison/equals
            if ch in "<>=":
                tokens.append(Token(TokenType.COMPARISON, ch, i))
                i += 1
                continue

            # Operators
            if ch in "+-*/^%":
                tokens.append(Token(TokenType.OPERATOR, ch, i))
                i += 1
                continue

            # Parentheses and punctuation
            if ch == "(":
                tokens.append(Token(TokenType.LPAREN, ch, i))
                i += 1
                continue
            if ch == ")":
                tokens.append(Token(TokenType.RPAREN, ch, i))
                i += 1
                continue
            if ch == ",":
                tokens.append(Token(TokenType.COMMA, ch, i))
                i += 1
                continue
            if ch == ":":
                tokens.append(Token(TokenType.COLON, ch, i))
                i += 1
                continue
            # Lotus-style range separator (..)
            if ch == "." and i + 1 < len(formula) and formula[i + 1] == ".":
                tokens.append(Token(TokenType.COLON, ":", i))
                i += 2
                continue

            # @ prefix for Lotus-style functions
            if ch == "@":
                i += 1
                continue

            # Number
            if ch.isdigit() or (ch == "." and i + 1 < len(formula) and formula[i + 1].isdigit()):
                match = self.NUMBER_PATTERN.match(formula, i)
                if match:
                    num_str = match.group(0)
                    num_value: int | float = (
                        float(num_str) if "." in num_str or "e" in num_str.lower() else int(num_str)
                    )
                    tokens.append(Token(TokenType.NUMBER, num_value, i))
                    i = match.end()
                    continue

            # Identifier (function name or cell reference)
            if ch.isalpha() or ch == "_" or ch == "$":
                j = i
                while j < len(formula) and (formula[j].isalnum() or formula[j] in "_$"):
                    j += 1
                name = formula[i:j]

                # Skip whitespace to check for function call
                k = j
                while k < len(formula) and formula[k].isspace():
                    k += 1

                if k < len(formula) and formula[k] == "(":
                    # It's a function
                    tokens.append(Token(TokenType.FUNCTION, name.upper(), i))
                else:
                    # Check if it's a named range
                    if self.spreadsheet.named_ranges.exists(name):
                        # Resolve named range to cell/range reference
                        named = self.spreadsheet.named_ranges.get(name)
                        if named:
                            from ..core.reference import RangeReference

                            if isinstance(named.reference, RangeReference):
                                tokens.append(
                                    Token(TokenType.RANGE, named.reference.to_string(), i)
                                )
                            else:
                                tokens.append(Token(TokenType.CELL, named.reference.to_string(), i))
                    else:
                        # Cell reference
                        tokens.append(Token(TokenType.CELL, name.upper(), i))
                i = j
                continue

            # Unknown character - skip
            i += 1

        tokens.append(Token(TokenType.EOF, None, len(formula)))
        return tokens

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

    def _parse_expression(self, min_prec: int = 0) -> Any:
        """Parse expression with operator precedence."""
        left = self._parse_comparison()

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
            right = self._parse_expression(prec + 1)

            try:
                left = op_fn(left, right)
            except ZeroDivisionError:
                left = "#DIV/0!"
            except Exception:
                left = "#ERR!"

        return left

    def _parse_comparison(self) -> Any:
        """Parse comparison operators."""
        left = self._parse_atom()

        while self._current().type == TokenType.COMPARISON:
            op_str = self._advance().value
            right = self._parse_atom()

            cmp_fn = self.COMPARISONS.get(op_str)
            if cmp_fn:
                left = cmp_fn(left, right)

        return left

    def _parse_atom(self) -> Any:
        """Parse atomic value."""
        token = self._current()

        # Unary minus
        if token.type == TokenType.OPERATOR and token.value == "-":
            self._advance()
            val = self._parse_atom()
            return -val if isinstance(val, (int, float)) else val

        # Unary plus
        if token.type == TokenType.OPERATOR and token.value == "+":
            self._advance()
            return self._parse_atom()

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
            return "#REF!"

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

        # EOF or unknown
        if token.type == TokenType.EOF:
            return ""

        self._advance()
        return ""

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
            return "#NAME?"

        try:
            # Special handling for aggregate functions that need flattened lists
            aggregate_fns = {
                "SUM",
                "AVG",
                "AVERAGE",
                "MIN",
                "MAX",
                "COUNT",
                "COUNTA",
                "AND",
                "OR",
                "STD",
                "VAR",
                "STDEV",
                "MEDIAN",
                "MODE",
            }

            if name in aggregate_fns:
                return fn(*args)
            return fn(*args)
        except Exception:
            return "#ERR!"

    def _get_cell_value(self, ref: str) -> Any:
        """Get value of a cell reference."""
        try:
            # Strip $ signs for lookup
            clean_ref = ref.replace("$", "")
            return self.spreadsheet.get_value_by_ref(clean_ref)
        except ValueError:
            return "#REF!"

    def _get_range_values(self, start_ref: str, end_ref: str) -> list[Any]:
        """Get flat list of values from a range."""
        try:
            start = start_ref.replace("$", "")
            end = end_ref.replace("$", "")
            return self.spreadsheet.get_range_flat(start, end)
        except ValueError:
            return ["#REF!"]

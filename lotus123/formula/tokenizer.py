"""Tokenizer for formula parsing and analysis."""

import re
from dataclasses import dataclass
from enum import Enum, auto
from typing import Any, TYPE_CHECKING

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
    raw_text: str = ""


class Tokenizer:
    """Tokenizer for spreadsheet formulas."""

    # Regex patterns
    NUMBER_PATTERN = re.compile(r"\d+\.?\d*([eE][+-]?\d+)?")

    def __init__(self, spreadsheet: Spreadsheet | None = None) -> None:
        self.spreadsheet = spreadsheet

    def tokenize(self, formula: str) -> list[Token]:
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
                    if formula[j] == '"' and j + 1 < len(formula):
                        j += 2
                    else:
                        j += 1
                value = formula[i + 1 : j].replace('"', '"')
                raw_text = formula[i : j + 1]
                tokens.append(Token(TokenType.STRING, value, i, raw_text))
                i = j + 1
                continue

            # Multi-char comparison operators
            if i + 1 < len(formula):
                two = formula[i : i + 2]
                if two in ("<>", "<=", ">=", "!=", "=="):
                    tokens.append(Token(TokenType.COMPARISON, two, i, two))
                    i += 2
                    continue

            # Single char comparison/equals
            if ch in "<>= ":
                tokens.append(Token(TokenType.COMPARISON, ch, i, ch))
                i += 1
                continue

            # Operators
            if ch in "+-*/^%":
                tokens.append(Token(TokenType.OPERATOR, ch, i, ch))
                i += 1
                continue

            # Parentheses and punctuation
            if ch == "(":
                tokens.append(Token(TokenType.LPAREN, ch, i, ch))
                i += 1
                continue
            if ch == ")":
                tokens.append(Token(TokenType.RPAREN, ch, i, ch))
                i += 1
                continue
            if ch == ",":
                tokens.append(Token(TokenType.COMMA, ch, i, ch))
                i += 1
                continue
            if ch == ":":
                tokens.append(Token(TokenType.COLON, ch, i, ch))
                i += 1
                continue
            # Lotus-style range separator (..)
            if ch == "." and i + 1 < len(formula) and formula[i + 1] == ".":
                tokens.append(Token(TokenType.COLON, ":", i, ".."))
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
                    tokens.append(Token(TokenType.NUMBER, num_value, i, num_str))
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
                    tokens.append(Token(TokenType.FUNCTION, name.upper(), i, name))
                else:
                    # Check if it's a named range (if spreadsheet is available)
                    is_named_range = False
                    if self.spreadsheet and self.spreadsheet.named_ranges.exists(name):
                        named = self.spreadsheet.named_ranges.get(name)
                        if named:
                            from ..core.reference import RangeReference

                            is_named_range = True
                            if isinstance(named.reference, RangeReference):
                                tokens.append(
                                    Token(TokenType.RANGE, named.reference.to_string(), i, name)
                                )
                            else:
                                tokens.append(
                                    Token(TokenType.CELL, named.reference.to_string(), i, name)
                                )

                    if not is_named_range:
                        # Cell reference
                        tokens.append(Token(TokenType.CELL, name.upper(), i, name))
                i = j
                continue

            # Unknown character - skip
            i += 1

        tokens.append(Token(TokenType.EOF, None, len(formula), ""))
        return tokens

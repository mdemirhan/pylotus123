"""Formula parser and evaluator for spreadsheet formulas."""
from __future__ import annotations
import re
import math
import operator
from typing import TYPE_CHECKING, Any, Callable

if TYPE_CHECKING:
    from .spreadsheet import Spreadsheet

CELL_REF_PATTERN = re.compile(r'\$?([A-Za-z]+)\$?(\d+)')
RANGE_PATTERN = re.compile(r'([A-Za-z]+\d+):([A-Za-z]+\d+)')
FUNC_PATTERN = re.compile(r'@?([A-Za-z_][A-Za-z0-9_]*)\s*\(')


class FormulaParser:
    """Parser and evaluator for spreadsheet formulas."""

    OPERATORS = {
        '+': (1, operator.add),
        '-': (1, operator.sub),
        '*': (2, operator.mul),
        '/': (2, operator.truediv),
        '^': (3, operator.pow),
        '%': (2, operator.mod),
    }

    COMPARISONS = {
        '=': operator.eq,
        '<>': operator.ne,
        '!=': operator.ne,
        '<': operator.lt,
        '>': operator.gt,
        '<=': operator.le,
        '>=': operator.ge,
    }

    def __init__(self, spreadsheet: Spreadsheet):
        self.spreadsheet = spreadsheet
        self.functions: dict[str, Callable] = {
            'SUM': self._fn_sum,
            'AVG': self._fn_avg,
            'AVERAGE': self._fn_avg,
            'MIN': self._fn_min,
            'MAX': self._fn_max,
            'COUNT': self._fn_count,
            'COUNTA': self._fn_counta,
            'ABS': lambda x: abs(x),
            'INT': lambda x: int(x),
            'ROUND': lambda x, d=0: round(x, int(d)),
            'SQRT': lambda x: math.sqrt(x),
            'SIN': lambda x: math.sin(x),
            'COS': lambda x: math.cos(x),
            'TAN': lambda x: math.tan(x),
            'LOG': lambda x: math.log10(x),
            'LN': lambda x: math.log(x),
            'EXP': lambda x: math.exp(x),
            'PI': lambda: math.pi,
            'IF': self._fn_if,
            'AND': self._fn_and,
            'OR': self._fn_or,
            'NOT': lambda x: not x,
            'LEN': lambda x: len(str(x)),
            'LEFT': lambda s, n: str(s)[:int(n)],
            'RIGHT': lambda s, n: str(s)[-int(n):],
            'MID': lambda s, start, n: str(s)[int(start)-1:int(start)-1+int(n)],
            'UPPER': lambda s: str(s).upper(),
            'LOWER': lambda s: str(s).lower(),
            'TRIM': lambda s: str(s).strip(),
            'CONCATENATE': lambda *args: ''.join(str(a) for a in args),
            'CONCAT': lambda *args: ''.join(str(a) for a in args),
            'VALUE': lambda s: float(s),
            'TEXT': lambda v, _: str(v),
            'NOW': lambda: __import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M'),
            'TODAY': lambda: __import__('datetime').date.today().isoformat(),
            'POWER': lambda x, y: x ** y,
            'MOD': lambda x, y: x % y,
        }

    def evaluate(self, formula: str) -> Any:
        """Evaluate a formula string and return the result."""
        formula = formula.strip()
        if not formula:
            return ""
        tokens = self._tokenize(formula)
        result = self._parse_expression(tokens)
        return result

    def _tokenize(self, formula: str) -> list:
        """Convert formula string to tokens."""
        tokens = []
        i = 0
        while i < len(formula):
            ch = formula[i]

            if ch.isspace():
                i += 1
                continue

            # String literal
            if ch == '"':
                j = i + 1
                while j < len(formula) and formula[j] != '"':
                    j += 1
                tokens.append(('STRING', formula[i+1:j]))
                i = j + 1
                continue

            # Multi-char comparison operators
            if i + 1 < len(formula):
                two = formula[i:i+2]
                if two in ('<>', '<=', '>=', '!='):
                    tokens.append(('CMP', two))
                    i += 2
                    continue

            # Single char comparison
            if ch in '<>=':
                tokens.append(('CMP', ch))
                i += 1
                continue

            # Operators and punctuation
            if ch in '+-*/^%':
                tokens.append(('OP', ch))
                i += 1
                continue

            if ch in '(),':
                tokens.append((ch, ch))
                i += 1
                continue

            if ch == ':':
                tokens.append((':', ':'))
                i += 1
                continue

            # @ prefix for functions (Lotus style)
            if ch == '@':
                i += 1
                continue

            # Number
            if ch.isdigit() or (ch == '.' and i + 1 < len(formula) and formula[i+1].isdigit()):
                j = i
                has_dot = False
                while j < len(formula) and (formula[j].isdigit() or (formula[j] == '.' and not has_dot)):
                    if formula[j] == '.':
                        has_dot = True
                    j += 1
                num_str = formula[i:j]
                tokens.append(('NUM', float(num_str) if '.' in num_str else int(num_str)))
                i = j
                continue

            # Identifier (function name or cell reference)
            if ch.isalpha() or ch == '_':
                j = i
                while j < len(formula) and (formula[j].isalnum() or formula[j] in '_$'):
                    j += 1
                name = formula[i:j]
                # Check if it's a function (followed by parenthesis)
                k = j
                while k < len(formula) and formula[k].isspace():
                    k += 1
                if k < len(formula) and formula[k] == '(':
                    tokens.append(('FUNC', name.upper()))
                else:
                    # Cell reference
                    tokens.append(('CELL', name.upper()))
                i = j
                continue

            i += 1

        return tokens

    def _parse_expression(self, tokens: list, min_prec: int = 0) -> Any:
        """Parse expression with operator precedence."""
        if not tokens:
            return ""

        left = self._parse_atom(tokens)

        while tokens:
            # Check for comparison
            if tokens[0][0] == 'CMP':
                op_str = tokens[0][1]
                tokens.pop(0)
                right = self._parse_expression(tokens, 0)
                cmp_fn = self.COMPARISONS.get(op_str)
                if cmp_fn:
                    left = cmp_fn(left, right)
                continue

            if tokens[0][0] != 'OP':
                break

            op_str = tokens[0][1]
            if op_str not in self.OPERATORS:
                break

            prec, op_fn = self.OPERATORS[op_str]
            if prec < min_prec:
                break

            tokens.pop(0)
            right = self._parse_expression(tokens, prec + 1)

            try:
                left = op_fn(left, right)
            except ZeroDivisionError:
                left = "#DIV/0!"
            except Exception:
                left = "#ERR!"

        return left

    def _parse_atom(self, tokens: list) -> Any:
        """Parse an atomic value (number, string, cell ref, function call, parenthesized expr)."""
        if not tokens:
            return ""

        tok_type, tok_val = tokens[0]

        # Unary minus
        if tok_type == 'OP' and tok_val == '-':
            tokens.pop(0)
            val = self._parse_atom(tokens)
            return -val if isinstance(val, (int, float)) else val

        # Unary plus
        if tok_type == 'OP' and tok_val == '+':
            tokens.pop(0)
            return self._parse_atom(tokens)

        if tok_type == 'NUM':
            tokens.pop(0)
            return tok_val

        if tok_type == 'STRING':
            tokens.pop(0)
            return tok_val

        if tok_type == 'CELL':
            tokens.pop(0)
            # Check if it's part of a range
            if tokens and tokens[0][0] == ':':
                tokens.pop(0)
                if tokens and tokens[0][0] == 'CELL':
                    end_ref = tokens.pop(0)[1]
                    return self._get_range_values(tok_val, end_ref)
            return self.spreadsheet.get_value_by_ref(tok_val)

        if tok_type == 'FUNC':
            tokens.pop(0)
            return self._parse_function(tok_val, tokens)

        if tok_type == '(':
            tokens.pop(0)
            result = self._parse_expression(tokens)
            if tokens and tokens[0][0] == ')':
                tokens.pop(0)
            return result

        tokens.pop(0)
        return ""

    def _parse_function(self, name: str, tokens: list) -> Any:
        """Parse and evaluate a function call."""
        args = []

        # Consume opening paren
        if tokens and tokens[0][0] == '(':
            tokens.pop(0)

        while tokens:
            if tokens[0][0] == ')':
                tokens.pop(0)
                break
            if tokens[0][0] == ',':
                tokens.pop(0)
                continue
            arg = self._parse_expression(tokens)
            args.append(arg)

        fn = self.functions.get(name)
        if not fn:
            return f"#NAME?:{name}"

        try:
            # Flatten range lists for aggregate functions
            if name in ('SUM', 'AVG', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'AND', 'OR'):
                flat_args = []
                for arg in args:
                    if isinstance(arg, list):
                        flat_args.extend(arg)
                    else:
                        flat_args.append(arg)
                return fn(flat_args)
            return fn(*args)
        except Exception as e:
            return f"#ERR!"

    def _get_range_values(self, start_ref: str, end_ref: str) -> list:
        """Get flat list of values from a range."""
        return self.spreadsheet.get_range_flat(start_ref, end_ref)

    # Aggregate functions
    def _fn_sum(self, values: list) -> float:
        return sum(v for v in values if isinstance(v, (int, float)))

    def _fn_avg(self, values: list) -> float:
        nums = [v for v in values if isinstance(v, (int, float))]
        return sum(nums) / len(nums) if nums else 0

    def _fn_min(self, values: list) -> float:
        nums = [v for v in values if isinstance(v, (int, float))]
        return min(nums) if nums else 0

    def _fn_max(self, values: list) -> float:
        nums = [v for v in values if isinstance(v, (int, float))]
        return max(nums) if nums else 0

    def _fn_count(self, values: list) -> int:
        return sum(1 for v in values if isinstance(v, (int, float)))

    def _fn_counta(self, values: list) -> int:
        return sum(1 for v in values if v != "")

    def _fn_if(self, condition: Any, true_val: Any, false_val: Any = "") -> Any:
        return true_val if condition else false_val

    def _fn_and(self, values: list) -> bool:
        return all(bool(v) for v in values)

    def _fn_or(self, values: list) -> bool:
        return any(bool(v) for v in values)

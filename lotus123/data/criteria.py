"""Criteria parsing for database queries.

Supports Lotus 1-2-3 style criteria:
- Exact match
- Wildcards (* and ?)
- Comparison operators (<, >, <=, >=, <>)
- Compound criteria (AND across columns, OR across rows)
- Formula-based criteria
"""

import re
from dataclasses import dataclass
from enum import Enum, auto
from typing import Any, Callable


class CriterionOperator(Enum):
    """Comparison operators for criteria."""

    EQUAL = auto()
    NOT_EQUAL = auto()
    LESS_THAN = auto()
    GREATER_THAN = auto()
    LESS_EQUAL = auto()
    GREATER_EQUAL = auto()
    CONTAINS = auto()  # Wildcard pattern
    STARTS_WITH = auto()
    ENDS_WITH = auto()


@dataclass
class Criterion:
    """A single criterion for matching."""

    column: int | None = None  # None for formula-based
    operator: CriterionOperator = CriterionOperator.EQUAL
    value: Any = None
    pattern: str | None = None  # For wildcard matching
    is_formula: bool = False
    formula: str = ""

    def matches(self, cell_value: Any) -> bool:
        """Check if a cell value matches this criterion."""
        if self.is_formula:
            # Formula criteria would need evaluation
            return True  # Placeholder

        if self.operator == CriterionOperator.EQUAL:
            return self._equal_match(cell_value)
        elif self.operator == CriterionOperator.NOT_EQUAL:
            return not self._equal_match(cell_value)
        elif self.operator == CriterionOperator.LESS_THAN:
            return self._compare(cell_value) < 0
        elif self.operator == CriterionOperator.GREATER_THAN:
            return self._compare(cell_value) > 0
        elif self.operator == CriterionOperator.LESS_EQUAL:
            return self._compare(cell_value) <= 0
        elif self.operator == CriterionOperator.GREATER_EQUAL:
            return self._compare(cell_value) >= 0
        elif self.operator == CriterionOperator.CONTAINS:
            return self._wildcard_match(cell_value)
        elif self.operator == CriterionOperator.STARTS_WITH:
            return str(cell_value).lower().startswith(str(self.value).lower())
        elif self.operator == CriterionOperator.ENDS_WITH:
            return str(cell_value).lower().endswith(str(self.value).lower())

        return False

    def _equal_match(self, cell_value: Any) -> bool:
        """Check equality, handling types appropriately."""
        if self.pattern:
            return self._wildcard_match(cell_value)

        # Numeric comparison
        if isinstance(self.value, (int, float)) and isinstance(cell_value, (int, float)):
            return cell_value == self.value

        # String comparison (case-insensitive)
        return str(cell_value).lower() == str(self.value).lower()

    def _compare(self, cell_value: Any) -> int:
        """Compare cell value to criterion value.

        Returns: negative if cell < criterion, 0 if equal, positive if cell > criterion
        """
        try:
            cell_num = float(cell_value) if cell_value else 0
            crit_num = float(self.value) if self.value else 0
            if cell_num < crit_num:
                return -1
            elif cell_num > crit_num:
                return 1
            return 0
        except (ValueError, TypeError):
            # String comparison
            cell_str = str(cell_value).lower()
            crit_str = str(self.value).lower()
            if cell_str < crit_str:
                return -1
            elif cell_str > crit_str:
                return 1
            return 0

    def _wildcard_match(self, cell_value: Any) -> bool:
        """Match using wildcards (* = any chars, ? = single char)."""
        if not self.pattern:
            return True

        # Convert Lotus wildcards to regex
        pattern = self.pattern
        pattern = pattern.replace(".", r"\.")
        pattern = pattern.replace("*", ".*")
        pattern = pattern.replace("?", ".")
        pattern = f"^{pattern}$"

        try:
            return bool(re.match(pattern, str(cell_value), re.IGNORECASE))
        except re.error:
            return False


class CriteriaParser:
    """Parser for Lotus 1-2-3 style criteria ranges.

    Criteria range format:
    - First row: field names (must match data headers)
    - Subsequent rows: criteria values

    Rules:
    - Multiple criteria in same row = AND
    - Multiple rows = OR
    - Empty cell = no constraint on that field
    """

    COMPARISON_PATTERN = re.compile(r"^(<>|<=|>=|<|>|=)(.*)$")

    def __init__(self) -> None:
        self._criteria: list[list[Criterion]] = []  # List of AND groups (OR between groups)

    def parse_range(self, headers: list[str], criteria_rows: list[list[Any]]) -> None:
        """Parse a criteria range.

        Args:
            headers: Field names from criteria range first row
            criteria_rows: Subsequent rows containing criteria values
        """
        self._criteria = []

        for row in criteria_rows:
            and_group = []
            for col, value in enumerate(row):
                if value is None or value == "":
                    continue

                criterion = self.parse_criterion(col, value)
                and_group.append(criterion)

            if and_group:
                self._criteria.append(and_group)

    def parse_criterion(self, column: int, value: Any) -> Criterion:
        """Parse a single criterion value."""
        value_str = str(value).strip()

        # Check for comparison operators
        match = self.COMPARISON_PATTERN.match(value_str)
        if match:
            op_str, val = match.groups()
            operator = self._parse_operator(op_str)
            return Criterion(column=column, operator=operator, value=self._parse_value(val))

        # Check for wildcards
        if "*" in value_str or "?" in value_str:
            return Criterion(column=column, operator=CriterionOperator.CONTAINS, pattern=value_str)

        # Check for formula
        if value_str.startswith("+") or value_str.startswith("="):
            return Criterion(column=column, is_formula=True, formula=value_str)

        # Exact match
        return Criterion(
            column=column, operator=CriterionOperator.EQUAL, value=self._parse_value(value_str)
        )

    def _parse_operator(self, op_str: str) -> CriterionOperator:
        """Convert operator string to enum."""
        operators = {
            "=": CriterionOperator.EQUAL,
            "<>": CriterionOperator.NOT_EQUAL,
            "<": CriterionOperator.LESS_THAN,
            ">": CriterionOperator.GREATER_THAN,
            "<=": CriterionOperator.LESS_EQUAL,
            ">=": CriterionOperator.GREATER_EQUAL,
        }
        return operators.get(op_str, CriterionOperator.EQUAL)

    def _parse_value(self, value: str) -> Any:
        """Parse a value string to appropriate type."""
        value = value.strip()
        if not value:
            return ""

        # Try numeric
        try:
            if "." in value:
                return float(value)
            return int(value)
        except ValueError:
            return value

    def matches(self, row_values: list[Any]) -> bool:
        """Check if a row matches the criteria.

        Args:
            row_values: Values from a data row

        Returns:
            True if row matches (any OR group matches)
        """
        if not self._criteria:
            return True  # No criteria = match all

        # OR between groups
        for and_group in self._criteria:
            # AND within group
            group_matches = True
            for criterion in and_group:
                if criterion.column is not None and criterion.column < len(row_values):
                    if not criterion.matches(row_values[criterion.column]):
                        group_matches = False
                        break

            if group_matches:
                return True

        return False

    def create_filter(self) -> Callable[[list[Any]], bool]:
        """Create a filter function for use with DatabaseOperations.query().

        Returns:
            Function that takes row_values and returns bool
        """
        return lambda row: self.matches(row)


def parse_simple_criteria(column: int, criteria_str: str) -> Criterion:
    """Parse a simple criteria string for a specific column.

    Convenience function for simple filtering.

    Examples:
        ">100" - greater than 100
        "John*" - starts with John
        "<>0" - not equal to 0
    """
    parser = CriteriaParser()
    return parser.parse_criterion(column, criteria_str)

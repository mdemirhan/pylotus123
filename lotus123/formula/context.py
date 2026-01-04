"""Evaluation context for formula evaluation."""

from dataclasses import dataclass, field


@dataclass
class EvaluationContext:
    """Context for formula evaluation.

    Tracks the current cell being evaluated and dependencies.
    """

    current_row: int = 0
    current_col: int = 0
    dependencies: set[tuple[int, int]] = field(default_factory=set)
    computing: set[tuple[int, int]] = field(default_factory=set)

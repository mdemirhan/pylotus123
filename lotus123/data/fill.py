"""Fill operations for sequences and patterns.

Implements /Data Fill command functionality.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum, auto
from typing import TYPE_CHECKING, Any
import re
import datetime

if TYPE_CHECKING:
    from ..core.spreadsheet import Spreadsheet


class FillType(Enum):
    """Type of fill operation."""
    LINEAR = auto()       # Numeric sequence with constant step
    GROWTH = auto()       # Geometric sequence (multiply by step)
    DATE = auto()         # Date sequence
    AUTO = auto()         # Auto-detect pattern
    COPY = auto()         # Just copy values


@dataclass
class FillSpec:
    """Specification for a fill operation."""
    fill_type: FillType = FillType.LINEAR
    start_value: float = 1
    step: float = 1
    stop_value: float | None = None
    date_unit: str = "day"  # day, week, month, year


class FillOperations:
    """Fill operations for ranges.

    Supports:
    - Linear numeric sequences
    - Geometric (growth) sequences
    - Date sequences
    - Pattern detection and continuation
    """

    def __init__(self, spreadsheet: Spreadsheet) -> None:
        self.spreadsheet = spreadsheet

    def fill_series(self, start_row: int, start_col: int,
                    end_row: int, end_col: int,
                    spec: FillSpec,
                    direction: str = "down") -> None:
        """Fill a range with a series.

        Args:
            start_row, start_col: Start of range
            end_row, end_col: End of range
            spec: Fill specification
            direction: "down", "right", "up", or "left"
        """
        # Normalize range
        if start_row > end_row:
            start_row, end_row = end_row, start_row
        if start_col > end_col:
            start_col, end_col = end_col, start_col

        if spec.fill_type == FillType.LINEAR:
            self._fill_linear(start_row, start_col, end_row, end_col, spec, direction)
        elif spec.fill_type == FillType.GROWTH:
            self._fill_growth(start_row, start_col, end_row, end_col, spec, direction)
        elif spec.fill_type == FillType.DATE:
            self._fill_date(start_row, start_col, end_row, end_col, spec, direction)
        elif spec.fill_type == FillType.COPY:
            self._fill_copy(start_row, start_col, end_row, end_col, direction)
        else:  # AUTO
            self._fill_auto(start_row, start_col, end_row, end_col, direction)

        self.spreadsheet._invalidate_cache()

    def _fill_linear(self, start_row: int, start_col: int,
                     end_row: int, end_col: int,
                     spec: FillSpec, direction: str) -> None:
        """Fill with linear sequence."""
        value = spec.start_value
        step = spec.step

        for row, col in self._iter_cells(start_row, start_col, end_row, end_col, direction):
            if spec.stop_value is not None:
                if (step > 0 and value > spec.stop_value) or (step < 0 and value < spec.stop_value):
                    break

            self.spreadsheet.set_cell(row, col, str(value))
            value += step

    def _fill_growth(self, start_row: int, start_col: int,
                     end_row: int, end_col: int,
                     spec: FillSpec, direction: str) -> None:
        """Fill with geometric sequence."""
        value = spec.start_value
        step = spec.step

        for row, col in self._iter_cells(start_row, start_col, end_row, end_col, direction):
            if spec.stop_value is not None:
                if (step > 1 and value > spec.stop_value) or (step < 1 and value < spec.stop_value):
                    break

            self.spreadsheet.set_cell(row, col, str(value))
            value *= step

    def _fill_date(self, start_row: int, start_col: int,
                   end_row: int, end_col: int,
                   spec: FillSpec, direction: str) -> None:
        """Fill with date sequence."""
        from ..core.formatting import serial_to_date, date_to_serial

        serial = spec.start_value
        step = int(spec.step)
        unit = spec.date_unit.lower()

        for row, col in self._iter_cells(start_row, start_col, end_row, end_col, direction):
            try:
                date = serial_to_date(serial)

                if unit == "day":
                    date = date + datetime.timedelta(days=0)  # Just use current
                    serial += step
                elif unit == "week":
                    serial += step * 7
                elif unit == "month":
                    new_month = date.month + step
                    new_year = date.year + (new_month - 1) // 12
                    new_month = ((new_month - 1) % 12) + 1
                    import calendar
                    max_day = calendar.monthrange(new_year, new_month)[1]
                    new_day = min(date.day, max_day)
                    date = datetime.date(new_year, new_month, new_day)
                    serial = date_to_serial(date)
                elif unit == "year":
                    date = datetime.date(date.year + step, date.month, date.day)
                    serial = date_to_serial(date)

                self.spreadsheet.set_cell(row, col, str(int(serial)))

            except (ValueError, OverflowError):
                break

    def _fill_copy(self, start_row: int, start_col: int,
                   end_row: int, end_col: int, direction: str) -> None:
        """Fill by copying source cells."""
        # Get source values
        if direction in ("down", "up"):
            # Source is first column
            source = []
            for c in range(start_col, end_col + 1):
                cell = self.spreadsheet.get_cell_if_exists(start_row, c)
                source.append(cell.raw_value if cell else "")
        else:
            # Source is first row
            source = []
            for r in range(start_row, end_row + 1):
                cell = self.spreadsheet.get_cell_if_exists(r, start_col)
                source.append(cell.raw_value if cell else "")

        # Fill
        idx = 0
        for row, col in self._iter_cells(start_row, start_col, end_row, end_col, direction):
            if direction in ("down", "up"):
                src_idx = col - start_col
            else:
                src_idx = row - start_row

            if src_idx < len(source):
                self.spreadsheet.set_cell(row, col, source[src_idx])
            idx += 1

    def _fill_auto(self, start_row: int, start_col: int,
                   end_row: int, end_col: int, direction: str) -> None:
        """Fill by detecting and continuing pattern."""
        # Get source values to detect pattern
        source_values = []
        if direction in ("down", "up"):
            # Use first 1-3 rows as source
            for r in range(start_row, min(start_row + 3, end_row + 1)):
                for c in range(start_col, end_col + 1):
                    source_values.append(self.spreadsheet.get_value(r, c))
        else:
            # Use first 1-3 columns as source
            for r in range(start_row, end_row + 1):
                for c in range(start_col, min(start_col + 3, end_col + 1)):
                    source_values.append(self.spreadsheet.get_value(r, c))

        # Detect pattern
        pattern = self._detect_pattern(source_values)

        if pattern['type'] == 'linear':
            spec = FillSpec(
                fill_type=FillType.LINEAR,
                start_value=pattern['start'],
                step=pattern['step'],
            )
            self._fill_linear(start_row, start_col, end_row, end_col, spec, direction)
        elif pattern['type'] == 'text_sequence':
            self._fill_text_sequence(start_row, start_col, end_row, end_col,
                                    pattern['values'], direction)
        else:
            # Default to copy
            self._fill_copy(start_row, start_col, end_row, end_col, direction)

    def _detect_pattern(self, values: list) -> dict:
        """Detect the pattern in a list of values."""
        if not values:
            return {'type': 'none'}

        # Check for numeric sequence
        numbers = []
        for v in values:
            if isinstance(v, (int, float)):
                numbers.append(float(v))
            elif isinstance(v, str):
                try:
                    numbers.append(float(v.replace(",", "")))
                except ValueError:
                    numbers = []
                    break

        if len(numbers) >= 2:
            # Check for constant step (linear)
            diffs = [numbers[i+1] - numbers[i] for i in range(len(numbers)-1)]
            if all(abs(d - diffs[0]) < 0.0001 for d in diffs):
                return {
                    'type': 'linear',
                    'start': numbers[0],
                    'step': diffs[0],
                }

            # Check for constant ratio (growth)
            if all(n != 0 for n in numbers[:-1]):
                ratios = [numbers[i+1] / numbers[i] for i in range(len(numbers)-1)]
                if all(abs(r - ratios[0]) < 0.0001 for r in ratios):
                    return {
                        'type': 'growth',
                        'start': numbers[0],
                        'ratio': ratios[0],
                    }

        # Check for text patterns (like Mon, Tue, Wed or Jan, Feb, Mar)
        strings = [str(v).strip() for v in values if v]
        if strings:
            # Known sequences
            days = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday']
            days_short = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat']
            months = ['january', 'february', 'march', 'april', 'may', 'june',
                     'july', 'august', 'september', 'october', 'november', 'december']
            months_short = ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
                           'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

            lower_strings = [s.lower() for s in strings]

            for seq_list in [days, days_short, months, months_short]:
                if lower_strings[0] in seq_list:
                    return {
                        'type': 'text_sequence',
                        'values': seq_list,
                        'start_idx': seq_list.index(lower_strings[0]),
                    }

        return {'type': 'none'}

    def _fill_text_sequence(self, start_row: int, start_col: int,
                           end_row: int, end_col: int,
                           sequence: list[str], direction: str) -> None:
        """Fill with a cyclical text sequence."""
        # Find starting index
        start_cell = self.spreadsheet.get_value(start_row, start_col)
        start_str = str(start_cell).lower()

        try:
            idx = sequence.index(start_str)
        except ValueError:
            idx = 0

        for row, col in self._iter_cells(start_row, start_col, end_row, end_col, direction):
            value = sequence[idx % len(sequence)]
            # Match case of original
            if start_cell and str(start_cell)[0].isupper():
                value = value.capitalize()
            self.spreadsheet.set_cell(row, col, value)
            idx += 1

    def _iter_cells(self, start_row: int, start_col: int,
                    end_row: int, end_col: int,
                    direction: str):
        """Iterate over cells in specified direction."""
        if direction == "down":
            for r in range(start_row, end_row + 1):
                for c in range(start_col, end_col + 1):
                    yield r, c
        elif direction == "up":
            for r in range(end_row, start_row - 1, -1):
                for c in range(start_col, end_col + 1):
                    yield r, c
        elif direction == "right":
            for c in range(start_col, end_col + 1):
                for r in range(start_row, end_row + 1):
                    yield r, c
        else:  # left
            for c in range(end_col, start_col - 1, -1):
                for r in range(start_row, end_row + 1):
                    yield r, c

    def fill_down(self, start_row: int, start_col: int,
                  end_row: int, end_col: int) -> None:
        """Fill down from first row to remaining rows."""
        for c in range(start_col, end_col + 1):
            source = self.spreadsheet.get_cell_if_exists(start_row, c)
            if not source:
                continue

            for r in range(start_row + 1, end_row + 1):
                # Copy with formula adjustment
                self.spreadsheet.copy_cell(start_row, c, r, c, adjust_refs=True)

        self.spreadsheet._invalidate_cache()

    def fill_right(self, start_row: int, start_col: int,
                   end_row: int, end_col: int) -> None:
        """Fill right from first column to remaining columns."""
        for r in range(start_row, end_row + 1):
            source = self.spreadsheet.get_cell_if_exists(r, start_col)
            if not source:
                continue

            for c in range(start_col + 1, end_col + 1):
                # Copy with formula adjustment
                self.spreadsheet.copy_cell(r, start_col, r, c, adjust_refs=True)

        self.spreadsheet._invalidate_cache()

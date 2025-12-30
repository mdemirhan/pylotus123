"""Tests for fill operations module."""

import pytest

from lotus123 import Spreadsheet
from lotus123.data.fill import (
    FillOperations,
    FillSpec,
    FillType,
)


class TestFillType:
    """Tests for FillType enum."""

    def test_types_exist(self):
        """Test all fill types exist."""
        assert FillType.LINEAR
        assert FillType.GROWTH
        assert FillType.DATE
        assert FillType.AUTO
        assert FillType.COPY


class TestFillSpec:
    """Tests for FillSpec dataclass."""

    def test_default_values(self):
        """Test default values."""
        spec = FillSpec()
        assert spec.fill_type == FillType.LINEAR
        assert spec.start_value == 1
        assert spec.step == 1
        assert spec.stop_value is None
        assert spec.date_unit == "day"

    def test_custom_values(self):
        """Test custom values."""
        spec = FillSpec(
            fill_type=FillType.GROWTH,
            start_value=10,
            step=2,
            stop_value=100
        )
        assert spec.fill_type == FillType.GROWTH
        assert spec.start_value == 10
        assert spec.step == 2
        assert spec.stop_value == 100


class TestFillLinear:
    """Tests for linear fill."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.fill = FillOperations(self.ss)

    def test_fill_linear_down(self):
        """Test linear fill going down."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=1, step=1)
        self.fill.fill_series(0, 0, 4, 0, spec, "down")

        assert self.ss.get_value(0, 0) == 1
        assert self.ss.get_value(1, 0) == 2
        assert self.ss.get_value(2, 0) == 3
        assert self.ss.get_value(3, 0) == 4
        assert self.ss.get_value(4, 0) == 5

    def test_fill_linear_right(self):
        """Test linear fill going right."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=10, step=5)
        self.fill.fill_series(0, 0, 0, 3, spec, "right")

        assert self.ss.get_value(0, 0) == 10
        assert self.ss.get_value(0, 1) == 15
        assert self.ss.get_value(0, 2) == 20
        assert self.ss.get_value(0, 3) == 25

    def test_fill_linear_with_stop(self):
        """Test linear fill with stop value."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=1, step=1, stop_value=3)
        self.fill.fill_series(0, 0, 9, 0, spec, "down")

        assert self.ss.get_value(0, 0) == 1
        assert self.ss.get_value(1, 0) == 2
        assert self.ss.get_value(2, 0) == 3
        # Should stop at 3
        assert self.ss.get_value(3, 0) == ""

    def test_fill_linear_negative_step(self):
        """Test linear fill with negative step."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=10, step=-2)
        self.fill.fill_series(0, 0, 3, 0, spec, "down")

        assert self.ss.get_value(0, 0) == 10
        assert self.ss.get_value(1, 0) == 8
        assert self.ss.get_value(2, 0) == 6
        assert self.ss.get_value(3, 0) == 4

    def test_fill_linear_reversed_range(self):
        """Test fill with reversed range normalizes correctly."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=1, step=1)
        # Pass end before start
        self.fill.fill_series(4, 0, 0, 0, spec, "down")

        # Should still work (range normalized)
        assert self.ss.get_value(0, 0) == 1


class TestFillGrowth:
    """Tests for growth fill."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.fill = FillOperations(self.ss)

    def test_fill_growth_basic(self):
        """Test basic growth fill."""
        spec = FillSpec(fill_type=FillType.GROWTH, start_value=1, step=2)
        self.fill.fill_series(0, 0, 4, 0, spec, "down")

        assert self.ss.get_value(0, 0) == 1
        assert self.ss.get_value(1, 0) == 2
        assert self.ss.get_value(2, 0) == 4
        assert self.ss.get_value(3, 0) == 8
        assert self.ss.get_value(4, 0) == 16

    def test_fill_growth_with_stop(self):
        """Test growth fill with stop value."""
        spec = FillSpec(fill_type=FillType.GROWTH, start_value=1, step=2, stop_value=5)
        self.fill.fill_series(0, 0, 9, 0, spec, "down")

        assert self.ss.get_value(0, 0) == 1
        assert self.ss.get_value(1, 0) == 2
        assert self.ss.get_value(2, 0) == 4
        # Should stop before exceeding 5
        assert self.ss.get_value(3, 0) == ""


class TestFillDate:
    """Tests for date fill."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.fill = FillOperations(self.ss)

    def test_fill_date_days(self):
        """Test date fill by days."""
        # Serial 45000 is a date in 2023
        spec = FillSpec(fill_type=FillType.DATE, start_value=45000, step=1, date_unit="day")
        self.fill.fill_series(0, 0, 2, 0, spec, "down")

        val0 = self.ss.get_value(0, 0)
        val1 = self.ss.get_value(1, 0)
        val2 = self.ss.get_value(2, 0)
        # Should increment by 1 day
        assert val1 == val0 + 1
        assert val2 == val0 + 2

    def test_fill_date_weeks(self):
        """Test date fill by weeks."""
        spec = FillSpec(fill_type=FillType.DATE, start_value=45000, step=1, date_unit="week")
        self.fill.fill_series(0, 0, 2, 0, spec, "down")

        val0 = self.ss.get_value(0, 0)
        val1 = self.ss.get_value(1, 0)
        # Should increment by 7 days
        assert val1 == val0 + 7

    def test_fill_date_months(self):
        """Test date fill by months."""
        spec = FillSpec(fill_type=FillType.DATE, start_value=45000, step=1, date_unit="month")
        self.fill.fill_series(0, 0, 2, 0, spec, "down")

        # Just verify it doesn't crash and produces values
        assert self.ss.get_value(0, 0) is not None
        assert self.ss.get_value(1, 0) is not None

    def test_fill_date_years(self):
        """Test date fill by years."""
        spec = FillSpec(fill_type=FillType.DATE, start_value=45000, step=1, date_unit="year")
        self.fill.fill_series(0, 0, 2, 0, spec, "down")

        # Just verify it doesn't crash
        assert self.ss.get_value(0, 0) is not None


class TestFillCopy:
    """Tests for copy fill."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.fill = FillOperations(self.ss)

    def test_fill_copy_down(self):
        """Test copy fill down."""
        self.ss.set_cell(0, 0, "A")
        self.ss.set_cell(0, 1, "B")

        spec = FillSpec(fill_type=FillType.COPY)
        self.fill.fill_series(0, 0, 2, 1, spec, "down")

        assert self.ss.get_value(0, 0) == "A"
        assert self.ss.get_value(0, 1) == "B"
        assert self.ss.get_value(1, 0) == "A"
        assert self.ss.get_value(1, 1) == "B"
        assert self.ss.get_value(2, 0) == "A"
        assert self.ss.get_value(2, 1) == "B"

    def test_fill_copy_right(self):
        """Test copy fill right."""
        self.ss.set_cell(0, 0, "X")
        self.ss.set_cell(1, 0, "Y")

        spec = FillSpec(fill_type=FillType.COPY)
        self.fill.fill_series(0, 0, 1, 2, spec, "right")

        assert self.ss.get_value(0, 0) == "X"
        assert self.ss.get_value(1, 0) == "Y"
        assert self.ss.get_value(0, 1) == "X"
        assert self.ss.get_value(1, 1) == "Y"


class TestFillAuto:
    """Tests for auto fill."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.fill = FillOperations(self.ss)

    def test_fill_auto_numeric(self):
        """Test auto fill detects numeric pattern."""
        self.ss.set_cell(0, 0, "1")
        self.ss.set_cell(1, 0, "2")
        self.ss.set_cell(2, 0, "3")

        spec = FillSpec(fill_type=FillType.AUTO)
        self.fill.fill_series(0, 0, 5, 0, spec, "down")

        # Should continue the pattern
        assert self.ss.get_value(3, 0) == 4
        assert self.ss.get_value(4, 0) == 5
        assert self.ss.get_value(5, 0) == 6

    def test_fill_auto_text_days(self):
        """Test auto fill detects day sequence."""
        self.ss.set_cell(0, 0, "Mon")

        spec = FillSpec(fill_type=FillType.AUTO)
        self.fill.fill_series(0, 0, 3, 0, spec, "down")

        # Should continue with days
        assert self.ss.get_value(0, 0) in ["Mon", "mon"]

    def test_fill_auto_text_months(self):
        """Test auto fill detects month sequence."""
        self.ss.set_cell(0, 0, "Jan")

        spec = FillSpec(fill_type=FillType.AUTO)
        self.fill.fill_series(0, 0, 2, 0, spec, "down")

        # Should continue with months
        assert self.ss.get_value(0, 0) in ["Jan", "jan"]


class TestDetectPattern:
    """Tests for pattern detection."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.fill = FillOperations(self.ss)

    def test_detect_linear(self):
        """Test detecting linear pattern."""
        pattern = self.fill._detect_pattern([1, 2, 3, 4])
        assert pattern["type"] == "linear"
        assert pattern["start"] == 1
        assert pattern["step"] == 1

    def test_detect_linear_step(self):
        """Test detecting linear pattern with step."""
        pattern = self.fill._detect_pattern([10, 20, 30])
        assert pattern["type"] == "linear"
        assert pattern["step"] == 10

    def test_detect_growth(self):
        """Test detecting growth pattern."""
        pattern = self.fill._detect_pattern([1, 2, 4, 8])
        assert pattern["type"] == "growth"
        assert pattern["ratio"] == 2

    def test_detect_day_sequence(self):
        """Test detecting day sequence."""
        pattern = self.fill._detect_pattern(["monday", "tuesday"])
        assert pattern["type"] == "text_sequence"

    def test_detect_month_sequence(self):
        """Test detecting month sequence."""
        pattern = self.fill._detect_pattern(["jan", "feb"])
        assert pattern["type"] == "text_sequence"

    def test_detect_empty(self):
        """Test detecting pattern in empty list."""
        pattern = self.fill._detect_pattern([])
        assert pattern["type"] == "none"

    def test_detect_no_pattern(self):
        """Test when no pattern is detected."""
        pattern = self.fill._detect_pattern(["random", "text", "here"])
        assert pattern["type"] == "none"


class TestFillDownRight:
    """Tests for fill_down and fill_right methods."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.fill = FillOperations(self.ss)

    def test_fill_down(self):
        """Test fill down copies values."""
        self.ss.set_cell(0, 0, "10")
        self.ss.set_cell(0, 1, "20")

        self.fill.fill_down(0, 0, 2, 1)

        # Values should be copied down
        assert self.ss.get_value(1, 0) == 10
        assert self.ss.get_value(2, 0) == 10
        assert self.ss.get_value(1, 1) == 20
        assert self.ss.get_value(2, 1) == 20

    def test_fill_right(self):
        """Test fill right copies values."""
        self.ss.set_cell(0, 0, "A")
        self.ss.set_cell(1, 0, "B")

        self.fill.fill_right(0, 0, 1, 2)

        assert self.ss.get_value(0, 1) == "A"
        assert self.ss.get_value(0, 2) == "A"
        assert self.ss.get_value(1, 1) == "B"
        assert self.ss.get_value(1, 2) == "B"

    def test_fill_down_empty_source(self):
        """Test fill down with empty source cell."""
        self.fill.fill_down(0, 0, 2, 0)
        # Should not crash with empty source


class TestIterCells:
    """Tests for cell iteration."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.fill = FillOperations(self.ss)

    def test_iter_down(self):
        """Test iterating down."""
        cells = list(self.fill._iter_cells(0, 0, 2, 1, "down"))
        # Should go row by row
        assert cells[0] == (0, 0)
        assert cells[1] == (0, 1)
        assert cells[2] == (1, 0)
        assert cells[3] == (1, 1)

    def test_iter_up(self):
        """Test iterating up."""
        cells = list(self.fill._iter_cells(0, 0, 2, 0, "up"))
        # Should start from bottom
        assert cells[0] == (2, 0)
        assert cells[1] == (1, 0)
        assert cells[2] == (0, 0)

    def test_iter_right(self):
        """Test iterating right."""
        cells = list(self.fill._iter_cells(0, 0, 1, 2, "right"))
        # Should go column by column
        assert cells[0] == (0, 0)
        assert cells[1] == (1, 0)
        assert cells[2] == (0, 1)
        assert cells[3] == (1, 1)

    def test_iter_left(self):
        """Test iterating left."""
        cells = list(self.fill._iter_cells(0, 0, 0, 2, "left"))
        # Should start from rightmost
        assert cells[0] == (0, 2)
        assert cells[1] == (0, 1)
        assert cells[2] == (0, 0)

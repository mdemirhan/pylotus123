"""Tests for fill operations."""

import pytest

from lotus123 import Spreadsheet
from lotus123.data.fill import FillOperations, FillSpec, FillType


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
            fill_type=FillType.GROWTH, start_value=10, step=2, stop_value=100, date_unit="month"
        )
        assert spec.fill_type == FillType.GROWTH
        assert spec.start_value == 10


class TestFillOperations:
    """Tests for FillOperations class."""

    def setup_method(self):
        """Set up test spreadsheet."""
        self.ss = Spreadsheet()
        self.fill_ops = FillOperations(self.ss)

    def test_fill_linear_down(self):
        """Test linear fill going down."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=1, step=1)
        self.fill_ops.fill_series(0, 0, 4, 0, spec, direction="down")

        assert self.ss.get_value(0, 0) == 1
        assert self.ss.get_value(1, 0) == 2
        assert self.ss.get_value(4, 0) == 5

    def test_fill_linear_with_step(self):
        """Test linear fill with custom step."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=0, step=5)
        self.fill_ops.fill_series(0, 0, 4, 0, spec, direction="down")

        assert self.ss.get_value(0, 0) == 0
        assert self.ss.get_value(1, 0) == 5
        assert self.ss.get_value(2, 0) == 10

    def test_fill_linear_with_stop(self):
        """Test linear fill stops at stop_value."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=1, step=1, stop_value=3)
        self.fill_ops.fill_series(0, 0, 9, 0, spec, direction="down")

        assert self.ss.get_value(0, 0) == 1
        assert self.ss.get_value(2, 0) == 3
        # Should stop at 3, row 3+ should be empty
        assert self.ss.get_value(3, 0) in [None, ""]

    def test_fill_linear_negative_step(self):
        """Test linear fill with negative step."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=10, step=-2)
        self.fill_ops.fill_series(0, 0, 4, 0, spec, direction="down")

        assert self.ss.get_value(0, 0) == 10
        assert self.ss.get_value(1, 0) == 8
        assert self.ss.get_value(2, 0) == 6

    def test_fill_growth(self):
        """Test geometric/growth fill."""
        spec = FillSpec(fill_type=FillType.GROWTH, start_value=1, step=2)
        self.fill_ops.fill_series(0, 0, 4, 0, spec, direction="down")

        assert self.ss.get_value(0, 0) == 1
        assert self.ss.get_value(1, 0) == 2
        assert self.ss.get_value(2, 0) == 4
        assert self.ss.get_value(3, 0) == 8

    def test_fill_growth_with_stop(self):
        """Test growth fill stops at stop_value."""
        spec = FillSpec(fill_type=FillType.GROWTH, start_value=1, step=2, stop_value=5)
        self.fill_ops.fill_series(0, 0, 9, 0, spec, direction="down")

        assert self.ss.get_value(0, 0) == 1
        assert self.ss.get_value(1, 0) == 2
        assert self.ss.get_value(2, 0) == 4

    def test_fill_copy(self):
        """Test copy fill."""
        # Set up source data
        self.ss.set_cell(0, 0, "Test")
        self.ss.set_cell(0, 1, "123")

        spec = FillSpec(fill_type=FillType.COPY)
        self.fill_ops.fill_series(0, 0, 3, 1, spec, direction="down")

        # Values should be copied
        assert self.ss.get_value(1, 0) == "Test"
        assert self.ss.get_value(1, 1) == 123

    def test_fill_direction_right(self):
        """Test fill going right."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=1, step=1)
        self.fill_ops.fill_series(0, 0, 0, 4, spec, direction="right")

        assert self.ss.get_value(0, 0) == 1
        assert self.ss.get_value(0, 1) == 2
        assert self.ss.get_value(0, 4) == 5

    def test_fill_direction_up(self):
        """Test fill going up."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=1, step=1)
        self.fill_ops.fill_series(0, 0, 4, 0, spec, direction="up")

        # Fill starts from bottom when going up
        assert self.ss.get_value(4, 0) == 1
        assert self.ss.get_value(3, 0) == 2

    def test_fill_direction_left(self):
        """Test fill going left."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=1, step=1)
        self.fill_ops.fill_series(0, 0, 0, 4, spec, direction="left")

        # Fill starts from right when going left
        assert self.ss.get_value(0, 4) == 1
        assert self.ss.get_value(0, 3) == 2

    def test_fill_normalizes_range(self):
        """Test that fill normalizes reversed ranges."""
        spec = FillSpec(fill_type=FillType.LINEAR, start_value=1, step=1)
        # Reversed range (end < start)
        self.fill_ops.fill_series(4, 4, 0, 0, spec, direction="down")

        # Should still work
        assert self.ss.get_value(0, 0) == 1

    def test_fill_down_method(self):
        """Test fill_down convenience method."""
        self.ss.set_cell(0, 0, "=A1+1")
        self.ss.set_cell(0, 1, "Test")

        self.fill_ops.fill_down(0, 0, 3, 1)

        # Formula should be copied and adjusted
        # First column has formula, second has text

    def test_fill_right_method(self):
        """Test fill_right convenience method."""
        self.ss.set_cell(0, 0, "=A1+1")
        self.ss.set_cell(1, 0, "Test")

        self.fill_ops.fill_right(0, 0, 1, 3)


class TestPatternDetection:
    """Tests for pattern detection in auto-fill."""

    def setup_method(self):
        """Set up test spreadsheet."""
        self.ss = Spreadsheet()
        self.fill_ops = FillOperations(self.ss)

    def test_detect_linear_pattern(self):
        """Test detecting linear numeric pattern."""
        values = [1, 2, 3]
        pattern = self.fill_ops._detect_pattern(values)
        assert pattern["type"] == "linear"
        assert pattern["start"] == 1
        assert pattern["step"] == 1

    def test_detect_linear_pattern_custom_step(self):
        """Test detecting linear pattern with custom step."""
        values = [10, 15, 20]
        pattern = self.fill_ops._detect_pattern(values)
        assert pattern["type"] == "linear"
        assert pattern["step"] == 5

    def test_detect_growth_pattern(self):
        """Test detecting geometric pattern."""
        values = [1, 2, 4, 8]
        pattern = self.fill_ops._detect_pattern(values)
        assert pattern["type"] == "growth"
        assert pattern["ratio"] == 2

    def test_detect_day_names(self):
        """Test detecting day name sequence."""
        values = ["Monday", "Tuesday"]
        pattern = self.fill_ops._detect_pattern(values)
        assert pattern["type"] == "text_sequence"

    def test_detect_month_names(self):
        """Test detecting month name sequence."""
        values = ["Jan", "Feb"]
        pattern = self.fill_ops._detect_pattern(values)
        assert pattern["type"] == "text_sequence"

    def test_detect_no_pattern(self):
        """Test no pattern detection."""
        values = ["random", "text", "here"]
        pattern = self.fill_ops._detect_pattern(values)
        assert pattern["type"] == "none"

    def test_detect_empty_values(self):
        """Test pattern detection with empty values."""
        pattern = self.fill_ops._detect_pattern([])
        assert pattern["type"] == "none"


class TestIterCells:
    """Tests for _iter_cells method."""

    def setup_method(self):
        """Set up test spreadsheet."""
        self.ss = Spreadsheet()
        self.fill_ops = FillOperations(self.ss)

    def test_iter_cells_down(self):
        """Test iterating cells going down."""
        cells = list(self.fill_ops._iter_cells(0, 0, 2, 0, "down"))
        assert cells == [(0, 0), (1, 0), (2, 0)]

    def test_iter_cells_right(self):
        """Test iterating cells going right."""
        cells = list(self.fill_ops._iter_cells(0, 0, 0, 2, "right"))
        assert cells == [(0, 0), (0, 1), (0, 2)]

    def test_iter_cells_up(self):
        """Test iterating cells going up."""
        cells = list(self.fill_ops._iter_cells(0, 0, 2, 0, "up"))
        assert cells == [(2, 0), (1, 0), (0, 0)]

    def test_iter_cells_left(self):
        """Test iterating cells going left."""
        cells = list(self.fill_ops._iter_cells(0, 0, 0, 2, "left"))
        assert cells == [(0, 2), (0, 1), (0, 0)]


class TestFillType:
    """Tests for FillType enum."""

    def test_all_types_exist(self):
        """Test all fill types exist."""
        assert FillType.LINEAR
        assert FillType.GROWTH
        assert FillType.DATE
        assert FillType.AUTO
        assert FillType.COPY

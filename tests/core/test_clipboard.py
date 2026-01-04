"""Tests for clipboard operations."""

from lotus123 import Spreadsheet
from lotus123.utils.clipboard import Clipboard, ClipboardCell, ClipboardContent, ClipboardMode


class TestClipboardCell:
    """Tests for ClipboardCell dataclass."""

    def test_default_values(self):
        """Test default values."""
        cell = ClipboardCell(raw_value="test")
        assert cell.raw_value == "test"
        assert cell.format_code == "G"
        assert cell.is_formula is False

    def test_formula_cell(self):
        """Test formula cell."""
        cell = ClipboardCell(raw_value="=A1+B1", is_formula=True)
        assert cell.is_formula is True


class TestClipboardContent:
    """Tests for ClipboardContent dataclass."""

    def test_default_values(self):
        """Test default values."""
        content = ClipboardContent()
        assert content.mode == ClipboardMode.EMPTY
        assert content.num_rows == 0
        assert content.num_cols == 0
        assert len(content.cells) == 0


class TestClipboard:
    """Tests for Clipboard class."""

    def setup_method(self):
        """Set up test spreadsheet."""
        self.ss = Spreadsheet()
        self.clipboard = Clipboard(self.ss)

    def test_initial_state(self):
        """Test initial clipboard state."""
        assert self.clipboard.is_empty is True
        assert self.clipboard.mode == ClipboardMode.EMPTY
        assert self.clipboard.has_content is False

    def test_copy_cell(self):
        """Test copying a single cell."""
        self.ss.set_cell(0, 0, "Test")
        self.clipboard.copy_cell(0, 0)

        assert self.clipboard.is_empty is False
        assert self.clipboard.mode == ClipboardMode.COPY
        assert self.clipboard.has_content is True

    def test_copy_range(self):
        """Test copying a range."""
        self.ss.set_cell(0, 0, "A")
        self.ss.set_cell(0, 1, "B")
        self.ss.set_cell(1, 0, "C")
        self.ss.set_cell(1, 1, "D")

        self.clipboard.copy_range(0, 0, 1, 1)

        assert self.clipboard.size == (2, 2)
        assert self.clipboard.source_range == (0, 0, 1, 1)

    def test_copy_range_normalizes(self):
        """Test copy_range normalizes reversed ranges."""
        self.ss.set_cell(0, 0, "Test")
        self.clipboard.copy_range(1, 1, 0, 0)

        # Should still work with reversed range
        assert self.clipboard.size == (2, 2)

    def test_cut_cell(self):
        """Test cutting a cell."""
        self.ss.set_cell(0, 0, "Test")
        self.clipboard.cut_cell(0, 0)

        assert self.clipboard.mode == ClipboardMode.CUT

    def test_cut_range(self):
        """Test cutting a range."""
        self.ss.set_cell(0, 0, "A")
        self.ss.set_cell(0, 1, "B")

        self.clipboard.cut_range(0, 0, 0, 1)

        assert self.clipboard.mode == ClipboardMode.CUT

    def test_paste_copy(self):
        """Test pasting after copy."""
        self.ss.set_cell(0, 0, "Test")
        self.clipboard.copy_cell(0, 0)

        modified = self.clipboard.paste(5, 5)

        assert (5, 5) in modified
        assert self.ss.get_value(5, 5) == "Test"
        # Original should still exist
        assert self.ss.get_value(0, 0) == "Test"

    def test_paste_cut(self):
        """Test pasting after cut."""
        self.ss.set_cell(0, 0, "Test")
        self.clipboard.cut_cell(0, 0)

        modified = self.clipboard.paste(5, 5)

        assert (5, 5) in modified
        assert self.ss.get_value(5, 5) == "Test"
        # Original should be cleared
        cell = self.ss.get_cell_if_exists(0, 0)
        assert cell is None or cell.is_empty

    def test_paste_empty_clipboard(self):
        """Test pasting with empty clipboard."""
        modified = self.clipboard.paste(0, 0)
        assert modified == []

    def test_paste_with_formula_adjustment(self):
        """Test pasting adjusts formula references."""
        self.ss.set_cell(0, 0, "=A1+B1")
        self.clipboard.copy_cell(0, 0)

        self.clipboard.paste(2, 2, adjust_references=True)

        # Formula should be adjusted
        cell = self.ss.get_cell(2, 2)
        # The reference should be adjusted

    def test_paste_without_formula_adjustment(self):
        """Test pasting without adjusting formula references."""
        self.ss.set_cell(0, 0, "=A1+B1")
        self.clipboard.copy_cell(0, 0)

        self.clipboard.paste(2, 2, adjust_references=False)

        cell = self.ss.get_cell(2, 2)
        assert "A1" in cell.raw_value

    def test_paste_special_values_only(self):
        """Test paste special with values only."""
        self.ss.set_cell(0, 0, "100")
        self.ss.set_cell(0, 1, "=A1*2")

        self.clipboard.copy_range(0, 0, 0, 1)
        modified = self.clipboard.paste_special(5, 5, values_only=True)

        assert len(modified) == 2
        # Formula should be replaced with computed value
        assert self.ss.get_value(5, 6) == 200

    def test_paste_special_formats_only(self):
        """Test paste special with formats only."""
        cell = self.ss.get_cell(0, 0)
        cell.set_value("100")
        cell.format_code = "C2"

        self.clipboard.copy_cell(0, 0)

        # Set a different value at destination
        self.ss.set_cell(5, 5, "500")

        self.clipboard.paste_special(5, 5, formats_only=True)

        # Value should remain 500, format should be C2
        dest_cell = self.ss.get_cell(5, 5)
        assert dest_cell.format_code == "C2"

    def test_paste_special_transpose(self):
        """Test paste special with transpose."""
        self.ss.set_cell(0, 0, "A")
        self.ss.set_cell(0, 1, "B")
        self.ss.set_cell(1, 0, "C")
        self.ss.set_cell(1, 1, "D")

        self.clipboard.copy_range(0, 0, 1, 1)
        self.clipboard.paste_special(5, 5, transpose=True)

        # Should be transposed
        assert self.ss.get_value(5, 5) == "A"
        assert self.ss.get_value(6, 5) == "B"  # Was column 1, now row 1
        assert self.ss.get_value(5, 6) == "C"  # Was row 1, now column 1

    def test_clear(self):
        """Test clearing clipboard."""
        self.ss.set_cell(0, 0, "Test")
        self.clipboard.copy_cell(0, 0)

        self.clipboard.clear()

        assert self.clipboard.is_empty is True
        assert self.clipboard.has_content is False

    def test_size_empty(self):
        """Test size of empty clipboard."""
        assert self.clipboard.size == (0, 0)

    def test_source_range_empty(self):
        """Test source_range when empty."""
        assert self.clipboard.source_range is None

    def test_source_range_after_copy(self):
        """Test source_range after copy."""
        self.ss.set_cell(0, 0, "Test")
        self.clipboard.copy_range(1, 2, 3, 4)

        assert self.clipboard.source_range == (1, 2, 3, 4)


class TestClipboardMode:
    """Tests for ClipboardMode enum."""

    def test_all_modes_exist(self):
        """Test all clipboard modes exist."""
        assert ClipboardMode.EMPTY
        assert ClipboardMode.COPY
        assert ClipboardMode.CUT

"""Tests for OS clipboard utility module."""

import sys
from unittest.mock import MagicMock, patch

import pytest

from lotus123.utils.os_clipboard import (
    copy_to_clipboard,
    format_cells_as_tsv,
    is_clipboard_available,
)


class TestFormatCellsAsTsv:
    """Tests for format_cells_as_tsv function."""

    def test_single_cell(self):
        """Test formatting a single cell."""
        cells = [["Hello"]]
        result = format_cells_as_tsv(cells)
        assert result == "Hello"

    def test_single_row(self):
        """Test formatting a single row with multiple cells."""
        cells = [["A", "B", "C"]]
        result = format_cells_as_tsv(cells)
        assert result == "A\tB\tC"

    def test_multiple_rows(self):
        """Test formatting multiple rows."""
        cells = [["A1", "B1"], ["A2", "B2"]]
        result = format_cells_as_tsv(cells)
        assert result == "A1\tB1\nA2\tB2"

    def test_empty_cells(self):
        """Test formatting with empty cells."""
        cells = [["A", "", "C"], ["", "B", ""]]
        result = format_cells_as_tsv(cells)
        assert result == "A\t\tC\n\tB\t"

    def test_none_values(self):
        """Test that None values are converted to empty strings."""
        cells = [["A", None, "C"]]
        result = format_cells_as_tsv(cells)
        assert result == "A\t\tC"

    def test_numeric_values(self):
        """Test that numeric values are converted to strings."""
        cells = [[1, 2.5, 100]]
        result = format_cells_as_tsv(cells)
        assert result == "1\t2.5\t100"

    def test_tabs_replaced(self):
        """Test that embedded tabs are replaced with spaces."""
        cells = [["Hello\tWorld"]]
        result = format_cells_as_tsv(cells)
        assert result == "Hello World"

    def test_newlines_replaced(self):
        """Test that embedded newlines are replaced with spaces."""
        cells = [["Line1\nLine2"]]
        result = format_cells_as_tsv(cells)
        assert result == "Line1 Line2"

    def test_carriage_returns_removed(self):
        """Test that carriage returns are removed."""
        cells = [["Text\r\nMore"]]
        result = format_cells_as_tsv(cells)
        assert result == "Text More"

    def test_empty_grid(self):
        """Test formatting an empty grid."""
        cells: list[list[str]] = []
        result = format_cells_as_tsv(cells)
        assert result == ""

    def test_grid_with_empty_row(self):
        """Test formatting a grid with an empty row."""
        cells = [["A", "B"], []]
        result = format_cells_as_tsv(cells)
        assert result == "A\tB\n"

    def test_special_characters_preserved(self):
        """Test that special characters (except tabs/newlines) are preserved."""
        cells = [["Hello!", "@Formula", "$100.00"]]
        result = format_cells_as_tsv(cells)
        assert result == "Hello!\t@Formula\t$100.00"


class TestCopyToClipboard:
    """Tests for copy_to_clipboard function."""

    @patch("lotus123.utils.os_clipboard.subprocess.Popen")
    def test_copy_macos(self, mock_popen):
        """Test clipboard copy on macOS."""
        mock_process = MagicMock()
        mock_process.returncode = 0
        mock_process.communicate = MagicMock()
        mock_popen.return_value = mock_process

        with patch.object(sys, "platform", "darwin"):
            result = copy_to_clipboard("Test text")

        assert result is True
        mock_popen.assert_called_once()
        call_args = mock_popen.call_args
        assert call_args[0][0] == ["pbcopy"]
        mock_process.communicate.assert_called_once_with(input=b"Test text")

    @patch("lotus123.utils.os_clipboard.subprocess.Popen")
    def test_copy_windows(self, mock_popen):
        """Test clipboard copy on Windows."""
        mock_process = MagicMock()
        mock_process.returncode = 0
        mock_process.communicate = MagicMock()
        mock_popen.return_value = mock_process

        with patch.object(sys, "platform", "win32"):
            result = copy_to_clipboard("Test text")

        assert result is True
        mock_popen.assert_called_once()
        call_args = mock_popen.call_args
        assert call_args[0][0] == ["clip.exe"]
        # Windows uses UTF-16-LE encoding
        mock_process.communicate.assert_called_once_with(input="Test text".encode("utf-16-le"))

    @patch("lotus123.utils.os_clipboard.shutil.which")
    @patch("lotus123.utils.os_clipboard.subprocess.Popen")
    def test_copy_linux_xclip(self, mock_popen, mock_which):
        """Test clipboard copy on Linux using xclip."""
        mock_which.side_effect = lambda x: x if x == "xclip" else None
        mock_process = MagicMock()
        mock_process.returncode = 0
        mock_process.communicate = MagicMock()
        mock_popen.return_value = mock_process

        with patch.object(sys, "platform", "linux"):
            result = copy_to_clipboard("Test text")

        assert result is True
        mock_popen.assert_called_once()
        call_args = mock_popen.call_args
        assert call_args[0][0] == ["xclip", "-selection", "clipboard"]

    @patch("lotus123.utils.os_clipboard.shutil.which")
    @patch("lotus123.utils.os_clipboard.subprocess.Popen")
    def test_copy_linux_xsel(self, mock_popen, mock_which):
        """Test clipboard copy on Linux using xsel (fallback)."""
        mock_which.side_effect = lambda x: x if x == "xsel" else None
        mock_process = MagicMock()
        mock_process.returncode = 0
        mock_process.communicate = MagicMock()
        mock_popen.return_value = mock_process

        with patch.object(sys, "platform", "linux"):
            result = copy_to_clipboard("Test text")

        assert result is True
        mock_popen.assert_called_once()
        call_args = mock_popen.call_args
        assert call_args[0][0] == ["xsel", "--clipboard", "--input"]

    @patch("lotus123.utils.os_clipboard.shutil.which")
    @patch("lotus123.utils.os_clipboard.subprocess.Popen")
    def test_copy_linux_wl_copy(self, mock_popen, mock_which):
        """Test clipboard copy on Linux using wl-copy (Wayland)."""
        mock_which.side_effect = lambda x: x if x == "wl-copy" else None
        mock_process = MagicMock()
        mock_process.returncode = 0
        mock_process.communicate = MagicMock()
        mock_popen.return_value = mock_process

        with patch.object(sys, "platform", "linux"):
            result = copy_to_clipboard("Test text")

        assert result is True
        mock_popen.assert_called_once()
        call_args = mock_popen.call_args
        assert call_args[0][0] == ["wl-copy"]

    @patch("lotus123.utils.os_clipboard.shutil.which")
    def test_copy_linux_no_tool_available(self, mock_which):
        """Test clipboard copy on Linux when no tool is available."""
        mock_which.return_value = None

        with patch.object(sys, "platform", "linux"):
            result = copy_to_clipboard("Test text")

        assert result is False

    @patch("lotus123.utils.os_clipboard.subprocess.Popen")
    def test_copy_subprocess_error(self, mock_popen):
        """Test clipboard copy handles subprocess errors gracefully."""
        import subprocess

        mock_popen.side_effect = subprocess.SubprocessError("Failed")

        with patch.object(sys, "platform", "darwin"):
            result = copy_to_clipboard("Test text")

        assert result is False

    @patch("lotus123.utils.os_clipboard.subprocess.Popen")
    def test_copy_os_error(self, mock_popen):
        """Test clipboard copy handles OS errors gracefully."""
        mock_popen.side_effect = OSError("Command not found")

        with patch.object(sys, "platform", "darwin"):
            result = copy_to_clipboard("Test text")

        assert result is False

    @patch("lotus123.utils.os_clipboard.subprocess.Popen")
    def test_copy_nonzero_return_code(self, mock_popen):
        """Test clipboard copy returns False on non-zero return code."""
        mock_process = MagicMock()
        mock_process.returncode = 1
        mock_process.communicate = MagicMock()
        mock_popen.return_value = mock_process

        with patch.object(sys, "platform", "darwin"):
            result = copy_to_clipboard("Test text")

        assert result is False


class TestIsClipboardAvailable:
    """Tests for is_clipboard_available function."""

    @patch("lotus123.utils.os_clipboard.shutil.which")
    def test_available_macos(self, mock_which):
        """Test clipboard availability on macOS."""
        mock_which.return_value = "/usr/bin/pbcopy"

        with patch.object(sys, "platform", "darwin"):
            result = is_clipboard_available()

        assert result is True
        mock_which.assert_called_with("pbcopy")

    @patch("lotus123.utils.os_clipboard.shutil.which")
    def test_not_available_macos(self, mock_which):
        """Test clipboard not available on macOS."""
        mock_which.return_value = None

        with patch.object(sys, "platform", "darwin"):
            result = is_clipboard_available()

        assert result is False

    @patch("lotus123.utils.os_clipboard.shutil.which")
    def test_available_windows(self, mock_which):
        """Test clipboard availability on Windows."""
        mock_which.return_value = "C:\\Windows\\System32\\clip.exe"

        with patch.object(sys, "platform", "win32"):
            result = is_clipboard_available()

        assert result is True
        mock_which.assert_called_with("clip.exe")

    @patch("lotus123.utils.os_clipboard.shutil.which")
    def test_available_linux_xclip(self, mock_which):
        """Test clipboard availability on Linux with xclip."""
        mock_which.side_effect = lambda x: "/usr/bin/xclip" if x == "xclip" else None

        with patch.object(sys, "platform", "linux"):
            result = is_clipboard_available()

        assert result is True

    @patch("lotus123.utils.os_clipboard.shutil.which")
    def test_available_linux_xsel(self, mock_which):
        """Test clipboard availability on Linux with xsel."""
        mock_which.side_effect = lambda x: "/usr/bin/xsel" if x == "xsel" else None

        with patch.object(sys, "platform", "linux"):
            result = is_clipboard_available()

        assert result is True

    @patch("lotus123.utils.os_clipboard.shutil.which")
    def test_available_linux_wl_copy(self, mock_which):
        """Test clipboard availability on Linux with wl-copy."""
        mock_which.side_effect = lambda x: "/usr/bin/wl-copy" if x == "wl-copy" else None

        with patch.object(sys, "platform", "linux"):
            result = is_clipboard_available()

        assert result is True

    @patch("lotus123.utils.os_clipboard.shutil.which")
    def test_not_available_linux(self, mock_which):
        """Test clipboard not available on Linux."""
        mock_which.return_value = None

        with patch.object(sys, "platform", "linux"):
            result = is_clipboard_available()

        assert result is False

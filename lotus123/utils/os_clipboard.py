"""OS clipboard integration for cross-platform copy/paste support.

Provides functions to copy text to the system clipboard, enabling
data exchange with other applications.
"""

import logging
import shutil
import subprocess
import sys

logger = logging.getLogger(__name__)


def copy_to_clipboard(text: str) -> bool:
    """Copy text to the OS clipboard.

    Uses platform-specific commands:
    - macOS: pbcopy
    - Linux: xclip or xsel
    - Windows: clip.exe

    Args:
        text: Text to copy to clipboard

    Returns:
        True if successful, False otherwise
    """
    if sys.platform == "darwin":
        return _copy_macos(text)
    elif sys.platform == "win32":
        return _copy_windows(text)
    else:
        # Linux and other Unix-like systems
        return _copy_linux(text)


def _copy_macos(text: str) -> bool:
    """Copy text using macOS pbcopy."""
    try:
        process = subprocess.Popen(
            ["pbcopy"],
            stdin=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        process.communicate(input=text.encode("utf-8"))
        return process.returncode == 0
    except (subprocess.SubprocessError, OSError) as e:
        logger.debug(f"Failed to copy to clipboard (macOS): {e}")
        return False


def _copy_windows(text: str) -> bool:
    """Copy text using Windows clip.exe."""
    try:
        process = subprocess.Popen(
            ["clip.exe"],
            stdin=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        process.communicate(input=text.encode("utf-16-le"))
        return process.returncode == 0
    except (subprocess.SubprocessError, OSError) as e:
        logger.debug(f"Failed to copy to clipboard (Windows): {e}")
        return False


def _copy_linux(text: str) -> bool:
    """Copy text using xclip or xsel on Linux."""
    # Try xclip first
    if shutil.which("xclip"):
        try:
            process = subprocess.Popen(
                ["xclip", "-selection", "clipboard"],
                stdin=subprocess.PIPE,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            process.communicate(input=text.encode("utf-8"))
            if process.returncode == 0:
                return True
        except (subprocess.SubprocessError, OSError) as e:
            logger.debug(f"xclip failed: {e}")

    # Try xsel as fallback
    if shutil.which("xsel"):
        try:
            process = subprocess.Popen(
                ["xsel", "--clipboard", "--input"],
                stdin=subprocess.PIPE,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            process.communicate(input=text.encode("utf-8"))
            if process.returncode == 0:
                return True
        except (subprocess.SubprocessError, OSError) as e:
            logger.debug(f"xsel failed: {e}")

    # Try wl-copy for Wayland
    if shutil.which("wl-copy"):
        try:
            process = subprocess.Popen(
                ["wl-copy"],
                stdin=subprocess.PIPE,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            process.communicate(input=text.encode("utf-8"))
            if process.returncode == 0:
                return True
        except (subprocess.SubprocessError, OSError) as e:
            logger.debug(f"wl-copy failed: {e}")

    logger.debug("No clipboard tool available on Linux")
    return False


def is_clipboard_available() -> bool:
    """Check if OS clipboard is available.

    Returns:
        True if clipboard commands are available
    """
    if sys.platform == "darwin":
        return shutil.which("pbcopy") is not None
    elif sys.platform == "win32":
        return shutil.which("clip.exe") is not None
    else:
        return (
            shutil.which("xclip") is not None
            or shutil.which("xsel") is not None
            or shutil.which("wl-copy") is not None
        )


def format_cells_as_tsv(cells: list[list[str]]) -> str:
    """Format a 2D array of cell values as TSV (tab-separated values).

    This format is compatible with Excel, Google Sheets, and other
    spreadsheet applications.

    Args:
        cells: 2D list of cell values

    Returns:
        TSV-formatted string
    """
    lines = []
    for row in cells:
        # Escape tabs and newlines in cell values
        escaped_row = []
        for cell in row:
            if cell is None:
                escaped_row.append("")
            else:
                # Replace tabs with spaces, newlines with space
                escaped = str(cell).replace("\t", " ").replace("\n", " ").replace("\r", "")
                escaped_row.append(escaped)
        lines.append("\t".join(escaped_row))
    return "\n".join(lines)

"""Color themes for the Lotus 1-2-3 clone.

Provides theme definitions including the classic Lotus blue theme,
Tomorrow Night, and Mocha color schemes.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum, auto


class ThemeType(Enum):
    """Available color themes."""

    LOTUS = auto()
    TOMORROW = auto()
    MOCHA = auto()
    NORD = auto()
    GRUVBOX = auto()
    TOKYO_NIGHT = auto()
    TEXTUAL_DARK = auto()


@dataclass
class Theme:
    """Color theme definition."""

    name: str
    background: str
    foreground: str
    header_bg: str
    header_fg: str
    cell_bg: str
    cell_fg: str
    selected_bg: str
    selected_fg: str
    border: str
    menu_bg: str
    menu_fg: str
    menu_highlight: str
    status_bg: str
    status_fg: str
    input_bg: str
    input_fg: str
    accent: str


# Theme definitions
THEMES: dict[ThemeType, Theme] = {
    ThemeType.LOTUS: Theme(
        name="Lotus 1-2-3",
        background="#000080",
        foreground="#ffffff",
        header_bg="#00aaaa",
        header_fg="#000000",
        cell_bg="#000080",
        cell_fg="#ffffff",
        selected_bg="#ffffff",
        selected_fg="#000000",
        border="#00aaaa",
        menu_bg="#00aaaa",
        menu_fg="#000000",
        menu_highlight="#ffffff",
        status_bg="#000080",
        status_fg="#ffffff",
        input_bg="#000080",
        input_fg="#ffffff",
        accent="#00aaaa",
    ),
    ThemeType.TOMORROW: Theme(
        name="Tomorrow Night",
        background="#1d1f21",
        foreground="#c5c8c6",
        header_bg="#373b41",
        header_fg="#c5c8c6",
        cell_bg="#1d1f21",
        cell_fg="#c5c8c6",
        selected_bg="#81a2be",
        selected_fg="#1d1f21",
        border="#373b41",
        menu_bg="#282a2e",
        menu_fg="#c5c8c6",
        menu_highlight="#81a2be",
        status_bg="#282a2e",
        status_fg="#969896",
        input_bg="#282a2e",
        input_fg="#c5c8c6",
        accent="#81a2be",
    ),
    ThemeType.MOCHA: Theme(
        name="Mocha",
        background="#1e1e2e",
        foreground="#cdd6f4",
        header_bg="#313244",
        header_fg="#cdd6f4",
        cell_bg="#1e1e2e",
        cell_fg="#cdd6f4",
        selected_bg="#89b4fa",
        selected_fg="#1e1e2e",
        border="#45475a",
        menu_bg="#313244",
        menu_fg="#cdd6f4",
        menu_highlight="#89b4fa",
        status_bg="#181825",
        status_fg="#a6adc8",
        input_bg="#313244",
        input_fg="#cdd6f4",
        accent="#89b4fa",
    ),
    ThemeType.NORD: Theme(
        name="Nord",
        background="#2e3440",
        foreground="#d8dee9",
        header_bg="#3b4252",
        header_fg="#eceff4",
        cell_bg="#2e3440",
        cell_fg="#d8dee9",
        selected_bg="#88c0d0",
        selected_fg="#2e3440",
        border="#4c566a",
        menu_bg="#3b4252",
        menu_fg="#eceff4",
        menu_highlight="#88c0d0",
        status_bg="#3b4252",
        status_fg="#d8dee9",
        input_bg="#3b4252",
        input_fg="#eceff4",
        accent="#88c0d0",
    ),
    ThemeType.GRUVBOX: Theme(
        name="Gruvbox",
        background="#282828",
        foreground="#ebdbb2",
        header_bg="#3c3836",
        header_fg="#ebdbb2",
        cell_bg="#282828",
        cell_fg="#ebdbb2",
        selected_bg="#fe8019",
        selected_fg="#282828",
        border="#504945",
        menu_bg="#3c3836",
        menu_fg="#ebdbb2",
        menu_highlight="#fe8019",
        status_bg="#3c3836",
        status_fg="#a89984",
        input_bg="#3c3836",
        input_fg="#ebdbb2",
        accent="#fe8019",
    ),
    ThemeType.TOKYO_NIGHT: Theme(
        name="Tokyo Night",
        background="#1a1b26",
        foreground="#a9b1d6",
        header_bg="#16161e",
        header_fg="#7dcfff",
        cell_bg="#1a1b26",
        cell_fg="#a9b1d6",
        selected_bg="#bb9af7",
        selected_fg="#1a1b26",
        border="#3b4261",
        menu_bg="#16161e",
        menu_fg="#7dcfff",
        menu_highlight="#bb9af7",
        status_bg="#16161e",
        status_fg="#565f89",
        input_bg="#24283b",
        input_fg="#c0caf5",
        accent="#bb9af7",
    ),
    ThemeType.TEXTUAL_DARK: Theme(
        name="Textual Dark",
        background="#121212",
        foreground="#e0e0e0",
        header_bg="#1e1e1e",
        header_fg="#ffffff",
        cell_bg="#121212",
        cell_fg="#e0e0e0",
        selected_bg="#0178d4",
        selected_fg="#ffffff",
        border="#333333",
        menu_bg="#1e1e1e",
        menu_fg="#ffffff",
        menu_highlight="#0178d4",
        status_bg="#1e1e1e",
        status_fg="#808080",
        input_bg="#1e1e1e",
        input_fg="#ffffff",
        accent="#0178d4",
    ),
}


def get_theme_type(name: str) -> ThemeType:
    """Get ThemeType from string name."""
    try:
        return ThemeType[name.upper()]
    except KeyError:
        return ThemeType.LOTUS


def get_theme(theme_type: ThemeType) -> Theme:
    """Get Theme instance from ThemeType."""
    return THEMES[theme_type]

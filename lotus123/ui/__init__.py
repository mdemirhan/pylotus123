"""UI components for the Lotus 1-2-3 clone."""

# Theme system
from .themes import Theme, ThemeType, THEMES, get_theme_type, get_theme

# Configuration
from .config import AppConfig, CONFIG_DIR, CONFIG_FILE

# Widgets
from .grid import SpreadsheetGrid
from .menu_bar import LotusMenu
from .status_bar import StatusBar, StatusBarWidget, ModeIndicator, Mode

# Dialogs
from .dialogs import (
    FileDialog,
    CommandInput,
    ThemeDialog,
    ThemeItem,
    ChartViewScreen,
)

# Window management (existing)
from .menu import (
    Menu,
    MenuItem,
    MenuAction,
    MenuState,
    MenuContext,
    MenuSystem,
)
from .window import (
    SplitType,
    TitleFreezeType,
    ViewPort,
    FrozenTitles,
    WindowSplit,
    WindowManager,
)

__all__ = [
    # Themes
    "Theme",
    "ThemeType",
    "THEMES",
    "get_theme_type",
    "get_theme",
    # Config
    "AppConfig",
    "CONFIG_DIR",
    "CONFIG_FILE",
    # Widgets
    "SpreadsheetGrid",
    "LotusMenu",
    "StatusBar",
    "StatusBarWidget",
    "ModeIndicator",
    "Mode",
    # Dialogs
    "FileDialog",
    "CommandInput",
    "ThemeDialog",
    "ThemeItem",
    "ChartViewScreen",
    # Menu system
    "Menu",
    "MenuItem",
    "MenuAction",
    "MenuState",
    "MenuContext",
    "MenuSystem",
    # Window management
    "SplitType",
    "TitleFreezeType",
    "ViewPort",
    "FrozenTitles",
    "WindowSplit",
    "WindowManager",
]

"""UI components for the Lotus 1-2-3 clone."""

# Theme system
# Configuration
from .config import CONFIG_DIR, CONFIG_FILE, AppConfig

# Dialogs
from .dialogs import ChartViewScreen, CommandInput, FileDialog, ThemeDialog, ThemeItem

# Widgets
from .grid import SpreadsheetGrid

# Window management (existing)
from .menu import Menu, MenuAction, MenuContext, MenuItem, MenuState, MenuSystem
from .menu_bar import LotusMenu
from .status_bar import Mode, ModeIndicator, StatusBar, StatusBarWidget
from .themes import THEMES, Theme, ThemeType, get_theme, get_theme_type
from .window import FrozenTitles, SplitType, TitleFreezeType, ViewPort, WindowManager, WindowSplit

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

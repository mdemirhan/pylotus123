
import pytest
from unittest.mock import MagicMock
from textual.app import App, ComposeResult
from lotus123.ui.grid import SpreadsheetGrid
from lotus123.ui.themes import THEMES, ThemeType

class GridApp(App):
    def __init__(self, sheet, theme):
        super().__init__()
        self.sheet = sheet
        self.grid_theme = theme

    def compose(self) -> ComposeResult:
        yield SpreadsheetGrid(self.sheet, self.grid_theme)

@pytest.mark.asyncio
async def test_grid_interaction():
    # Setup mocks
    mock_sheet = MagicMock()
    mock_sheet.rows = 100
    mock_sheet.cols = 26
    mock_sheet.get_col_width.return_value = 9
    mock_sheet.get_display_value.return_value = "Test"
    
    theme = THEMES[ThemeType.LOTUS]
    
    app = GridApp(mock_sheet, theme)
    async with app.run_test() as pilot:
        grid = app.query_one(SpreadsheetGrid)
        
        # Test Initialization
        assert grid.cursor_row == 0
        assert grid.cursor_col == 0
        assert not grid.has_selection
        
        # Test Selection Logic
        grid.start_selection()
        grid.cursor_row = 5
        grid.cursor_col = 5
        # In a real app, reactives update
        # await pilot.pause() # maybe needed?
        
        assert grid.has_selection
        assert grid.selection_range == (0, 0, 5, 5)
        
        # Test Cursor Bounds
        # Move up from (5,5)
        grid.move_cursor(-1, -1)
        assert grid.cursor_row == 4
        
        # Test Top Boundary
        grid.cursor_row = 0
        grid.move_cursor(-1, -1)
        assert grid.cursor_row == 0
        
        grid.cursor_row = 99
        grid.move_cursor(1, 0)
        assert grid.cursor_row == 99
        
        # Test Goto Cell
        grid.goto_cell("C5")
        assert grid.cursor_row == 4
        assert grid.cursor_col == 2

@pytest.mark.asyncio
async def test_grid_rendering():
    # Setup mocks
    mock_sheet = MagicMock()
    mock_sheet.rows = 100
    mock_sheet.cols = 26
    mock_sheet.get_col_width.return_value = 9
    mock_sheet.get_display_value.return_value = "Test"
    
    theme = THEMES[ThemeType.LOTUS]
    app = GridApp(mock_sheet, theme)
    
    async with app.run_test() as pilot:
        grid = app.query_one(SpreadsheetGrid)
        
        # Trigger refresh
        grid.refresh_grid()
        # Verify no crash. We can't easily assert textual render output without snapshot,
        # but execution covers the lines.
        
        # Resize
        app.check_width = 40
        app.check_height = 12
        # On resize logic is internal to Textual, but we can call on_resize manually if needed
        # or rely on app resize events.
        # grid._calculate_visible_area() should run on mount
        assert grid.visible_rows > 0

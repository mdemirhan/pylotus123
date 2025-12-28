"""Tests for window management system."""
import pytest
from lotus123.ui.window import (
    ViewPort, FrozenTitles, WindowSplit, WindowManager,
    SplitType, TitleFreezeType
)


class TestViewPort:
    """Tests for ViewPort class."""

    def test_default_values(self):
        """Test default viewport values."""
        vp = ViewPort()
        assert vp.top_row == 0
        assert vp.left_col == 0
        assert vp.visible_rows == 20
        assert vp.visible_cols == 10
        assert vp.cursor_row == 0
        assert vp.cursor_col == 0

    def test_scroll_to_visible(self):
        """Test scroll_to when cell is already visible."""
        vp = ViewPort()
        vp.scroll_to(5, 5)
        assert vp.top_row == 0
        assert vp.left_col == 0

    def test_scroll_to_below(self):
        """Test scroll_to when cell is below visible area."""
        vp = ViewPort(visible_rows=10)
        vp.scroll_to(15, 0)
        assert vp.top_row == 6  # 15 - 10 + 1

    def test_scroll_to_above(self):
        """Test scroll_to when cell is above visible area."""
        vp = ViewPort(top_row=10)
        vp.scroll_to(5, 0)
        assert vp.top_row == 5

    def test_scroll_to_right(self):
        """Test scroll_to when cell is to the right."""
        vp = ViewPort(visible_cols=5)
        vp.scroll_to(0, 8)
        assert vp.left_col == 4  # 8 - 5 + 1

    def test_scroll_to_left(self):
        """Test scroll_to when cell is to the left."""
        vp = ViewPort(left_col=5)
        vp.scroll_to(0, 2)
        assert vp.left_col == 2

    def test_move_cursor(self):
        """Test move_cursor updates cursor and scrolls."""
        vp = ViewPort()
        vp.move_cursor(5, 3)
        assert vp.cursor_row == 5
        assert vp.cursor_col == 3

    def test_is_visible_true(self):
        """Test is_visible returns True for visible cell."""
        vp = ViewPort(visible_rows=10, visible_cols=5)
        assert vp.is_visible(5, 3) is True

    def test_is_visible_false(self):
        """Test is_visible returns False for hidden cell."""
        vp = ViewPort(visible_rows=10, visible_cols=5)
        assert vp.is_visible(15, 3) is False
        assert vp.is_visible(5, 8) is False

    def test_get_visible_range(self):
        """Test get_visible_range returns correct bounds."""
        vp = ViewPort(top_row=5, left_col=2, visible_rows=10, visible_cols=5)
        start_row, end_row, start_col, end_col = vp.get_visible_range()
        assert start_row == 5
        assert end_row == 14
        assert start_col == 2
        assert end_col == 6


class TestFrozenTitles:
    """Tests for FrozenTitles class."""

    def test_default_state(self):
        """Test default frozen titles state."""
        ft = FrozenTitles()
        assert ft.freeze_type == TitleFreezeType.NONE
        assert ft.frozen_rows == 0
        assert ft.frozen_cols == 0

    def test_freeze_horizontal(self):
        """Test freezing rows."""
        ft = FrozenTitles()
        ft.freeze_horizontal(5)
        assert ft.freeze_type == TitleFreezeType.HORIZONTAL
        assert ft.frozen_rows == 5
        assert ft.freeze_row == 5

    def test_freeze_vertical(self):
        """Test freezing columns."""
        ft = FrozenTitles()
        ft.freeze_vertical(3)
        assert ft.freeze_type == TitleFreezeType.VERTICAL
        assert ft.frozen_cols == 3
        assert ft.freeze_col == 3

    def test_freeze_both(self):
        """Test freezing both rows and columns."""
        ft = FrozenTitles()
        ft.freeze_both(5, 3)
        assert ft.freeze_type == TitleFreezeType.BOTH
        assert ft.frozen_rows == 5
        assert ft.frozen_cols == 3

    def test_freeze_horizontal_then_vertical(self):
        """Test freezing horizontal then vertical."""
        ft = FrozenTitles()
        ft.freeze_horizontal(5)
        ft.freeze_vertical(3)
        assert ft.freeze_type == TitleFreezeType.BOTH

    def test_clear(self):
        """Test clearing frozen titles."""
        ft = FrozenTitles()
        ft.freeze_both(5, 3)
        ft.clear()
        assert ft.freeze_type == TitleFreezeType.NONE
        assert ft.frozen_rows == 0
        assert ft.frozen_cols == 0

    def test_has_frozen_rows(self):
        """Test has_frozen_rows property."""
        ft = FrozenTitles()
        assert ft.has_frozen_rows is False
        ft.freeze_horizontal(5)
        assert ft.has_frozen_rows is True

    def test_has_frozen_cols(self):
        """Test has_frozen_cols property."""
        ft = FrozenTitles()
        assert ft.has_frozen_cols is False
        ft.freeze_vertical(3)
        assert ft.has_frozen_cols is True


class TestWindowSplit:
    """Tests for WindowSplit class."""

    def test_default_state(self):
        """Test default window split state."""
        ws = WindowSplit()
        assert ws.split_type == SplitType.NONE
        assert ws.split_position == 0
        assert ws.synchronized is True

    def test_split_horizontal(self):
        """Test horizontal split."""
        ws = WindowSplit()
        ws.split_horizontal(10)
        assert ws.split_type == SplitType.HORIZONTAL
        assert ws.split_position == 10

    def test_split_vertical(self):
        """Test vertical split."""
        ws = WindowSplit()
        ws.split_vertical(5)
        assert ws.split_type == SplitType.VERTICAL
        assert ws.split_position == 5

    def test_clear_split(self):
        """Test clearing split."""
        ws = WindowSplit()
        ws.split_horizontal(10)
        ws.clear()
        assert ws.split_type == SplitType.NONE
        assert ws.split_position == 0

    def test_is_split(self):
        """Test is_split property."""
        ws = WindowSplit()
        assert ws.is_split is False
        ws.split_horizontal(10)
        assert ws.is_split is True


class TestWindowManager:
    """Tests for WindowManager class."""

    def test_default_state(self):
        """Test default window manager state."""
        wm = WindowManager()
        assert wm.active_pane == 0
        assert wm.titles.freeze_type == TitleFreezeType.NONE
        assert wm.split.split_type == SplitType.NONE

    def test_freeze_titles_horizontal(self):
        """Test freezing horizontal titles."""
        wm = WindowManager()
        wm.freeze_titles_horizontal(5)
        assert wm.titles.frozen_rows == 5

    def test_freeze_titles_vertical(self):
        """Test freezing vertical titles."""
        wm = WindowManager()
        wm.freeze_titles_vertical(3)
        assert wm.titles.frozen_cols == 3

    def test_freeze_titles_both(self):
        """Test freezing both titles."""
        wm = WindowManager()
        wm.freeze_titles_both(5, 3)
        assert wm.titles.frozen_rows == 5
        assert wm.titles.frozen_cols == 3

    def test_clear_titles(self):
        """Test clearing titles."""
        wm = WindowManager()
        wm.freeze_titles_both(5, 3)
        wm.clear_titles()
        assert wm.titles.freeze_type == TitleFreezeType.NONE

    def test_split_horizontal(self):
        """Test horizontal window split."""
        wm = WindowManager()
        wm.split_horizontal(10)
        assert wm.split.split_type == SplitType.HORIZONTAL
        assert wm.secondary.top_row == 10
        assert wm.secondary.cursor_row == 10

    def test_split_vertical(self):
        """Test vertical window split."""
        wm = WindowManager()
        wm.split_vertical(5)
        assert wm.split.split_type == SplitType.VERTICAL
        assert wm.secondary.left_col == 5
        assert wm.secondary.cursor_col == 5

    def test_clear_split(self):
        """Test clearing split."""
        wm = WindowManager()
        wm.split_horizontal(10)
        wm.clear_split()
        assert wm.split.split_type == SplitType.NONE
        assert wm.active_pane == 0

    def test_switch_pane_with_split(self):
        """Test switching panes when split."""
        wm = WindowManager()
        wm.split_horizontal(10)
        assert wm.active_pane == 0
        wm.switch_pane()
        assert wm.active_pane == 1
        wm.switch_pane()
        assert wm.active_pane == 0

    def test_switch_pane_without_split(self):
        """Test switching panes without split does nothing."""
        wm = WindowManager()
        assert wm.active_pane == 0
        wm.switch_pane()
        assert wm.active_pane == 0

    def test_active_viewport(self):
        """Test active_viewport property."""
        wm = WindowManager()
        assert wm.active_viewport == wm.primary

        wm.split_horizontal(10)
        wm.switch_pane()
        assert wm.active_viewport == wm.secondary

    def test_scroll_without_split(self):
        """Test scrolling without split."""
        wm = WindowManager()
        wm.scroll(delta_row=5, delta_col=3)
        assert wm.primary.top_row == 5
        assert wm.primary.left_col == 3

    def test_scroll_synchronized_horizontal(self):
        """Test synchronized scroll with horizontal split."""
        wm = WindowManager()
        wm.split_horizontal(10)
        wm.scroll(delta_col=5)
        # Both panes should scroll horizontally
        assert wm.primary.left_col == 5
        assert wm.secondary.left_col == 5

    def test_scroll_synchronized_vertical(self):
        """Test synchronized scroll with vertical split."""
        wm = WindowManager()
        wm.split_vertical(5)
        wm.scroll(delta_row=5)
        # Both panes should scroll vertically
        assert wm.primary.top_row == 5
        assert wm.secondary.top_row == 5

    def test_scroll_unsynchronized(self):
        """Test unsynchronized scrolling."""
        wm = WindowManager()
        wm.split_horizontal(10)
        wm.unsync_scrolling()
        wm.scroll(delta_col=5)
        # Only active pane should scroll
        assert wm.primary.left_col == 5
        assert wm.secondary.left_col == 0

    def test_move_cursor(self):
        """Test cursor movement."""
        wm = WindowManager()
        wm.move_cursor(5, 3)
        assert wm.primary.cursor_row == 5
        assert wm.primary.cursor_col == 3

    def test_move_cursor_respects_frozen_rows(self):
        """Test cursor can't enter frozen row area."""
        wm = WindowManager()
        wm.freeze_titles_horizontal(5)
        wm.move_cursor(2, 3)
        assert wm.primary.cursor_row == 5  # Clamped to freeze boundary

    def test_move_cursor_respects_frozen_cols(self):
        """Test cursor can't enter frozen column area."""
        wm = WindowManager()
        wm.freeze_titles_vertical(3)
        wm.move_cursor(5, 1)
        assert wm.primary.cursor_col == 3  # Clamped to freeze boundary

    def test_get_cursor_position(self):
        """Test getting cursor position."""
        wm = WindowManager()
        wm.move_cursor(5, 3)
        row, col = wm.get_cursor_position()
        assert row == 5
        assert col == 3

    def test_get_pane_count(self):
        """Test getting pane count."""
        wm = WindowManager()
        assert wm.get_pane_count() == 1
        wm.split_horizontal(10)
        assert wm.get_pane_count() == 2

    def test_get_visible_regions_no_freeze_no_split(self):
        """Test visible regions without freeze or split."""
        wm = WindowManager()
        regions = wm.get_visible_regions()
        assert len(regions) == 1
        assert regions[0]['type'] == 'main'

    def test_get_visible_regions_with_frozen_rows(self):
        """Test visible regions with frozen rows."""
        wm = WindowManager()
        wm.freeze_titles_horizontal(5)
        regions = wm.get_visible_regions()
        types = [r['type'] for r in regions]
        assert 'frozen_row' in types
        assert 'main' in types

    def test_get_visible_regions_with_frozen_cols(self):
        """Test visible regions with frozen columns."""
        wm = WindowManager()
        wm.freeze_titles_vertical(3)
        regions = wm.get_visible_regions()
        types = [r['type'] for r in regions]
        assert 'frozen_col' in types
        assert 'main' in types

    def test_get_visible_regions_with_both_frozen(self):
        """Test visible regions with both frozen."""
        wm = WindowManager()
        wm.freeze_titles_both(5, 3)
        regions = wm.get_visible_regions()
        types = [r['type'] for r in regions]
        assert 'frozen_corner' in types
        assert 'frozen_row' in types
        assert 'frozen_col' in types
        assert 'main' in types

    def test_get_visible_regions_with_split(self):
        """Test visible regions with split."""
        wm = WindowManager()
        wm.split_horizontal(10)
        regions = wm.get_visible_regions()
        types = [r['type'] for r in regions]
        assert 'main' in types
        assert 'secondary' in types

    def test_resize(self):
        """Test window resize."""
        wm = WindowManager()
        wm.resize(30, 15)
        assert wm.primary.visible_rows == 30
        assert wm.primary.visible_cols == 15

    def test_resize_with_horizontal_split(self):
        """Test resize with horizontal split."""
        wm = WindowManager()
        wm.split_horizontal(10)
        wm.resize(30, 15)
        assert wm.primary.visible_rows == 10
        assert wm.secondary.visible_rows == 20
        assert wm.primary.visible_cols == 15
        assert wm.secondary.visible_cols == 15

    def test_resize_with_vertical_split(self):
        """Test resize with vertical split."""
        wm = WindowManager()
        wm.split_vertical(5)
        wm.resize(30, 15)
        assert wm.primary.visible_rows == 30
        assert wm.secondary.visible_rows == 30
        assert wm.primary.visible_cols == 5
        assert wm.secondary.visible_cols == 10

    def test_get_status_empty(self):
        """Test status string when nothing frozen or split."""
        wm = WindowManager()
        assert wm.get_status() == ""

    def test_get_status_with_titles(self):
        """Test status string with frozen titles."""
        wm = WindowManager()
        wm.freeze_titles_horizontal(5)
        status = wm.get_status()
        assert "Titles:5R" in status

    def test_get_status_with_split(self):
        """Test status string with split."""
        wm = WindowManager()
        wm.split_horizontal(10)
        status = wm.get_status()
        assert "Split:H@10" in status

    def test_get_status_unsync(self):
        """Test status shows unsync."""
        wm = WindowManager()
        wm.split_horizontal(10)
        wm.unsync_scrolling()
        status = wm.get_status()
        assert "Unsync" in status

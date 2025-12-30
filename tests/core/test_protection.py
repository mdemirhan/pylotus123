"""Tests for protection module."""

import pytest

from lotus123 import Spreadsheet
from lotus123.core.protection import ProtectionManager, ProtectionSettings


class TestProtectionSettings:
    """Tests for ProtectionSettings dataclass."""

    def test_default_values(self):
        """Test default values."""
        settings = ProtectionSettings()
        assert settings.enabled is False
        assert settings.password_hash == ""
        assert settings.allow_formatting is False
        assert settings.allow_insert_rows is False
        assert settings.allow_insert_cols is False
        assert settings.allow_delete_rows is False
        assert settings.allow_delete_cols is False
        assert settings.allow_sort is False

    def test_to_dict(self):
        """Test serialization to dict."""
        settings = ProtectionSettings(enabled=True, allow_sort=True)
        data = settings.to_dict()
        assert data["enabled"] is True
        assert data["allow_sort"] is True

    def test_from_dict(self):
        """Test deserialization from dict."""
        data = {
            "enabled": True,
            "password_hash": "abc123",
            "allow_formatting": True
        }
        settings = ProtectionSettings.from_dict(data)
        assert settings.enabled is True
        assert settings.password_hash == "abc123"
        assert settings.allow_formatting is True

    def test_from_dict_defaults(self):
        """Test from_dict with missing keys uses defaults."""
        settings = ProtectionSettings.from_dict({})
        assert settings.enabled is False
        assert settings.password_hash == ""


class TestProtectionManager:
    """Tests for ProtectionManager class."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.manager = ProtectionManager(self.ss)

    def test_initial_state(self):
        """Test initial manager state."""
        assert self.manager.is_enabled is False
        assert self.manager.settings.enabled is False

    def test_enable_without_password(self):
        """Test enabling protection without password."""
        self.manager.enable()
        assert self.manager.is_enabled is True

    def test_enable_with_password(self):
        """Test enabling protection with password."""
        self.manager.enable("secret")
        assert self.manager.is_enabled is True
        assert self.manager.settings.password_hash != ""

    def test_disable_without_password(self):
        """Test disabling protection without password."""
        self.manager.enable()
        result = self.manager.disable()
        assert result is True
        assert self.manager.is_enabled is False

    def test_disable_with_correct_password(self):
        """Test disabling with correct password."""
        self.manager.enable("secret")
        result = self.manager.disable("secret")
        assert result is True
        assert self.manager.is_enabled is False

    def test_disable_with_wrong_password(self):
        """Test disabling with wrong password fails."""
        self.manager.enable("secret")
        result = self.manager.disable("wrong")
        assert result is False
        assert self.manager.is_enabled is True


class TestCellProtection:
    """Tests for cell protection methods."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.manager = ProtectionManager(self.ss)

    def test_protect_cell(self):
        """Test protecting a cell."""
        self.manager.unprotect_cell(0, 0)
        self.manager.protect_cell(0, 0)
        assert (0, 0) not in self.manager.get_unprotected_cells()

    def test_unprotect_cell(self):
        """Test unprotecting a cell."""
        self.manager.unprotect_cell(0, 0)
        assert (0, 0) in self.manager.get_unprotected_cells()

    def test_is_cell_protected_when_disabled(self):
        """Test is_cell_protected returns False when protection disabled."""
        assert self.manager.is_cell_protected(0, 0) is False

    def test_is_cell_protected_when_enabled(self):
        """Test is_cell_protected when protection enabled."""
        self.manager.enable()
        # All cells protected by default
        assert self.manager.is_cell_protected(0, 0) is True

    def test_is_cell_protected_unprotected_cell(self):
        """Test unprotected cell is not protected."""
        self.manager.enable()
        self.manager.unprotect_cell(0, 0)
        assert self.manager.is_cell_protected(0, 0) is False

    def test_can_edit_cell(self):
        """Test can_edit_cell method."""
        self.manager.enable()
        self.manager.unprotect_cell(0, 0)
        assert self.manager.can_edit_cell(0, 0) is True
        assert self.manager.can_edit_cell(0, 1) is False

    def test_protect_range(self):
        """Test protecting a range."""
        self.manager.unprotect_range(0, 0, 2, 2)
        self.manager.protect_range(0, 0, 1, 1)
        assert (0, 0) not in self.manager.get_unprotected_cells()
        assert (2, 2) in self.manager.get_unprotected_cells()

    def test_unprotect_range(self):
        """Test unprotecting a range."""
        self.manager.unprotect_range(0, 0, 2, 2)
        unprotected = self.manager.get_unprotected_cells()
        assert (0, 0) in unprotected
        assert (2, 2) in unprotected


class TestPermissions:
    """Tests for permission methods."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.manager = ProtectionManager(self.ss)

    def test_can_insert_row_when_disabled(self):
        """Test can_insert_row when protection disabled."""
        assert self.manager.can_insert_row() is True

    def test_can_insert_row_when_enabled(self):
        """Test can_insert_row when protection enabled."""
        self.manager.enable()
        assert self.manager.can_insert_row() is False

    def test_can_insert_row_when_allowed(self):
        """Test can_insert_row when explicitly allowed."""
        self.manager.enable()
        self.manager.settings.allow_insert_rows = True
        assert self.manager.can_insert_row() is True

    def test_can_insert_col_when_disabled(self):
        """Test can_insert_col when protection disabled."""
        assert self.manager.can_insert_col() is True

    def test_can_insert_col_when_enabled(self):
        """Test can_insert_col when protection enabled."""
        self.manager.enable()
        assert self.manager.can_insert_col() is False

    def test_can_insert_col_when_allowed(self):
        """Test can_insert_col when explicitly allowed."""
        self.manager.enable()
        self.manager.settings.allow_insert_cols = True
        assert self.manager.can_insert_col() is True

    def test_can_delete_row_when_disabled(self):
        """Test can_delete_row when protection disabled."""
        assert self.manager.can_delete_row() is True

    def test_can_delete_row_when_enabled(self):
        """Test can_delete_row when protection enabled."""
        self.manager.enable()
        assert self.manager.can_delete_row() is False

    def test_can_delete_row_when_allowed(self):
        """Test can_delete_row when explicitly allowed."""
        self.manager.enable()
        self.manager.settings.allow_delete_rows = True
        assert self.manager.can_delete_row() is True

    def test_can_delete_col_when_disabled(self):
        """Test can_delete_col when protection disabled."""
        assert self.manager.can_delete_col() is True

    def test_can_delete_col_when_enabled(self):
        """Test can_delete_col when protection enabled."""
        self.manager.enable()
        assert self.manager.can_delete_col() is False

    def test_can_delete_col_when_allowed(self):
        """Test can_delete_col when explicitly allowed."""
        self.manager.enable()
        self.manager.settings.allow_delete_cols = True
        assert self.manager.can_delete_col() is True

    def test_can_sort_when_disabled(self):
        """Test can_sort when protection disabled."""
        assert self.manager.can_sort() is True

    def test_can_sort_when_enabled(self):
        """Test can_sort when protection enabled."""
        self.manager.enable()
        assert self.manager.can_sort() is False

    def test_can_sort_when_allowed(self):
        """Test can_sort when explicitly allowed."""
        self.manager.enable()
        self.manager.settings.allow_sort = True
        assert self.manager.can_sort() is True


class TestInputCells:
    """Tests for input cell methods."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.manager = ProtectionManager(self.ss)

    def test_get_input_cells_empty(self):
        """Test get_input_cells with no unprotected cells."""
        cells = self.manager.get_input_cells()
        assert cells == []

    def test_get_input_cells_sorted(self):
        """Test get_input_cells returns sorted list."""
        self.manager.unprotect_cell(2, 1)
        self.manager.unprotect_cell(0, 0)
        self.manager.unprotect_cell(1, 2)
        cells = self.manager.get_input_cells()
        assert cells == [(0, 0), (1, 2), (2, 1)]

    def test_next_input_cell_empty(self):
        """Test next_input_cell with no unprotected cells."""
        result = self.manager.next_input_cell(0, 0)
        assert result is None

    def test_next_input_cell_basic(self):
        """Test next_input_cell finds next cell."""
        self.manager.unprotect_cell(0, 0)
        self.manager.unprotect_cell(0, 2)
        self.manager.unprotect_cell(1, 0)
        result = self.manager.next_input_cell(0, 0)
        assert result == (0, 2)

    def test_next_input_cell_wraps(self):
        """Test next_input_cell wraps to beginning."""
        self.manager.unprotect_cell(0, 0)
        self.manager.unprotect_cell(1, 0)
        result = self.manager.next_input_cell(1, 0)
        assert result == (0, 0)


class TestAdjustments:
    """Tests for row/column adjustment methods."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.manager = ProtectionManager(self.ss)
        # Set up some unprotected cells
        self.manager.unprotect_cell(1, 1)
        self.manager.unprotect_cell(2, 2)
        self.manager.unprotect_cell(3, 1)

    def test_adjust_for_insert_row(self):
        """Test adjusting for row insertion."""
        self.manager.adjust_for_insert_row(2)
        cells = self.manager.get_unprotected_cells()
        assert (1, 1) in cells  # Before insert, unchanged
        assert (3, 2) in cells  # Was (2, 2), shifted down
        assert (4, 1) in cells  # Was (3, 1), shifted down

    def test_adjust_for_delete_row(self):
        """Test adjusting for row deletion."""
        self.manager.adjust_for_delete_row(2)
        cells = self.manager.get_unprotected_cells()
        assert (1, 1) in cells  # Before delete, unchanged
        assert (2, 2) not in cells  # Was deleted
        assert (2, 1) in cells  # Was (3, 1), shifted up

    def test_adjust_for_insert_col(self):
        """Test adjusting for column insertion."""
        self.manager.adjust_for_insert_col(2)
        cells = self.manager.get_unprotected_cells()
        assert (1, 1) in cells  # Before insert, unchanged
        assert (2, 3) in cells  # Was (2, 2), shifted right
        assert (3, 1) in cells  # Col before insert, unchanged

    def test_adjust_for_delete_col(self):
        """Test adjusting for column deletion."""
        self.manager.adjust_for_delete_col(2)
        cells = self.manager.get_unprotected_cells()
        assert (1, 1) in cells  # Before delete, unchanged
        assert (2, 2) not in cells  # Was deleted
        assert (3, 1) in cells  # Col before delete, unchanged


class TestSerialization:
    """Tests for serialization methods."""

    def setup_method(self):
        """Set up test fixtures."""
        self.ss = Spreadsheet()
        self.manager = ProtectionManager(self.ss)

    def test_to_dict(self):
        """Test serialization to dict."""
        self.manager.enable("secret")
        self.manager.unprotect_cell(0, 0)
        self.manager.unprotect_cell(1, 1)

        data = self.manager.to_dict()
        assert data["settings"]["enabled"] is True
        assert len(data["unprotected_cells"]) == 2

    def test_from_dict(self):
        """Test deserialization from dict."""
        data = {
            "settings": {"enabled": True, "allow_sort": True},
            "unprotected_cells": [[0, 0], [1, 1]]
        }
        self.manager.from_dict(data)

        assert self.manager.is_enabled is True
        assert self.manager.settings.allow_sort is True
        assert (0, 0) in self.manager.get_unprotected_cells()
        assert (1, 1) in self.manager.get_unprotected_cells()

    def test_clear(self):
        """Test clearing protection."""
        self.manager.enable()
        self.manager.unprotect_cell(0, 0)

        self.manager.clear()

        assert self.manager.is_enabled is False
        assert len(self.manager.get_unprotected_cells()) == 0

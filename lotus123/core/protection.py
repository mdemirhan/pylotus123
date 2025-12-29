"""Cell and worksheet protection management."""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .spreadsheet import Spreadsheet


@dataclass
class ProtectionSettings:
    """Worksheet-level protection settings.

    Attributes:
        enabled: Whether protection is active
        password_hash: Hashed password (if any)
        allow_formatting: Allow format changes on protected cells
        allow_insert_rows: Allow inserting rows
        allow_insert_cols: Allow inserting columns
        allow_delete_rows: Allow deleting rows
        allow_delete_cols: Allow deleting columns
        allow_sort: Allow sorting ranges
    """

    enabled: bool = False
    password_hash: str = ""
    allow_formatting: bool = False
    allow_insert_rows: bool = False
    allow_insert_cols: bool = False
    allow_delete_rows: bool = False
    allow_delete_cols: bool = False
    allow_sort: bool = False

    def to_dict(self) -> dict:
        """Serialize to dictionary."""
        return {
            "enabled": self.enabled,
            "password_hash": self.password_hash,
            "allow_formatting": self.allow_formatting,
            "allow_insert_rows": self.allow_insert_rows,
            "allow_insert_cols": self.allow_insert_cols,
            "allow_delete_rows": self.allow_delete_rows,
            "allow_delete_cols": self.allow_delete_cols,
            "allow_sort": self.allow_sort,
        }

    @classmethod
    def from_dict(cls, data: dict) -> ProtectionSettings:
        """Deserialize from dictionary."""
        return cls(
            enabled=data.get("enabled", False),
            password_hash=data.get("password_hash", ""),
            allow_formatting=data.get("allow_formatting", False),
            allow_insert_rows=data.get("allow_insert_rows", False),
            allow_insert_cols=data.get("allow_insert_cols", False),
            allow_delete_rows=data.get("allow_delete_rows", False),
            allow_delete_cols=data.get("allow_delete_cols", False),
            allow_sort=data.get("allow_sort", False),
        )


class ProtectionManager:
    """Manages cell and worksheet protection.

    In Lotus 1-2-3 style:
    - By default, all cells are protected when protection is enabled
    - Users explicitly unprotect cells they want to allow editing
    - The /Range Input command allows entry only in unprotected cells
    """

    def __init__(self, spreadsheet: Spreadsheet) -> None:
        self._spreadsheet = spreadsheet
        self._settings = ProtectionSettings()
        self._unprotected_cells: set[tuple[int, int]] = set()

    @property
    def is_enabled(self) -> bool:
        """Check if worksheet protection is enabled."""
        return self._settings.enabled

    @property
    def settings(self) -> ProtectionSettings:
        """Get protection settings."""
        return self._settings

    def enable(self, password: str = "") -> None:
        """Enable worksheet protection.

        Args:
            password: Optional password (will be hashed)
        """
        self._settings.enabled = True
        if password:
            self._settings.password_hash = self._hash_password(password)

    def disable(self, password: str = "") -> bool:
        """Disable worksheet protection.

        Args:
            password: Password if required

        Returns:
            True if disabled, False if password wrong
        """
        if self._settings.password_hash:
            if self._hash_password(password) != self._settings.password_hash:
                return False
        self._settings.enabled = False
        return True

    def protect_cell(self, row: int, col: int) -> None:
        """Mark a cell as protected (remove from unprotected set)."""
        self._unprotected_cells.discard((row, col))

    def unprotect_cell(self, row: int, col: int) -> None:
        """Mark a cell as unprotected (allow editing when protection enabled)."""
        self._unprotected_cells.add((row, col))

    def protect_range(self, start_row: int, start_col: int, end_row: int, end_col: int) -> None:
        """Mark a range of cells as protected."""
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                self.protect_cell(r, c)

    def unprotect_range(self, start_row: int, start_col: int, end_row: int, end_col: int) -> None:
        """Mark a range of cells as unprotected."""
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                self.unprotect_cell(r, c)

    def is_cell_protected(self, row: int, col: int) -> bool:
        """Check if a specific cell is protected.

        A cell is protected if:
        - Protection is enabled AND
        - Cell is not in the unprotected set
        """
        if not self._settings.enabled:
            return False
        return (row, col) not in self._unprotected_cells

    def can_edit_cell(self, row: int, col: int) -> bool:
        """Check if a cell can be edited.

        This is the inverse of is_cell_protected.
        """
        return not self.is_cell_protected(row, col)

    def can_insert_row(self) -> bool:
        """Check if row insertion is allowed."""
        if not self._settings.enabled:
            return True
        return self._settings.allow_insert_rows

    def can_insert_col(self) -> bool:
        """Check if column insertion is allowed."""
        if not self._settings.enabled:
            return True
        return self._settings.allow_insert_cols

    def can_delete_row(self) -> bool:
        """Check if row deletion is allowed."""
        if not self._settings.enabled:
            return True
        return self._settings.allow_delete_rows

    def can_delete_col(self) -> bool:
        """Check if column deletion is allowed."""
        if not self._settings.enabled:
            return True
        return self._settings.allow_delete_cols

    def can_sort(self) -> bool:
        """Check if sorting is allowed."""
        if not self._settings.enabled:
            return True
        return self._settings.allow_sort

    def get_unprotected_cells(self) -> set[tuple[int, int]]:
        """Get all unprotected cell coordinates."""
        return self._unprotected_cells.copy()

    def get_input_cells(self) -> list[tuple[int, int]]:
        """Get list of cells available for input (unprotected).

        Returns cells in reading order (left-to-right, top-to-bottom).
        """
        return sorted(self._unprotected_cells, key=lambda x: (x[0], x[1]))

    def next_input_cell(self, row: int, col: int) -> tuple[int, int] | None:
        """Get next unprotected cell after the given position.

        Args:
            row: Current row
            col: Current column

        Returns:
            Next unprotected cell or None if at end
        """
        cells = self.get_input_cells()
        if not cells:
            return None

        # Find first cell after current position
        for r, c in cells:
            if r > row or (r == row and c > col):
                return r, c

        # Wrap to beginning
        return cells[0] if cells else None

    def adjust_for_insert_row(self, at_row: int) -> None:
        """Adjust unprotected cells when a row is inserted."""
        new_set = set()
        for r, c in self._unprotected_cells:
            if r >= at_row:
                new_set.add((r + 1, c))
            else:
                new_set.add((r, c))
        self._unprotected_cells = new_set

    def adjust_for_delete_row(self, at_row: int) -> None:
        """Adjust unprotected cells when a row is deleted."""
        new_set = set()
        for r, c in self._unprotected_cells:
            if r == at_row:
                continue  # Cell is deleted
            elif r > at_row:
                new_set.add((r - 1, c))
            else:
                new_set.add((r, c))
        self._unprotected_cells = new_set

    def adjust_for_insert_col(self, at_col: int) -> None:
        """Adjust unprotected cells when a column is inserted."""
        new_set = set()
        for r, c in self._unprotected_cells:
            if c >= at_col:
                new_set.add((r, c + 1))
            else:
                new_set.add((r, c))
        self._unprotected_cells = new_set

    def adjust_for_delete_col(self, at_col: int) -> None:
        """Adjust unprotected cells when a column is deleted."""
        new_set = set()
        for r, c in self._unprotected_cells:
            if c == at_col:
                continue  # Cell is deleted
            elif c > at_col:
                new_set.add((r, c - 1))
            else:
                new_set.add((r, c))
        self._unprotected_cells = new_set

    def clear(self) -> None:
        """Clear all protection settings and unprotected cells."""
        self._settings = ProtectionSettings()
        self._unprotected_cells.clear()

    @staticmethod
    def _hash_password(password: str) -> str:
        """Hash a password for storage.

        Note: This is a simple hash for demo purposes.
        In production, use proper password hashing.
        """
        import hashlib

        return hashlib.sha256(password.encode()).hexdigest()

    def to_dict(self) -> dict:
        """Serialize to dictionary."""
        return {
            "settings": self._settings.to_dict(),
            "unprotected_cells": list(self._unprotected_cells),
        }

    def from_dict(self, data: dict) -> None:
        """Load from dictionary."""
        self._settings = ProtectionSettings.from_dict(data.get("settings", {}))
        self._unprotected_cells = set(tuple(cell) for cell in data.get("unprotected_cells", []))

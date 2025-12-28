"""Named range management for the spreadsheet."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Iterator

from .reference import RangeReference, CellReference


# Valid name pattern: starts with letter, can contain letters, numbers, underscores
NAME_PATTERN = re.compile(r'^[A-Za-z][A-Za-z0-9_]*$')


@dataclass
class NamedRange:
    """A named reference to a cell or range.

    Attributes:
        name: The name of the range (case-insensitive)
        reference: The range or cell reference
        description: Optional description/comment
    """
    name: str
    reference: RangeReference | CellReference
    description: str = ""

    def __post_init__(self) -> None:
        """Normalize the name to uppercase."""
        self.name = self.name.upper()

    @property
    def is_single_cell(self) -> bool:
        """Check if this name refers to a single cell."""
        return isinstance(self.reference, CellReference)

    def to_dict(self) -> dict[str, Any]:
        """Serialize to dictionary."""
        data: dict[str, Any] = {
            "name": self.name,
            "reference": self.reference.to_string(),
            "is_range": isinstance(self.reference, RangeReference),
        }
        if self.description:
            data["description"] = self.description
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> NamedRange:
        """Deserialize from dictionary."""
        ref_str = data["reference"]
        reference: CellReference | RangeReference
        if data.get("is_range", ":" in ref_str):
            reference = RangeReference.parse(ref_str)
        else:
            reference = CellReference.parse(ref_str)
        return cls(
            name=data["name"],
            reference=reference,
            description=data.get("description", ""),
        )


class NamedRangeManager:
    """Manages named ranges for a spreadsheet.

    Provides methods to create, delete, and lookup named ranges.
    Names are case-insensitive and must start with a letter.
    """

    def __init__(self) -> None:
        self._names: dict[str, NamedRange] = {}

    def add(self, name: str, reference: RangeReference | CellReference,
            description: str = "") -> NamedRange:
        """Add or update a named range.

        Args:
            name: Name for the range (must start with letter)
            reference: Cell or range reference
            description: Optional description

        Returns:
            The created NamedRange

        Raises:
            ValueError: If name is invalid
        """
        if not self.is_valid_name(name):
            raise ValueError(
                f"Invalid name '{name}': must start with letter, "
                "contain only letters, numbers, underscores"
            )

        named_range = NamedRange(name, reference, description)
        self._names[named_range.name] = named_range
        return named_range

    def add_from_string(self, name: str, ref_str: str, description: str = "") -> NamedRange:
        """Add a named range from a reference string.

        Args:
            name: Name for the range
            ref_str: Reference string like 'A1' or 'A1:B10'
            description: Optional description

        Returns:
            The created NamedRange
        """
        reference: CellReference | RangeReference
        if ":" in ref_str:
            reference = RangeReference.parse(ref_str)
        else:
            reference = CellReference.parse(ref_str)
        return self.add(name, reference, description)

    def delete(self, name: str) -> bool:
        """Delete a named range.

        Args:
            name: Name to delete

        Returns:
            True if deleted, False if not found
        """
        name = name.upper()
        if name in self._names:
            del self._names[name]
            return True
        return False

    def get(self, name: str) -> NamedRange | None:
        """Get a named range by name.

        Args:
            name: Name to lookup

        Returns:
            NamedRange if found, None otherwise
        """
        return self._names.get(name.upper())

    def get_reference(self, name: str) -> RangeReference | CellReference | None:
        """Get the reference for a named range.

        Args:
            name: Name to lookup

        Returns:
            The reference if found, None otherwise
        """
        named = self.get(name)
        return named.reference if named else None

    def exists(self, name: str) -> bool:
        """Check if a name exists."""
        return name.upper() in self._names

    def list_all(self) -> list[NamedRange]:
        """Get all named ranges sorted by name."""
        return sorted(self._names.values(), key=lambda x: x.name)

    def find_by_cell(self, row: int, col: int) -> list[NamedRange]:
        """Find all named ranges that contain a specific cell.

        Args:
            row: Row index
            col: Column index

        Returns:
            List of named ranges containing the cell
        """
        result = []
        for named in self._names.values():
            if isinstance(named.reference, CellReference):
                if named.reference.row == row and named.reference.col == col:
                    result.append(named)
            else:
                if named.reference.contains(row, col):
                    result.append(named)
        return result

    def adjust_for_insert_row(self, at_row: int) -> None:
        """Adjust all named ranges when a row is inserted.

        Args:
            at_row: Row index where insertion occurs
        """
        for named in self._names.values():
            ref = named.reference
            if isinstance(ref, CellReference):
                if ref.row >= at_row:
                    named.reference = CellReference(
                        ref.row + 1, ref.col, ref.col_absolute, ref.row_absolute
                    )
            else:
                start, end = ref.start, ref.end
                if start.row >= at_row:
                    start = CellReference(
                        start.row + 1, start.col, start.col_absolute, start.row_absolute
                    )
                if end.row >= at_row:
                    end = CellReference(
                        end.row + 1, end.col, end.col_absolute, end.row_absolute
                    )
                named.reference = RangeReference(start, end)

    def adjust_for_delete_row(self, at_row: int) -> list[str]:
        """Adjust all named ranges when a row is deleted.

        Args:
            at_row: Row index being deleted

        Returns:
            List of names that were invalidated (pointed to deleted row)
        """
        invalidated = []
        for named in list(self._names.values()):
            ref = named.reference
            if isinstance(ref, CellReference):
                if ref.row == at_row:
                    invalidated.append(named.name)
                    del self._names[named.name]
                elif ref.row > at_row:
                    named.reference = CellReference(
                        ref.row - 1, ref.col, ref.col_absolute, ref.row_absolute
                    )
            else:
                start, end = ref.start, ref.end
                if start.row > at_row:
                    start = CellReference(
                        start.row - 1, start.col, start.col_absolute, start.row_absolute
                    )
                if end.row > at_row:
                    end = CellReference(
                        end.row - 1, end.col, end.col_absolute, end.row_absolute
                    )
                named.reference = RangeReference(start, end)
        return invalidated

    def adjust_for_insert_col(self, at_col: int) -> None:
        """Adjust all named ranges when a column is inserted."""
        for named in self._names.values():
            ref = named.reference
            if isinstance(ref, CellReference):
                if ref.col >= at_col:
                    named.reference = CellReference(
                        ref.row, ref.col + 1, ref.col_absolute, ref.row_absolute
                    )
            else:
                start, end = ref.start, ref.end
                if start.col >= at_col:
                    start = CellReference(
                        start.row, start.col + 1, start.col_absolute, start.row_absolute
                    )
                if end.col >= at_col:
                    end = CellReference(
                        end.row, end.col + 1, end.col_absolute, end.row_absolute
                    )
                named.reference = RangeReference(start, end)

    def adjust_for_delete_col(self, at_col: int) -> list[str]:
        """Adjust all named ranges when a column is deleted.

        Returns:
            List of names that were invalidated
        """
        invalidated = []
        for named in list(self._names.values()):
            ref = named.reference
            if isinstance(ref, CellReference):
                if ref.col == at_col:
                    invalidated.append(named.name)
                    del self._names[named.name]
                elif ref.col > at_col:
                    named.reference = CellReference(
                        ref.row, ref.col - 1, ref.col_absolute, ref.row_absolute
                    )
            else:
                start, end = ref.start, ref.end
                if start.col > at_col:
                    start = CellReference(
                        start.row, start.col - 1, start.col_absolute, start.row_absolute
                    )
                if end.col > at_col:
                    end = CellReference(
                        end.row, end.col - 1, end.col_absolute, end.row_absolute
                    )
                named.reference = RangeReference(start, end)
        return invalidated

    def clear(self) -> None:
        """Remove all named ranges."""
        self._names.clear()

    def __len__(self) -> int:
        return len(self._names)

    def __iter__(self) -> Iterator[NamedRange]:
        return iter(self._names.values())

    def __contains__(self, name: str) -> bool:
        return self.exists(name)

    @staticmethod
    def is_valid_name(name: str) -> bool:
        """Check if a name is valid.

        Names must:
        - Start with a letter
        - Contain only letters, numbers, underscores
        - Not be a cell reference (like A1, AA10)
        """
        if not NAME_PATTERN.match(name):
            return False

        # Check it's not a cell reference
        # Cell references are letters followed by numbers
        import re
        if re.match(r'^[A-Za-z]+\d+$', name):
            return False

        return True

    def to_dict(self) -> dict[str, dict[str, Any]]:
        """Serialize all named ranges."""
        return {name: nr.to_dict() for name, nr in self._names.items()}

    def from_dict(self, data: dict[str, dict[str, Any]]) -> None:
        """Load named ranges from dictionary."""
        self._names.clear()
        for name, nr_data in data.items():
            self._names[name.upper()] = NamedRange.from_dict(nr_data)

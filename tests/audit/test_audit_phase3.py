from lotus123.core.spreadsheet import Spreadsheet
from lotus123.formula.recalc import RecalcMode


class TestPhase3Audit:
    """Systematic audit tests for edge cases and complexity."""

    def setup_method(self):
        self.sheet = Spreadsheet()
        self.engine = self.sheet._recalc_engine
        self.sheet.set_recalc_mode(RecalcMode.MANUAL)

    def test_circular_reference_self(self):
        """Audit: Direct self-reference."""
        # A1 = A1
        self.sheet.set_cell(0, 0, "=A1")
        self.sheet.recalculate()

        # Should detect cycle and return #CIRC!
        val = self.sheet.get_value(0, 0)
        assert val == "#CIRC!" or val == "#CIRC", f"Expected #CIRC!, got {val}"

        # Check RecalcEngine stats
        assert self.sheet._recalc_engine is not None
        assert (0, 0) in self.sheet._recalc_engine.get_circular_references()

    def test_circular_reference_chain(self):
        """Audit: Chain cycle A1->B1->A1."""
        self.sheet.set_cell(0, 0, "=B1")
        self.sheet.set_cell(0, 1, "=A1")

        self.sheet.recalculate()

        val_a = self.sheet.get_value(0, 0)
        val_b = self.sheet.get_value(0, 1)

        assert val_a == "#CIRC!"
        assert val_b == "#CIRC!"

    def test_breaking_circular_reference(self):
        """Audit: Breaking a cycle restores values."""
        self.sheet.set_cell(0, 0, "=B1")
        self.sheet.set_cell(0, 1, "=A1")
        self.sheet.recalculate()
        assert self.sheet.get_value(0, 0) == "#CIRC!"

        # Break cycle: B1 = 10
        self.sheet.set_cell(0, 1, "10")
        self.sheet.recalculate()

        assert self.sheet.get_value(0, 1) == 10.0
        assert self.sheet.get_value(0, 0) == 10.0
        assert self.sheet._recalc_engine is not None
        assert not self.sheet._recalc_engine.get_circular_references()

    def test_structural_update_with_invalid_ref(self):
        """Audit: Formula referring to deleted row becomes error."""
        # A1 = B1. B1 = 10.
        self.sheet.set_cell(0, 0, "=B1")
        self.sheet.set_cell(0, 1, "10")

        # Delete Col 1 (B)
        self.sheet.delete_col(1)

        # A1 should be #REF!
        val = self.sheet.get_cell(0, 0).raw_value
        assert val == "=#REF!", f"Expected =#REF!, got {val}"

    def test_overlapping_range_functions(self):
        """Audit: Functions over overlapping ranges."""
        # A1=1, A2=2, A3=3
        self.sheet.set_cell(0, 0, "1")
        self.sheet.set_cell(1, 0, "2")
        self.sheet.set_cell(2, 0, "3")

        # SUM(A1:A3)
        self.sheet.set_cell(0, 1, "=SUM(A1:A3)")

        self.sheet.recalculate()
        assert self.sheet.get_value(0, 1) == 6.0

        # Insert row at 1. Range should expand?
        # A1:A3 becomes A1:A4 (A1..new..A3..A4)
        self.sheet.insert_row(1)
        # A1=1, A2=empty, A3=2, A4=3

        assert self.sheet.get_cell(0, 1).raw_value == "=SUM(A1:A4)"
        self.sheet.recalculate()
        assert self.sheet.get_value(0, 1) == 6.0  # 1+0+2+3

    def test_named_range_deletion_behavior(self):
        """Audit: Named range pointing to deleted data."""
        self.sheet.set_cell(0, 0, "100")
        self.sheet.named_ranges.add_from_string("MYVAR", "A1")

        assert self.sheet.named_ranges.exists("MYVAR")

        # Delete A1's row
        self.sheet.delete_row(0)

        # Named range should be gone
        assert not self.sheet.named_ranges.exists("MYVAR")

    def test_large_dependency_chain_with_updates(self):
        """Audit: Deep dependency chain (stack overflow check)."""
        # Chain of 50 cells
        for i in range(50):
            self.sheet.set_cell(i, 0, f"={i}")  # Simple values to start

        # Link them: A2=A1, A3=A2...
        self.sheet.set_cell(0, 0, "1")
        for i in range(1, 50):
            self.sheet.set_cell(i, 0, f"=A{i}")

        self.sheet.recalculate()
        assert self.sheet.get_value(49, 0) == 1.0

        # Update root
        self.sheet.set_cell(0, 0, "2")
        self.sheet.recalculate()
        assert self.sheet.get_value(49, 0) == 2.0

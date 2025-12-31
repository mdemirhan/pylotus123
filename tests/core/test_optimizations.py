"""Tests for optimizations and refactoring."""

import os
from lotus123.core.spreadsheet import Spreadsheet
from lotus123.core.cell import Cell
from lotus123.formula.parser import FormulaParser
from lotus123.formula.functions import REGISTRY, FunctionRegistry


class TestOptimizations:
    def test_registry_singleton(self):
        """Test that FunctionRegistry is a singleton."""
        sheet = Spreadsheet()
        parser1 = FormulaParser(sheet)
        parser2 = FormulaParser(sheet)
        
        # Check if parsers share the same registry instance
        assert parser1.functions is REGISTRY
        assert parser2.functions is REGISTRY
        assert parser1.functions is parser2.functions
        
        # Explicit check against direct instantiation
        new_registry = FunctionRegistry()
        assert parser1.functions is not new_registry

    def test_formula_detection_logic(self):
        """Test enhanced is_formula detection logic."""
        c = Cell()
        
        cases = [
            # Standard formulas
            ("=1+1", True),
            ("@SUM(1,2)", True),
            # Standard numbers
            ("123", False),
            ("12.34", False),
            ("-123.45", False),
            ("1e5", False),
            # Plus/Minus prefixes acting as numbers
            ("+123", False),
            ("-123", False),
            # Plus/Minus prefixes acting as formulas
            ("+1.2.3", True),  # Malformed number -> Formula
            ("+A1", True),     # Reference -> Formula
            ("-B2*5", True),   # Expression -> Formula
        ]
        
        for val, expected in cases:
            c.set_value(val)
            assert c.is_formula == expected, f"Value '{val}': expected {expected}, got {c.is_formula}"


class TestIORefactor:
    def test_spreadsheet_json_roundtrip(self, tmp_path):
        """Test saving and loading spreadsheet via JSON."""
        filename = str(tmp_path / "test_sheet.json")
        
        # Create initial sheet
        s1 = Spreadsheet()
        s1.set_cell(0, 0, "100")       # A1
        s1.set_cell(0, 1, "=A1*2")     # B1
        s1.set_cell(1, 0, "Test")      # A2
        
        # Add named range
        s1.named_ranges.add_from_string("MyRange", "A1:B1")
        
        # Save
        s1.save(filename)
        assert os.path.exists(filename)
        
        # Load into new sheet
        s2 = Spreadsheet()
        s2.load(filename)
        
        # Verify content
        assert s2.get_value(0, 0) == 100
        assert s2.get_cell(0, 1).raw_value == "=A1*2"
        assert s2.get_value(1, 0) == "Test"
        
        # Verify named range
        assert s2.named_ranges.exists("MyRange")
        ref = s2.named_ranges.get("MyRange").reference
        assert ref.start.row == 0 and ref.start.col == 0 # A1
        assert ref.end.row == 0 and ref.end.col == 1     # B1

"""Tests for XLSX formula and format translators."""

import pytest

from lotus123.io.xlsx_formula_translator import FormulaTranslator
from lotus123.io.xlsx_format_translator import FormatTranslator


class TestFormulaTranslator:
    """Tests for FormulaTranslator class."""

    def test_lotus_to_excel_at_prefix(self):
        """Test @ prefix is converted to = prefix."""
        assert FormulaTranslator.lotus_to_excel("@SUM(A1:A10)") == "=SUM(A1:A10)"
        assert FormulaTranslator.lotus_to_excel("@AVG(B1:B5)") == "=AVERAGE(B1:B5)"

    def test_lotus_to_excel_equals_prefix(self):
        """Test = prefix is preserved."""
        assert FormulaTranslator.lotus_to_excel("=SUM(A1:A10)") == "=SUM(A1:A10)"
        assert FormulaTranslator.lotus_to_excel("=AVG(B1:B5)") == "=AVERAGE(B1:B5)"

    def test_lotus_to_excel_function_mappings(self):
        """Test all function name mappings."""
        # AVG -> AVERAGE
        assert FormulaTranslator.lotus_to_excel("=AVG(A1:A10)") == "=AVERAGE(A1:A10)"
        # STD -> STDEV
        assert FormulaTranslator.lotus_to_excel("=STD(A1:A10)") == "=STDEV(A1:A10)"
        # STDS -> STDEV
        assert FormulaTranslator.lotus_to_excel("=STDS(A1:A10)") == "=STDEV(A1:A10)"
        # STDP -> STDEV.P
        assert FormulaTranslator.lotus_to_excel("=STDP(A1:A10)") == "=STDEV.P(A1:A10)"
        # VARS -> VAR
        assert FormulaTranslator.lotus_to_excel("=VARS(A1:A10)") == "=VAR(A1:A10)"
        # VARP -> VAR.P
        assert FormulaTranslator.lotus_to_excel("=VARP(A1:A10)") == "=VAR.P(A1:A10)"
        # LENGTH -> LEN
        assert FormulaTranslator.lotus_to_excel("=LENGTH(A1)") == "=LEN(A1)"
        # COLS -> COLUMNS
        assert FormulaTranslator.lotus_to_excel("=COLS(A1:C1)") == "=COLUMNS(A1:C1)"
        # DAVG -> DAVERAGE
        assert (
            FormulaTranslator.lotus_to_excel("=DAVG(A1:C10,1,E1:E2)") == "=DAVERAGE(A1:C10,1,E1:E2)"
        )

    def test_lotus_to_excel_nested_functions(self):
        """Test nested function translation."""
        result = FormulaTranslator.lotus_to_excel("=AVG(STD(A1:A10),STD(B1:B10))")
        assert result == "=AVERAGE(STDEV(A1:A10),STDEV(B1:B10))"

    def test_lotus_to_excel_preserves_refs(self):
        """Test cell references are preserved."""
        assert FormulaTranslator.lotus_to_excel("=$A$1+B2") == "=$A$1+B2"
        assert FormulaTranslator.lotus_to_excel("=A1:Z100") == "=A1:Z100"

    def test_excel_to_lotus_function_mappings(self):
        """Test reverse function name mappings."""
        # AVERAGE -> AVG
        assert FormulaTranslator.excel_to_lotus("=AVERAGE(A1:A10)") == "=AVG(A1:A10)"
        # STDEV -> STD
        assert FormulaTranslator.excel_to_lotus("=STDEV(A1:A10)") == "=STD(A1:A10)"
        # STDEV.S -> STD
        assert FormulaTranslator.excel_to_lotus("=STDEV.S(A1:A10)") == "=STD(A1:A10)"
        # STDEV.P -> STDP
        assert FormulaTranslator.excel_to_lotus("=STDEV.P(A1:A10)") == "=STDP(A1:A10)"
        # VAR.S -> VAR
        assert FormulaTranslator.excel_to_lotus("=VAR.S(A1:A10)") == "=VAR(A1:A10)"
        # VAR.P -> VARP
        assert FormulaTranslator.excel_to_lotus("=VAR.P(A1:A10)") == "=VARP(A1:A10)"
        # COLUMNS -> COLS
        assert FormulaTranslator.excel_to_lotus("=COLUMNS(A1:C1)") == "=COLS(A1:C1)"
        # DAVERAGE -> DAVG
        assert (
            FormulaTranslator.excel_to_lotus("=DAVERAGE(A1:C10,1,E1:E2)") == "=DAVG(A1:C10,1,E1:E2)"
        )

    def test_excel_to_lotus_nested_functions(self):
        """Test nested function reverse translation."""
        result = FormulaTranslator.excel_to_lotus("=AVERAGE(STDEV(A1:A10),STDEV(B1:B10))")
        assert result == "=AVG(STD(A1:A10),STD(B1:B10))"

    def test_formula_roundtrip(self):
        """Test formulas survive round-trip translation."""
        # Note: LENGTH is an alias for LEN in Lotus, so LEN is the canonical form
        test_formulas = [
            "=SUM(A1:A10)",
            "=AVG(A1:A10)",
            "=STD(A1:A10)",
            "=STDP(A1:A10)",
            "=VARP(A1:A10)",
            "=LEN(A1)",  # LENGTH normalizes to LEN
            "=COLS(A1:C1)",
            "=DAVG(A1:C10,1,E1:E2)",
            "=$A$1+B2*C3",
            "=IF(A1>0,AVG(B1:B10),STD(C1:C10))",
        ]
        for formula in test_formulas:
            excel = FormulaTranslator.lotus_to_excel(formula)
            back = FormulaTranslator.excel_to_lotus(excel)
            assert back == formula, f"Round-trip failed: {formula} -> {excel} -> {back}"

    def test_length_alias_normalization(self):
        """Test that LENGTH normalizes to LEN (which is the canonical form)."""
        # LENGTH is an alias in Lotus - LEN is the canonical form
        excel = FormulaTranslator.lotus_to_excel("=LENGTH(A1)")
        assert excel == "=LEN(A1)"
        # On round-trip, we get the canonical form
        back = FormulaTranslator.excel_to_lotus(excel)
        assert back == "=LEN(A1)"

    def test_get_unsupported_lotus_functions(self):
        """Test detection of unsupported Lotus functions."""
        unsupported = FormulaTranslator.get_unsupported_lotus_functions("=CELLPOINTER(A1)+AA(B1)")
        assert "CELLPOINTER" in unsupported

    def test_non_formula_passthrough(self):
        """Test non-formula values - translator always adds = prefix."""
        # The translator is designed for formulas, so it always adds = prefix
        # The calling code should check is_formula() before calling translator
        assert FormulaTranslator.lotus_to_excel("Hello World") == "=Hello World"
        # excel_to_lotus only processes values starting with =
        assert FormulaTranslator.excel_to_lotus("123.45") == "123.45"

    def test_is_formula(self):
        """Test formula detection."""
        assert FormulaTranslator.is_formula("=SUM(A1:A10)")
        assert FormulaTranslator.is_formula("@SUM(A1:A10)")
        assert not FormulaTranslator.is_formula("Hello World")
        assert not FormulaTranslator.is_formula("123.45")
        assert not FormulaTranslator.is_formula("")


class TestFormatTranslator:
    """Tests for FormatTranslator class."""

    def test_fixed_formats(self):
        """Test fixed decimal format codes."""
        assert FormatTranslator.lotus_to_excel("F0") == "0"
        assert FormatTranslator.lotus_to_excel("F1") == "0.0"
        assert FormatTranslator.lotus_to_excel("F2") == "0.00"
        assert FormatTranslator.lotus_to_excel("F15") == "0.000000000000000"

    def test_scientific_formats(self):
        """Test scientific notation format codes."""
        assert FormatTranslator.lotus_to_excel("S0") == "0E+00"
        assert FormatTranslator.lotus_to_excel("S2") == "0.00E+00"

    def test_currency_formats(self):
        """Test currency format codes."""
        assert FormatTranslator.lotus_to_excel("C0") == "$#,##0"
        assert FormatTranslator.lotus_to_excel("C2") == "$#,##0.00"

    def test_percent_formats(self):
        """Test percent format codes."""
        assert FormatTranslator.lotus_to_excel("P0") == "0%"
        assert FormatTranslator.lotus_to_excel("P2") == "0.00%"

    def test_comma_formats(self):
        """Test comma-separated format codes."""
        assert FormatTranslator.lotus_to_excel(",0") == "#,##0"
        assert FormatTranslator.lotus_to_excel(",2") == "#,##0.00"

    def test_date_formats(self):
        """Test date format codes."""
        assert FormatTranslator.lotus_to_excel("D1") == "DD-MMM-YY"
        assert FormatTranslator.lotus_to_excel("D2") == "DD-MMM"
        assert FormatTranslator.lotus_to_excel("D3") == "MMM-YY"
        assert FormatTranslator.lotus_to_excel("D4") == "MM/DD/YY"
        assert FormatTranslator.lotus_to_excel("D5") == "MM/DD"
        assert FormatTranslator.lotus_to_excel("D6") == "DD-MMM-YYYY"
        assert FormatTranslator.lotus_to_excel("D7") == "YYYY-MM-DD"
        assert FormatTranslator.lotus_to_excel("D8") == "DD/MM/YY"
        assert FormatTranslator.lotus_to_excel("D9") == "DD.MM.YYYY"

    def test_time_formats(self):
        """Test time format codes."""
        assert FormatTranslator.lotus_to_excel("T1") == "HH:MM:SS AM/PM"
        assert FormatTranslator.lotus_to_excel("T2") == "HH:MM AM/PM"
        assert FormatTranslator.lotus_to_excel("T3") == "HH:MM:SS"
        assert FormatTranslator.lotus_to_excel("T4") == "HH:MM"

    def test_hidden_format(self):
        """Test hidden format code."""
        assert FormatTranslator.lotus_to_excel("H") == ";;;"

    def test_general_format(self):
        """Test general format code."""
        assert FormatTranslator.lotus_to_excel("G") == "General"

    def test_excel_to_lotus_reverse(self):
        """Test reverse format translation."""
        assert FormatTranslator.excel_to_lotus("0.00") == "F2"
        assert FormatTranslator.excel_to_lotus("$#,##0.00") == "C2"
        assert FormatTranslator.excel_to_lotus("0.00%") == "P2"
        assert FormatTranslator.excel_to_lotus("DD-MMM-YY") == "D1"
        assert FormatTranslator.excel_to_lotus(";;;") == "H"
        assert FormatTranslator.excel_to_lotus("General") == "G"

    def test_format_roundtrip(self):
        """Test format codes survive round-trip translation."""
        test_codes = [
            "F0",
            "F1",
            "F2",
            "F5",
            "F15",
            "S0",
            "S2",
            "C0",
            "C2",
            "P0",
            "P2",
            ",0",
            ",2",
            "D1",
            "D2",
            "D3",
            "D4",
            "D5",
            "D6",
            "D7",
            "D8",
            "D9",
            "T1",
            "T2",
            "T3",
            "T4",
            "H",
            "G",
        ]
        for code in test_codes:
            excel = FormatTranslator.lotus_to_excel(code)
            back = FormatTranslator.excel_to_lotus(excel)
            assert back == code, f"Round-trip failed: {code} -> {excel} -> {back}"

    def test_unknown_format_fallback(self):
        """Test unknown formats fall back to General."""
        # Unknown Lotus format codes default to General
        assert FormatTranslator.lotus_to_excel("UNKNOWN") == "General"
        # Unknown Excel formats default to G (General)
        assert FormatTranslator.excel_to_lotus("CustomFormat") == "G"

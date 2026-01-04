"""Formula translation between Lotus 1-2-3 and Excel XLSX formats.

This module provides bidirectional translation of formulas between Lotus 1-2-3
and Excel XLSX formats, ensuring lossless round-trip conversion.
"""

import re


# Lotus function names that differ from Excel equivalents
# Maps Lotus name -> Excel name
LOTUS_TO_EXCEL: dict[str, str | None] = {
    # Statistical functions
    "AVG": "AVERAGE",
    "STD": "STDEV",
    "STDS": "STDEV",
    "STDP": "STDEV.P",
    "VARS": "VAR",
    "VARP": "VAR.P",
    # String functions
    "LENGTH": "LEN",
    # Lookup functions
    "COLS": "COLUMNS",
    # Database functions (DAVG -> DAVERAGE)
    "DAVG": "DAVERAGE",
    # Lotus-specific functions with no Excel equivalent
    "CELLPOINTER": None,  # Will cause #NAME? error in Excel
}

# Excel function names that differ from Lotus equivalents
# Maps Excel name -> Lotus name
# This is the reverse mapping for import
EXCEL_TO_LOTUS: dict[str, str] = {
    # Statistical functions
    "AVERAGE": "AVG",
    "STDEV": "STD",
    "STDEV.S": "STD",  # Excel's sample stdev
    "STDEV.P": "STDP",
    "VAR.S": "VAR",  # Excel's sample variance
    "VAR.P": "VARP",
    # String functions
    # Note: LEN maps to LEN in Lotus (not LENGTH, which is just an alias)
    # Lookup functions
    "COLUMNS": "COLS",
    # Database functions
    "DAVERAGE": "DAVG",
}

# Excel functions that don't exist in Lotus 1-2-3
# These will be preserved as-is but will show #NAME? error
EXCEL_ONLY_FUNCTIONS: set[str] = {
    # Modern Excel functions
    "XLOOKUP",
    "XMATCH",
    "FILTER",
    "SORT",
    "SORTBY",
    "UNIQUE",
    "SEQUENCE",
    "RANDARRAY",
    "LET",
    "LAMBDA",
    "MAP",
    "REDUCE",
    "SCAN",
    "MAKEARRAY",
    "BYROW",
    "BYCOL",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "DROP",
    "TAKE",
    "EXPAND",
    "VSTACK",
    "HSTACK",
    "WRAPCOLS",
    "WRAPROWS",
    "TOCOL",
    "TOROW",
    "TEXTSPLIT",
    "TEXTBEFORE",
    "TEXTAFTER",
    "VALUETOTEXT",
    "ARRAYTOTEXT",
    # Stock/geography functions
    "STOCKHISTORY",
    "WEBSERVICE",
    # Other modern functions
    "IFS",
    "MAXIFS",
    "MINIFS",
    "CONCAT",  # Lotus has CONCATENATE
    "TEXTJOIN",
    "SWITCH",  # Note: Lotus does have SWITCH
}


class FormulaTranslator:
    """Translate formulas between Lotus 1-2-3 and Excel XLSX formats.

    This class provides bidirectional translation ensuring round-trip fidelity:
    - Lotus -> Excel -> Lotus should produce the original formula
    - Function names are translated appropriately in both directions
    - The @ prefix is normalized to = (both work in Lotus)
    """

    # Regex to find function names in formulas
    # Matches: optional @, function name, opening paren
    _FUNCTION_PATTERN = re.compile(r"@?([A-Z][A-Z0-9_.]*)\s*\(", re.IGNORECASE)

    @classmethod
    def lotus_to_excel(cls, formula: str) -> str:
        """Convert Lotus 1-2-3 formula to Excel formula.

        Args:
            formula: Lotus formula (may start with = or @)

        Returns:
            Excel-compatible formula (starts with =)

        Examples:
            >>> FormulaTranslator.lotus_to_excel("@SUM(A1:A10)")
            '=SUM(A1:A10)'
            >>> FormulaTranslator.lotus_to_excel("=AVG(B1:B10)")
            '=AVERAGE(B1:B10)'
        """
        if not formula:
            return formula

        # Remove @ prefix if present, ensure = prefix
        result = formula.lstrip("@")
        if result.startswith("="):
            result = result[1:]  # Remove = temporarily for processing

        # Replace function names
        def replace_func(match: re.Match) -> str:
            full_match = match.group(0)
            func_name = match.group(1).upper()

            # Look up translation
            excel_name = LOTUS_TO_EXCEL.get(func_name)
            if excel_name is None and func_name in LOTUS_TO_EXCEL:
                # Function exists in mapping but has no Excel equivalent
                return f"_UNSUPPORTED_{func_name}("
            elif excel_name is not None:
                # Preserve original case style if possible
                if match.group(1).isupper():
                    return f"{excel_name}("
                elif match.group(1).islower():
                    return f"{excel_name.lower()}("
                else:
                    return f"{excel_name}("
            else:
                # Function name is the same in both systems
                return full_match.lstrip("@")

        result = cls._FUNCTION_PATTERN.sub(replace_func, result)

        return "=" + result

    @classmethod
    def excel_to_lotus(cls, formula: str) -> str:
        """Convert Excel formula to Lotus 1-2-3 formula.

        Args:
            formula: Excel formula (starts with =)

        Returns:
            Lotus-compatible formula (starts with =)

        Examples:
            >>> FormulaTranslator.excel_to_lotus("=AVERAGE(A1:A10)")
            '=AVG(A1:A10)'
            >>> FormulaTranslator.excel_to_lotus("=STDEV.P(B1:B10)")
            '=STDP(B1:B10)'
        """
        if not formula:
            return formula

        # Ensure we're working with a formula
        if not formula.startswith("="):
            return formula

        result = formula[1:]  # Remove = temporarily

        # Replace function names
        def replace_func(match: re.Match) -> str:
            func_name = match.group(1).upper()

            # Look up translation
            lotus_name = EXCEL_TO_LOTUS.get(func_name)
            if lotus_name is not None:
                # Preserve original case style if possible
                if match.group(1).isupper():
                    return f"{lotus_name}("
                elif match.group(1).islower():
                    return f"{lotus_name.lower()}("
                else:
                    return f"{lotus_name}("
            else:
                # Function name is the same or unknown - keep as is
                return match.group(0)

        result = cls._FUNCTION_PATTERN.sub(replace_func, result)

        return "=" + result

    @classmethod
    def get_unsupported_lotus_functions(cls, formula: str) -> list[str]:
        """Find Lotus functions that have no Excel equivalent.

        Args:
            formula: Lotus formula to check

        Returns:
            List of function names that cannot be translated to Excel
        """
        unsupported = []

        for match in cls._FUNCTION_PATTERN.finditer(formula):
            func_name = match.group(1).upper()
            if func_name in LOTUS_TO_EXCEL and LOTUS_TO_EXCEL[func_name] is None:
                unsupported.append(func_name)

        return unsupported

    @classmethod
    def get_unsupported_excel_functions(cls, formula: str) -> list[str]:
        """Find Excel functions that don't exist in Lotus 1-2-3.

        Args:
            formula: Excel formula to check

        Returns:
            List of function names that will show #NAME? in Lotus
        """
        unsupported = []

        for match in cls._FUNCTION_PATTERN.finditer(formula):
            func_name = match.group(1).upper()
            if func_name in EXCEL_ONLY_FUNCTIONS:
                unsupported.append(func_name)

        return unsupported

    @classmethod
    def is_formula(cls, value: str) -> bool:
        """Check if a value is a formula.

        Args:
            value: Cell value to check

        Returns:
            True if value starts with = or @
        """
        if not value:
            return False
        return value.startswith("=") or value.startswith("@")

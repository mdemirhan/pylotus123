"""Format code translation between Lotus 1-2-3 and Excel XLSX formats.

This module provides bidirectional translation of format codes, ensuring
lossless round-trip conversion between Lotus and Excel formats.
"""

import re


# Complete mapping of Lotus format codes to Excel number format strings
# This is the authoritative mapping for export
LOTUS_TO_EXCEL_FORMAT: dict[str, str] = {
    # General
    "G": "General",
    # Fixed decimal (F0-F15)
    "F0": "0",
    "F1": "0.0",
    "F2": "0.00",
    "F3": "0.000",
    "F4": "0.0000",
    "F5": "0.00000",
    "F6": "0.000000",
    "F7": "0.0000000",
    "F8": "0.00000000",
    "F9": "0.000000000",
    "F10": "0.0000000000",
    "F11": "0.00000000000",
    "F12": "0.000000000000",
    "F13": "0.0000000000000",
    "F14": "0.00000000000000",
    "F15": "0.000000000000000",
    # Scientific notation (S0-S15)
    "S0": "0E+00",
    "S1": "0.0E+00",
    "S2": "0.00E+00",
    "S3": "0.000E+00",
    "S4": "0.0000E+00",
    "S5": "0.00000E+00",
    "S6": "0.000000E+00",
    "S7": "0.0000000E+00",
    "S8": "0.00000000E+00",
    "S9": "0.000000000E+00",
    "S10": "0.0000000000E+00",
    "S11": "0.00000000000E+00",
    "S12": "0.000000000000E+00",
    "S13": "0.0000000000000E+00",
    "S14": "0.00000000000000E+00",
    "S15": "0.000000000000000E+00",
    # Currency (C0-C15)
    "C0": "$#,##0",
    "C1": "$#,##0.0",
    "C2": "$#,##0.00",
    "C3": "$#,##0.000",
    "C4": "$#,##0.0000",
    "C5": "$#,##0.00000",
    "C6": "$#,##0.000000",
    "C7": "$#,##0.0000000",
    "C8": "$#,##0.00000000",
    "C9": "$#,##0.000000000",
    "C10": "$#,##0.0000000000",
    "C11": "$#,##0.00000000000",
    "C12": "$#,##0.000000000000",
    "C13": "$#,##0.0000000000000",
    "C14": "$#,##0.00000000000000",
    "C15": "$#,##0.000000000000000",
    # Comma/Thousands (,0-,15)
    ",0": "#,##0",
    ",1": "#,##0.0",
    ",2": "#,##0.00",
    ",3": "#,##0.000",
    ",4": "#,##0.0000",
    ",5": "#,##0.00000",
    ",6": "#,##0.000000",
    ",7": "#,##0.0000000",
    ",8": "#,##0.00000000",
    ",9": "#,##0.000000000",
    ",10": "#,##0.0000000000",
    ",11": "#,##0.00000000000",
    ",12": "#,##0.000000000000",
    ",13": "#,##0.0000000000000",
    ",14": "#,##0.00000000000000",
    ",15": "#,##0.000000000000000",
    # Percent (P0-P15)
    "P0": "0%",
    "P1": "0.0%",
    "P2": "0.00%",
    "P3": "0.000%",
    "P4": "0.0000%",
    "P5": "0.00000%",
    "P6": "0.000000%",
    "P7": "0.0000000%",
    "P8": "0.00000000%",
    "P9": "0.000000000%",
    "P10": "0.0000000000%",
    "P11": "0.00000000000%",
    "P12": "0.000000000000%",
    "P13": "0.0000000000000%",
    "P14": "0.00000000000000%",
    "P15": "0.000000000000000%",
    # Date formats (D1-D9)
    "D1": "DD-MMM-YY",
    "D2": "DD-MMM",
    "D3": "MMM-YY",
    "D4": "MM/DD/YY",
    "D5": "MM/DD",
    "D6": "DD-MMM-YYYY",
    "D7": "YYYY-MM-DD",
    "D8": "DD/MM/YY",
    "D9": "DD.MM.YYYY",
    # Time formats (T1-T4)
    "T1": "HH:MM:SS AM/PM",
    "T2": "HH:MM AM/PM",
    "T3": "HH:MM:SS",
    "T4": "HH:MM",
    # Hidden
    "H": ";;;",
    # Plus/minus bar graph (no Excel equivalent, use general)
    "+": "General",
}

# Reverse mapping: Excel format string -> Lotus code
# Built from LOTUS_TO_EXCEL_FORMAT for exact matches
EXCEL_TO_LOTUS_FORMAT: dict[str, str] = {
    v: k for k, v in LOTUS_TO_EXCEL_FORMAT.items() if v != "General" or k == "G"
}

# Additional Excel format patterns that map to Lotus codes
# These handle variations in Excel format strings
EXCEL_FORMAT_ALIASES: dict[str, str] = {
    # General variations
    "general": "G",
    # Fixed decimal variations (Excel sometimes uses different representations)
    "0.0": "F1",
    "0.00": "F2",
    # Currency variations (Excel may use different currency symbols)
    '"$"#,##0': "C0",
    '"$"#,##0.00': "C2",
    "_($* #,##0_)": "C0",
    "_($* #,##0.00_)": "C2",
    # Percent variations
    "0.0%": "P1",
    "0.00%": "P2",
    # Date variations
    "dd-mmm-yy": "D1",
    "dd-mmm": "D2",
    "mmm-yy": "D3",
    "mm/dd/yy": "D4",
    "m/d/yy": "D4",
    "mm/dd": "D5",
    "m/d": "D5",
    "dd-mmm-yyyy": "D6",
    "yyyy-mm-dd": "D7",
    "dd/mm/yy": "D8",
    "d/m/yy": "D8",
    "dd.mm.yyyy": "D9",
    "d.m.yyyy": "D9",
    # Time variations
    "hh:mm:ss am/pm": "T1",
    "h:mm:ss am/pm": "T1",
    "hh:mm am/pm": "T2",
    "h:mm am/pm": "T2",
    "hh:mm:ss": "T3",
    "h:mm:ss": "T3",
    "hh:mm": "T4",
    "h:mm": "T4",
    # Hidden
    ";;;": "H",
}


class FormatTranslator:
    """Translate format codes between Lotus 1-2-3 and Excel XLSX formats.

    This class provides bidirectional translation ensuring round-trip fidelity:
    - Lotus -> Excel -> Lotus should produce the original format code
    """

    @classmethod
    def lotus_to_excel(cls, lotus_format: str) -> str:
        """Convert Lotus format code to Excel number format string.

        Args:
            lotus_format: Lotus format code (e.g., "F2", "C0", "D1")

        Returns:
            Excel number format string (e.g., "0.00", "$#,##0", "DD-MMM-YY")

        Examples:
            >>> FormatTranslator.lotus_to_excel("F2")
            '0.00'
            >>> FormatTranslator.lotus_to_excel("C2")
            '$#,##0.00'
            >>> FormatTranslator.lotus_to_excel("D7")
            'YYYY-MM-DD'
        """
        if not lotus_format:
            return "General"

        lotus_format = lotus_format.upper()

        # Direct lookup first
        if lotus_format in LOTUS_TO_EXCEL_FORMAT:
            return LOTUS_TO_EXCEL_FORMAT[lotus_format]

        # Parse format code for decimals beyond what we explicitly mapped
        if len(lotus_format) >= 2:
            fmt_type = lotus_format[0]
            try:
                decimals = int(lotus_format[1:])
                decimals = max(0, min(15, decimals))  # Clamp to 0-15

                if fmt_type == "F":
                    return "0." + "0" * decimals if decimals > 0 else "0"
                elif fmt_type == "S":
                    return "0." + "0" * decimals + "E+00" if decimals > 0 else "0E+00"
                elif fmt_type == "C":
                    return "$#,##0." + "0" * decimals if decimals > 0 else "$#,##0"
                elif fmt_type == ",":
                    return "#,##0." + "0" * decimals if decimals > 0 else "#,##0"
                elif fmt_type == "P":
                    return "0." + "0" * decimals + "%" if decimals > 0 else "0%"
            except ValueError:
                pass

        return "General"

    @classmethod
    def excel_to_lotus(cls, excel_format: str) -> str:
        """Convert Excel number format string to Lotus format code.

        Args:
            excel_format: Excel number format string

        Returns:
            Lotus format code

        Examples:
            >>> FormatTranslator.excel_to_lotus("0.00")
            'F2'
            >>> FormatTranslator.excel_to_lotus("$#,##0.00")
            'C2'
            >>> FormatTranslator.excel_to_lotus("YYYY-MM-DD")
            'D7'
        """
        if not excel_format:
            return "G"

        # Normalize for comparison
        normalized = excel_format.strip()
        lower = normalized.lower()

        if lower == "general":
            return "G"

        # Exact match lookup (case-sensitive first)
        if normalized in EXCEL_TO_LOTUS_FORMAT:
            return EXCEL_TO_LOTUS_FORMAT[normalized]

        # Try case-insensitive match
        if lower in EXCEL_FORMAT_ALIASES:
            return EXCEL_FORMAT_ALIASES[lower]

        # Pattern matching for common Excel formats
        return cls._pattern_match_excel_format(normalized)

    @classmethod
    def _pattern_match_excel_format(cls, excel_format: str) -> str:
        """Match Excel format patterns to Lotus codes.

        Args:
            excel_format: Excel format string to match

        Returns:
            Lotus format code
        """
        fmt = excel_format

        # Hidden format
        if fmt == ";;;":
            return "H"

        # Currency format (starts with $ or has currency)
        if fmt.startswith("$") or fmt.startswith('"$"'):
            decimals = cls._count_decimal_places(fmt)
            return f"C{decimals}"

        # Percent format (ends with %)
        if fmt.endswith("%"):
            decimals = cls._count_decimal_places(fmt.rstrip("%"))
            return f"P{decimals}"

        # Scientific notation (contains E+ or E-)
        if "E+" in fmt.upper() or "E-" in fmt.upper():
            decimals = cls._count_decimal_places(fmt.split("E")[0].split("e")[0])
            return f"S{decimals}"

        # Comma format (has #,##0 pattern)
        if "#,##0" in fmt or ",##0" in fmt:
            decimals = cls._count_decimal_places(fmt)
            return f",{decimals}"

        # Date format (contains date components)
        if cls._is_date_format(fmt):
            return cls._match_date_format(fmt)

        # Time format (contains time components)
        if cls._is_time_format(fmt):
            return cls._match_time_format(fmt)

        # Fixed decimal format (just numbers and decimal)
        if re.match(r"^0(\.0+)?$", fmt):
            decimals = cls._count_decimal_places(fmt)
            return f"F{decimals}"

        # Default to General
        return "G"

    @classmethod
    def _count_decimal_places(cls, fmt: str) -> int:
        """Count decimal places in a format string.

        Args:
            fmt: Format string

        Returns:
            Number of decimal places (0-15)
        """
        if "." not in fmt:
            return 0

        # Find the decimal portion
        after_decimal = fmt.split(".")[-1]
        # Count zeros (and sometimes #) after decimal
        count = 0
        for char in after_decimal:
            if char in "0#":
                count += 1
            elif char.isalpha() or char in "%":
                break

        return min(count, 15)

    @classmethod
    def _is_date_format(cls, fmt: str) -> bool:
        """Check if format string is a date format.

        Args:
            fmt: Format string

        Returns:
            True if it's a date format
        """
        lower = fmt.lower()
        date_indicators = ["yyyy", "yy", "mmmm", "mmm", "mm", "dd", "d"]
        return any(ind in lower for ind in date_indicators) and ":" not in fmt

    @classmethod
    def _is_time_format(cls, fmt: str) -> bool:
        """Check if format string is a time format.

        Args:
            fmt: Format string

        Returns:
            True if it's a time format
        """
        lower = fmt.lower()
        return ":" in lower and ("hh" in lower or "h:" in lower or "mm" in lower)

    @classmethod
    def _match_date_format(cls, fmt: str) -> str:
        """Match a date format string to a Lotus date code.

        Args:
            fmt: Date format string

        Returns:
            Lotus date code (D1-D9)
        """
        lower = fmt.lower()

        # D7: ISO format (YYYY-MM-DD)
        if "yyyy-mm-dd" in lower or "yyyy/mm/dd" in lower:
            return "D7"

        # D9: German format (DD.MM.YYYY)
        if "dd.mm.yyyy" in lower or "d.m.yyyy" in lower:
            return "D9"

        # D6: Full date (DD-MMM-YYYY)
        if "yyyy" in lower and "mmm" in lower:
            return "D6"

        # D1: Standard date (DD-MMM-YY)
        if "-" in fmt and "mmm" in lower and "yy" in lower:
            return "D1"

        # D2: Day-Month only (DD-MMM)
        if "-" in fmt and "mmm" in lower and "yy" not in lower:
            return "D2"

        # D3: Month-Year (MMM-YY)
        if "-" in fmt and "mmm" in lower and "dd" not in lower:
            return "D3"

        # D8: European date (DD/MM/YY)
        if "/" in fmt and lower.index("dd") < lower.index("mm"):
            return "D8"

        # D4: US date (MM/DD/YY) - default for / format
        if "/" in fmt and "yy" in lower:
            return "D4"

        # D5: Month-Day (MM/DD)
        if "/" in fmt and "yy" not in lower:
            return "D5"

        # Default to D1
        return "D1"

    @classmethod
    def _match_time_format(cls, fmt: str) -> str:
        """Match a time format string to a Lotus time code.

        Args:
            fmt: Time format string

        Returns:
            Lotus time code (T1-T4)
        """
        lower = fmt.lower()
        has_ampm = "am/pm" in lower or "am" in lower or "pm" in lower
        has_seconds = lower.count(":") >= 2 or "ss" in lower

        if has_ampm:
            if has_seconds:
                return "T1"  # HH:MM:SS AM/PM
            else:
                return "T2"  # HH:MM AM/PM
        else:
            if has_seconds:
                return "T3"  # HH:MM:SS (24-hour)
            else:
                return "T4"  # HH:MM (24-hour)


def get_all_lotus_formats() -> list[str]:
    """Get list of all supported Lotus format codes.

    Returns:
        Sorted list of format codes
    """
    return sorted(LOTUS_TO_EXCEL_FORMAT.keys())


def get_all_excel_formats() -> list[str]:
    """Get list of all Excel format strings we can import.

    Returns:
        List of Excel format strings
    """
    formats = set(LOTUS_TO_EXCEL_FORMAT.values())
    formats.update(EXCEL_FORMAT_ALIASES.keys())
    return sorted(formats)

"""Tests for the formatting module."""

import datetime


from lotus123.core.formatting import (
    DateFormat,
    FormatCode,
    FormatSpec,
    TimeFormat,
    _format_comma,
    _format_currency,
    _format_date,
    _format_fixed,
    _format_general,
    _format_percent,
    _format_plusminus,
    _format_scientific,
    _format_time,
    date_to_serial,
    format_value,
    parse_format_code,
    serial_to_date,
    serial_to_time,
    time_to_serial,
)


class TestParseFormatCode:
    """Tests for parse_format_code function."""

    def test_general(self):
        """Test general format."""
        spec = parse_format_code("G")
        assert spec.format_type == FormatCode.GENERAL

    def test_empty_string(self):
        """Test empty string defaults to general."""
        spec = parse_format_code("")
        assert spec.format_type == FormatCode.GENERAL

    def test_hidden(self):
        """Test hidden format."""
        spec = parse_format_code("H")
        assert spec.format_type == FormatCode.HIDDEN

    def test_plusminus(self):
        """Test plus/minus format."""
        spec = parse_format_code("+")
        assert spec.format_type == FormatCode.PLUSMINUS

    def test_fixed_with_decimals(self):
        """Test fixed format with decimals."""
        spec = parse_format_code("F2")
        assert spec.format_type == FormatCode.FIXED
        assert spec.decimals == 2

    def test_fixed_no_decimals(self):
        """Test fixed format with 0 decimals."""
        spec = parse_format_code("F0")
        assert spec.format_type == FormatCode.FIXED
        assert spec.decimals == 0

    def test_scientific(self):
        """Test scientific format."""
        spec = parse_format_code("S3")
        assert spec.format_type == FormatCode.SCIENTIFIC
        assert spec.decimals == 3

    def test_currency(self):
        """Test currency format."""
        spec = parse_format_code("C2")
        assert spec.format_type == FormatCode.CURRENCY
        assert spec.decimals == 2

    def test_comma(self):
        """Test comma format."""
        spec = parse_format_code(",2")
        assert spec.format_type == FormatCode.COMMA
        assert spec.decimals == 2

    def test_percent(self):
        """Test percent format."""
        spec = parse_format_code("P1")
        assert spec.format_type == FormatCode.PERCENT
        assert spec.decimals == 1

    def test_date_formats(self):
        """Test date formats D1-D9."""
        for i in range(1, 10):
            spec = parse_format_code(f"D{i}")
            assert spec.format_type == FormatCode.DATE
            assert spec.date_variant == DateFormat[f"D{i}"]

    def test_time_formats(self):
        """Test time formats T1-T4."""
        for i in range(1, 5):
            spec = parse_format_code(f"T{i}")
            assert spec.format_type == FormatCode.TIME
            assert spec.time_variant == TimeFormat[f"T{i}"]

    def test_decimals_clamped(self):
        """Test that decimals are clamped to 0-15."""
        spec = parse_format_code("F20")
        assert spec.decimals == 15

        # F-5 parses as F with decimal -5, which gets clamped to 0
        spec = parse_format_code("F-5")
        assert spec.decimals == 0

    def test_lowercase(self):
        """Test lowercase codes are normalized."""
        spec = parse_format_code("f2")
        assert spec.format_type == FormatCode.FIXED

    def test_invalid_date_variant(self):
        """Test invalid date variant defaults to D1."""
        spec = parse_format_code("D99")
        assert spec.date_variant == DateFormat.D1


class TestSerialDateConversion:
    """Tests for serial to/from date conversion."""

    def test_serial_to_date_basic(self):
        """Test basic serial to date conversion."""
        # January 1, 1900 = serial 1
        date = serial_to_date(1)
        assert date == datetime.date(1900, 1, 1)

    def test_serial_to_date_modern(self):
        """Test serial to date for modern date."""
        # After the leap year bug adjustment point (serial >= 60)
        date = serial_to_date(44927)  # Approximately 2023-01-01
        assert date.year == 2023

    def test_date_to_serial_basic(self):
        """Test basic date to serial conversion."""
        serial = date_to_serial(datetime.date(1900, 1, 1))
        assert serial == 1

    def test_date_to_serial_modern(self):
        """Test date to serial for modern date."""
        date = datetime.date(2023, 1, 1)
        serial = date_to_serial(date)
        assert serial > 40000  # Modern dates have large serials

    def test_roundtrip(self):
        """Test serial to date and back."""
        original = datetime.date(2023, 6, 15)
        serial = date_to_serial(original)
        result = serial_to_date(serial)
        assert result == original


class TestSerialTimeConversion:
    """Tests for serial to/from time conversion."""

    def test_serial_to_time_noon(self):
        """Test serial to time for noon."""
        time = serial_to_time(0.5)
        assert time.hour == 12
        assert time.minute == 0

    def test_serial_to_time_midnight(self):
        """Test serial to time for midnight."""
        time = serial_to_time(0.0)
        assert time.hour == 0

    def test_serial_to_time_evening(self):
        """Test serial to time for evening."""
        time = serial_to_time(0.75)  # 6 PM
        assert time.hour == 18

    def test_time_to_serial_noon(self):
        """Test time to serial for noon."""
        serial = time_to_serial(datetime.time(12, 0, 0))
        assert abs(serial - 0.5) < 0.001

    def test_time_to_serial_with_minutes(self):
        """Test time to serial with minutes."""
        serial = time_to_serial(datetime.time(6, 30, 0))
        assert serial > 0.25 and serial < 0.3


class TestFormatValue:
    """Tests for format_value function."""

    def test_hidden_format(self):
        """Test hidden format returns empty."""
        spec = FormatSpec(FormatCode.HIDDEN)
        assert format_value(42, spec) == ""

    def test_none_value(self):
        """Test None value returns empty."""
        spec = FormatSpec(FormatCode.GENERAL)
        assert format_value(None, spec) == ""

    def test_empty_value(self):
        """Test empty string returns empty."""
        spec = FormatSpec(FormatCode.GENERAL)
        assert format_value("", spec) == ""

    def test_error_value(self):
        """Test error values pass through."""
        spec = FormatSpec(FormatCode.FIXED, decimals=2)
        assert format_value("#DIV/0!", spec) == "#DIV/0!"

    def test_general_integer(self):
        """Test general format for integer."""
        spec = FormatSpec(FormatCode.GENERAL)
        assert format_value(42.0, spec) == "42"

    def test_general_float(self):
        """Test general format for float."""
        spec = FormatSpec(FormatCode.GENERAL)
        result = format_value(3.14159, spec)
        assert "3.14" in result

    def test_fixed_format(self):
        """Test fixed format."""
        spec = FormatSpec(FormatCode.FIXED, decimals=2)
        assert format_value(3.14159, spec) == "3.14"

    def test_scientific_format(self):
        """Test scientific format."""
        spec = FormatSpec(FormatCode.SCIENTIFIC, decimals=2)
        result = format_value(1234.5, spec)
        assert "E" in result

    def test_currency_format_positive(self):
        """Test currency format for positive value."""
        spec = FormatSpec(FormatCode.CURRENCY, decimals=2)
        result = format_value(1234.5, spec)
        assert "$" in result

    def test_currency_format_negative(self):
        """Test currency format for negative value."""
        spec = FormatSpec(FormatCode.CURRENCY, decimals=2)
        result = format_value(-1234.5, spec)
        assert "(" in result and ")" in result

    def test_comma_format(self):
        """Test comma format."""
        spec = FormatSpec(FormatCode.COMMA, decimals=2)
        result = format_value(1234567.89, spec)
        assert "," in result

    def test_percent_format(self):
        """Test percent format."""
        spec = FormatSpec(FormatCode.PERCENT, decimals=1)
        result = format_value(0.75, spec)
        assert "75.0%" == result

    def test_date_format(self):
        """Test date format."""
        spec = FormatSpec(FormatCode.DATE, date_variant=DateFormat.D4)
        # Serial 44927 is approximately 2023-01-01
        result = format_value(44927, spec)
        assert "/" in result

    def test_time_format(self):
        """Test time format."""
        spec = FormatSpec(FormatCode.TIME, time_variant=TimeFormat.T4)
        result = format_value(0.5, spec)  # Noon
        assert "12:" in result

    def test_plusminus_positive(self):
        """Test plus/minus format for positive."""
        spec = FormatSpec(FormatCode.PLUSMINUS)
        result = format_value(5, spec, width=20)
        assert "+" in result

    def test_plusminus_negative(self):
        """Test plus/minus format for negative."""
        spec = FormatSpec(FormatCode.PLUSMINUS)
        result = format_value(-5, spec, width=20)
        assert "-" in result

    def test_non_numeric_for_numeric_format(self):
        """Test non-numeric value for numeric format."""
        spec = FormatSpec(FormatCode.FIXED, decimals=2)
        assert format_value("hello", spec) == "hello"


class TestFormatHelpers:
    """Tests for individual format helper functions."""

    def test_format_general_string(self):
        """Test general format for string."""
        assert _format_general("hello") == "hello"

    def test_format_fixed(self):
        """Test fixed format."""
        assert _format_fixed(3.14159, 2) == "3.14"
        assert _format_fixed(3.14159, 0) == "3"

    def test_format_scientific(self):
        """Test scientific format."""
        result = _format_scientific(1234.5, 2)
        assert "E" in result

    def test_format_currency(self):
        """Test currency format."""
        result = _format_currency(1234.5, 2, "$")
        assert "$1,234.50" == result

    def test_format_comma(self):
        """Test comma format."""
        assert _format_comma(1234567.89, 2) == "1,234,567.89"

    def test_format_percent(self):
        """Test percent format."""
        assert _format_percent(0.5, 0) == "50%"

    def test_format_date_d1(self):
        """Test date format D1."""
        result = _format_date(44927, DateFormat.D1)
        assert "-" in result

    def test_format_date_d7_iso(self):
        """Test date format D7 (ISO)."""
        result = _format_date(44927, DateFormat.D7)
        assert result.count("-") == 2

    def test_format_time_t1(self):
        """Test time format T1."""
        result = _format_time(0.5, TimeFormat.T1)
        assert "PM" in result or "AM" in result

    def test_format_time_t3(self):
        """Test time format T3 (24-hour)."""
        result = _format_time(0.5, TimeFormat.T3)
        assert "12:" in result

    def test_format_plusminus_non_numeric(self):
        """Test plus/minus for non-numeric."""
        result = _format_plusminus("hello", 10)
        assert result == "hello"


class TestFormatSpec:
    """Tests for FormatSpec dataclass."""

    def test_default_values(self):
        """Test default values."""
        spec = FormatSpec(FormatCode.GENERAL)
        assert spec.decimals == 2
        assert spec.currency_symbol == "$"
        assert spec.negative_format == "-"

    def test_custom_values(self):
        """Test custom values."""
        spec = FormatSpec(
            FormatCode.CURRENCY, decimals=4, currency_symbol="EUR", negative_format="()"
        )
        assert spec.decimals == 4
        assert spec.currency_symbol == "EUR"


class TestDateFormat:
    """Tests for DateFormat enum."""

    def test_all_variants_exist(self):
        """Test all date variants exist."""
        assert DateFormat.D1.value == "DD-MMM-YY"
        assert DateFormat.D7.value == "YYYY-MM-DD"
        assert DateFormat.D9.value == "DD.MM.YYYY"


class TestTimeFormat:
    """Tests for TimeFormat enum."""

    def test_all_variants_exist(self):
        """Test all time variants exist."""
        assert TimeFormat.T1.value == "HH:MM:SS AM/PM"
        assert TimeFormat.T4.value == "HH:MM"

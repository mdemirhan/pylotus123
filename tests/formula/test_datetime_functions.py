"""Tests for datetime functions."""

import datetime


from lotus123.formula.functions.datetime import (
    _date_to_serial,
    _serial_to_date,
    _serial_to_time,
    _time_to_serial,
    _parse_date_string,
    _to_number,
    fn_date,
    fn_datevalue,
    fn_day,
    fn_days,
    fn_edate,
    fn_eomonth,
    fn_hour,
    fn_minute,
    fn_month,
    fn_now,
    fn_second,
    fn_time,
    fn_timevalue,
    fn_today,
    fn_weekday,
    fn_year,
    fn_yearfrac,
)


class TestHelperFunctions:
    """Tests for helper functions."""

    def test_serial_to_date(self):
        """Test serial number to date conversion."""
        # Day 1 is January 1, 1900
        date = _serial_to_date(1)
        assert date == datetime.date(1900, 1, 1)

    def test_serial_to_date_leap_year_bug(self):
        """Test handling of leap year bug after day 60."""
        # After serial 60, subtract 1 due to leap year bug
        date = _serial_to_date(61)
        assert date.year == 1900
        assert date.month == 3
        assert date.day == 1

    def test_date_to_serial(self):
        """Test date to serial number conversion."""
        serial = _date_to_serial(datetime.date(1900, 1, 1))
        assert serial == 1

    def test_date_to_serial_leap_year_bug(self):
        """Test serial adds 1 after Feb 28, 1900."""
        serial = _date_to_serial(datetime.date(1900, 3, 1))
        assert serial >= 61

    def test_serial_to_time(self):
        """Test serial number to time conversion."""
        # 0.5 = noon (12:00:00)
        time = _serial_to_time(0.5)
        assert time.hour == 12
        assert time.minute == 0

    def test_serial_to_time_quarter_day(self):
        """Test 0.25 = 6:00 AM."""
        time = _serial_to_time(0.25)
        assert time.hour == 6
        assert time.minute == 0

    def test_time_to_serial(self):
        """Test time to serial conversion."""
        serial = _time_to_serial(datetime.time(12, 0, 0))
        assert abs(serial - 0.5) < 0.0001

    def test_time_to_serial_midnight(self):
        """Test midnight is 0."""
        serial = _time_to_serial(datetime.time(0, 0, 0))
        assert serial == 0

    def test_parse_date_string_iso(self):
        """Test ISO date format."""
        date = _parse_date_string("2023-06-15")
        assert date == datetime.date(2023, 6, 15)

    def test_parse_date_string_us(self):
        """Test US date format."""
        date = _parse_date_string("06/15/2023")
        assert date == datetime.date(2023, 6, 15)

    def test_parse_date_string_invalid(self):
        """Test invalid date string."""
        date = _parse_date_string("invalid")
        assert date is None

    def test_to_number_int(self):
        """Test converting int."""
        assert _to_number(42) == 42.0

    def test_to_number_float(self):
        """Test converting float."""
        assert _to_number(3.14) == 3.14

    def test_to_number_string(self):
        """Test converting numeric string."""
        assert _to_number("123") == 123.0

    def test_to_number_string_with_comma(self):
        """Test converting string with commas."""
        assert _to_number("1,234") == 1234.0

    def test_to_number_invalid(self):
        """Test invalid conversion returns 0."""
        assert _to_number("abc") == 0.0


class TestDateFunctions:
    """Tests for date functions."""

    def test_fn_date_basic(self):
        """Test basic DATE function."""
        serial = fn_date(2023, 6, 15)
        date = _serial_to_date(serial)
        assert date.year == 2023
        assert date.month == 6
        assert date.day == 15

    def test_fn_date_two_digit_year_old(self):
        """Test 2-digit year >= 30 maps to 1900s."""
        serial = fn_date(50, 1, 1)
        date = _serial_to_date(serial)
        assert date.year == 1950

    def test_fn_date_two_digit_year_new(self):
        """Test 2-digit year < 30 maps to 2000s."""
        serial = fn_date(25, 1, 1)
        date = _serial_to_date(serial)
        assert date.year == 2025

    def test_fn_date_month_overflow(self):
        """Test month > 12 rolls to next year."""
        serial = fn_date(2023, 14, 1)  # 14th month = Feb next year
        date = _serial_to_date(serial)
        assert date.year == 2024
        assert date.month == 2

    def test_fn_date_month_underflow(self):
        """Test month < 1 rolls to previous year."""
        serial = fn_date(2023, -1, 1)  # -1 month = Nov 2022
        date = _serial_to_date(serial)
        assert date.year == 2022
        assert date.month == 11

    def test_fn_date_invalid(self):
        """Test invalid date returns 0."""
        assert fn_date(2023, 2, 30) == 0  # Feb 30 doesn't exist

    def test_fn_datevalue_iso(self):
        """Test DATEVALUE with ISO format."""
        serial = fn_datevalue("2023-06-15")
        date = _serial_to_date(serial)
        assert date == datetime.date(2023, 6, 15)

    def test_fn_datevalue_us(self):
        """Test DATEVALUE with US format."""
        serial = fn_datevalue("06/15/2023")
        date = _serial_to_date(serial)
        assert date == datetime.date(2023, 6, 15)

    def test_fn_datevalue_invalid(self):
        """Test DATEVALUE with invalid string."""
        assert fn_datevalue("invalid") == 0

    def test_fn_day(self):
        """Test DAY function."""
        serial = fn_date(2023, 6, 15)
        assert fn_day(serial) == 15

    def test_fn_day_invalid(self):
        """Test DAY with invalid serial."""
        assert fn_day(float("inf")) == 0

    def test_fn_month(self):
        """Test MONTH function."""
        serial = fn_date(2023, 6, 15)
        assert fn_month(serial) == 6

    def test_fn_month_invalid(self):
        """Test MONTH with invalid serial."""
        assert fn_month(float("inf")) == 0

    def test_fn_year(self):
        """Test YEAR function."""
        serial = fn_date(2023, 6, 15)
        assert fn_year(serial) == 2023

    def test_fn_year_invalid(self):
        """Test YEAR with invalid serial."""
        assert fn_year(float("inf")) == 0

    def test_fn_weekday_type1(self):
        """Test WEEKDAY with type 1 (Sunday=1)."""
        # Jan 1, 2023 was a Sunday
        serial = fn_date(2023, 1, 1)
        assert fn_weekday(serial, 1) == 1

    def test_fn_weekday_type2(self):
        """Test WEEKDAY with type 2 (Monday=1)."""
        serial = fn_date(2023, 1, 2)  # Monday
        assert fn_weekday(serial, 2) == 1

    def test_fn_weekday_type3(self):
        """Test WEEKDAY with type 3 (Monday=0)."""
        serial = fn_date(2023, 1, 2)  # Monday
        assert fn_weekday(serial, 3) == 0

    def test_fn_weekday_invalid(self):
        """Test WEEKDAY with invalid serial."""
        assert fn_weekday(float("inf")) == 0

    def test_fn_today(self):
        """Test TODAY returns current date."""
        serial = fn_today()
        date = _serial_to_date(serial)
        assert date == datetime.date.today()

    def test_fn_now(self):
        """Test NOW returns current datetime."""
        serial = fn_now()
        # Just check it's a reasonable value (after year 2000)
        assert serial > 36526  # 2000-01-01


class TestTimeFunctions:
    """Tests for time functions."""

    def test_fn_time_basic(self):
        """Test basic TIME function."""
        serial = fn_time(14, 30, 45)
        time = _serial_to_time(serial)
        assert time.hour == 14
        assert time.minute == 30
        assert time.second == 45

    def test_fn_time_overflow_seconds(self):
        """Test TIME with seconds overflow."""
        serial = fn_time(0, 0, 90)  # 90 seconds = 1 min 30 sec
        time = _serial_to_time(serial)
        assert time.minute == 1
        assert time.second == 30

    def test_fn_time_overflow_minutes(self):
        """Test TIME with minutes overflow."""
        serial = fn_time(0, 90, 0)  # 90 minutes = 1 hour 30 min
        time = _serial_to_time(serial)
        assert time.hour == 1
        assert time.minute == 30

    def test_fn_time_overflow_hours(self):
        """Test TIME with hours overflow wraps around."""
        serial = fn_time(25, 0, 0)  # 25 hours = 1 AM
        time = _serial_to_time(serial)
        assert time.hour == 1

    def test_fn_timevalue_24h(self):
        """Test TIMEVALUE with 24-hour format."""
        serial = fn_timevalue("14:30:00")
        time = _serial_to_time(serial)
        assert time.hour == 14
        assert time.minute == 30

    def test_fn_timevalue_short(self):
        """Test TIMEVALUE with HH:MM format."""
        serial = fn_timevalue("14:30")
        time = _serial_to_time(serial)
        assert time.hour == 14
        assert time.minute == 30

    def test_fn_timevalue_12h(self):
        """Test TIMEVALUE with 12-hour format."""
        serial = fn_timevalue("02:30:00 PM")
        time = _serial_to_time(serial)
        assert time.hour == 14

    def test_fn_timevalue_invalid(self):
        """Test TIMEVALUE with invalid string."""
        assert fn_timevalue("invalid") == 0.0

    def test_fn_hour(self):
        """Test HOUR function."""
        serial = fn_time(14, 30, 45)
        assert fn_hour(serial) == 14

    def test_fn_hour_invalid(self):
        """Test HOUR with invalid serial."""
        assert fn_hour(float("inf")) == 0

    def test_fn_minute(self):
        """Test MINUTE function."""
        serial = fn_time(14, 30, 45)
        assert fn_minute(serial) == 30

    def test_fn_minute_invalid(self):
        """Test MINUTE with invalid serial."""
        assert fn_minute(float("inf")) == 0

    def test_fn_second(self):
        """Test SECOND function."""
        serial = fn_time(14, 30, 45)
        assert fn_second(serial) == 45

    def test_fn_second_invalid(self):
        """Test SECOND with invalid serial."""
        assert fn_second(float("inf")) == 0


class TestDateCalculations:
    """Tests for date calculation functions."""

    def test_fn_days_positive(self):
        """Test DAYS with positive difference."""
        start = fn_date(2023, 1, 1)
        end = fn_date(2023, 1, 10)
        assert fn_days(end, start) == 9

    def test_fn_days_negative(self):
        """Test DAYS with negative difference."""
        start = fn_date(2023, 1, 10)
        end = fn_date(2023, 1, 1)
        assert fn_days(end, start) == -9

    def test_fn_edate_add_months(self):
        """Test EDATE adding months."""
        start = fn_date(2023, 1, 15)
        result = fn_edate(start, 3)
        date = _serial_to_date(result)
        assert date.month == 4
        assert date.day == 15

    def test_fn_edate_subtract_months(self):
        """Test EDATE subtracting months."""
        start = fn_date(2023, 4, 15)
        result = fn_edate(start, -3)
        date = _serial_to_date(result)
        assert date.month == 1

    def test_fn_edate_day_overflow(self):
        """Test EDATE handles day overflow."""
        # Jan 31 + 1 month = Feb 28 (not 31)
        start = fn_date(2023, 1, 31)
        result = fn_edate(start, 1)
        date = _serial_to_date(result)
        assert date.month == 2
        assert date.day == 28

    def test_fn_edate_invalid(self):
        """Test EDATE with invalid input."""
        assert fn_edate(float("inf"), 1) == 0

    def test_fn_eomonth_basic(self):
        """Test EOMONTH basic usage."""
        start = fn_date(2023, 1, 15)
        result = fn_eomonth(start, 0)  # End of same month
        date = _serial_to_date(result)
        assert date.month == 1
        assert date.day == 31

    def test_fn_eomonth_add_months(self):
        """Test EOMONTH adding months."""
        start = fn_date(2023, 1, 15)
        result = fn_eomonth(start, 1)  # End of Feb
        date = _serial_to_date(result)
        assert date.month == 2
        assert date.day == 28

    def test_fn_eomonth_subtract_months(self):
        """Test EOMONTH subtracting months."""
        start = fn_date(2023, 3, 15)
        result = fn_eomonth(start, -1)  # End of Feb
        date = _serial_to_date(result)
        assert date.month == 2

    def test_fn_eomonth_invalid(self):
        """Test EOMONTH with invalid input."""
        assert fn_eomonth(float("inf"), 1) == 0

    def test_fn_yearfrac_basis_0(self):
        """Test YEARFRAC with 30/360 basis."""
        start = fn_date(2023, 1, 1)
        end = fn_date(2023, 7, 1)  # 180 days in 30/360
        result = fn_yearfrac(start, end, 0)
        assert abs(result - 0.5) < 0.1

    def test_fn_yearfrac_basis_1(self):
        """Test YEARFRAC with actual/actual basis."""
        start = fn_date(2023, 1, 1)
        end = fn_date(2024, 1, 1)  # Full year
        result = fn_yearfrac(start, end, 1)
        assert abs(result - 1.0) < 0.1

    def test_fn_yearfrac_basis_2(self):
        """Test YEARFRAC with actual/360 basis."""
        start = fn_date(2023, 1, 1)
        end = fn_date(2023, 12, 27)  # ~360 days later
        result = fn_yearfrac(start, end, 2)
        assert abs(result - 1.0) < 0.1

    def test_fn_yearfrac_basis_3(self):
        """Test YEARFRAC with actual/365 basis."""
        start = fn_date(2023, 1, 1)
        end = fn_date(2024, 1, 1)
        result = fn_yearfrac(start, end, 3)
        assert result > 0.9

    def test_fn_yearfrac_basis_4(self):
        """Test YEARFRAC with 30/360 European basis."""
        start = fn_date(2023, 1, 1)
        end = fn_date(2023, 7, 1)
        result = fn_yearfrac(start, end, 4)
        assert abs(result - 0.5) < 0.1

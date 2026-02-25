"""Tests for utility functions — pure unit tests, no CLI invocations."""

import datetime

import pytest

from agent_xlsx.formatters.json_formatter import _serialise
from agent_xlsx.formatters.token_optimizer import cap_list, summarise_formulas
from agent_xlsx.utils.dates import (
    detect_date_columns,
    excel_serial_to_isodate,
)
from agent_xlsx.utils.errors import (
    InvalidColumnError,
    RangeInvalidError,
)
from agent_xlsx.utils.validation import (
    _normalise_shell_ref,
    col_letter_to_index,
    file_size_human,
    index_to_col_letter,
    parse_multi_range,
    parse_range,
    resolve_column_filter,
)

# ---------------------------------------------------------------------------
# validation.py — parse_range
# ---------------------------------------------------------------------------


class TestParseRange:
    """Tests for parse_range: Excel range string parsing."""

    def test_single_cell(self):
        result = parse_range("A1")
        assert result == {"sheet": None, "start": "A1", "end": None}

    def test_cell_range(self):
        result = parse_range("A1:C10")
        assert result == {"sheet": None, "start": "A1", "end": "C10"}

    def test_sheet_prefixed_range(self):
        result = parse_range("Sheet1!A1:C10")
        assert result == {"sheet": "Sheet1", "start": "A1", "end": "C10"}

    def test_sheet_prefixed_single_cell(self):
        result = parse_range("Sales!B5")
        assert result == {"sheet": "Sales", "start": "B5", "end": None}

    def test_case_insensitive_columns(self):
        result = parse_range("a1:c10")
        assert result["start"] == "A1"
        assert result["end"] == "C10"

    def test_three_letter_column(self):
        result = parse_range("AAA1:ZZZ999")
        assert result == {"sheet": None, "start": "AAA1", "end": "ZZZ999"}

    def test_sheet_with_spaces(self):
        """Sheet names can contain spaces when part of the bang-separated ref."""
        result = parse_range("My Sheet!A1:B2")
        assert result["sheet"] == "My Sheet"

    def test_sheet_with_year_name(self):
        """Numeric sheet names like '2022' are valid."""
        result = parse_range("2022!B1:C5")
        assert result == {"sheet": "2022", "start": "B1", "end": "C5"}

    def test_invalid_range_raises(self):
        with pytest.raises(RangeInvalidError):
            parse_range("not_a_range")

    def test_empty_string_raises(self):
        with pytest.raises(RangeInvalidError):
            parse_range("")

    def test_whitespace_stripped(self):
        result = parse_range("  A1:B2  ")
        assert result == {"sheet": None, "start": "A1", "end": "B2"}


# ---------------------------------------------------------------------------
# validation.py — parse_multi_range
# ---------------------------------------------------------------------------


class TestParseMultiRange:
    """Tests for parse_multi_range: comma-separated range parsing."""

    def test_single_range(self):
        result = parse_multi_range("A1:C10")
        assert len(result) == 1
        assert result[0] == {"sheet": None, "start": "A1", "end": "C10"}

    def test_two_ranges(self):
        result = parse_multi_range("A1:C10,E1:G10")
        assert len(result) == 2
        assert result[0]["start"] == "A1"
        assert result[1]["start"] == "E1"

    def test_sheet_context_carries_forward(self):
        """Sheet prefix from the first range carries to subsequent ranges."""
        result = parse_multi_range("2022!H54:AT54,H149:AT149")
        assert result[0]["sheet"] == "2022"
        assert result[1]["sheet"] == "2022"

    def test_second_range_overrides_sheet(self):
        """Explicit sheet on second range overrides inherited context."""
        result = parse_multi_range("Sheet1!A1:B2,Sheet2!C3:D4")
        assert result[0]["sheet"] == "Sheet1"
        assert result[1]["sheet"] == "Sheet2"

    def test_spaces_around_commas(self):
        result = parse_multi_range("A1:B2 , C3:D4")
        assert len(result) == 2
        assert result[0]["start"] == "A1"
        assert result[1]["start"] == "C3"


# ---------------------------------------------------------------------------
# validation.py — col_letter_to_index / index_to_col_letter
# ---------------------------------------------------------------------------


class TestColumnConversion:
    """Tests for column letter <-> index conversion."""

    def test_a_is_zero(self):
        assert col_letter_to_index("A") == 0

    def test_z_is_25(self):
        assert col_letter_to_index("Z") == 25

    def test_aa_is_26(self):
        assert col_letter_to_index("AA") == 26

    def test_az_is_51(self):
        assert col_letter_to_index("AZ") == 51

    def test_case_insensitive(self):
        assert col_letter_to_index("a") == col_letter_to_index("A")

    def test_index_to_a(self):
        assert index_to_col_letter(0) == "A"

    def test_index_to_z(self):
        assert index_to_col_letter(25) == "Z"

    def test_index_to_aa(self):
        assert index_to_col_letter(26) == "AA"

    def test_roundtrip(self):
        """Converting index -> letter -> index should be identity."""
        for i in range(0, 100):
            assert col_letter_to_index(index_to_col_letter(i)) == i


# ---------------------------------------------------------------------------
# validation.py — _normalise_shell_ref
# ---------------------------------------------------------------------------


class TestNormaliseShellRef:
    """Tests for _normalise_shell_ref: zsh escape normalisation."""

    def test_escaped_bang(self):
        assert _normalise_shell_ref("2022\\!B1") == "2022!B1"

    def test_no_escaping_needed(self):
        assert _normalise_shell_ref("Sheet1!A1") == "Sheet1!A1"

    def test_multiple_escapes(self):
        assert _normalise_shell_ref("a\\!b\\!c") == "a!b!c"


# ---------------------------------------------------------------------------
# validation.py — file_size_human
# ---------------------------------------------------------------------------


class TestFileSizeHuman:
    """Tests for file_size_human: human-readable file sizes."""

    def test_bytes(self, tmp_path):
        f = tmp_path / "tiny.txt"
        f.write_bytes(b"x" * 100)
        assert file_size_human(str(f)) == "100 B"

    def test_kilobytes(self, tmp_path):
        f = tmp_path / "kb.txt"
        f.write_bytes(b"x" * 5120)  # 5 KB
        assert file_size_human(str(f)) == "5.0 KB"

    def test_megabytes(self, tmp_path):
        f = tmp_path / "mb.txt"
        f.write_bytes(b"x" * (2 * 1024 * 1024))  # 2 MB
        assert file_size_human(str(f)) == "2.0 MB"


# ---------------------------------------------------------------------------
# validation.py — resolve_column_filter
# ---------------------------------------------------------------------------


class TestResolveColumnFilter:
    """Tests for resolve_column_filter: column spec resolution."""

    def test_exact_name_match(self):
        result = resolve_column_filter("Name,Score", ["Name", "Score", "Grade"])
        assert result == ["Name", "Score"]

    def test_column_letter_match(self):
        result = resolve_column_filter("A,C", ["Name", "Score", "Grade"])
        assert result == ["Name", "Grade"]

    def test_case_insensitive_letter(self):
        result = resolve_column_filter("a,c", ["Name", "Score", "Grade"])
        assert result == ["Name", "Grade"]

    def test_invalid_column_raises(self):
        with pytest.raises(InvalidColumnError):
            resolve_column_filter("X,Y", ["Name", "Score"])

    def test_deduplication(self):
        """Same column referenced twice should appear once."""
        result = resolve_column_filter("A,A", ["Name", "Score"])
        assert result == ["Name"]

    def test_header_name_fallback_resolution(self):
        """When df_columns are letters but headers are provided, header names resolve."""
        result = resolve_column_filter(
            "Revenue",
            df_columns=["A", "B", "C"],
            headers=["Date", "Revenue", "Qty"],
        )
        assert result == ["B"]


# ---------------------------------------------------------------------------
# dates.py — excel_serial_to_isodate
# ---------------------------------------------------------------------------


class TestExcelSerialToIsodate:
    """Tests for excel_serial_to_isodate: serial number conversion."""

    def test_date_only(self):
        # 45000 = 2023-03-15 (Excel serial from 1899-12-30 epoch)
        result = excel_serial_to_isodate(45000.0)
        assert result == "2023-03-15"

    def test_date_with_time(self):
        # .5 = noon
        result = excel_serial_to_isodate(45000.5)
        assert result is not None
        assert "T" in result  # contains time component
        assert result.startswith("2023-03-15")

    def test_nan_returns_none(self):
        assert excel_serial_to_isodate(float("nan")) is None

    def test_zero_returns_as_is(self):
        assert excel_serial_to_isodate(0) == 0

    def test_negative_returns_as_is(self):
        assert excel_serial_to_isodate(-1) == -1

    def test_known_date(self):
        # 1 = 1899-12-31 (the day after the epoch)
        assert excel_serial_to_isodate(1.0) == "1899-12-31"


# ---------------------------------------------------------------------------
# dates.py — detect_date_columns
# ---------------------------------------------------------------------------


class TestDetectDateColumns:
    """Tests for detect_date_columns: date format detection from workbooks."""

    def test_detects_date_formatted_column(self, rich_xlsx):
        """The rich_xlsx fixture has a 'Date' column with yyyy-mm-dd format."""
        result = detect_date_columns(str(rich_xlsx), sheet_name="Sales")
        assert "Date" in result
        assert result["Date"] is True

    def test_non_date_columns_excluded(self, rich_xlsx):
        result = detect_date_columns(str(rich_xlsx), sheet_name="Sales")
        assert "Product" not in result
        assert "Revenue" not in result

    def test_missing_sheet_returns_empty(self, rich_xlsx):
        result = detect_date_columns(str(rich_xlsx), sheet_name="NonExistent")
        assert result == {}


# ---------------------------------------------------------------------------
# token_optimizer.py — cap_list
# ---------------------------------------------------------------------------


class TestCapList:
    """Tests for cap_list: list truncation with metadata."""

    def test_under_cap(self):
        result = cap_list([1, 2, 3], max_count=5)
        assert result["items"] == [1, 2, 3]
        assert result["total"] == 3
        assert result["truncated"] is False

    def test_at_cap(self):
        result = cap_list([1, 2, 3], max_count=3)
        assert result["items"] == [1, 2, 3]
        assert result["truncated"] is False

    def test_over_cap(self):
        result = cap_list([1, 2, 3, 4, 5], max_count=2)
        assert result["items"] == [1, 2]
        assert result["total"] == 5
        assert result["truncated"] is True

    def test_empty_list(self):
        result = cap_list([], max_count=10)
        assert result["items"] == []
        assert result["total"] == 0
        assert result["truncated"] is False


# ---------------------------------------------------------------------------
# token_optimizer.py — summarise_formulas
# ---------------------------------------------------------------------------


class TestSummariseFormulas:
    """Tests for summarise_formulas: formula aggregation."""

    def test_basic_summary(self):
        cells = [
            {"cell": "B2", "formula": "=SUM(A1:A10)"},
            {"cell": "B3", "formula": "=AVERAGE(A1:A10)"},
            {"cell": "C2", "formula": "=A2+B2"},
        ]
        result = summarise_formulas(cells)
        assert result["formula_count"] == 3
        assert set(result["formula_columns"]) == {"B", "C"}
        assert result["truncated"] is False

    def test_truncation(self):
        # Create more cells than the default max_count of 50
        cells = [{"cell": f"A{i}", "formula": f"=SUM(B{i})"} for i in range(1, 101)]
        result = summarise_formulas(cells, max_count=50)
        assert result["formula_count"] == 100
        assert len(result["sample_formulas"]) == 50
        assert result["truncated"] is True

    def test_empty_cells(self):
        result = summarise_formulas([])
        assert result["formula_count"] == 0
        assert result["formula_columns"] == []
        assert result["truncated"] is False


# ---------------------------------------------------------------------------
# json_formatter.py — _serialise
# ---------------------------------------------------------------------------


class TestSerialise:
    """Tests for _serialise: JSON serialisation of non-standard types."""

    def test_datetime_midnight(self):
        """Midnight datetimes normalise to date-only strings."""
        dt = datetime.datetime(2024, 1, 15, 0, 0, 0)
        assert _serialise(dt) == "2024-01-15"

    def test_datetime_with_time(self):
        dt = datetime.datetime(2024, 1, 15, 14, 30, 0)
        assert _serialise(dt) == "2024-01-15T14:30:00"

    def test_date(self):
        d = datetime.date(2024, 6, 1)
        assert _serialise(d) == "2024-06-01"

    def test_timedelta(self):
        td = datetime.timedelta(days=5, hours=3)
        assert _serialise(td) == "5 days, 3:00:00"

    def test_float_like_whole_number(self):
        """Objects with __float__ that are whole numbers become ints."""

        class FakeNum:
            def __float__(self):
                return 42.0

        assert _serialise(FakeNum()) == 42

    def test_float_like_fractional(self):
        """Objects with __float__ that are fractional stay as floats."""

        class FakeNum:
            def __float__(self):
                return 3.14

        assert _serialise(FakeNum()) == 3.14

    def test_fallback_to_str(self):
        """Unknown types fall back to str()."""

        class Custom:
            def __str__(self):
                return "custom-repr"

        assert _serialise(Custom()) == "custom-repr"

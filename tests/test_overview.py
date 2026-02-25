"""Tests for the overview command — structural metadata overview of a workbook."""

import json

from typer.testing import CliRunner

from agent_xlsx.cli import app

runner = CliRunner()


# ---------------------------------------------------------------------------
# Default output structure
# ---------------------------------------------------------------------------


def test_overview_default_structure(sample_xlsx):
    """Default overview returns expected top-level keys and data origin tag."""
    result = runner.invoke(app, ["overview", str(sample_xlsx)])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    # output_spreadsheet_data wraps with _data_origin
    assert data["_data_origin"] == "untrusted_spreadsheet"

    # Top-level keys present
    for key in (
        "file",
        "size_bytes",
        "overview_time_ms",
        "sheets",
        "named_ranges",
        "named_range_count",
        "has_vba",
        "vba_module_count",
        "total_formula_count",
        "total_chart_count",
    ):
        assert key in data, f"Missing key: {key}"

    assert data["file"] == "sample.xlsx"
    assert isinstance(data["sheets"], list)
    assert len(data["sheets"]) == 1


def test_overview_sheet_info_structure(sample_xlsx):
    """Each sheet entry has the expected structural fields."""
    result = runner.invoke(app, ["overview", str(sample_xlsx)])
    data = json.loads(result.stdout)
    sheet = data["sheets"][0]

    assert sheet["name"] == "Sheet1"
    assert sheet["index"] == 0
    assert "dimensions" in sheet
    assert "row_count" in sheet
    assert "col_count" in sheet
    assert "has_formulas" in sheet
    assert "formula_count" in sheet
    assert "has_charts" in sheet
    assert "chart_count" in sheet
    assert "has_tables" in sheet


# ---------------------------------------------------------------------------
# Formula detection and counting
# ---------------------------------------------------------------------------


def test_overview_formula_count_accuracy(rich_xlsx):
    """Formula count matches the number of formulas in the rich fixture."""
    result = runner.invoke(app, ["overview", str(rich_xlsx)])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    # rich_xlsx has 3 formulas in Summary sheet: B3, B4, B5
    assert data["total_formula_count"] == 3

    # Sales sheet has no formulas
    sales_sheet = next(s for s in data["sheets"] if s["name"] == "Sales")
    assert sales_sheet["has_formulas"] is False
    assert sales_sheet["formula_count"] == 0

    # Summary sheet has all 3 formulas
    summary_sheet = next(s for s in data["sheets"] if s["name"] == "Summary")
    assert summary_sheet["has_formulas"] is True
    assert summary_sheet["formula_count"] == 3


def test_overview_no_formulas_in_plain_workbook(sample_xlsx):
    """Workbook with no formulas reports zero formula count."""
    result = runner.invoke(app, ["overview", str(sample_xlsx)])
    data = json.loads(result.stdout)

    assert data["total_formula_count"] == 0
    assert data["sheets"][0]["has_formulas"] is False
    assert data["sheets"][0]["formula_count"] == 0


# ---------------------------------------------------------------------------
# --include-formulas flag
# ---------------------------------------------------------------------------


def test_overview_include_formulas(rich_xlsx):
    """--include-formulas adds sample_formulas with pattern deduplication."""
    result = runner.invoke(app, ["overview", str(rich_xlsx), "--include-formulas"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    summary_sheet = next(s for s in data["sheets"] if s["name"] == "Summary")
    assert "sample_formulas" in summary_sheet

    formulas = summary_sheet["sample_formulas"]
    assert isinstance(formulas, list)
    assert len(formulas) > 0

    # Each formula entry should have pattern, example_cell, example, count
    for f in formulas:
        assert "pattern" in f
        assert "example_cell" in f
        assert "example" in f
        assert "count" in f


# ---------------------------------------------------------------------------
# --include-formatting flag
# ---------------------------------------------------------------------------


def test_overview_include_formatting(rich_xlsx):
    """--include-formatting reports merged cell info for sheets that have them."""
    result = runner.invoke(app, ["overview", str(rich_xlsx), "--include-formatting"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    # Summary sheet has merged cells (A1:C1)
    summary_sheet = next(s for s in data["sheets"] if s["name"] == "Summary")
    assert "has_merged_cells" in summary_sheet
    assert summary_sheet["has_merged_cells"] is True
    assert summary_sheet["merged_cell_count"] >= 1

    # Sales sheet has no merged cells
    sales_sheet = next(s for s in data["sheets"] if s["name"] == "Sales")
    assert sales_sheet["has_merged_cells"] is False


def test_overview_no_formatting_by_default(rich_xlsx):
    """Without --include-formatting, merged cell info is absent."""
    result = runner.invoke(app, ["overview", str(rich_xlsx)])
    data = json.loads(result.stdout)

    for sheet in data["sheets"]:
        assert "has_merged_cells" not in sheet


# ---------------------------------------------------------------------------
# Named ranges
# ---------------------------------------------------------------------------


def test_overview_named_ranges_empty(sample_xlsx):
    """Workbook without named ranges reports empty list and zero count."""
    result = runner.invoke(app, ["overview", str(sample_xlsx)])
    data = json.loads(result.stdout)

    assert data["named_ranges"] == []
    assert data["named_range_count"] == 0


# ---------------------------------------------------------------------------
# VBA detection
# ---------------------------------------------------------------------------


def test_overview_has_vba_false_for_xlsx(sample_xlsx):
    """A .xlsx file reports has_vba: false."""
    result = runner.invoke(app, ["overview", str(sample_xlsx)])
    data = json.loads(result.stdout)

    assert data["has_vba"] is False
    assert data["vba_module_count"] == 0


# ---------------------------------------------------------------------------
# Empty workbook edge case
# ---------------------------------------------------------------------------


def test_overview_empty_workbook(tmp_path):
    """Overview handles a workbook with no data gracefully."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Empty"
    # No data written — completely empty sheet
    p = tmp_path / "empty.xlsx"
    wb.save(p)

    result = runner.invoke(app, ["overview", str(p)])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    assert data["file"] == "empty.xlsx"
    assert data["total_formula_count"] == 0
    assert data["total_chart_count"] == 0
    assert len(data["sheets"]) == 1
    assert data["sheets"][0]["has_formulas"] is False

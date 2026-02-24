"""Tests for the format command — multi-range support."""

import json

import pytest
from openpyxl import Workbook
from typer.testing import CliRunner

from agent_xlsx.cli import app

runner = CliRunner()


@pytest.fixture()
def styled_xlsx(tmp_path):
    """Workbook with data for formatting tests."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Header A"
    ws["B1"] = "Header B"
    ws["C1"] = "Header C"
    ws["A2"] = 1
    ws["B2"] = 2
    ws["C2"] = 3
    ws["A3"] = 4
    ws["B3"] = 5
    ws["C3"] = 6
    ws["A4"] = ""
    ws["B4"] = "Summary"
    ws["C4"] = ""
    ws["A5"] = 10
    ws["B5"] = 20
    ws["C5"] = 30
    p = tmp_path / "styled.xlsx"
    wb.save(p)
    return p


# ---------------------------------------------------------------------------
# Multi-range apply
# ---------------------------------------------------------------------------


def test_format_multi_range_apply_bold(styled_xlsx):
    """Multi-range applies same formatting to all ranges."""
    result = runner.invoke(
        app,
        ["format", str(styled_xlsx), "A1:C1,B4", "--font", '{"bold": true}'],
    )
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["status"] == "success"
    assert data["ranges_formatted"] == 2
    assert data["total_cells_formatted"] == 4  # 3 cells in A1:C1 + 1 in B4

    # Verify formatting was applied to both ranges
    r1 = runner.invoke(app, ["format", str(styled_xlsx), "A1", "--read"])
    assert r1.exit_code == 0
    assert json.loads(r1.stdout)["font"]["bold"] is True

    r2 = runner.invoke(app, ["format", str(styled_xlsx), "B4", "--read"])
    assert r2.exit_code == 0
    assert json.loads(r2.stdout)["font"]["bold"] is True


def test_format_multi_range_apply_number_format(styled_xlsx):
    """Multi-range number format applies to all ranges."""
    result = runner.invoke(
        app,
        ["format", str(styled_xlsx), "A2:C2,A5:C5", "--number-format", "#,##0.00"],
    )
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["ranges_formatted"] == 2
    assert data["total_cells_formatted"] == 6


# ---------------------------------------------------------------------------
# Multi-range read
# ---------------------------------------------------------------------------


def test_format_multi_range_read(styled_xlsx):
    """Multi-range read returns formatting per range."""
    # Apply bold to A1 first
    runner.invoke(
        app,
        ["format", str(styled_xlsx), "A1", "--font", '{"bold": true}'],
    )

    result = runner.invoke(
        app,
        ["format", str(styled_xlsx), "A1,B4", "--read"],
    )
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["total_ranges"] == 2
    assert len(data["results"]) == 2
    # A1 should be bold, B4 should not
    assert data["results"][0]["formatting"]["font"]["bold"] is True
    assert data["results"][1]["formatting"]["font"]["bold"] is not True


# ---------------------------------------------------------------------------
# Multi-range copy
# ---------------------------------------------------------------------------


def test_format_multi_range_copy(styled_xlsx):
    """Multi-range copy applies source formatting to all target ranges."""
    # Make A1 bold
    runner.invoke(
        app,
        ["format", str(styled_xlsx), "A1", "--font", '{"bold": true}'],
    )

    # Copy A1's formatting to B4 and C5
    result = runner.invoke(
        app,
        ["format", str(styled_xlsx), "B4,C5", "--copy-from", "A1"],
    )
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["status"] == "success"
    assert data["ranges_formatted"] == 2

    # Verify both cells got bold
    for cell_ref in ("B4", "C5"):
        r = runner.invoke(app, ["format", str(styled_xlsx), cell_ref, "--read"])
        assert json.loads(r.stdout)["font"]["bold"] is True


# ---------------------------------------------------------------------------
# Multi-range with sheet prefix
# ---------------------------------------------------------------------------


def test_format_multi_range_with_sheet_prefix(styled_xlsx):
    """Sheet prefix in first range carries forward to subsequent ranges."""
    result = runner.invoke(
        app,
        ["format", str(styled_xlsx), "Sheet1!A1:C1,B4", "--font", '{"bold": true}'],
    )
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["ranges_formatted"] == 2


# ---------------------------------------------------------------------------
# Single-range backward compatibility
# ---------------------------------------------------------------------------


def test_format_single_range_unchanged(styled_xlsx):
    """Single-range formatting still works (backward compat)."""
    result = runner.invoke(
        app,
        ["format", str(styled_xlsx), "A1:C1", "--font", '{"bold": true}'],
    )
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["status"] == "success"
    assert data["cells_formatted"] == 3

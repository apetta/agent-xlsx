"""Tests for the recalc command — formula error checking and engine resolution."""

import json
from unittest.mock import patch

from openpyxl import Workbook
from typer.testing import CliRunner

from agent_xlsx.cli import app

runner = CliRunner()


# ---------------------------------------------------------------------------
# --check-only: clean workbook (no errors)
# ---------------------------------------------------------------------------


def test_recalc_check_only_clean(sample_xlsx):
    """--check-only on a workbook with no errors returns success status."""
    result = runner.invoke(app, ["recalc", str(sample_xlsx), "--check-only"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    assert data["status"] == "success"
    assert data["mode"] == "check_only"
    assert data["total_errors"] == 0
    assert "error_summary" not in data
    assert isinstance(data["check_time_ms"], (int, float))


# ---------------------------------------------------------------------------
# --check-only: errors found
# ---------------------------------------------------------------------------


def test_recalc_check_only_finds_errors(formula_error_xlsx):
    """--check-only detects cached error string values."""
    result = runner.invoke(app, ["recalc", str(formula_error_xlsx), "--check-only"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    assert data["status"] == "errors_found"
    assert data["total_errors"] == 4  # #REF! x2, #DIV/0! x1, #NAME? x1


# ---------------------------------------------------------------------------
# --check-only: error summary structure
# ---------------------------------------------------------------------------


def test_recalc_check_only_error_summary_structure(formula_error_xlsx):
    """Error summary groups by error type with count and locations."""
    result = runner.invoke(app, ["recalc", str(formula_error_xlsx), "--check-only"])
    data = json.loads(result.stdout)

    assert "error_summary" in data
    summary = data["error_summary"]

    # Should have entries for each error type found
    assert "#REF!" in summary
    assert "#DIV/0!" in summary
    assert "#NAME?" in summary

    # Each entry has count and locations
    for error_type, info in summary.items():
        assert "count" in info
        assert "locations" in info
        assert isinstance(info["locations"], list)
        # Locations include sheet name and cell coordinate
        for loc in info["locations"]:
            assert "!" in loc  # format: "SheetName!CellRef"

    # #REF! appears twice in the fixture
    assert summary["#REF!"]["count"] == 2
    assert len(summary["#REF!"]["locations"]) == 2


# ---------------------------------------------------------------------------
# --check-only: formula counting
# ---------------------------------------------------------------------------


def test_recalc_check_only_counts_formulas(formula_error_xlsx):
    """--check-only reports total formula count from a separate pass."""
    result = runner.invoke(app, ["recalc", str(formula_error_xlsx), "--check-only"])
    data = json.loads(result.stdout)

    # formula_error_xlsx has 2 actual formulas: B2=SUM(A2:A5), B3=A2+A3
    assert data["total_formulas"] == 2


# ---------------------------------------------------------------------------
# --check-only: location cap at 20
# ---------------------------------------------------------------------------


def test_recalc_check_only_location_cap(tmp_path):
    """Locations per error type are capped at 20 entries."""
    wb = Workbook()
    ws = wb.active
    ws.title = "ManyErrors"
    # Create 30 #REF! cached error values
    for i in range(1, 31):
        ws.cell(row=i, column=1, value="#REF!")
    p = tmp_path / "many_errors.xlsx"
    wb.save(p)

    result = runner.invoke(app, ["recalc", str(p), "--check-only"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    assert data["total_errors"] == 30
    # Locations capped at 20
    assert len(data["error_summary"]["#REF!"]["locations"]) == 20
    # But count reflects all 30
    assert data["error_summary"]["#REF!"]["count"] == 30


# ---------------------------------------------------------------------------
# Engine error path: all backends unavailable
# ---------------------------------------------------------------------------


def test_recalc_no_engine_available(sample_xlsx):
    """Without --check-only and no engine available, returns structured error."""
    with patch(
        "agent_xlsx.utils.engine.resolve_engine",
        side_effect=_import_no_backend_error(),
    ):
        result = runner.invoke(app, ["recalc", str(sample_xlsx)])

    # handle_error catches AgentExcelError and exits with code 1
    assert result.exit_code == 1
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "NO_RENDERING_BACKEND"


def test_recalc_check_only_skips_engine(sample_xlsx):
    """--check-only does not attempt engine resolution at all."""
    # Patch resolve_engine to raise — if called, the test will fail
    with patch(
        "agent_xlsx.utils.engine.resolve_engine",
        side_effect=AssertionError("resolve_engine should not be called"),
    ):
        result = runner.invoke(app, ["recalc", str(sample_xlsx), "--check-only"])

    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["mode"] == "check_only"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _import_no_backend_error():
    """Return a NoRenderingBackendError instance for mocking."""
    from agent_xlsx.utils.errors import NoRenderingBackendError

    return NoRenderingBackendError("recalc")

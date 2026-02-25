"""Tests for objects command: error paths and input validation.

All tests mock engine availability since no rendering engine is available in CI.
"""

import json

from typer.testing import CliRunner

from agent_xlsx.cli import app

runner = CliRunner()


# ---------------------------------------------------------------------------
# File validation errors
# ---------------------------------------------------------------------------


def test_objects_file_not_found():
    """Non-existent file produces a FILE_NOT_FOUND error."""
    result = runner.invoke(app, ["objects", "/tmp/nonexistent_xyz.xlsx"])
    assert result.exit_code == 1
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "FILE_NOT_FOUND"


def test_objects_invalid_format(tmp_path):
    """A non-Excel file extension produces an INVALID_FORMAT error."""
    txt_file = tmp_path / "notes.txt"
    txt_file.write_text("not an excel file")
    result = runner.invoke(app, ["objects", str(txt_file)])
    assert result.exit_code == 1
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "INVALID_FORMAT"


# ---------------------------------------------------------------------------
# Engine selection: LibreOffice blocked (objects passes libreoffice=False)
# ---------------------------------------------------------------------------


def test_objects_engine_libreoffice_blocked(sample_xlsx):
    """--engine libreoffice raises EXCEL_REQUIRED because objects has no LO adapter."""
    result = runner.invoke(app, ["objects", str(sample_xlsx), "--engine", "libreoffice"])
    assert result.exit_code == 1
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "EXCEL_REQUIRED"


def test_objects_engine_lo_alias_blocked(sample_xlsx):
    """--engine lo (alias) also raises EXCEL_REQUIRED — same as 'libreoffice'."""
    result = runner.invoke(app, ["objects", str(sample_xlsx), "--engine", "lo"])
    assert result.exit_code == 1
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "EXCEL_REQUIRED"


# ---------------------------------------------------------------------------
# Engine selection: no engine available (auto-detection with all unavailable)
# ---------------------------------------------------------------------------


def test_objects_no_engine_available(monkeypatch, sample_xlsx):
    """When all engines are unavailable, auto-detection raises EXCEL_REQUIRED.

    Because objects passes libreoffice=False to resolve_engine, the auto path
    skips LibreOffice and raises ExcelRequiredError (not NoRenderingBackendError).
    """
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter.is_aspose_available", lambda: False)
    monkeypatch.setattr("agent_xlsx.adapters.xlwings_adapter.is_excel_available", lambda: False)
    result = runner.invoke(app, ["objects", str(sample_xlsx)])
    assert result.exit_code == 1
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "EXCEL_REQUIRED"


# ---------------------------------------------------------------------------
# Engine selection: explicit engine unavailable
# ---------------------------------------------------------------------------


def test_objects_engine_excel_unavailable(monkeypatch, sample_xlsx):
    """--engine excel with Excel not installed raises EXCEL_REQUIRED."""
    monkeypatch.setattr("agent_xlsx.adapters.xlwings_adapter.is_excel_available", lambda: False)
    result = runner.invoke(app, ["objects", str(sample_xlsx), "--engine", "excel"])
    assert result.exit_code == 1
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "EXCEL_REQUIRED"


def test_objects_engine_aspose_unavailable(monkeypatch, sample_xlsx):
    """--engine aspose with Aspose not installed raises ASPOSE_NOT_INSTALLED."""
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter.is_aspose_available", lambda: False)
    result = runner.invoke(app, ["objects", str(sample_xlsx), "--engine", "aspose"])
    assert result.exit_code == 1
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "ASPOSE_NOT_INSTALLED"

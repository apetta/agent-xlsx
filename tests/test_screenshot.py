"""Tests for screenshot command: error paths and input validation.

All tests mock engine availability since no rendering engine is guaranteed in CI.
"""

import json

from typer.testing import CliRunner

from agent_xlsx.cli import app

runner = CliRunner()


def _disable_all_engines(monkeypatch):
    """Monkeypatch all rendering engines as unavailable."""
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter._ASPOSE_AVAILABLE", None)
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter.is_aspose_available", lambda: False)
    monkeypatch.setattr("agent_xlsx.adapters.xlwings_adapter.is_excel_available", lambda: False)
    monkeypatch.setattr(
        "agent_xlsx.adapters.libreoffice_adapter.is_libreoffice_available",
        lambda: False,
    )


# ---------------------------------------------------------------------------
# File validation errors
# ---------------------------------------------------------------------------


def test_screenshot_file_not_found(monkeypatch):
    """Non-existent file produces FILE_NOT_FOUND error JSON."""
    _disable_all_engines(monkeypatch)

    result = runner.invoke(app, ["screenshot", "/tmp/nonexistent_file.xlsx"])
    assert result.exit_code == 1

    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "FILE_NOT_FOUND"


def test_screenshot_invalid_format(tmp_path, monkeypatch):
    """A non-Excel file extension produces INVALID_FORMAT error JSON."""
    _disable_all_engines(monkeypatch)

    txt_file = tmp_path / "notes.txt"
    txt_file.write_text("not an excel file")

    result = runner.invoke(app, ["screenshot", str(txt_file)])
    assert result.exit_code == 1

    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "INVALID_FORMAT"


# ---------------------------------------------------------------------------
# Range validation errors
# ---------------------------------------------------------------------------


def test_screenshot_invalid_range(sample_xlsx, monkeypatch):
    """Malformed range string produces RANGE_INVALID error JSON."""
    # Need at least one engine "available" so resolve_engine doesn't fail first,
    # but parse_range runs before engine resolution, so no engine mock needed.
    # Actually, looking at the command flow: validate_file -> parse_range -> resolve_engine.
    # parse_range happens before resolve_engine, so we just need a valid file.
    _disable_all_engines(monkeypatch)

    result = runner.invoke(app, ["screenshot", str(sample_xlsx), "ZZZ"])
    assert result.exit_code == 1

    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "RANGE_INVALID"


def test_screenshot_invalid_range_positional_with_sheet_prefix(sample_xlsx, monkeypatch):
    """Malformed positional range with sheet prefix still fails RANGE_INVALID.

    Validates that parse_range correctly rejects ranges where the cell
    reference portion is invalid, even when a sheet prefix is present.
    """
    _disable_all_engines(monkeypatch)

    result = runner.invoke(app, ["screenshot", str(sample_xlsx), "Sheet1!NOTARANGE"])
    assert result.exit_code == 1

    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "RANGE_INVALID"


# ---------------------------------------------------------------------------
# Engine availability errors
# ---------------------------------------------------------------------------


def test_screenshot_no_engine_available(sample_xlsx, monkeypatch):
    """All engines unavailable produces NO_RENDERING_BACKEND error JSON."""
    _disable_all_engines(monkeypatch)

    result = runner.invoke(app, ["screenshot", str(sample_xlsx)])
    assert result.exit_code == 1

    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "NO_RENDERING_BACKEND"
    assert "suggestions" in data


def test_screenshot_engine_excel_unavailable(sample_xlsx, monkeypatch):
    """--engine excel with Excel unavailable produces EXCEL_REQUIRED error JSON."""
    monkeypatch.setattr("agent_xlsx.adapters.xlwings_adapter.is_excel_available", lambda: False)

    result = runner.invoke(app, ["screenshot", str(sample_xlsx), "--engine", "excel"])
    assert result.exit_code == 1

    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "EXCEL_REQUIRED"


def test_screenshot_engine_aspose_unavailable(sample_xlsx, monkeypatch):
    """--engine aspose with Aspose unavailable produces ASPOSE_NOT_INSTALLED error JSON."""
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter._ASPOSE_AVAILABLE", None)
    monkeypatch.setattr("agent_xlsx.adapters.aspose_adapter.is_aspose_available", lambda: False)

    result = runner.invoke(app, ["screenshot", str(sample_xlsx), "--engine", "aspose"])
    assert result.exit_code == 1

    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "ASPOSE_NOT_INSTALLED"


def test_screenshot_engine_libreoffice_unavailable(sample_xlsx, monkeypatch):
    """--engine libreoffice with LO unavailable produces LIBREOFFICE_REQUIRED error JSON."""
    monkeypatch.setattr(
        "agent_xlsx.adapters.libreoffice_adapter.is_libreoffice_available",
        lambda: False,
    )

    result = runner.invoke(app, ["screenshot", str(sample_xlsx), "--engine", "libreoffice"])
    assert result.exit_code == 1

    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "LIBREOFFICE_REQUIRED"


# ---------------------------------------------------------------------------
# Engine alias and env var interaction
# ---------------------------------------------------------------------------


def test_screenshot_engine_lo_alias_unavailable(sample_xlsx, monkeypatch):
    """--engine lo (alias) with LO unavailable also produces LIBREOFFICE_REQUIRED error."""
    monkeypatch.setattr(
        "agent_xlsx.adapters.libreoffice_adapter.is_libreoffice_available",
        lambda: False,
    )

    result = runner.invoke(app, ["screenshot", str(sample_xlsx), "--engine", "lo"])
    assert result.exit_code == 1

    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "LIBREOFFICE_REQUIRED"


def test_screenshot_env_engine_override_no_backend(sample_xlsx, monkeypatch):
    """AGENT_XLSX_ENGINE env var set to unavailable engine produces correct error."""
    _disable_all_engines(monkeypatch)
    monkeypatch.setenv("AGENT_XLSX_ENGINE", "excel")

    result = runner.invoke(app, ["screenshot", str(sample_xlsx)])
    assert result.exit_code == 1

    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "EXCEL_REQUIRED"

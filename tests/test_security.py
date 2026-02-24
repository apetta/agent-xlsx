"""Security regression tests for agent-xlsx.

Groups:
    A — _data_origin tagging: all spreadsheet-content commands must tag JSON output
    B — VBA execution gates: extension check, macro-name validation, MACRO_BLOCKED
    C — ASPOSE_LICENSE_DATA warning: fires on detection (not just success), once per process
"""

import json

from typer.testing import CliRunner

from agent_xlsx.cli import app

runner = CliRunner()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _assert_tagged(output: str) -> dict:
    """Parse JSON stdout and assert _data_origin tag is present."""
    data = json.loads(output)
    assert data.get("_data_origin") == "untrusted_spreadsheet", (
        f"_data_origin missing or wrong in output. Keys present: {list(data.keys())}"
    )
    return data


def _assert_not_tagged(output: str) -> dict:
    """Parse JSON stdout and assert _data_origin tag is absent (structural response)."""
    data = json.loads(output)
    assert "_data_origin" not in data, (
        f"_data_origin unexpectedly present in structural response. Keys: {list(data.keys())}"
    )
    return data


# ---------------------------------------------------------------------------
# Group A — _data_origin tagging
# ---------------------------------------------------------------------------


def test_read_has_data_origin(sample_xlsx):
    result = runner.invoke(app, ["read", str(sample_xlsx)])
    assert result.exit_code == 0, result.stdout
    _assert_tagged(result.stdout)


def test_search_has_data_origin(sample_xlsx):
    result = runner.invoke(app, ["search", str(sample_xlsx), "header"])
    assert result.exit_code == 0, result.stdout
    _assert_tagged(result.stdout)


def test_probe_has_data_origin(sample_xlsx):
    result = runner.invoke(app, ["probe", str(sample_xlsx)])
    assert result.exit_code == 0, result.stdout
    _assert_tagged(result.stdout)


def test_overview_has_data_origin(sample_xlsx):
    result = runner.invoke(app, ["overview", str(sample_xlsx)])
    assert result.exit_code == 0, result.stdout
    _assert_tagged(result.stdout)


def test_export_json_stdout_has_data_origin(sample_xlsx):
    """export --format json without --output must tag the stdout JSON."""
    result = runner.invoke(app, ["export", str(sample_xlsx), "--format", "json"])
    assert result.exit_code == 0, result.stdout
    _assert_tagged(result.stdout)


def test_export_csv_json_envelope_has_data_origin(sample_xlsx):
    """export --format csv --json-envelope must return tagged JSON with data field."""
    result = runner.invoke(app, ["export", str(sample_xlsx), "--format", "csv", "--json-envelope"])
    assert result.exit_code == 0, result.stdout
    data = _assert_tagged(result.stdout)
    assert data["format"] == "csv"
    assert "data" in data


def test_export_markdown_json_envelope_has_data_origin(sample_xlsx):
    """export --format markdown --json-envelope must return tagged JSON with data field."""
    result = runner.invoke(
        app, ["export", str(sample_xlsx), "--format", "markdown", "--json-envelope"]
    )
    assert result.exit_code == 0, result.stdout
    data = _assert_tagged(result.stdout)
    assert data["format"] == "markdown"
    assert "data" in data


def test_inspect_names_has_data_origin(sample_xlsx):
    result = runner.invoke(app, ["inspect", str(sample_xlsx), "--names"])
    assert result.exit_code == 0, result.stdout
    _assert_tagged(result.stdout)


def test_inspect_charts_has_data_origin(sample_xlsx):
    result = runner.invoke(app, ["inspect", str(sample_xlsx), "--charts"])
    assert result.exit_code == 0, result.stdout
    _assert_tagged(result.stdout)


def test_inspect_vba_flag_has_data_origin(sample_xlsx):
    result = runner.invoke(app, ["inspect", str(sample_xlsx), "--vba"])
    assert result.exit_code == 0, result.stdout
    _assert_tagged(result.stdout)


def test_inspect_default_has_data_origin(sample_xlsx):
    """Default inspect (no flags) must tag output."""
    result = runner.invoke(app, ["inspect", str(sample_xlsx)])
    assert result.exit_code == 0, result.stdout
    _assert_tagged(result.stdout)


def test_format_read_has_data_origin(sample_xlsx):
    result = runner.invoke(app, ["format", str(sample_xlsx), "A1", "--read"])
    assert result.exit_code == 0, result.stdout
    _assert_tagged(result.stdout)


def test_inspect_format_cell_has_data_origin(sample_xlsx):
    result = runner.invoke(app, ["inspect", str(sample_xlsx), "--format", "A1"])
    assert result.exit_code == 0, result.stdout
    _assert_tagged(result.stdout)


def test_write_has_no_data_origin(sample_xlsx, tmp_path):
    """write returns a structural operation result — must NOT carry _data_origin."""
    out = tmp_path / "out.xlsx"
    result = runner.invoke(app, ["write", str(sample_xlsx), "Z99", "testval", "-o", str(out)])
    assert result.exit_code == 0, result.stdout
    _assert_not_tagged(result.stdout)


def test_sheet_list_has_no_data_origin(sample_xlsx):
    """sheet --list returns structural metadata — must NOT carry _data_origin."""
    result = runner.invoke(app, ["sheet", str(sample_xlsx), "--list"])
    assert result.exit_code == 0, result.stdout
    _assert_not_tagged(result.stdout)


# ---------------------------------------------------------------------------
# Group B — VBA execution gates
# ---------------------------------------------------------------------------


def test_vba_run_rejects_xlsx_extension(sample_xlsx):
    """Gate 1: .xlsx files must be rejected with INVALID_FORMAT."""
    result = runner.invoke(app, ["vba", str(sample_xlsx), "--run", "Module1.Test"])
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "INVALID_FORMAT"


def test_vba_run_rejects_macro_name_path_traversal(sample_xlsm):
    """Gate 2: path-traversal sequences in macro names must be rejected."""
    result = runner.invoke(app, ["vba", str(sample_xlsm), "--run", "../../evil"])
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "INVALID_MACRO_NAME"


def test_vba_run_rejects_macro_name_slash(sample_xlsm):
    """Gate 2: forward slashes in macro names must be rejected."""
    result = runner.invoke(app, ["vba", str(sample_xlsm), "--run", "a/b"])
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "INVALID_MACRO_NAME"


def test_vba_run_rejects_macro_name_shell_meta(sample_xlsm):
    """Gate 2: shell meta-characters in macro names must be rejected."""
    result = runner.invoke(app, ["vba", str(sample_xlsm), "--run", "Macro;rm -rf /"])
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "INVALID_MACRO_NAME"


def test_vba_run_blocked_on_high_risk(sample_xlsm, monkeypatch):
    """Gate 3: macros flagged risk_level=high must return MACRO_BLOCKED."""
    monkeypatch.setattr(
        "agent_xlsx.commands.vba._security_analysis",
        lambda filepath: {
            "risk_level": "high",
            "auto_execute": [],
            "suspicious": [{"keyword": "Shell"}],
            "iocs": [],
        },
    )
    result = runner.invoke(app, ["vba", str(sample_xlsm), "--run", "Module1.Test"])
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "MACRO_BLOCKED"


def test_vba_run_blocked_error_contains_security_check(sample_xlsm, monkeypatch):
    """MACRO_BLOCKED error body must include the full security_check report."""
    monkeypatch.setattr(
        "agent_xlsx.commands.vba._security_analysis",
        lambda filepath: {
            "risk_level": "high",
            "auto_execute": [],
            "suspicious": [{"keyword": "Shell"}],
            "iocs": [],
        },
    )
    result = runner.invoke(app, ["vba", str(sample_xlsm), "--run", "Module1.Test"])
    data = json.loads(result.stdout)
    assert "security_check" in data
    assert data["security_check"]["risk_level"] == "high"


def test_vba_run_allow_risky_bypasses_macro_blocked(sample_xlsm, monkeypatch):
    """--allow-risky must skip Gate 3 and not return MACRO_BLOCKED."""
    monkeypatch.setattr(
        "agent_xlsx.commands.vba._security_analysis",
        lambda filepath: {
            "risk_level": "high",
            "auto_execute": [],
            "suspicious": [],
            "iocs": [],
        },
    )
    # Also stub run_macro so we don't need a live Excel instance
    monkeypatch.setattr(
        "agent_xlsx.adapters.xlwings_adapter.run_macro",
        lambda filepath, macro_name, args, save: {"status": "success", "result": None},
    )
    result = runner.invoke(app, ["vba", str(sample_xlsm), "--run", "Module1.Test", "--allow-risky"])
    data = json.loads(result.stdout)
    assert data.get("code") != "MACRO_BLOCKED"
    assert data.get("error") is not True


# ---------------------------------------------------------------------------
# Group C — ASPOSE_LICENSE_DATA warning
# ---------------------------------------------------------------------------


def test_aspose_warning_fires_on_detection_with_invalid_data(monkeypatch, capsys):
    """Warning must fire even when base64 data is invalid (apply fails)."""
    import agent_xlsx.adapters.aspose_adapter as adapter

    # Reset module-level state so this test is isolated
    adapter._LICENSE_APPLIED = False
    adapter._LICENSE_DATA_WARNED = False

    monkeypatch.setenv("ASPOSE_LICENSE_DATA", "!!!notvalidbase64!!!")

    adapter._apply_license()

    captured = capsys.readouterr()
    assert "ASPOSE_LICENSE_DATA" in captured.err
    assert "Warning" in captured.err


def test_aspose_warning_fires_only_once_per_process(monkeypatch, capsys):
    """Warning must fire at most once per process even on repeated calls."""
    import agent_xlsx.adapters.aspose_adapter as adapter

    adapter._LICENSE_APPLIED = False
    adapter._LICENSE_DATA_WARNED = False

    monkeypatch.setenv("ASPOSE_LICENSE_DATA", "!!!notvalidbase64!!!")

    adapter._apply_license()
    adapter._apply_license()
    adapter._apply_license()

    captured = capsys.readouterr()
    assert captured.err.count("ASPOSE_LICENSE_DATA") == 1


def test_aspose_warning_fires_when_path_overrides_data(monkeypatch, capsys):
    """Warning must fire even when ASPOSE_LICENSE_PATH takes priority over ASPOSE_LICENSE_DATA."""
    import agent_xlsx.adapters.aspose_adapter as adapter

    adapter._LICENSE_APPLIED = False
    adapter._LICENSE_DATA_WARNED = False

    monkeypatch.setenv("ASPOSE_LICENSE_DATA", "sensitivedata")
    monkeypatch.setenv("ASPOSE_LICENSE_PATH", "/nonexistent/path.lic")

    adapter._apply_license()

    captured = capsys.readouterr()
    assert "ASPOSE_LICENSE_DATA" in captured.err
    assert "Warning" in captured.err


def test_aspose_no_warning_without_env_var(monkeypatch, capsys):
    """No warning must appear when ASPOSE_LICENSE_DATA is not set."""
    import agent_xlsx.adapters.aspose_adapter as adapter

    adapter._LICENSE_APPLIED = False
    adapter._LICENSE_DATA_WARNED = False

    monkeypatch.delenv("ASPOSE_LICENSE_DATA", raising=False)
    monkeypatch.delenv("ASPOSE_LICENSE_PATH", raising=False)

    adapter._apply_license()

    captured = capsys.readouterr()
    assert "ASPOSE_LICENSE_DATA" not in captured.err

"""VBA command functional tests — list, read, read-all, and security operations.

Uses the committed fixture at tests/fixtures/vba_sample.xlsm (via vba_xlsm conftest
fixture) which contains:
    - Module1.bas: Sub TestMacro(), Function AddNumbers(), Sub Auto_Open()
    - ThisWorkbook.cls, Sheet1.cls: auto-generated document modules (Aspose)

oletools flags Aspose-generated hex attribute strings as suspicious → risk_level=high.
"""

import json

from typer.testing import CliRunner

from agent_xlsx.cli import app

runner = CliRunner()


# ---------------------------------------------------------------------------
# List (default behaviour and explicit --list)
# ---------------------------------------------------------------------------


def test_vba_list_default(vba_xlsm):
    """Invoking with no flags defaults to --list; output includes has_vba and modules."""
    result = runner.invoke(app, ["vba", str(vba_xlsm)])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["has_vba"] is True
    assert "modules" in data
    assert isinstance(data["modules"], list)


def test_vba_list_has_vba_true(vba_xlsm):
    """has_vba field must be true for a macro-enabled workbook."""
    result = runner.invoke(app, ["vba", str(vba_xlsm), "--list"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["has_vba"] is True


def test_vba_list_module_names(vba_xlsm):
    """Module1.bas must appear in the modules list with type 'standard'."""
    result = runner.invoke(app, ["vba", str(vba_xlsm), "--list"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    module_names = [m["name"] for m in data["modules"]]
    assert "Module1.bas" in module_names

    module1 = next(m for m in data["modules"] if m["name"] == "Module1.bas")
    assert module1["type"] == "standard"


def test_vba_list_auto_execute(vba_xlsm):
    """auto_execute array must contain 'Auto_Open' trigger."""
    result = runner.invoke(app, ["vba", str(vba_xlsm), "--list"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert "Auto_Open" in data["auto_execute"]


def test_vba_list_risk_level(vba_xlsm):
    """risk_level must be 'high' — Aspose-generated hex attributes trigger oletools flags."""
    result = runner.invoke(app, ["vba", str(vba_xlsm), "--list"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["risk_level"] == "high"


# ---------------------------------------------------------------------------
# Read single module
# ---------------------------------------------------------------------------


def test_vba_read_module(vba_xlsm):
    """--read Module1.bas returns code containing TestMacro, AddNumbers, Auto_Open."""
    result = runner.invoke(app, ["vba", str(vba_xlsm), "--read", "Module1.bas"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    assert data["module"] == "Module1.bas"
    assert "code" in data
    # Verify all three procedures are present in the code text
    assert "TestMacro" in data["code"]
    assert "AddNumbers" in data["code"]
    assert "Auto_Open" in data["code"]


def test_vba_read_procedures(vba_xlsm):
    """procedures list from --read Module1.bas includes the three known procedures."""
    result = runner.invoke(app, ["vba", str(vba_xlsm), "--read", "Module1.bas"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    assert "procedures" in data
    assert "TestMacro" in data["procedures"]
    assert "AddNumbers" in data["procedures"]
    assert "Auto_Open" in data["procedures"]


def test_vba_read_not_found(vba_xlsm):
    """--read with a nonexistent module name must fail with VBA_NOT_FOUND."""
    result = runner.invoke(app, ["vba", str(vba_xlsm), "--read", "NoModule"])
    assert result.exit_code == 1
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "VBA_NOT_FOUND"


# ---------------------------------------------------------------------------
# Read all modules
# ---------------------------------------------------------------------------


def test_vba_read_all(vba_xlsm):
    """--read-all returns module_count >= 3 and a modules array with code for each."""
    result = runner.invoke(app, ["vba", str(vba_xlsm), "--read-all"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    assert data["module_count"] >= 3
    assert isinstance(data["modules"], list)
    assert len(data["modules"]) >= 3

    # Every module entry must have code and module name
    for mod in data["modules"]:
        assert "module" in mod
        assert "code" in mod


# ---------------------------------------------------------------------------
# Security analysis
# ---------------------------------------------------------------------------


def test_vba_security(vba_xlsm):
    """--security returns has_vba, auto_execute with Auto_Open, and a risk_level."""
    result = runner.invoke(app, ["vba", str(vba_xlsm), "--security"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)

    assert data["has_vba"] is True
    assert "Auto_Open" in data["auto_execute"]
    assert "risk_level" in data


def test_vba_security_auto_open_detected(vba_xlsm):
    """--security must detect Auto_Open in the auto_execute list."""
    result = runner.invoke(app, ["vba", str(vba_xlsm), "--security"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert "Auto_Open" in data["auto_execute"]


# ---------------------------------------------------------------------------
# No-VBA fallback (plain .xlsx)
# ---------------------------------------------------------------------------


def test_vba_no_vba_xlsx(sample_xlsx):
    """Invoking on a plain .xlsx must report has_vba=false and empty modules."""
    result = runner.invoke(app, ["vba", str(sample_xlsx)])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["has_vba"] is False
    assert data["modules"] == []

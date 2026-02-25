"""Tests for license command: --status, --set, --clear workflows."""

import json
from unittest.mock import patch

from typer.testing import CliRunner

from agent_xlsx.cli import app

runner = CliRunner()


# ---------------------------------------------------------------------------
# Helpers — redirect config storage to tmp_path
# ---------------------------------------------------------------------------


def _config_patches(tmp_path):
    """Return a tuple of mock patches that redirect config to tmp_path."""
    config_dir = tmp_path / ".agent-xlsx"
    config_file = config_dir / "config.json"
    return (
        patch("agent_xlsx.utils.config.CONFIG_DIR", config_dir),
        patch("agent_xlsx.utils.config.CONFIG_FILE", config_file),
    )


# ---------------------------------------------------------------------------
# --status (default behaviour)
# ---------------------------------------------------------------------------


def test_license_status_default_installed(tmp_path):
    """Default invocation (no flags) shows status when Aspose is installed and licensed."""
    p_dir, p_file = _config_patches(tmp_path)
    with (
        p_dir,
        p_file,
        patch(
            "agent_xlsx.adapters.aspose_adapter.get_license_status",
            return_value={"installed": True, "licensed": True, "evaluation_mode": False},
        ),
    ):
        result = runner.invoke(app, ["license"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["installed"] is True
    assert data["licensed"] is True


def test_license_status_not_installed(tmp_path):
    """--status when Aspose is not installed includes install suggestions."""
    p_dir, p_file = _config_patches(tmp_path)
    with (
        p_dir,
        p_file,
        patch(
            "agent_xlsx.adapters.aspose_adapter.get_license_status",
            return_value={"installed": False, "licensed": False, "evaluation_mode": False},
        ),
    ):
        result = runner.invoke(app, ["license", "--status"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["installed"] is False
    assert "suggestions" in data
    # Suggestions should mention how to install
    assert any("install" in s.lower() for s in data["suggestions"])


def test_license_status_evaluation_mode(tmp_path):
    """--status in evaluation mode includes licence suggestions."""
    p_dir, p_file = _config_patches(tmp_path)
    with (
        p_dir,
        p_file,
        patch(
            "agent_xlsx.adapters.aspose_adapter.get_license_status",
            return_value={"installed": True, "licensed": False, "evaluation_mode": True},
        ),
    ):
        result = runner.invoke(app, ["license", "--status"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["evaluation_mode"] is True
    assert "suggestions" in data
    assert any("licence" in s.lower() or "license" in s.lower() for s in data["suggestions"])


# ---------------------------------------------------------------------------
# --set
# ---------------------------------------------------------------------------


def test_license_set_file_not_found(tmp_path):
    """--set with a non-existent path emits FILE_NOT_FOUND error."""
    p_dir, p_file = _config_patches(tmp_path)
    with p_dir, p_file:
        result = runner.invoke(app, ["license", "--set", "/nonexistent/fake.lic"])
    assert result.exit_code == 1
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "FILE_NOT_FOUND"


def test_license_set_saves_config(tmp_path):
    """--set with a valid file saves the licence path to config and shows status."""
    # Create a fake .lic file
    lic_file = tmp_path / "Aspose.Cells.lic"
    lic_file.write_text("fake-licence-content")

    p_dir, p_file = _config_patches(tmp_path)
    with (
        p_dir,
        p_file,
        patch(
            "agent_xlsx.adapters.aspose_adapter.get_license_status",
            return_value={"installed": True, "licensed": True, "evaluation_mode": False},
        ),
    ):
        result = runner.invoke(app, ["license", "--set", str(lic_file)])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    # Should contain the saved config path
    assert "config_path" in data
    assert str(lic_file.resolve()) in data["config_path"]
    assert "message" in data

    # Verify config file was actually written
    config_file = tmp_path / ".agent-xlsx" / "config.json"
    assert config_file.exists()
    saved = json.loads(config_file.read_text())
    assert saved["aspose_license_path"] == str(lic_file.resolve())


# ---------------------------------------------------------------------------
# --clear
# ---------------------------------------------------------------------------


def test_license_clear(tmp_path):
    """--clear removes aspose_license_path from config."""
    # Pre-populate config with a licence path
    config_dir = tmp_path / ".agent-xlsx"
    config_dir.mkdir()
    config_file = config_dir / "config.json"
    config_file.write_text(json.dumps({"aspose_license_path": "/old/path.lic"}))

    p_dir, p_file = _config_patches(tmp_path)
    with p_dir, p_file:
        result = runner.invoke(app, ["license", "--clear"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["status"] == "cleared"

    # Verify the licence path was removed from config
    saved = json.loads(config_file.read_text())
    assert "aspose_license_path" not in saved


def test_license_clear_idempotent(tmp_path):
    """--clear on an already-empty config succeeds without error."""
    p_dir, p_file = _config_patches(tmp_path)
    with p_dir, p_file:
        result = runner.invoke(app, ["license", "--clear"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["status"] == "cleared"


# ---------------------------------------------------------------------------
# Config reflects saved path
# ---------------------------------------------------------------------------


def test_config_reflects_saved_path(tmp_path):
    """After --set, load_config returns the saved licence path."""
    lic_file = tmp_path / "my.lic"
    lic_file.write_text("licence-data")

    p_dir, p_file = _config_patches(tmp_path)
    with (
        p_dir,
        p_file,
        patch(
            "agent_xlsx.adapters.aspose_adapter.get_license_status",
            return_value={"installed": True, "licensed": True, "evaluation_mode": False},
        ),
    ):
        runner.invoke(app, ["license", "--set", str(lic_file)])

    # Now verify load_config picks up the saved path
    from agent_xlsx.utils.config import load_config

    with p_dir, p_file:
        config = load_config()
    assert config["aspose_license_path"] == str(lic_file.resolve())

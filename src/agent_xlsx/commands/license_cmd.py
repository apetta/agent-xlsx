"""Licence management for Aspose.Cells engine."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import typer

from agent_xlsx.cli import app
from agent_xlsx.formatters.json_formatter import output
from agent_xlsx.utils.errors import handle_error


@app.command(name="license")
@handle_error
def license_cmd(
    set_path: Optional[str] = typer.Option(None, "--set", help="Path to Aspose.Cells .lic file"),
    status: bool = typer.Option(False, "--status", help="Show current licence status"),
    clear: bool = typer.Option(False, "--clear", help="Remove saved licence path"),
) -> None:
    """Manage Aspose.Cells licence for rendering engine.

    Use --status to check if Aspose is installed and licensed.
    Use --set to save a licence file path for automatic loading.
    Use --clear to remove the saved licence path.
    """
    if set_path:
        # Validate file exists
        lic_path = Path(set_path).resolve()
        if not lic_path.exists():
            output(
                {
                    "error": True,
                    "code": "FILE_NOT_FOUND",
                    "message": f"Licence file not found: {set_path}",
                }
            )
            raise SystemExit(1)

        # Save to config
        from agent_xlsx.utils.config import load_config, save_config

        config = load_config()
        config["aspose_license_path"] = str(lic_path)
        save_config(config)

        # Verify it works
        from agent_xlsx.adapters.aspose_adapter import get_license_status

        result = get_license_status()
        result["config_path"] = str(lic_path)
        result["message"] = "Licence path saved to ~/.agent-xlsx/config.json"
        output(result)
        return

    if clear:
        from agent_xlsx.utils.config import load_config, save_config

        config = load_config()
        config.pop("aspose_license_path", None)
        save_config(config)
        output({"status": "cleared", "message": "Licence path removed from config"})
        return

    # Default: show status (also when --status is explicitly passed)
    from agent_xlsx.adapters.aspose_adapter import get_license_status

    result = get_license_status()
    if not result.get("installed"):
        result["suggestions"] = [
            "Install with: uv add --optional aspose aspose-cells-python",
            "Or: uv pip install aspose-cells-python",
        ]
    elif result.get("evaluation_mode"):
        result["suggestions"] = [
            "Set licence: agent-xlsx license --set /path/to/Aspose.Cells.lic",
            "Or set ASPOSE_LICENSE_PATH environment variable",
            "Get 30-day trial: https://purchase.aspose.com/temporary-license",
        ]
    output(result)

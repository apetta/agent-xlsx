"""VBA operations command — list, read, read-all, macro execution, and security analysis.

Uses oletools for VBA extraction and analysis. Works headless on all platforms
without requiring Microsoft Excel. Macro execution uses xlwings and requires
a local Excel installation.
"""

import re as _re
from typing import Any, Optional

import typer

from agent_xlsx.cli import app
from agent_xlsx.formatters.json_formatter import output, output_spreadsheet_data
from agent_xlsx.utils.constants import VBA_EXTENSIONS
from agent_xlsx.utils.errors import AgentExcelError, handle_error
from agent_xlsx.utils.validation import validate_file


class VbaModuleNotFoundError(AgentExcelError):
    """Raised when a requested VBA module does not exist."""

    def __init__(self, name: str, available: list[str]):
        super().__init__(
            "VBA_NOT_FOUND",
            f"VBA module '{name}' not found",
            [f"Available modules: {', '.join(available)}"]
            if available
            else ["No VBA modules found in this file"],
        )


class MacroBlockedError(AgentExcelError):
    """Raised when a VBA security pre-check blocks macro execution (risk_level=high)."""

    def __init__(self, security_report: dict[str, Any]):
        super().__init__(
            "MACRO_BLOCKED",
            f"Macro execution blocked: security analysis returned "
            f"risk_level '{security_report.get('risk_level', 'unknown')}'.",
            [
                "Run 'agent-xlsx vba <file> --security' to review the full report.",
                "Pass '--allow-risky' to override if you trust this file.",
            ],
        )
        self._security_report = security_report

    def to_dict(self) -> dict[str, Any]:
        result = super().to_dict()
        result["security_check"] = {
            "risk_level": self._security_report.get("risk_level"),
            "auto_execute_triggers": self._security_report.get("auto_execute", []),
            "suspicious": self._security_report.get("suspicious", []),
            "iocs": self._security_report.get("iocs", []),
        }
        return result


def _list_modules(filepath: str) -> dict[str, Any]:
    """List all VBA modules with metadata and basic security summary."""
    from agent_xlsx.adapters.oletools_adapter import analyse_vba_security, extract_vba_modules

    modules = extract_vba_modules(filepath)
    security = analyse_vba_security(filepath)

    return {
        "file": filepath,
        "has_vba": len(modules) > 0,
        "modules": modules,
        "auto_execute": security.get("auto_execute", []),
        "risk_level": security.get("risk_level", "low"),
    }


def _read_module(filepath: str, module_name: str) -> dict[str, Any]:
    """Read source code for a single VBA module."""
    from agent_xlsx.adapters.oletools_adapter import extract_vba_modules, read_vba_code

    results = read_vba_code(filepath, module_name=module_name)

    if not results:
        available = [m["name"] for m in extract_vba_modules(filepath)]
        raise VbaModuleNotFoundError(module_name, available)

    return results[0]


def _read_all(filepath: str) -> dict[str, Any]:
    """Read source code for all VBA modules."""
    from agent_xlsx.adapters.oletools_adapter import read_vba_code

    results = read_vba_code(filepath)

    return {
        "file": filepath,
        "module_count": len(results),
        "modules": results,
    }


def _security_analysis(filepath: str) -> dict[str, Any]:
    """Perform VBA security analysis."""
    from agent_xlsx.adapters.oletools_adapter import analyse_vba_security

    return analyse_vba_security(filepath)


@app.command()
@handle_error
def vba(
    file: str = typer.Argument(..., help="Path to the Excel file (.xlsm/.xlsb)"),
    list_modules: bool = typer.Option(False, "--list", "-l", help="List VBA modules"),
    read: Optional[str] = typer.Option(
        None, "--read", "-r", help="Read a specific VBA module's code"
    ),
    read_all: bool = typer.Option(False, "--read-all", help="Read all VBA module code"),
    security: bool = typer.Option(False, "--security", help="Run VBA security analysis"),
    run: Optional[str] = typer.Option(
        None, "--run", help="Execute a VBA macro (e.g. 'Module1.MyMacro')"
    ),
    args: Optional[str] = typer.Option(
        None, "--args", help="JSON-encoded arguments for the macro (e.g. '[\"arg1\", 42]')"
    ),
    save: bool = typer.Option(False, "--save", help="Save workbook after macro execution"),
    allow_risky: bool = typer.Option(
        False,
        "--allow-risky",
        help="Override automatic security block for high-risk macros (use only when the file source is explicitly trusted)",  # noqa: E501
    ),
) -> None:
    """VBA operations: list, read, run, and analyse macros.

    Uses oletools for extraction and security analysis. Macro execution
    uses xlwings and requires a local Excel installation.
    """
    path = validate_file(file)
    filepath = str(path)

    if run:
        # Gate 1: only VBA-capable formats support macro execution
        if path.suffix.lower() not in VBA_EXTENSIONS:
            raise AgentExcelError(
                "INVALID_FORMAT",
                f"Macro execution requires a VBA-capable format (.xlsm or .xlsb); "
                f"'{path.suffix}' does not support macros.",
                ["Use a .xlsm or .xlsb file to run macros"],
            )

        # Gate 2: validate macro name — reject path separators and shell meta-characters
        if not _re.match(r"^[A-Za-z_][A-Za-z0-9_. ]*$", run):
            raise AgentExcelError(
                "INVALID_MACRO_NAME",
                f"Macro name '{run}' contains invalid characters.",
                ["Use format 'Module1.MacroName' or just 'MacroName'"],
            )

        # Gate 3: automatic security pre-check — runs silently before every execution
        security_report = _security_analysis(filepath)
        if security_report.get("risk_level") == "high" and not allow_risky:
            raise MacroBlockedError(security_report)

        import json as json_mod

        from agent_xlsx.adapters.xlwings_adapter import run_macro

        parsed_args = None
        if args:
            parsed_args = json_mod.loads(args)

        result = run_macro(
            filepath=filepath,
            macro_name=run,
            args=parsed_args,
            save=save,
        )
        # Always attach security telemetry — visible to the calling agent
        result["security_check"] = {
            "risk_level": security_report.get("risk_level", "unknown"),
            "auto_execute_triggers": security_report.get("auto_execute", []),
            "suspicious_count": len(security_report.get("suspicious", [])),
        }
        output(result)
        return

    if read:
        result = _read_module(filepath, read)
    elif read_all:
        result = _read_all(filepath)
    elif security:
        result = _security_analysis(filepath)
    else:
        # Default to --list behaviour
        result = _list_modules(filepath)

    output_spreadsheet_data(result)

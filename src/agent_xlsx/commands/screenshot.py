"""Visual sheet capture command with Excel/LibreOffice backend auto-detection.

Exports workbook sheets to HD PNG with full visual fidelity — charts, shapes,
arrows, conditional formatting, and all drawing objects are preserved.

Uses Aspose when available, then xlwings (Excel), then LibreOffice headless.
Supports per-sheet and range-level capture.
"""

from typing import Optional

import typer

from agent_xlsx.cli import app
from agent_xlsx.formatters.json_formatter import output
from agent_xlsx.utils.errors import ExcelRequiredError, NoRenderingBackendError, handle_error
from agent_xlsx.utils.validation import parse_range, validate_file


@app.command()
@handle_error
def screenshot(
    file: str = typer.Argument(..., help="Path to the Excel file"),
    range_: Optional[str] = typer.Argument(
        None,
        metavar="RANGE",
        help="Range e.g. 'Sheet1!A1:C10'",
    ),
    sheet: Optional[str] = typer.Option(
        None, "--sheet", "-s", help="Sheet name(s), comma-separated"
    ),
    range_opt: Optional[str] = typer.Option(
        None, "--range", "-r", help="Cell range (e.g. 'A1:F20')"
    ),
    output_path: Optional[str] = typer.Option(
        None, "--output", "-o", help="Output file/directory path"
    ),
    dpi: int = typer.Option(200, "--dpi", help="DPI for PNG rendering (Aspose/LibreOffice only)"),
    timeout: int = typer.Option(
        30, "--timeout", help="Timeout in seconds (LibreOffice backend only)"
    ),
    engine: str = typer.Option(
        "auto",
        "--engine",
        "-e",
        help="Rendering backend: auto (default), excel, aspose, or libreoffice",
    ),
    base64_output: bool = typer.Option(
        False,
        "--base64",
        help="Return image data as base64 in JSON instead of saving files",
    ),
) -> None:
    """Capture sheet(s) or a range as HD PNG for visual understanding.

    Uses Aspose when available, then Excel (via xlwings), then LibreOffice.
    Use --engine to force a specific backend.

    Produces HD per-sheet PNG images that Claude Code can view natively.
    Use --base64 to return image data inline in the JSON response without
    persisting files.
    """
    import tempfile

    path = validate_file(file)

    # When --base64 is used without --output, render to a temp dir
    temp_dir = None
    if base64_output and output_path is None:
        temp_dir = tempfile.mkdtemp(prefix="agent_xlsx_b64_")
        output_path = temp_dir

    # Resolve range — positional arg takes precedence over --range flag
    effective_range: Optional[str] = None
    if range_:
        parsed = parse_range(range_)
        if parsed["sheet"]:
            sheet = sheet or parsed["sheet"]
        effective_range = f"{parsed['start']}:{parsed['end']}" if parsed["end"] else parsed["start"]
    elif range_opt:
        effective_range = range_opt

    # Backend selection: explicit --engine or auto-detect
    from agent_xlsx.adapters.libreoffice_adapter import is_libreoffice_available
    from agent_xlsx.adapters.xlwings_adapter import is_excel_available

    use_engine = None  # "excel", "aspose", or "libreoffice"
    engine_lower = engine.lower()

    if engine_lower == "excel":
        if not is_excel_available():
            raise ExcelRequiredError("screenshot")
        use_engine = "excel"
    elif engine_lower == "aspose":
        from agent_xlsx.adapters.aspose_adapter import is_aspose_available
        from agent_xlsx.utils.errors import AsposeNotInstalledError

        if not is_aspose_available():
            raise AsposeNotInstalledError()
        use_engine = "aspose"
    elif engine_lower in ("libreoffice", "lo"):
        if not is_libreoffice_available():
            from agent_xlsx.utils.errors import LibreOfficeNotFoundError

            raise LibreOfficeNotFoundError()
        use_engine = "libreoffice"
    elif engine_lower == "auto":
        from agent_xlsx.adapters.aspose_adapter import is_aspose_available

        if is_aspose_available():
            use_engine = "aspose"
        elif is_excel_available():
            use_engine = "excel"
        elif is_libreoffice_available():
            use_engine = "libreoffice"
        else:
            raise NoRenderingBackendError("screenshot")
    else:
        raise NoRenderingBackendError("screenshot")

    if use_engine == "excel":
        from agent_xlsx.adapters.xlwings_adapter import screenshot as _screenshot

        result = _screenshot(
            filepath=str(path),
            sheet_name=sheet,
            range_str=effective_range,
            output_path=output_path,
        )
    elif use_engine == "aspose":
        from agent_xlsx.adapters.aspose_adapter import screenshot as _screenshot

        result = _screenshot(
            filepath=str(path),
            sheet_name=sheet,
            range_str=effective_range,
            output_path=output_path,
            dpi=dpi,
        )
    else:
        from agent_xlsx.adapters.libreoffice_adapter import screenshot as _screenshot

        result = _screenshot(
            filepath=str(path),
            sheet_name=sheet,
            range_str=effective_range,
            output_path=output_path,
            dpi=dpi,
            timeout=timeout,
        )

    # Quality gate — detect degenerate captures (tiny images)
    _validate_capture_quality(result)

    # Encode file(s) as base64 and embed in result
    if base64_output:
        import base64
        import shutil
        from pathlib import Path

        def _encode_file(file_path: str) -> str:
            return base64.b64encode(Path(file_path).read_bytes()).decode("ascii")

        if "sheets" in result:
            for sheet_entry in result["sheets"]:
                sheet_entry["base64"] = _encode_file(sheet_entry["path"])
        elif "path" in result:
            result["base64"] = _encode_file(result["path"])

        # Clean up temp dir if we created one
        if temp_dir is not None:
            shutil.rmtree(temp_dir, ignore_errors=True)
            # Remove file paths that no longer exist
            if "sheets" in result:
                for sheet_entry in result["sheets"]:
                    del sheet_entry["path"]
            elif "path" in result:
                del result["path"]

    output(result)


def _validate_capture_quality(result: dict) -> None:
    """Flag degenerate captures (tiny images) so agents don't waste time on garbage."""
    from agent_xlsx.utils.constants import MIN_CAPTURE_HEIGHT, MIN_CAPTURE_WIDTH

    def _check(entry: dict) -> None:
        w = entry.get("width", 0)
        h = entry.get("height", 0)
        if w < MIN_CAPTURE_WIDTH or h < MIN_CAPTURE_HEIGHT:
            entry["warning"] = (
                f"Degenerate capture ({w}\u00d7{h}px). "
                "The captured range is likely too small or empty. "
                "Try without a range argument to capture the full sheet."
            )
            result["status"] = "partial"

    if "sheets" in result:
        for sheet_entry in result["sheets"]:
            _check(sheet_entry)
    else:
        _check(result)

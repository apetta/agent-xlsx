"""Read and apply cell formatting."""

import json as json_mod
from typing import Optional

import typer

from agent_xlsx.cli import app
from agent_xlsx.formatters import json_formatter
from agent_xlsx.formatters.json_formatter import output_spreadsheet_data
from agent_xlsx.utils.errors import AgentExcelError, handle_error
from agent_xlsx.utils.validation import _normalise_shell_ref, validate_file


@app.command("format")
@handle_error
def format_cmd(
    file: str = typer.Argument(..., help="Path to the Excel file"),
    cell: str = typer.Argument(..., help="Cell or range reference (e.g. 'A1', '2022!A1', or 'Sheet1!A1:D10')"),
    read: bool = typer.Option(False, "--read", help="Read formatting at the cell"),
    font: Optional[str] = typer.Option(
        None,
        "--font",
        help='Font options as JSON (e.g. \'{"bold": true, "size": 14}\')',
    ),
    fill: Optional[str] = typer.Option(
        None,
        "--fill",
        help='Fill options as JSON (e.g. \'{"color": "FFFF00"}\')',
    ),
    border: Optional[str] = typer.Option(
        None,
        "--border",
        help='Border options as JSON (e.g. \'{"style": "thin"}\')',
    ),
    number_format: Optional[str] = typer.Option(
        None,
        "--number-format",
        help='Number format string (e.g. "#,##0.00")',
    ),
    copy_from: Optional[str] = typer.Option(
        None,
        "--copy-from",
        help="Copy formatting from another cell",
    ),
    output: Optional[str] = typer.Option(
        None,
        "--output",
        "-o",
        help="Save to a new file instead of overwriting",
    ),
    sheet: Optional[str] = typer.Option(None, "--sheet", "-s", help="Target sheet name"),
) -> None:
    """Read or apply cell formatting."""
    path = validate_file(file)

    cell = _normalise_shell_ref(cell)

    from agent_xlsx.adapters import openpyxl_adapter as oxl

    # --- Read mode ---
    if read:
        # Determine sheet for reading
        read_sheet = sheet
        read_cell = cell
        if "!" in cell:
            parts = cell.split("!", 1)
            read_sheet = parts[0]
            read_cell = parts[1]
        if not read_sheet:
            read_sheet = _default_sheet(str(path))
        result = oxl.get_cell_formatting(str(path), read_sheet, read_cell)
        output_spreadsheet_data(result)
        return

    # --- Copy mode ---
    if copy_from:
        target_sheet = sheet
        target_cell = cell
        if "!" in cell:
            parts = cell.split("!", 1)
            target_sheet = parts[0]
            target_cell = parts[1]
        if not target_sheet:
            target_sheet = _default_sheet(str(path))
        result = oxl.copy_formatting(
            str(path),
            sheet_name=target_sheet,
            source_ref=copy_from,
            target_ref=target_cell,
            output_path=output,
        )
        json_formatter.output(result)
        return

    # --- Apply mode ---
    has_formatting = any([font, fill, border, number_format])
    if not has_formatting:
        raise AgentExcelError(
            "MISSING_FORMAT",
            "No formatting options provided",
            [
                "Use --read to read formatting",
                "Use --font, --fill, --border, or --number-format to apply formatting",
                "Use --copy-from to copy formatting from another cell",
            ],
        )

    font_opts = _parse_json_opt(font, "font") if font else None
    fill_opts = _parse_json_opt(fill, "fill") if fill else None
    border_opts = _parse_json_opt(border, "border") if border else None

    target_sheet = sheet
    target_cell = cell
    if "!" in cell:
        parts = cell.split("!", 1)
        target_sheet = parts[0]
        target_cell = parts[1]
    if not target_sheet:
        target_sheet = _default_sheet(str(path))

    result = oxl.apply_formatting(
        str(path),
        sheet_name=target_sheet,
        cell_ref=target_cell,
        font_opts=font_opts,
        fill_opts=fill_opts,
        border_opts=border_opts,
        number_format=number_format,
        output_path=output,
    )
    json_formatter.output(result)


def _parse_json_opt(json_str: str, label: str) -> dict:
    """Parse a JSON string option, raising a clean error on failure."""
    try:
        data = json_mod.loads(json_str)
    except json_mod.JSONDecodeError as exc:
        raise AgentExcelError(
            "INVALID_JSON",
            f"Failed to parse --{label} JSON: {exc}",
            [f"Provide valid JSON for --{label}"],
        )
    if not isinstance(data, dict):
        raise AgentExcelError(
            "INVALID_JSON",
            f"--{label} must be a JSON object, not {type(data).__name__}",
            [f"e.g. --{label} '{{\"bold\": true}}'"],
        )
    return data


def _default_sheet(filepath: str) -> str:
    """Return the name of the first sheet in the workbook."""
    from openpyxl import load_workbook

    wb = load_workbook(filepath, read_only=True)
    try:
        return wb.sheetnames[0]
    finally:
        wb.close()

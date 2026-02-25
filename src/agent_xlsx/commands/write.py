"""Write values or formulas to specific cells/ranges."""

import csv
import json as json_mod
from pathlib import Path
from typing import Optional

import typer

from agent_xlsx.cli import app
from agent_xlsx.formatters import json_formatter
from agent_xlsx.utils.errors import AgentExcelError, handle_error
from agent_xlsx.utils.validation import _normalise_shell_ref, validate_file_for_write


@app.command("write")
@handle_error
def write_cmd(
    file: str = typer.Argument(..., help="Path to the Excel file"),
    cell: str = typer.Argument(
        ..., help="Cell reference (e.g. 'A1', '2022!A1', or 'Sheet1!A1:C3')"
    ),  # noqa: E501
    value: Optional[str] = typer.Argument(None, help="Value to write (for single cell)"),
    value_opt: Optional[str] = typer.Option(
        None,
        "--value",
        "-v",
        help="Value to write (alternative to positional arg — "
        "handles negative numbers like --value '-4.095')",
    ),
    formula: bool = typer.Option(
        False,
        "--formula",
        help="Single cell: adds '=' prefix if missing. "
        "Batch (--json/--from-csv): strings starting with '=' are "
        "written as formulas, all other values written as-is.",
    ),
    json: Optional[str] = typer.Option(None, "--json", help="JSON array data for range write"),
    from_json: Optional[str] = typer.Option(
        None,
        "--from-json",
        help="Path to JSON file containing 2D array data for range write",
    ),
    from_csv: Optional[str] = typer.Option(None, "--from-csv", help="CSV file to read data from"),
    number_format: Optional[str] = typer.Option(
        None,
        "--number-format",
        help="Number format (e.g. '0.00%%')",
    ),
    output: Optional[str] = typer.Option(
        None,
        "--output",
        "-o",
        help="Save to a new file instead of overwriting",
    ),
    sheet: Optional[str] = typer.Option(None, "--sheet", "-s", help="Target sheet name"),
) -> None:
    """Write values or formulas to specific cells or ranges."""
    path, _is_new = validate_file_for_write(file)

    # --value option takes precedence over positional value arg
    # (named options handle negative numbers that Click misparses as flags)
    effective_value = value_opt if value_opt is not None else value

    # Parse Sheet!Cell syntax (e.g. "2022!B1" → sheet=2022, cell=B1)
    cell = _normalise_shell_ref(cell)
    if "!" in cell:
        parts = cell.split("!", 1)
        if not sheet:  # --sheet flag takes precedence
            sheet = parts[0]
        cell = parts[1]

    from agent_xlsx.adapters import openpyxl_adapter as oxl

    # Build the cell data list for the adapter
    write_data: list[dict] = []

    if from_json:
        # Read JSON data from a file (avoids shell escaping issues for large payloads)
        json_path = Path(from_json)
        if not json_path.exists():
            raise AgentExcelError(
                "FILE_NOT_FOUND",
                f"JSON file '{from_json}' not found",
                ["Check the JSON file path"],
            )
        json_str = json_path.read_text(encoding="utf-8")
        write_data = _json_to_cells(cell, json_str, formula=formula)
    elif json:
        # Parse JSON array and map to cells starting at the given cell ref
        write_data = _json_to_cells(cell, json, formula=formula)
    elif from_csv:
        # Read CSV file and map rows to cells
        csv_path = Path(from_csv)
        if not csv_path.exists():
            raise AgentExcelError(
                "FILE_NOT_FOUND",
                f"CSV file '{from_csv}' does not exist",
                ["Check the CSV file path"],
            )
        write_data = _csv_to_cells(cell, csv_path, formula=formula)
    elif effective_value is not None:
        # Single cell write
        cell_value = effective_value
        if formula and not cell_value.startswith("="):
            cell_value = f"={cell_value}"
        elif not formula:
            # Try to coerce to number
            cell_value = _coerce_value(cell_value)

        entry: dict = {"cell": cell.upper(), "value": cell_value}
        if number_format:
            entry["number_format"] = number_format
        write_data = [entry]
    else:
        raise AgentExcelError(
            "MISSING_VALUE",
            "No value provided to write",
            [
                "Provide a value as the third argument (or --value for negatives)",
                "Use --json or --from-json for array data",
                "Use --from-csv to read from a CSV file",
            ],
        )

    # Apply number_format to all cells if specified and not already set
    if number_format:
        for entry in write_data:
            if "number_format" not in entry:
                entry["number_format"] = number_format

    result = oxl.write_cells(
        str(path),
        sheet_name=sheet,
        data=write_data,
        output_path=output,
    )

    # Add the range to the result
    if len(write_data) == 1:
        result["range"] = write_data[0]["cell"]
    else:
        first = write_data[0]["cell"]
        last = write_data[-1]["cell"]
        result["range"] = f"{first}:{last}"

    json_formatter.relativize_path(result)
    json_formatter.output(result)


def _coerce_value(val: str):
    """Attempt to coerce a string to int or float if it looks numeric."""
    try:
        int_val = int(val)
        return int_val
    except ValueError:
        pass
    try:
        float_val = float(val)
        return float_val
    except ValueError:
        pass
    return val


def _col_offset(col_letter: str, offset: int) -> str:
    """Shift a column letter by offset positions."""
    from agent_xlsx.utils.validation import col_letter_to_index, index_to_col_letter

    idx = col_letter_to_index(col_letter)
    return index_to_col_letter(idx + offset)


def _parse_cell_ref(ref: str) -> tuple[str, int]:
    """Split 'AB12' into ('AB', 12)."""
    import re

    m = re.match(r"^([A-Z]+)(\d+)$", ref.upper())
    if not m:
        raise AgentExcelError(
            "RANGE_INVALID",
            f"Invalid cell reference: '{ref}'",
            ["Use Excel notation e.g. 'A1'"],
        )
    return m.group(1), int(m.group(2))


def _json_to_cells(start_cell: str, json_str: str, *, formula: bool = False) -> list[dict]:
    """Convert a JSON 2D array to a list of cell write entries, starting at start_cell.

    Strings starting with '=' are auto-detected as formulas by openpyxl.
    The formula parameter is accepted for API consistency but has no effect
    on JSON data (values are already typed).
    """
    try:
        data = json_mod.loads(json_str)
    except json_mod.JSONDecodeError as exc:
        raise AgentExcelError(
            "INVALID_JSON",
            f"Failed to parse JSON: {exc}",
            ['Provide a valid JSON 2D array e.g. \'[["a","b"],["c","d"]]\''],
        )

    if not isinstance(data, list):
        raise AgentExcelError(
            "INVALID_JSON",
            "JSON data must be a 2D array (list of lists)",
            ['e.g. \'[["a","b"],["c","d"]]\''],
        )

    # Parse the starting cell reference (ignore range end if provided)
    ref = start_cell.split(":")[0]
    start_col, start_row = _parse_cell_ref(ref)

    cells: list[dict] = []
    for row_idx, row in enumerate(data):
        if not isinstance(row, list):
            row = [row]
        for col_idx, val in enumerate(row):
            col = _col_offset(start_col, col_idx)
            row_num = start_row + row_idx
            cells.append({"cell": f"{col}{row_num}", "value": val})
    return cells


def _csv_to_cells(start_cell: str, csv_path: Path, *, formula: bool = False) -> list[dict]:
    """Read a CSV file and map rows to cell write entries starting at start_cell.

    When formula=True, strings starting with '=' are preserved as-is
    (skipping numeric coercion) so openpyxl writes them as formulas.
    All other values are coerced normally.
    """
    ref = start_cell.split(":")[0]
    start_col, start_row = _parse_cell_ref(ref)

    cells: list[dict] = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row_idx, row in enumerate(reader):
            for col_idx, val in enumerate(row):
                col = _col_offset(start_col, col_idx)
                row_num = start_row + row_idx
                if formula and val.startswith("="):
                    pass  # Preserve formula string as-is (skip numeric coercion)
                else:
                    val = _coerce_value(val)
                cells.append({"cell": f"{col}{row_num}", "value": val})
    return cells

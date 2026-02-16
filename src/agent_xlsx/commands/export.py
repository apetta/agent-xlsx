"""Export sheet data to JSON, CSV, or Markdown format via Polars."""

import sys
import time
from pathlib import Path
from typing import Any, Optional

import typer

from agent_xlsx.adapters.polars_adapter import get_sheet_names, read_sheet_data
from agent_xlsx.cli import app
from agent_xlsx.formatters.json_formatter import output
from agent_xlsx.utils.errors import SheetNotFoundError, handle_error
from agent_xlsx.utils.validation import validate_file


@app.command()
@handle_error
def export(
    file: str = typer.Argument(..., help="Path to the Excel file"),
    sheet: Optional[str] = typer.Option(None, "--sheet", "-s", help="Sheet to export"),
    format_: str = typer.Option(
        "json",
        "--format",
        "-f",
        help="Output format: json, csv, markdown",
    ),
    output_path: Optional[str] = typer.Option(None, "--output", "-o", help="Output file path"),
) -> None:
    """Export a sheet to JSON, CSV, or Markdown format.

    Uses Polars + fastexcel for fast data extraction.
    """
    path = validate_file(file)
    start = time.perf_counter()

    target_sheet = sheet or 0

    # Validate sheet name
    if isinstance(target_sheet, str):
        available = get_sheet_names(str(path))
        if target_sheet not in available:
            raise SheetNotFoundError(target_sheet, available)

    df = read_sheet_data(filepath=path, sheet_name=target_sheet)

    elapsed_ms = round((time.perf_counter() - start) * 1000, 1)

    fmt = format_.lower()
    if fmt == "csv":
        _export_csv(df, output_path)
    elif fmt == "markdown":
        _export_markdown(df, output_path)
    else:
        _export_json(df, output_path, path, target_sheet, elapsed_ms)


def _export_json(
    df: Any,
    output_path: Optional[str],
    source: Path,
    sheet: Any,
    elapsed_ms: float,
) -> None:
    """Export as JSON — either to file or stdout."""
    import json

    from agent_xlsx.formatters.json_formatter import _serialise

    rows = []
    for row in df.rows(named=True):
        clean = {}
        for k, v in row.items():
            if v is None:
                clean[k] = None
            elif hasattr(v, "isoformat"):
                clean[k] = v.isoformat()
            elif isinstance(v, float) and v != v:
                clean[k] = None
            else:
                clean[k] = v
        rows.append(clean)

    data = {
        "source": Path(source).name if hasattr(source, "name") else str(source),
        "sheet": sheet if isinstance(sheet, str) else None,
        "row_count": len(rows),
        "columns": df.columns,
        "data": rows,
        "export_time_ms": elapsed_ms,
    }

    if output_path:
        with open(output_path, "w") as f:
            json.dump(data, f, indent=2, default=_serialise)
        output(
            {
                "status": "success",
                "format": "json",
                "output": output_path,
                "row_count": len(rows),
                "export_time_ms": elapsed_ms,
            }
        )
    else:
        output(data)


def _export_csv(df: Any, output_path: Optional[str]) -> None:
    """Export as CSV — either to file or stdout."""
    csv_str = df.write_csv()

    if output_path:
        with open(output_path, "w") as f:
            f.write(csv_str)
        output(
            {
                "status": "success",
                "format": "csv",
                "output": output_path,
                "row_count": len(df),
            }
        )
    else:
        sys.stdout.write(csv_str)


def _export_markdown(df: Any, output_path: Optional[str]) -> None:
    """Export as Markdown table — either to file or stdout."""
    lines = _df_to_markdown(df)
    md_str = "\n".join(lines) + "\n"

    if output_path:
        with open(output_path, "w") as f:
            f.write(md_str)
        output(
            {
                "status": "success",
                "format": "markdown",
                "output": output_path,
                "row_count": len(df),
            }
        )
    else:
        sys.stdout.write(md_str)


def _df_to_markdown(df) -> list[str]:
    """Convert a Polars DataFrame to Markdown table lines."""
    headers = df.columns
    rows = df.rows()

    # Format cells
    def fmt(val) -> str:
        if val is None:
            return ""
        if isinstance(val, float) and val != val:
            return ""
        if hasattr(val, "isoformat"):
            return val.isoformat()
        return str(val)

    header_line = "| " + " | ".join(headers) + " |"
    separator = "| " + " | ".join("---" for _ in headers) + " |"
    data_lines = ["| " + " | ".join(fmt(v) for v in row) + " |" for row in rows]

    return [header_line, separator, *data_lines]

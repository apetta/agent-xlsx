"""Read data from Excel ranges — fast path via Polars, openpyxl fallback for formulas."""

import time
from pathlib import Path
from typing import Optional

import typer

from agent_xlsx.adapters.polars_adapter import get_sheet_names, read_sheet_data
from agent_xlsx.cli import app
from agent_xlsx.formatters.json_formatter import output
from agent_xlsx.utils.constants import DEFAULT_LIMIT, DEFAULT_OFFSET, MAX_READ_ROWS
from agent_xlsx.utils.dates import convert_date_values, detect_date_columns
from agent_xlsx.utils.errors import SheetNotFoundError, handle_error
from agent_xlsx.utils.validation import parse_range, validate_file


@app.command()
@handle_error
def read(
    file: str = typer.Argument(..., help="Path to the Excel file"),
    range_: Optional[str] = typer.Argument(
        None,
        metavar="RANGE",
        help="Range e.g. 'Sheet1!A1:C10'",
    ),
    sheet: Optional[str] = typer.Option(None, "--sheet", "-s", help="Sheet name"),
    limit: int = typer.Option(DEFAULT_LIMIT, "--limit", "-l", help="Maximum rows to return"),
    offset: int = typer.Option(DEFAULT_OFFSET, "--offset", help="Rows to skip"),
    format_: str = typer.Option("json", "--format", "-f", help="Output format: json or csv"),
    formulas: bool = typer.Option(
        False,
        "--formulas",
        help="Include formula strings (slower)",
    ),
    sort: Optional[str] = typer.Option(
        None,
        "--sort",
        help="Sort by column name",
    ),
    descending: bool = typer.Option(
        False,
        "--descending",
        help="Sort in descending order",
    ),
) -> None:
    """Read data from an Excel range or sheet.

    Default fast path uses Polars + fastexcel (7-10x faster than openpyxl).
    Use --formulas to fall back to openpyxl for formula string extraction.
    """
    path = validate_file(file)
    start = time.perf_counter()

    # Cap limit to prevent massive outputs
    effective_limit = min(limit, MAX_READ_ROWS)

    if formulas:
        _read_with_formulas(path, range_, sheet, effective_limit, offset)
        return

    # Fast path — Polars + fastexcel
    target_sheet = sheet or 0
    use_columns = None
    range_info = None

    if range_:
        range_info = parse_range(range_)
        if range_info["sheet"]:
            target_sheet = range_info["sheet"]

    # Validate sheet name if it's a string
    if isinstance(target_sheet, str):
        available = get_sheet_names(str(path))
        if target_sheet not in available:
            raise SheetNotFoundError(target_sheet, available)

    # Build column filter from range
    if range_info and range_info["start"]:
        start_col_letters = "".join(c for c in range_info["start"] if c.isalpha())
        if range_info["end"]:
            end_col_letters = "".join(c for c in range_info["end"] if c.isalpha())
            use_columns = f"{start_col_letters}:{end_col_letters}"

    df = read_sheet_data(
        filepath=path,
        sheet_name=target_sheet,
        skip_rows=offset,
        n_rows=effective_limit,
        use_columns=use_columns,
    )

    # Apply row range filtering if range specifies rows
    if range_info and range_info["start"]:
        start_row = int("".join(c for c in range_info["start"] if c.isdigit()))
        if range_info["end"]:
            end_row = int("".join(c for c in range_info["end"] if c.isdigit()))
            # Rows in range are 1-based; row 1 is header, data starts at row 2 (index 0)
            row_start = max(start_row - 2, 0)  # -2: 1 for header, 1 for 0-based
            row_end = end_row - 1  # -1 for header
            df = df.slice(row_start, row_end - row_start)

    # Sort if requested
    if sort:
        if sort in df.columns:
            df = df.sort(sort, descending=descending)

    # Build output
    elapsed_ms = round((time.perf_counter() - start) * 1000, 1)

    if format_ == "csv":
        _output_csv(df)
        return

    sheet_name_str = target_sheet if isinstance(target_sheet, str) else df.columns[0]
    range_str = range_ or f"{sheet_name_str}"

    # Convert to rows
    rows = _df_to_serialisable_rows(df)

    # Convert Excel serial numbers to ISO dates in date columns
    sheet_name_str = target_sheet if isinstance(target_sheet, str) else str(target_sheet)
    try:
        sheet_arg = sheet_name_str if isinstance(target_sheet, str) else None
        date_cols = detect_date_columns(str(path), sheet_arg)
        if date_cols:
            rows = convert_date_values(rows, list(df.columns), set(date_cols.keys()))
    except Exception:
        pass  # date detection is best-effort; don't break reads

    result = {
        "range": range_str,
        "dimensions": {"rows": len(df), "cols": len(df.columns)},
        "headers": df.columns,
        "data": rows,
        "row_count": len(df),
        "truncated": len(df) >= effective_limit,
        "backend": "polars+fastexcel",
        "read_time_ms": elapsed_ms,
    }

    output(result)


def _read_with_formulas(
    path: Path,
    range_str: Optional[str],
    sheet: Optional[str],
    limit: int,
    offset: int,
) -> None:
    """Read with formula strings via openpyxl (slower path)."""
    start = time.perf_counter()

    from openpyxl import load_workbook

    wb = load_workbook(str(path), read_only=True, data_only=False)

    target_sheet = sheet
    range_info = None

    if range_str:
        range_info = parse_range(range_str)
        if range_info["sheet"]:
            target_sheet = range_info["sheet"]

    if target_sheet:
        if target_sheet not in wb.sheetnames:
            raise SheetNotFoundError(target_sheet, wb.sheetnames)
        ws = wb[target_sheet]
    else:
        ws = wb.active
        target_sheet = ws.title

    # Determine cell range
    if range_info and range_info["start"]:
        cell_range = range_info["start"]
        if range_info["end"]:
            cell_range = f"{range_info['start']}:{range_info['end']}"
        rows = list(ws[cell_range])
    else:
        rows = list(ws.iter_rows(min_row=1 + offset, max_row=1 + offset + limit))

    cells: list[dict] = []
    for row in rows:
        for cell in row:
            value = cell.value
            formula = None
            if isinstance(value, str) and value.startswith("="):
                formula = value
                # Try to get computed value from data_only workbook
                value = formula  # We don't have the computed value in non-data_only mode

            cells.append(
                {
                    "cell": cell.coordinate,
                    "value": value,
                    "formula": formula,
                }
            )

    wb.close()

    elapsed_ms = round((time.perf_counter() - start) * 1000, 1)

    result = {
        "range": range_str or target_sheet,
        "cells": cells[: limit * 20],  # Cap cell output
        "cell_count": len(cells),
        "truncated": len(cells) > limit * 20,
        "backend": "openpyxl",
        "read_time_ms": elapsed_ms,
    }

    output(result)


def _output_csv(df) -> None:
    """Output DataFrame as CSV to stdout."""
    import sys

    csv_str = df.write_csv()
    sys.stdout.write(csv_str)


def _df_to_serialisable_rows(df) -> list[list]:
    """Convert DataFrame rows to JSON-serialisable lists."""
    rows = df.rows()
    result = []
    for row in rows:
        clean = []
        for val in row:
            if val is None:
                clean.append(None)
            elif hasattr(val, "isoformat"):
                clean.append(val.isoformat())
            elif isinstance(val, float) and val != val:
                clean.append(None)
            else:
                clean.append(val)
        result.append(clean)
    return result

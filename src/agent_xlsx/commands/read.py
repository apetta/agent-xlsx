"""Read data from Excel ranges — fast path via Polars, openpyxl fallback for formulas."""

import time
from pathlib import Path
from typing import Optional

import polars as pl
import typer

from agent_xlsx.adapters.polars_adapter import get_sheet_names, read_exact_range, read_sheet_data
from agent_xlsx.cli import app
from agent_xlsx.formatters.json_formatter import output, output_spreadsheet_data
from agent_xlsx.utils.constants import DEFAULT_LIMIT, DEFAULT_OFFSET, MAX_READ_ROWS
from agent_xlsx.utils.dataframe import apply_compact
from agent_xlsx.utils.dates import detect_date_column_indices, excel_serial_to_isodate
from agent_xlsx.utils.errors import SheetNotFoundError, handle_error
from agent_xlsx.utils.validation import (
    col_letter_to_index,
    parse_multi_range,
    parse_range,
    validate_file,
)


@app.command()
@handle_error
def read(
    file: str = typer.Argument(..., help="Path to the Excel file"),
    range_: Optional[str] = typer.Argument(
        None,
        metavar="RANGE",
        help="Range e.g. 'Sheet1!A1:C10' or comma-separated 'Sheet1!A1:C10,E1:G10'",
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
    no_header: bool = typer.Option(
        False,
        "--no-header",
        help="Treat row 1 as data, use column letters (A, B, C) as headers. "
        "Use for non-tabular sheets like P&L reports and dashboards.",
    ),
    compact: bool = typer.Option(
        True,
        "--compact/--no-compact",
        help="Drop fully-null columns from output to reduce token waste (default: on).",
    ),
    all_sheets: bool = typer.Option(
        False,
        "--all-sheets",
        help="Read the same range(s) from every sheet in the workbook.",
    ),
) -> None:
    """Read data from an Excel range or sheet.

    Default fast path uses Polars + fastexcel (7-10x faster than openpyxl).
    Use --formulas to fall back to openpyxl for formula string extraction.

    Supports comma-separated ranges (e.g. 'Sheet1!A1:C10,E1:G10') for
    multi-range reads, and --all-sheets to read the same range(s) from
    every sheet.
    """
    path = validate_file(file)
    start = time.perf_counter()

    # Cap limit to prevent massive outputs
    effective_limit = min(limit, MAX_READ_ROWS)

    if formulas:
        _read_with_formulas(path, range_, sheet, effective_limit, offset, compact)
        return

    # --- Determine ranges to read ---
    is_multi_range = range_ is not None and "," in range_

    if is_multi_range:
        ranges = parse_multi_range(range_)
    elif range_:
        ranges = [parse_range(range_)]
    else:
        ranges = []  # No range — full sheet read

    # --- Determine target sheets ---
    available = get_sheet_names(str(path))

    if all_sheets:
        target_sheets: list[str | int] = list(available)
    elif ranges and ranges[0]["sheet"]:
        ts = ranges[0]["sheet"]
        if ts not in available:
            raise SheetNotFoundError(ts, available)
        target_sheets = [ts]
    elif sheet:
        if sheet not in available:
            raise SheetNotFoundError(sheet, available)
        target_sheets = [sheet]
    else:
        target_sheets = [0]  # Default to first sheet

    # --- Multi-result path (multi-range OR all-sheets) ---
    is_multi = is_multi_range or all_sheets

    if is_multi:
        results = []
        for target_sheet in target_sheets:
            sheet_name = target_sheet if isinstance(target_sheet, str) else available[target_sheet]
            if ranges:
                for ri in ranges:
                    df = _read_single_range(
                        path, target_sheet, ri, no_header, effective_limit, offset
                    )
                    df = apply_compact(df, compact)
                    if sort and sort in df.columns:
                        df = df.sort(sort, descending=descending)

                    rows = _df_to_serialisable_rows(df)
                    rows = _apply_date_conversion(rows, df, path, target_sheet)

                    range_str = (
                        f"{ri['start']}:{ri['end']}" if ri.get("end") else ri.get("start", "")
                    )
                    results.append(
                        {
                            "range": range_str,
                            "sheet": sheet_name,
                            "headers": df.columns,
                            "data": rows,
                            "row_count": len(df),
                        }
                    )
            else:
                # No range — full sheet read per sheet
                df = read_sheet_data(
                    filepath=path,
                    sheet_name=target_sheet,
                    skip_rows=offset,
                    n_rows=effective_limit,
                    no_header=no_header,
                )
                df = apply_compact(df, compact)
                if sort and sort in df.columns:
                    df = df.sort(sort, descending=descending)

                rows = _df_to_serialisable_rows(df)
                rows = _apply_date_conversion(rows, df, path, target_sheet)

                results.append(
                    {
                        "range": sheet_name,
                        "sheet": sheet_name,
                        "headers": df.columns,
                        "data": rows,
                        "row_count": len(df),
                    }
                )

        elapsed_ms = round((time.perf_counter() - start) * 1000, 1)
        result = {
            "results": results,
            "total_ranges": len(results),
            "compact": compact,
            "read_time_ms": elapsed_ms,
        }
        output_spreadsheet_data(result)
        return

    # --- Single-range path (existing behaviour, backward compatible) ---
    target_sheet = target_sheets[0]
    range_info = ranges[0] if ranges else None

    if range_info and range_info["start"]:
        df = _read_single_range(path, target_sheet, range_info, no_header, effective_limit, offset)
    else:
        df = read_sheet_data(
            filepath=path,
            sheet_name=target_sheet,
            skip_rows=offset,
            n_rows=effective_limit,
            no_header=no_header,
        )

    df = apply_compact(df, compact)

    if sort and sort in df.columns:
        df = df.sort(sort, descending=descending)

    elapsed_ms = round((time.perf_counter() - start) * 1000, 1)

    if format_ == "csv":
        _output_csv(df)
        return

    sheet_name_str = target_sheet if isinstance(target_sheet, str) else df.columns[0]
    range_str = range_ or f"{sheet_name_str}"

    rows = _df_to_serialisable_rows(df)
    rows = _apply_date_conversion(rows, df, path, target_sheet)

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

    output_spreadsheet_data(result)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _read_single_range(
    path: Path,
    target_sheet: str | int,
    range_info: dict,
    no_header: bool,
    effective_limit: int,
    offset: int,
) -> pl.DataFrame:
    """Read a single parsed range from a sheet, returning a DataFrame."""
    if range_info and range_info["start"]:
        start_col = "".join(c for c in range_info["start"] if c.isalpha())
        start_row = int("".join(c for c in range_info["start"] if c.isdigit()))

        if range_info["end"]:
            end_col = "".join(c for c in range_info["end"] if c.isalpha())
            end_row = int("".join(c for c in range_info["end"] if c.isdigit()))
        else:
            end_col = start_col
            end_row = start_row

        return read_exact_range(
            filepath=path,
            sheet_name=target_sheet,
            start_col_idx=col_letter_to_index(start_col),
            end_col_idx=col_letter_to_index(end_col),
            start_row=start_row,
            end_row=end_row,
        )

    return read_sheet_data(
        filepath=path,
        sheet_name=target_sheet,
        skip_rows=offset,
        n_rows=effective_limit,
        no_header=no_header,
    )


def _apply_date_conversion(
    rows: list[list],
    df: pl.DataFrame,
    path: Path,
    target_sheet: str | int,
) -> list[list]:
    """Best-effort conversion of Excel serial numbers to ISO dates.

    Uses index-based detection so it works in both normal and --no-header mode.
    """
    try:
        sheet_arg = target_sheet if isinstance(target_sheet, str) else None
        date_indices = detect_date_column_indices(str(path), sheet_arg)
        if not date_indices:
            return rows
        for row in rows:
            for idx in date_indices:
                if idx >= len(row):
                    continue
                val = row[idx]
                # --no-header makes all columns String; coerce numeric strings
                if isinstance(val, str):
                    try:
                        val = float(val)
                    except (ValueError, TypeError):
                        continue
                if isinstance(val, (int, float)) and val == val and val > 0:
                    converted = excel_serial_to_isodate(float(val))
                    if converted is not None:
                        row[idx] = converted
    except Exception:
        pass  # date detection is best-effort; don't break reads
    return rows


def _read_with_formulas(
    path: Path,
    range_str: Optional[str],
    sheet: Optional[str],
    limit: int,
    offset: int,
    compact: bool = True,
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
        if range_info["end"]:
            cell_range = f"{range_info['start']}:{range_info['end']}"
        else:
            # Single cell — expand to self-range for consistent openpyxl
            # tuple-of-tuples return (ws["A1"] returns a bare Cell object)
            cell_range = f"{range_info['start']}:{range_info['start']}"
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

    # Strip blank cells (both value and formula null) in compact mode
    if compact:
        cells = [c for c in cells if c["value"] is not None or c["formula"] is not None]

    elapsed_ms = round((time.perf_counter() - start) * 1000, 1)

    result = {
        "range": range_str or target_sheet,
        "cells": cells[: limit * 20],  # Cap cell output
        "cell_count": len(cells),
        "truncated": len(cells) > limit * 20,
        "backend": "openpyxl",
        "read_time_ms": elapsed_ms,
    }

    output_spreadsheet_data(result)


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

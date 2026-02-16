"""Primary data adapter using Polars + fastexcel (Calamine) for blazingly fast Excel reads."""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

import fastexcel
import polars as pl

from agent_xlsx.utils.constants import (
    CHUNK_SIZE_ROWS,
    CHUNK_THRESHOLD_BYTES,
    MAX_SAMPLE_ROWS,
    MAX_SEARCH_RESULTS,
)
from agent_xlsx.utils.dates import (
    convert_date_values,
    detect_date_columns,
    excel_serial_to_isodate,
)
from agent_xlsx.utils.errors import SheetNotFoundError
from agent_xlsx.utils.validation import file_size_bytes, index_to_col_letter


def get_sheet_names(filepath: str | Path) -> list[str]:
    """Return sheet names from a workbook via fastexcel (instant, no data parsed)."""
    reader = fastexcel.read_excel(str(filepath))
    return reader.sheet_names


def get_sheet_dimensions(filepath: str | Path, sheet_name: str | int = 0) -> dict[str, Any]:
    """Return sheet dimensions without loading data (n_rows=0 trick)."""
    reader = fastexcel.read_excel(str(filepath))
    sheet = reader.load_sheet(sheet_name, n_rows=0)
    return {
        "name": sheet.name,
        "rows": sheet.total_height,
        "cols": sheet.width,
        "headers": [c.name for c in sheet.available_columns()],
        "visible": sheet.visible == "visible",
    }


def read_sheet_data(
    filepath: str | Path,
    sheet_name: str | int = 0,
    skip_rows: int = 0,
    n_rows: int | None = None,
    use_columns: str | list[str] | None = None,
) -> pl.DataFrame:
    """Read sheet data into a Polars DataFrame via fastexcel (zero-copy Arrow).

    For large files (>100MB), automatically chunks the read to stay within memory budget.
    """
    fpath = str(filepath)
    size = file_size_bytes(fpath)

    # Polars read_excel requires string sheet names — resolve integer indices
    resolved_name = _resolve_sheet_name(fpath, sheet_name)

    read_opts: dict[str, Any] = {}
    if skip_rows:
        read_opts["skip_rows"] = skip_rows
    if n_rows is not None:
        read_opts["n_rows"] = n_rows
    if use_columns is not None:
        read_opts["use_columns"] = use_columns

    if size < CHUNK_THRESHOLD_BYTES or n_rows is not None:
        # Direct read — fastexcel handles this efficiently
        return pl.read_excel(fpath, sheet_name=resolved_name, read_options=read_opts or None)

    # Chunked read for very large files
    return _read_chunked(fpath, resolved_name, skip_rows, n_rows)


def _read_chunked(
    filepath: str,
    sheet_name: str | int,
    skip_rows: int,
    n_rows: int | None,
) -> pl.DataFrame:
    """Read in chunks using fastexcel's skip_rows + n_rows for very large files."""
    reader = fastexcel.read_excel(filepath)
    probe = reader.load_sheet(sheet_name, n_rows=0)
    total_rows = probe.total_height

    target_rows = n_rows if n_rows is not None else total_rows - skip_rows
    chunks: list[pl.DataFrame] = []
    rows_read = 0

    offset = skip_rows
    while rows_read < target_rows and offset < total_rows:
        chunk_n = min(CHUNK_SIZE_ROWS, target_rows - rows_read)
        sheet = reader.load_sheet(
            sheet_name,
            skip_rows=offset,
            n_rows=chunk_n,
        )
        chunk_df = pl.DataFrame(sheet)
        chunks.append(chunk_df)
        rows_read += len(chunk_df)
        offset += chunk_n

    if not chunks:
        # Return empty DataFrame with proper schema
        sheet = reader.load_sheet(sheet_name, n_rows=0)
        return pl.DataFrame(sheet)

    return pl.concat(chunks)


def probe_workbook(
    filepath: str | Path,
    sheet_name: str | None = None,
    sample_rows: int = 0,
    stats: bool = False,
    include_types: bool = False,
) -> dict[str, Any]:
    """Ultra-fast workbook profiling via fastexcel + Polars.

    Default is a lean metadata-only probe (sheet names, dimensions, headers) with
    zero data parsing.  Pass ``include_types``, ``sample_rows``, or ``stats`` to
    opt into progressively richer detail.
    """
    start = time.perf_counter()
    fpath = str(filepath)
    reader = fastexcel.read_excel(fpath)

    all_sheet_names = reader.sheet_names
    target_sheets = [sheet_name] if sheet_name else all_sheet_names

    # Validate requested sheet exists
    if sheet_name and sheet_name not in all_sheet_names:
        raise SheetNotFoundError(sheet_name, all_sheet_names)

    needs_data = include_types or sample_rows > 0 or stats
    sheets_result: list[dict[str, Any]] = []

    for idx, name in enumerate(all_sheet_names):
        if name not in target_sheets:
            continue

        # Fast metadata via n_rows=0 — no data parsed
        meta = reader.load_sheet(name, n_rows=0)
        sheet_info: dict[str, Any] = {
            "name": name,
            "index": idx,
            "visible": meta.visible == "visible",
            "rows": meta.total_height,
            "cols": meta.width,
        }

        if not needs_data:
            # Lean path — headers from metadata only, zero data parsing
            sheet_info["headers"] = [c.name for c in meta.available_columns()]
            headers = sheet_info["headers"]
            sheet_info["last_col"] = index_to_col_letter(len(headers) - 1) if headers else "A"
            sheet_info["column_map"] = {h: index_to_col_letter(i) for i, h in enumerate(headers)}
            sheets_result.append(sheet_info)
            continue

        # Load full sheet data for profiling
        sheet = reader.load_sheet(name)
        df = pl.DataFrame(sheet)

        sheet_info["headers"] = df.columns
        headers = sheet_info["headers"]
        sheet_info["last_col"] = index_to_col_letter(len(headers) - 1) if headers else "A"
        sheet_info["column_map"] = {h: index_to_col_letter(i) for i, h in enumerate(headers)}

        if include_types:
            sheet_info["column_types"] = {
                col: _polars_dtype_to_str(dtype) for col, dtype in zip(df.columns, df.dtypes)
            }

            # Detect date columns masquerading as float64
            float_cols = [
                col for col, t in sheet_info["column_types"].items()
                if t in ("float64", "float32")
            ]
            if float_cols:
                date_col_names = detect_date_columns(fpath, name)
                for col in float_cols:
                    if col in date_col_names:
                        sheet_info["column_types"][col] = "date"

            # Null counts per column
            if len(df) > 0:
                null_row = df.null_count().row(0)
                sheet_info["null_counts"] = dict(zip(df.columns, null_row))
            else:
                sheet_info["null_counts"] = {col: 0 for col in df.columns}

        # Sample data (head + tail)
        capped_sample = min(sample_rows, MAX_SAMPLE_ROWS)
        if capped_sample > 0 and len(df) > 0:
            date_col_set = {
                col for col, t in sheet_info.get("column_types", {}).items()
                if t == "date"
            }
            head_rows = _df_to_rows(df.head(capped_sample))
            tail_rows = _df_to_rows(df.tail(capped_sample))
            if date_col_set:
                head_rows = convert_date_values(head_rows, df.columns, date_col_set)
                tail_rows = convert_date_values(tail_rows, df.columns, date_col_set)
            sheet_info["sample"] = {
                "head": head_rows,
                "tail": tail_rows,
            }

        # Extended statistics via Polars describe()
        if stats and len(df) > 0:
            date_col_set = {
                col for col, t in sheet_info.get("column_types", {}).items()
                if t == "date"
            }
            numeric_cols = [col for col, dtype in zip(df.columns, df.dtypes)
                            if dtype.is_numeric() and col not in date_col_set]
            if numeric_cols:
                sheet_info["numeric_summary"] = _build_numeric_summary(df, numeric_cols)

            string_cols = [col for col, dtype in zip(df.columns, df.dtypes) if dtype == pl.Utf8]
            if string_cols:
                sheet_info["string_summary"] = _build_string_summary(df, string_cols)

            # Date column summary (min/max as ISO dates)
            if date_col_set:
                date_summary = {}
                for col in date_col_set:
                    series = df[col].drop_nulls()
                    if len(series) > 0:
                        min_val = _safe_scalar(series.min())
                        max_val = _safe_scalar(series.max())
                        if isinstance(min_val, (int, float)) and min_val == min_val:
                            date_summary[col] = {
                                "min": excel_serial_to_isodate(float(min_val)),
                                "max": excel_serial_to_isodate(float(max_val)),
                                "count": len(series),
                            }
                if date_summary:
                    sheet_info["date_summary"] = date_summary

        sheets_result.append(sheet_info)

    elapsed_ms = round((time.perf_counter() - start) * 1000, 1)

    # Workbook-level metadata from fastexcel
    result: dict[str, Any] = {
        "file": Path(filepath).name,
        "size_bytes": file_size_bytes(fpath),
        "format": Path(filepath).suffix.lstrip(".").lower(),
        "probe_time_ms": elapsed_ms,
        "sheets": sheets_result,
    }

    # Named ranges and tables (fast via fastexcel)
    try:
        defined = reader.defined_names()
        result["named_ranges"] = [n["name"] for n in defined] if defined else []
    except Exception:
        result["named_ranges"] = []

    try:
        result["tables"] = reader.table_names() or []
    except Exception:
        result["tables"] = []

    # VBA detection — check file extension
    result["has_vba"] = Path(filepath).suffix.lower() in {".xlsm", ".xlsb"}

    return result


def search_values(
    filepath: str | Path,
    query: str,
    sheet_name: str | None = None,
    regex: bool = False,
    ignore_case: bool = False,
) -> list[dict[str, Any]]:
    """Search for values across sheets using Polars string matching.

    Returns list of match dicts: {sheet, column, row, value}.
    """
    fpath = str(filepath)
    reader = fastexcel.read_excel(fpath)
    all_sheets = reader.sheet_names

    if sheet_name:
        if sheet_name not in all_sheets:
            raise SheetNotFoundError(sheet_name, all_sheets)
        target_sheets = [sheet_name]
    else:
        target_sheets = all_sheets

    matches: list[dict[str, Any]] = []

    for name in target_sheets:
        sheet = reader.load_sheet(name)
        df = pl.DataFrame(sheet)

        if len(df) == 0:
            continue

        for col in df.columns:
            # Cast column to string for searching
            str_col = df[col].cast(pl.Utf8, strict=False)

            if regex:
                if ignore_case:
                    mask = str_col.str.contains(f"(?i){query}")
                else:
                    mask = str_col.str.contains(query)
            elif ignore_case:
                mask = str_col.str.to_lowercase().str.contains(query.lower(), literal=True)
            else:
                mask = str_col.str.contains(query, literal=True)

            # Get matching row indices
            match_indices = df.with_row_index("__row_idx__").filter(mask)["__row_idx__"].to_list()

            for row_idx in match_indices:
                cell_value = df[col][row_idx]
                # Convert to Python native
                if hasattr(cell_value, "item"):
                    cell_value = cell_value.item()

                col_idx = df.columns.index(col)
                # +2: 1-based indexing + header row
                cell_ref = f"{index_to_col_letter(col_idx)}{row_idx + 2}"

                matches.append(
                    {
                        "sheet": name,
                        "column": col,
                        "row": row_idx + 1,  # 1-based for user display
                        "cell": cell_ref,
                        "value": cell_value,
                    }
                )

                if len(matches) >= MAX_SEARCH_RESULTS:
                    return matches

    return matches


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _resolve_sheet_name(filepath: str, sheet_name: str | int) -> str:
    """Resolve an integer sheet index to a string name. Polars requires string names."""
    if isinstance(sheet_name, str):
        return sheet_name
    reader = fastexcel.read_excel(filepath)
    names = reader.sheet_names
    if sheet_name < 0 or sheet_name >= len(names):
        raise SheetNotFoundError(str(sheet_name), names)
    return names[sheet_name]


def _polars_dtype_to_str(dtype: pl.DataType) -> str:
    """Convert Polars dtype to a short human-readable string."""
    mapping = {
        pl.Utf8: "string",
        pl.String: "string",
        pl.Int8: "int8",
        pl.Int16: "int16",
        pl.Int32: "int32",
        pl.Int64: "int64",
        pl.UInt8: "uint8",
        pl.UInt16: "uint16",
        pl.UInt32: "uint32",
        pl.UInt64: "uint64",
        pl.Float32: "float32",
        pl.Float64: "float64",
        pl.Boolean: "boolean",
        pl.Date: "date",
        pl.Time: "time",
        pl.Null: "null",
    }
    # Handle parameterised types like Datetime(...)
    for base_type, label in mapping.items():
        if dtype == base_type:
            return label

    dtype_str = str(dtype).lower()
    if "datetime" in dtype_str:
        return "datetime"
    if "duration" in dtype_str:
        return "duration"

    return str(dtype).lower()


def _df_to_rows(df: pl.DataFrame) -> list[list[Any]]:
    """Convert DataFrame to a list of rows, handling special types for JSON serialisation."""
    rows = df.rows()
    result = []
    for row in rows:
        clean_row = []
        for val in row:
            if val is None:
                clean_row.append(None)
            elif hasattr(val, "isoformat"):
                clean_row.append(val.isoformat())
            elif isinstance(val, float) and val != val:
                clean_row.append(None)  # NaN → null
            else:
                clean_row.append(val)
        result.append(clean_row)
    return result


def _build_numeric_summary(df: pl.DataFrame, numeric_cols: list[str]) -> dict[str, Any]:
    """Build numeric summary statistics using Polars."""
    summary: dict[str, Any] = {}
    for col in numeric_cols:
        series = df[col].drop_nulls()
        if len(series) == 0:
            continue
        summary[col] = {
            "min": _safe_scalar(series.min()),
            "max": _safe_scalar(series.max()),
            "mean": _safe_scalar(series.mean()),
            "median": _safe_scalar(series.median()),
            "std": _safe_scalar(series.std()),
        }
    return summary


def _build_string_summary(df: pl.DataFrame, string_cols: list[str]) -> dict[str, Any]:
    """Build string column summary with unique counts and top values."""
    summary: dict[str, Any] = {}
    for col in string_cols:
        series = df[col].drop_nulls()
        if len(series) == 0:
            continue
        n_unique = series.n_unique()
        # Get top values by frequency (up to 10)
        top = series.value_counts(sort=True).head(10).get_column(col).to_list()
        summary[col] = {
            "unique": n_unique,
            "top_values": top,
        }
    return summary


def _safe_scalar(val: Any) -> Any:
    """Convert Polars scalar to JSON-safe Python type."""
    if val is None:
        return None
    if isinstance(val, float) and val != val:
        return None
    if hasattr(val, "item"):
        return val.item()
    if isinstance(val, float) and val == int(val):
        return int(val)
    if isinstance(val, float):
        return round(val, 6)
    return val

"""Shared validation helpers for file paths and ranges."""

from __future__ import annotations

import os
import re
from pathlib import Path

from agent_xlsx.utils.constants import EXCEL_EXTENSIONS, WRITABLE_EXTENSIONS
from agent_xlsx.utils.errors import (
    AgentExcelError,
    ExcelFileNotFoundError,
    InvalidFormatError,
    RangeInvalidError,
)


def validate_file(filepath: str) -> Path:
    """Validate that the file exists and has a supported extension. Returns resolved Path."""
    p = Path(filepath).resolve()
    if not p.exists():
        raise ExcelFileNotFoundError(filepath)
    if p.suffix.lower() not in EXCEL_EXTENSIONS:
        raise InvalidFormatError(filepath)
    return p


def validate_file_for_write(filepath: str) -> tuple[Path, bool]:
    """Validate a file path for write operations.

    Unlike validate_file(), allows non-existent files — they will be auto-created.
    Returns (resolved_path, is_new_file).
    """
    p = Path(filepath).resolve()
    if p.exists():
        if p.suffix.lower() not in EXCEL_EXTENSIONS:
            raise InvalidFormatError(filepath)
        return p, False
    # New file — must be a writable extension
    if p.suffix.lower() not in WRITABLE_EXTENSIONS:
        raise AgentExcelError(
            "INVALID_FORMAT",
            f"Cannot create '{filepath}' — only .xlsx and .xlsm files can be created",
            [f"Writable formats: {', '.join(sorted(WRITABLE_EXTENSIONS))}"],
        )
    return p, True


def file_size_bytes(filepath: str | Path) -> int:
    """Return file size in bytes."""
    return os.path.getsize(filepath)


def file_size_human(filepath: str | Path) -> str:
    """Return file size as a human-readable string (e.g. '107.7 KB', '76.2 MB')."""
    size = file_size_bytes(filepath)
    if size < 1024:
        return f"{size} B"
    elif size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    elif size < 1024 * 1024 * 1024:
        return f"{size / (1024 * 1024):.1f} MB"
    return f"{size / (1024 * 1024 * 1024):.1f} GB"


# Pattern: optional "SheetName!" prefix, then cell range like A1:C10 or just A1
_RANGE_RE = re.compile(
    r"^(?:(?P<sheet>.+?)!)?"
    r"(?P<start>[A-Z]{1,3}\d+)"
    r"(?::(?P<end>[A-Z]{1,3}\d+))?$",
    re.IGNORECASE,
)


def _normalise_shell_ref(ref: str) -> str:
    """Normalise shell-escaped cell/range references.

    Zsh escapes ``!`` to ``\\!`` even in single-quoted strings when passed
    through subprocess chains (e.g. ``uv run``).  This strips the escape so
    ``2022\\!B1`` becomes ``2022!B1``.
    """
    return ref.replace("\\!", "!")


def parse_range(range_str: str) -> dict[str, str | None]:
    """Parse an Excel range string like 'Sheet1!A1:C10'.

    Returns dict with keys: sheet, start, end (end may be None for single cell).
    """
    m = _RANGE_RE.match(_normalise_shell_ref(range_str).strip())
    if not m:
        raise RangeInvalidError(range_str)
    return {
        "sheet": m.group("sheet"),
        "start": m.group("start").upper(),
        "end": m.group("end").upper() if m.group("end") else None,
    }


def parse_multi_range(range_str: str) -> list[dict[str, str | None]]:
    """Parse comma-separated ranges like ``Sheet!A1:C10,E1:G10`` into a list.

    The sheet prefix from the first range carries forward to subsequent
    ranges that don't specify one, e.g.
    ``"2022!H54:AT54,H149:AT149"`` → both ranges on sheet ``"2022"``.
    """
    parts = range_str.split(",")
    results: list[dict[str, str | None]] = []
    sheet_ctx: str | None = None
    for part in parts:
        part = part.strip()
        if "!" in part:
            parsed = parse_range(part)
        elif sheet_ctx:
            parsed = parse_range(f"{sheet_ctx}!{part}")
        else:
            parsed = parse_range(part)
        if parsed["sheet"]:
            sheet_ctx = parsed["sheet"]
        results.append(parsed)
    return results


def col_letter_to_index(col: str) -> int:
    """Convert Excel column letter(s) to 0-based index. A=0, B=1, Z=25, AA=26."""
    result = 0
    for ch in col.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


def index_to_col_letter(index: int) -> str:
    """Convert 0-based index to Excel column letter(s). 0=A, 25=Z, 26=AA."""
    result = ""
    idx = index + 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def resolve_column_filter(
    columns_str: str,
    df_columns: list[str],
    headers: list[str] | None = None,
) -> list[str]:
    """Resolve comma-separated column specs to DataFrame column names.

    Accepts column letters (A, B, C) or header names ("Indicator Name").
    When *headers* is provided (row-1 header names), header names are
    resolved to their column letters before matching against df_columns.
    This enables header-name resolution even when df_columns are letters
    (e.g. range-scoped searches).
    """
    from agent_xlsx.utils.errors import InvalidColumnError

    requested = [c.strip() for c in columns_str.split(",") if c.strip()]
    resolved: list[str] = []
    invalid: list[str] = []

    # Map header names → column letters for fallback resolution
    header_map = {h: index_to_col_letter(i) for i, h in enumerate(headers)} if headers else {}

    for ref in requested:
        # Exact DataFrame column name match (header name or letter)
        if ref in df_columns:
            if ref not in resolved:
                resolved.append(ref)
            continue

        # Column letter match (e.g. "A", "BC", case-insensitive)
        upper_ref = ref.upper()
        if upper_ref.isalpha():
            idx = col_letter_to_index(upper_ref)
            if 0 <= idx < len(df_columns):
                name = df_columns[idx]
                if name not in resolved:
                    resolved.append(name)
                continue

        # Header name fallback: resolve name → column letter, check df_columns
        if ref in header_map:
            col_letter = header_map[ref]
            if col_letter in df_columns and col_letter not in resolved:
                resolved.append(col_letter)
                continue

        invalid.append(ref)

    if invalid:
        # Include header names in error message for discoverability
        avail = (
            list(df_columns) + [h for h in headers if h not in df_columns]
            if headers
            else list(df_columns)
        )
        raise InvalidColumnError(invalid, avail)

    return resolved


def resolve_column_letters(columns_str: str, headers: list[str] | None = None) -> set[str]:
    """Resolve column specs to uppercase column letters for openpyxl filtering.

    Accepts column letters (A, B, C) or header names (when headers provided).
    Returns a set of uppercase column letters.
    """
    from agent_xlsx.utils.errors import InvalidColumnError

    requested = [c.strip() for c in columns_str.split(",") if c.strip()]
    letters: set[str] = set()
    invalid: list[str] = []
    header_map = {h: index_to_col_letter(i) for i, h in enumerate(headers)} if headers else {}

    for ref in requested:
        # Header name lookup first (names can be purely alphabetic like "Formula")
        if ref in header_map:
            letters.add(header_map[ref])
            continue
        # Column letter (e.g. "A", "BC")
        upper_ref = ref.upper()
        if upper_ref.isalpha():
            letters.add(upper_ref)
            continue
        invalid.append(ref)

    if invalid:
        available = list(header_map.keys()) if headers else ["(column letters only)"]
        raise InvalidColumnError(invalid, available)

    return letters

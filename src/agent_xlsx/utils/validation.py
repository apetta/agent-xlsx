"""Shared validation helpers for file paths and ranges."""

from __future__ import annotations

import os
import re
from pathlib import Path

from agent_xlsx.utils.constants import EXCEL_EXTENSIONS
from agent_xlsx.utils.errors import ExcelFileNotFoundError, InvalidFormatError, RangeInvalidError


def validate_file(filepath: str) -> Path:
    """Validate that the file exists and has a supported extension. Returns resolved Path."""
    p = Path(filepath).resolve()
    if not p.exists():
        raise ExcelFileNotFoundError(filepath)
    if p.suffix.lower() not in EXCEL_EXTENSIONS:
        raise InvalidFormatError(filepath)
    return p


def file_size_bytes(filepath: str | Path) -> int:
    """Return file size in bytes."""
    return os.path.getsize(filepath)


# Pattern: optional "SheetName!" prefix, then cell range like A1:C10 or just A1
_RANGE_RE = re.compile(
    r"^(?:(?P<sheet>.+?)!)?"
    r"(?P<start>[A-Z]{1,3}\d+)"
    r"(?::(?P<end>[A-Z]{1,3}\d+))?$",
    re.IGNORECASE,
)


def parse_range(range_str: str) -> dict[str, str | None]:
    """Parse an Excel range string like 'Sheet1!A1:C10'.

    Returns dict with keys: sheet, start, end (end may be None for single cell).
    """
    m = _RANGE_RE.match(range_str.strip())
    if not m:
        raise RangeInvalidError(range_str)
    return {
        "sheet": m.group("sheet"),
        "start": m.group("start").upper(),
        "end": m.group("end").upper() if m.group("end") else None,
    }


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

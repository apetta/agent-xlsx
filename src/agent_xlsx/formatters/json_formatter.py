"""Token-efficient JSON output formatting."""

from __future__ import annotations

import json
import sys
from pathlib import Path as _Path
from typing import Any

# Module-level flag toggled by the global --no-meta CLI option
_suppress_meta: bool = False


def set_suppress_meta(value: bool) -> None:
    """Toggle metadata suppression (called by the --no-meta global callback)."""
    global _suppress_meta
    _suppress_meta = value


def should_include_meta() -> bool:
    """Whether to include per-call metadata (_data_origin, file_size_human)."""
    return not _suppress_meta


def output(data: dict[str, Any]) -> None:
    """Print a dict as compact JSON to stdout."""
    json.dump(data, sys.stdout, indent=2, default=_serialise)
    sys.stdout.write("\n")


def output_spreadsheet_data(data: dict[str, Any]) -> None:
    """Output spreadsheet-sourced content, automatically tagged as untrusted external data.

    Prepends ``_data_origin`` so any consuming LLM has per-call provenance context —
    a prompt-injection boundary that requires no skill-level instructions.
    When ``--no-meta`` is active, the tag is suppressed to reduce token waste
    on repeated calls against the same file.
    """
    if _suppress_meta:
        output(data)
    else:
        output({"_data_origin": "untrusted_spreadsheet", **data})


def relativize_path(result: dict[str, Any], key: str = "output_file") -> dict[str, Any]:
    """Convert an absolute path to a concise relative path for JSON output.

    In-place writes return just the filename (the agent already knows the target).
    Different-file writes return a relative path from cwd when possible.
    """
    if key in result and result[key]:
        p = _Path(result[key])
        if p.is_absolute():
            try:
                result[key] = str(p.relative_to(_Path.cwd()))
            except ValueError:
                result[key] = p.name
    return result


def _serialise(obj: Any) -> Any:
    """Handle non-standard types during JSON serialisation."""
    import datetime

    if isinstance(obj, datetime.datetime):
        # Normalise: midnight datetimes → date-only string for consistency
        if obj.hour == 0 and obj.minute == 0 and obj.second == 0 and obj.microsecond == 0:
            return obj.strftime("%Y-%m-%d")
        return obj.isoformat()
    if isinstance(obj, datetime.date):
        return obj.isoformat()
    if isinstance(obj, datetime.timedelta):
        return str(obj)
    if hasattr(obj, "__float__"):
        v = float(obj)
        # Return int if it's a whole number for cleaner output
        if v == int(v) and not (v != v):  # NaN check
            return int(v)
        return v
    return str(obj)

"""Token-efficient JSON output formatting."""

from __future__ import annotations

import json
import sys
from typing import Any


def output(data: dict[str, Any]) -> None:
    """Print a dict as compact JSON to stdout."""
    json.dump(data, sys.stdout, indent=2, default=_serialise)
    sys.stdout.write("\n")


def output_spreadsheet_data(data: dict[str, Any]) -> None:
    """Output spreadsheet-sourced content, automatically tagged as untrusted external data.

    Prepends ``_data_origin`` so any consuming LLM has per-call provenance context —
    a prompt-injection boundary that requires no skill-level instructions.
    """
    output({"_data_origin": "untrusted_spreadsheet", **data})


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

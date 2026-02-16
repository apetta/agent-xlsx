"""Token-efficient JSON output formatting."""

from __future__ import annotations

import json
import sys
from typing import Any


def output(data: dict[str, Any]) -> None:
    """Print a dict as compact JSON to stdout."""
    json.dump(data, sys.stdout, indent=2, default=_serialise)
    sys.stdout.write("\n")


def _serialise(obj: Any) -> Any:
    """Handle non-standard types during JSON serialisation."""
    import datetime

    if isinstance(obj, (datetime.date, datetime.datetime)):
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

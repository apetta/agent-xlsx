"""oletools adapter for VBA extraction and security analysis."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from agent_xlsx.utils.constants import MAX_VBA_LINES

# ---------------------------------------------------------------------------
# VBA Detection
# ---------------------------------------------------------------------------


def has_vba(filepath: str | Path) -> bool:
    """Quick check whether the file contains VBA macros."""
    from oletools.olevba import VBA_Parser

    filepath = str(Path(filepath).resolve())
    try:
        parser = VBA_Parser(filepath)
        result = parser.detect_vba_macros()
        parser.close()
        return bool(result)
    except Exception:
        return False


# ---------------------------------------------------------------------------
# VBA Module Extraction
# ---------------------------------------------------------------------------

# Map oletools stream-path keywords to human-friendly module types
_MODULE_TYPE_MAP = {
    "ThisWorkbook": "document",
    "Sheet": "document",
    "Module": "standard",
    "Class": "class",
    "UserForm": "form",
}


def _classify_module(vba_filename: str, stream_path: str) -> str:
    """Infer module type from the stream path and VBA filename."""
    for key, value in _MODULE_TYPE_MAP.items():
        if key.lower() in vba_filename.lower() or key.lower() in stream_path.lower():
            return value
    return "standard"


def extract_vba_modules(filepath: str | Path) -> list[dict[str, Any]]:
    """List all VBA modules with metadata (name, type, line count)."""
    from oletools.olevba import VBA_Parser

    filepath = str(Path(filepath).resolve())
    parser = VBA_Parser(filepath)

    modules: list[dict[str, Any]] = []
    if parser.detect_vba_macros():
        for _filename, stream_path, vba_filename, vba_code in parser.extract_macros():
            lines = vba_code.count("\n") + 1 if vba_code else 0
            modules.append(
                {
                    "name": vba_filename,
                    "type": _classify_module(vba_filename, stream_path),
                    "lines": lines,
                }
            )

    parser.close()
    return modules


# ---------------------------------------------------------------------------
# VBA Code Reading
# ---------------------------------------------------------------------------


def read_vba_code(
    filepath: str | Path,
    module_name: str | None = None,
) -> list[dict[str, Any]]:
    """Read VBA source code.

    If *module_name* is given, return only that module's code.
    Otherwise return all modules (lines capped per MAX_VBA_LINES).
    """
    from oletools.olevba import VBA_Parser

    filepath = str(Path(filepath).resolve())
    parser = VBA_Parser(filepath)

    results: list[dict[str, Any]] = []
    if parser.detect_vba_macros():
        for _filename, stream_path, vba_filename, vba_code in parser.extract_macros():
            if module_name and vba_filename != module_name:
                continue

            code = vba_code or ""
            lines = code.split("\n")
            truncated = len(lines) > MAX_VBA_LINES
            if truncated:
                lines = lines[:MAX_VBA_LINES]
                code = "\n".join(lines)

            # Extract Sub/Function names for quick reference
            procedures: list[str] = []
            for line in (vba_code or "").split("\n"):
                stripped = line.strip()
                for kw in ("Sub ", "Function "):
                    if any(
                        (
                            stripped.startswith(kw),
                            stripped.startswith(f"Public {kw}"),
                            stripped.startswith(f"Private {kw}"),
                        )
                    ):
                        # Extract procedure name (before the opening paren)
                        name_part = stripped.split("(")[0]
                        proc_name = name_part.split()[-1]
                        procedures.append(proc_name)

            results.append(
                {
                    "module": vba_filename,
                    "type": _classify_module(vba_filename, stream_path),
                    "code": code,
                    "line_count": len((vba_code or "").split("\n")),
                    "truncated": truncated,
                    "procedures": procedures,
                }
            )

            if module_name:
                break

    parser.close()
    return results


# ---------------------------------------------------------------------------
# VBA Security Analysis
# ---------------------------------------------------------------------------


def analyse_vba_security(filepath: str | Path) -> dict[str, Any]:
    """Analyse VBA code for security concerns.

    Returns a report with auto_execute triggers, suspicious patterns,
    indicators of compromise (IOCs), and an overall risk_level.
    """
    from oletools.olevba import VBA_Parser

    filepath = str(Path(filepath).resolve())
    parser = VBA_Parser(filepath)

    report: dict[str, Any] = {
        "auto_execute": [],
        "suspicious": [],
        "iocs": [],
        "risk_level": "low",
    }

    if not parser.detect_vba_macros():
        report["has_vba"] = False
        parser.close()
        return report

    report["has_vba"] = True

    results = parser.analyze_macros()
    has_suspicious = False
    has_auto_exec = False

    for kw_type, keyword, description in results:
        if kw_type == "AutoExec":
            report["auto_execute"].append(keyword)
            has_auto_exec = True
        elif kw_type == "Suspicious":
            report["suspicious"].append({"keyword": keyword, "description": description})
            has_suspicious = True
        elif kw_type == "IOC":
            report["iocs"].append({"keyword": keyword, "description": description})

    # Determine risk level
    if has_suspicious:
        report["risk_level"] = "high"
    elif has_auto_exec:
        report["risk_level"] = "medium"
    else:
        report["risk_level"] = "low"

    parser.close()
    return report

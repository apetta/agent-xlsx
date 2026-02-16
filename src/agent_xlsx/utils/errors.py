"""Consistent error types and JSON error formatting."""

from __future__ import annotations

import functools
import json
import sys
from typing import Any


class AgentExcelError(Exception):
    """Base error with structured JSON output."""

    def __init__(self, code: str, message: str, suggestions: list[str] | None = None):
        self.code = code
        self.message = message
        self.suggestions = suggestions or []
        super().__init__(message)

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "error": True,
            "code": self.code,
            "message": self.message,
        }
        if self.suggestions:
            result["suggestions"] = self.suggestions
        return result


class ExcelFileNotFoundError(AgentExcelError):
    def __init__(self, path: str):
        super().__init__(
            "FILE_NOT_FOUND",
            f"The file '{path}' does not exist",
            ["Check the file path is correct", "Ensure the file has a supported extension"],
        )


class InvalidFormatError(AgentExcelError):
    def __init__(self, path: str):
        super().__init__(
            "INVALID_FORMAT",
            f"'{path}' is not a supported Excel file",
            ["Supported formats: .xlsx, .xlsm, .xlsb, .xls, .ods"],
        )


class SheetNotFoundError(AgentExcelError):
    def __init__(self, name: str, available: list[str]):
        super().__init__(
            "SHEET_NOT_FOUND",
            f"Sheet '{name}' not found",
            [f"Available sheets: {', '.join(available)}"],
        )


class RangeInvalidError(AgentExcelError):
    def __init__(self, range_str: str):
        super().__init__(
            "RANGE_INVALID",
            f"Invalid range: '{range_str}'",
            ["Use Excel notation e.g. 'A1:C10' or 'Sheet1!A1:C10'"],
        )


class ExcelRequiredError(AgentExcelError):
    def __init__(self, operation: str):
        super().__init__(
            "EXCEL_REQUIRED",
            f"'{operation}' requires Microsoft Excel",
            [
                "Install Microsoft Excel for this feature",
                "Use 'agent-xlsx probe' for fast text-based profiling",
                "Use 'agent-xlsx read' for data extraction",
            ],
        )


class LibreOfficeNotFoundError(AgentExcelError):
    def __init__(self):
        super().__init__(
            "LIBREOFFICE_REQUIRED",
            "LibreOffice is required for this operation but was not found",
            [
                "Install LibreOffice: apt install libreoffice-calc",
                "Or on macOS: brew install --cask libreoffice",
                "Ensure 'soffice' or 'libreoffice' is on PATH",
            ],
        )


class NoRenderingBackendError(AgentExcelError):
    def __init__(self, operation: str):
        super().__init__(
            "NO_RENDERING_BACKEND",
            f"'{operation}' requires Microsoft Excel or LibreOffice, but neither was found",
            [
                "Install Microsoft Excel (macOS/Windows) for best fidelity",
                "Or install aspose-cells-python (all platforms):"
                " uv pip install aspose-cells-python",
                "Or install LibreOffice (all platforms): apt install libreoffice-calc",
                "Use 'agent-xlsx read' for text-based data extraction instead",
            ],
        )


class AsposeNotInstalledError(AgentExcelError):
    def __init__(self):
        super().__init__(
            "ASPOSE_NOT_INSTALLED",
            "Aspose.Cells for Python is not installed",
            [
                "Install with: uv pip install aspose-cells-python",
                "Use --engine excel or --engine libreoffice instead",
                "See: https://pypi.org/project/aspose-cells-python/",
            ],
        )


class MemoryExceededError(AgentExcelError):
    def __init__(self, used_mb: float, limit_mb: float):
        super().__init__(
            "MEMORY_EXCEEDED",
            f"Memory usage {used_mb:.0f}MB exceeds limit {limit_mb:.0f}MB",
            ["Use --limit to read fewer rows", "Use 'probe' for a lightweight summary"],
        )


def handle_error(func):
    """Decorator that catches AgentExcelError and prints JSON to stdout."""

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except AgentExcelError as e:
            json.dump(e.to_dict(), sys.stdout, indent=2)
            sys.stdout.write("\n")
            raise SystemExit(1)

    return wrapper

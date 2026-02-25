"""Aspose.Cells adapter for high-fidelity rendering without a live Excel instance.

Provides screenshots (PNG), formula recalculation, object inspection, and
chart export using the Aspose.Cells for Python library. Works headlessly on any
platform — no Excel or LibreOffice installation required.

Requires: aspose-cells-python (optional dependency).
"""

from __future__ import annotations

import base64
import io
import time
from pathlib import Path
from typing import Any

from agent_xlsx.utils.errors import AgentExcelError, SheetNotFoundError
from agent_xlsx.utils.validation import index_to_col_letter

# ---------------------------------------------------------------------------
# Licence management (module-level, applied once per process)
# ---------------------------------------------------------------------------

_LICENSE_APPLIED = False
_LICENSE_DATA_WARNED = False  # tracks whether the ASPOSE_LICENSE_DATA warning has fired


_ASPOSE_AVAILABLE: bool | None = None


def is_aspose_available() -> bool:
    """Check whether the aspose-cells-python package is importable.

    Uses a two-stage check: fast importlib.util.find_spec first, then
    a subprocess probe to detect CoreCLR sandbox crashes that would
    otherwise kill the parent process. Result is cached for the process
    lifetime.
    """
    global _ASPOSE_AVAILABLE
    if _ASPOSE_AVAILABLE is not None:
        return _ASPOSE_AVAILABLE

    # Fast path: check if package is even installed
    try:
        import importlib.util

        if importlib.util.find_spec("aspose.cells") is None:
            _ASPOSE_AVAILABLE = False
            return False
    except (ImportError, ModuleNotFoundError, ValueError):
        _ASPOSE_AVAILABLE = False
        return False

    # Slow path: import in subprocess to detect CoreCLR sandbox crashes
    import subprocess
    import sys

    try:
        result = subprocess.run(
            [sys.executable, "-c", "import aspose.cells"],
            capture_output=True,
            timeout=10,
        )
        _ASPOSE_AVAILABLE = result.returncode == 0
    except (subprocess.TimeoutExpired, OSError):
        _ASPOSE_AVAILABLE = False
    return _ASPOSE_AVAILABLE


def _apply_license() -> bool:
    """Apply the Aspose licence if configured. Returns True when licensed."""
    global _LICENSE_APPLIED, _LICENSE_DATA_WARNED
    if _LICENSE_APPLIED:
        return True

    from agent_xlsx.utils.config import get_aspose_license_path

    lic_path = get_aspose_license_path()
    if lic_path is None:
        return False

    # Warn on *presence* of ASPOSE_LICENSE_DATA regardless of which source was selected.
    # The env var is visible in ps aux even when ASPOSE_LICENSE_PATH takes precedence.
    # Goes to stderr to keep stdout clean. Fires once per process.
    import os as _os

    if _os.environ.get("ASPOSE_LICENSE_DATA") and not _LICENSE_DATA_WARNED:
        _LICENSE_DATA_WARNED = True
        import sys as _sys

        _sys.stderr.write(
            "[agent-xlsx] Warning: ASPOSE_LICENSE_DATA is present in the environment "
            "and visible in process listings (ps aux). "
            "Consider removing it or switching to ASPOSE_LICENSE_PATH exclusively.\n"
        )

    try:
        from aspose.cells import License

        lic = License()
        if lic_path.startswith("base64:"):
            raw = base64.b64decode(lic_path[7:])
            lic.set_license(io.BytesIO(raw))
        else:
            lic.set_license(lic_path)
        _LICENSE_APPLIED = True
        return True
    except Exception:
        return False


def get_license_status() -> dict[str, Any]:
    """Return current Aspose installation and licence status."""
    installed = is_aspose_available()
    if not installed:
        return {"installed": False, "licensed": False, "evaluation_mode": False}
    licensed = _apply_license()
    return {
        "installed": True,
        "licensed": licensed,
        "evaluation_mode": not licensed,
    }


def _eval_fields(licensed: bool) -> dict[str, Any]:
    """Return evaluation-mode notice fields when unlicensed."""
    if licensed:
        return {}
    return {
        "evaluation_mode": True,
        "evaluation_notice": (
            "Aspose evaluation mode: rendered images contain a watermark. "
            "Set a licence via 'agent-xlsx license --set <path>' or "
            "ASPOSE_LICENSE_PATH env var."
        ),
    }


def _get_worksheet(wb, name: str):
    """Resolve a worksheet by name, raising SheetNotFoundError if missing."""
    ws = wb.worksheets.get(name)
    if ws is None:
        available = [wb.worksheets[i].name for i in range(len(wb.worksheets))]
        raise SheetNotFoundError(name, available)
    return ws


# ---------------------------------------------------------------------------
# Screenshots
# ---------------------------------------------------------------------------


def screenshot(
    filepath: str | Path,
    sheet_name: str | None = None,
    range_str: str | None = None,
    output_path: str | Path | None = None,
    dpi: int = 200,
) -> dict[str, Any]:
    """Export workbook sheet(s) or a range to PNG via Aspose.Cells."""
    import aspose.cells as ac
    from aspose.cells.drawing import ImageType
    from aspose.cells.rendering import ImageOrPrintOptions, SheetRender

    filepath = Path(filepath).resolve()
    start = time.perf_counter()
    licensed = _apply_license()

    wb = ac.Workbook(str(filepath))

    # Determine target sheets
    if sheet_name:
        target_sheets = [s.strip() for s in sheet_name.split(",")]
    else:
        target_sheets = [wb.worksheets[i].name for i in range(len(wb.worksheets))]  # ty: ignore[invalid-argument-type]  # aspose stub missing __len__

    # Output directory
    if output_path is None:
        output_dir = Path("/tmp/agent-xlsx")
    else:
        output_dir = Path(output_path)
        if output_dir.suffix:
            output_dir = output_dir.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    stem = filepath.stem
    sheets_result: list[dict[str, Any]] = []

    for target in target_sheets:
        ws = _get_worksheet(wb, target)

        ws.auto_fit_columns()
        ws.auto_fit_rows()

        # Compute data extent
        max_data_row = ws.cells.max_data_row  # 0-based
        max_data_col = ws.cells.max_data_column  # 0-based
        max_col_letter = index_to_col_letter(max_data_col)
        max_row_num = max_data_row + 1  # 1-based for Excel notation

        # Set print area — CRITICAL to avoid "Unable to allocate pixels"
        if range_str:
            ws.page_setup.print_area = range_str
        else:
            ws.page_setup.print_area = f"A1:{max_col_letter}{max_row_num}"

        # Configure rendering options
        opts = ImageOrPrintOptions()
        opts.image_type = ImageType.PNG
        opts.horizontal_resolution = dpi
        opts.vertical_resolution = dpi
        opts.one_page_per_sheet = True
        opts.set_desired_size(4000, 4000, True)

        sr = SheetRender(ws, opts)

        # Build output filename
        safe_name = target.replace("/", "_").replace("\\", "_").replace(" ", "_")
        if range_str:
            safe_range = range_str.replace(":", "-").replace("$", "")
            png_path = output_dir / f"{stem}_{safe_name}_{safe_range}.png"
        else:
            png_path = output_dir / f"{stem}_{safe_name}.png"

        sr.to_image(0, str(png_path))

        # Get image dimensions via PIL
        from PIL import Image

        img = Image.open(str(png_path))
        width, height = img.size
        img.close()

        resolved_range = range_str if range_str else f"A1:{max_col_letter}{max_row_num}"

        sheets_result.append(
            {
                "name": target,
                "path": str(png_path),
                "resolved_range": resolved_range,
                "size_bytes": png_path.stat().st_size,
                "width": width,
                "height": height,
            }
        )

    elapsed_ms = round((time.perf_counter() - start) * 1000, 1)

    if len(sheets_result) == 1:
        result = {
            "status": "success",
            "format": "png",
            "path": sheets_result[0]["path"],
            "sheet": sheets_result[0]["name"],
            "resolved_range": sheets_result[0]["resolved_range"],
            "size_bytes": sheets_result[0]["size_bytes"],
            "width": sheets_result[0]["width"],
            "height": sheets_result[0]["height"],
            "capture_time_ms": elapsed_ms,
            "engine": "aspose",
            **_eval_fields(licensed),
        }
        if range_str:
            result["range"] = range_str
        return result
    else:
        return {
            "status": "success",
            "format": "png",
            "sheets": sheets_result,
            "capture_time_ms": elapsed_ms,
            "engine": "aspose",
            **_eval_fields(licensed),
        }


# ---------------------------------------------------------------------------
# Recalculation
# ---------------------------------------------------------------------------


def recalculate(
    filepath: str | Path,
    output_path: str | Path | None = None,
) -> dict[str, Any]:
    """Trigger a full recalculation of all formulae via Aspose.Cells and save."""
    import aspose.cells as ac

    filepath = Path(filepath).resolve()
    start = time.perf_counter()
    licensed = _apply_license()

    wb = ac.Workbook(str(filepath))
    wb.calculate_formula()

    target = Path(output_path) if output_path else filepath
    wb.save(str(target))

    elapsed_ms = round((time.perf_counter() - start) * 1000, 1)
    return {
        "status": "success",
        "engine": "aspose",
        "output_file": str(target),
        "recalc_time_ms": elapsed_ms,
        **_eval_fields(licensed),
    }


# ---------------------------------------------------------------------------
# Object inspection (charts, shapes, pictures)
# ---------------------------------------------------------------------------


def get_objects(
    filepath: str | Path,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """List charts, shapes, and pictures in the workbook."""
    import aspose.cells as ac

    filepath = Path(filepath).resolve()
    wb = ac.Workbook(str(filepath))

    # Determine target sheets
    if sheet_name:
        target_names = [s.strip() for s in sheet_name.split(",")]
    else:
        target_names = [wb.worksheets[i].name for i in range(len(wb.worksheets))]  # ty: ignore[invalid-argument-type]  # aspose stub missing __len__

    sheets_info: list[dict[str, Any]] = []

    for name in target_names:
        ws = _get_worksheet(wb, name)

        # Charts
        charts: list[dict[str, Any]] = []
        for i in range(len(ws.charts)):
            chart = ws.charts[i]
            info: dict[str, Any] = {
                "name": chart.name,
                "position": {
                    "left": chart.left,
                    "top": chart.top,
                    "width": chart.width,
                    "height": chart.height,
                },
            }
            try:
                info["chart_type"] = str(chart.type)
            except Exception:
                pass
            charts.append(info)

        # Shapes
        shapes: list[dict[str, Any]] = []
        for shape in ws.shapes:
            shapes.append(
                {
                    "name": shape.name,
                    "type": str(shape.type),
                    "position": {
                        "left": shape.left,
                        "top": shape.top,
                        "width": shape.width,
                        "height": shape.height,
                    },
                }
            )

        # Pictures
        pictures: list[dict[str, Any]] = []
        for pic in ws.pictures:
            pic_info: dict[str, Any] = {
                "name": pic.name,
                "position": {
                    "left": pic.left,
                    "top": pic.top,
                    "width": pic.width,
                    "height": pic.height,
                },
            }
            pictures.append(pic_info)

        sheets_info.append(
            {
                "name": name,
                "charts": charts,
                "shapes": shapes,
                "pictures": pictures,
            }
        )

    return {
        "status": "success",
        "file": str(filepath),
        "sheets": sheets_info,
    }


# ---------------------------------------------------------------------------
# Chart export
# ---------------------------------------------------------------------------


def export_chart(
    filepath: str | Path,
    chart_name: str,
    sheet_name: str | None = None,
    output_path: str | Path | None = None,
) -> dict[str, Any]:
    """Export a named chart to PNG."""
    import aspose.cells as ac

    filepath = Path(filepath).resolve()

    if output_path is None:
        output_dir = Path("/tmp/agent-xlsx")
    else:
        output_dir = Path(output_path)
        if output_dir.suffix:
            output_dir = output_dir.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    wb = ac.Workbook(str(filepath))

    # Determine search scope
    if sheet_name:
        target_names = [s.strip() for s in sheet_name.split(",")]
    else:
        target_names = [wb.worksheets[i].name for i in range(len(wb.worksheets))]  # ty: ignore[invalid-argument-type]  # aspose stub missing __len__

    # Search for chart by name
    target_chart = None
    for name in target_names:
        ws = _get_worksheet(wb, name)
        for i in range(len(ws.charts)):
            chart = ws.charts[i]
            if chart.name == chart_name:
                target_chart = chart
                break
        if target_chart:
            break

    if target_chart is None:
        raise AgentExcelError(
            "CHART_NOT_FOUND",
            f"Chart '{chart_name}' was not found in the workbook",
            ["Use 'agent-xlsx objects' to list available charts"],
        )

    safe_name = chart_name.replace("/", "_").replace("\\", "_").replace("..", "_")
    out_path = output_dir / f"{safe_name}.png"

    target_chart.to_image(str(out_path))

    return {
        "status": "success",
        "chart": chart_name,
        "path": str(out_path),
        "format": "png",
        "size_bytes": out_path.stat().st_size,
    }

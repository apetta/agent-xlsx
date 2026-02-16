"""xlwings adapter for operations requiring a live Microsoft Excel instance.

Provides high-fidelity screenshots, formula recalculation, VBA macro execution,
and object inspection by driving Excel through the xlwings COM/AppleScript bridge.

Requires Microsoft Excel to be installed on the host machine (macOS or Windows).
"""

from __future__ import annotations

import contextlib
import time
from pathlib import Path
from typing import Any

from agent_xlsx.utils.errors import AgentExcelError, ExcelRequiredError

# ---------------------------------------------------------------------------
# Excel session management
# ---------------------------------------------------------------------------


@contextlib.contextmanager
def _excel_session(filepath: str | Path | None = None, visible: bool = False):
    """Context manager: open Excel app + workbook, yield (app, wb), clean up.

    Starts a headless Excel instance with alerts and screen updating disabled
    for maximum performance. If *filepath* is provided the workbook is opened
    and yielded alongside the app; otherwise *wb* is ``None``.
    """
    import xlwings as xw

    app = xw.App(visible=visible)
    app.display_alerts = False
    app.screen_updating = False
    wb = None
    try:
        if filepath:
            wb = app.books.open(str(Path(filepath).resolve()))
        yield app, wb
    finally:
        if wb:
            try:
                wb.close()
            except Exception:
                pass
        app.quit()


# ---------------------------------------------------------------------------
# Discovery
# ---------------------------------------------------------------------------


def is_excel_available() -> bool:
    """Check whether xlwings can connect to a live Excel instance."""
    try:
        import xlwings as xw

        app = xw.App(visible=False)
        app.quit()
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Screenshots
# ---------------------------------------------------------------------------


def screenshot(
    filepath: str | Path,
    sheet_name: str | None = None,
    range_str: str | None = None,
    output_path: str | Path | None = None,
) -> dict[str, Any]:
    """Export workbook sheet(s) or a range to PNG via xlwings.

    Renders each target sheet's used range (or a specific range) to a PNG image.
    """
    filepath = Path(filepath).resolve()
    start = time.perf_counter()

    if output_path is None:
        output_dir = Path("/tmp/agent-xlsx")
    else:
        output_dir = Path(output_path)
        if output_dir.suffix:
            output_dir = output_dir.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    stem = filepath.stem

    try:
        import sys
        need_visible = sys.platform == "darwin"

        with _excel_session(filepath, visible=need_visible) as (app, wb):
            # CopyPicture requires Excel to have rendered cells
            app.screen_updating = True
            if need_visible:
                time.sleep(0.5)  # Allow Excel to paint on macOS

            # -----------------------------------------------------------
            # PNG path — per-sheet rendering
            # -----------------------------------------------------------
            if sheet_name:
                target_sheets = [s.strip() for s in sheet_name.split(",")]
            else:
                target_sheets = [s.name for s in wb.sheets]

            sheets_result: list[dict[str, Any]] = []

            for target in target_sheets:
                sheet = wb.sheets[target]
                sheet.autofit('c')  # auto-fit column widths to prevent ######## display

                if range_str:
                    capture_range = sheet.range(range_str)
                else:
                    # current_region gives the contiguous data block from A1,
                    # bounded by the first empty row/column — far tighter than
                    # used_range which includes every cell Excel ever touched.
                    capture_range = sheet.range("A1").current_region

                # Capture resolved range address for the response
                resolved_range = capture_range.address.replace("$", "")

                safe_name = (
                    target.replace("/", "_").replace("\\", "_").replace(" ", "_")
                )
                if range_str:
                    safe_range = range_str.replace(":", "-").replace("$", "")
                    png_path = output_dir / f"{stem}_{safe_name}_{safe_range}.png"
                else:
                    png_path = output_dir / f"{stem}_{safe_name}.png"
                capture_range.to_png(str(png_path))

                from PIL import Image

                img = Image.open(str(png_path))
                width, height = img.size
                img.close()

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
                result: dict[str, Any] = {
                    "status": "success",
                    "format": "png",
                    "path": sheets_result[0]["path"],
                    "sheet": sheets_result[0]["name"],
                    "resolved_range": sheets_result[0]["resolved_range"],
                    "size_bytes": sheets_result[0]["size_bytes"],
                    "width": sheets_result[0]["width"],
                    "height": sheets_result[0]["height"],
                    "capture_time_ms": elapsed_ms,
                    "engine": "xlwings",
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
                    "engine": "xlwings",
                }

    except ExcelRequiredError:
        raise
    except AgentExcelError:
        raise
    except Exception as exc:
        # xlwings connection failures (e.g. Excel not installed) surface here
        raise ExcelRequiredError("screenshot") from exc


# ---------------------------------------------------------------------------
# Recalculation
# ---------------------------------------------------------------------------


def recalculate(
    filepath: str | Path,
    output_path: str | Path | None = None,
) -> dict[str, Any]:
    """Trigger a full recalculation of all formulas via Excel and save."""
    filepath = Path(filepath).resolve()
    start = time.perf_counter()

    try:
        with _excel_session(filepath) as (app, wb):
            app.calculate()

            target = Path(output_path) if output_path else filepath
            wb.save(str(target))

            elapsed_ms = round((time.perf_counter() - start) * 1000, 1)
            return {
                "status": "success",
                "engine": "xlwings",
                "output_file": str(target),
                "recalc_time_ms": elapsed_ms,
            }

    except ExcelRequiredError:
        raise
    except AgentExcelError:
        raise
    except Exception as exc:
        raise ExcelRequiredError("recalculate") from exc


# ---------------------------------------------------------------------------
# VBA macro execution
# ---------------------------------------------------------------------------


def run_macro(
    filepath: str | Path,
    macro_name: str,
    args: list[Any] | None = None,
    save: bool = False,
) -> dict[str, Any]:
    """Execute a VBA macro within an open workbook via xlwings."""
    filepath = Path(filepath).resolve()
    start = time.perf_counter()

    try:
        with _excel_session(filepath) as (_app, wb):
            result = wb.macro(macro_name)(*(args or []))

            if save:
                wb.save()

            elapsed_ms = round((time.perf_counter() - start) * 1000, 1)
            return {
                "status": "success",
                "macro": macro_name,
                "return_value": result,
                "saved": save,
                "execution_time_ms": elapsed_ms,
            }

    except ExcelRequiredError:
        raise
    except AgentExcelError:
        raise
    except Exception as exc:
        raise ExcelRequiredError("run_macro") from exc


# ---------------------------------------------------------------------------
# Object inspection (charts, shapes, pictures)
# ---------------------------------------------------------------------------


def get_objects(
    filepath: str | Path,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """List charts, shapes, and pictures in the workbook."""
    filepath = Path(filepath).resolve()

    try:
        with _excel_session(filepath) as (_app, wb):
            if sheet_name:
                target_sheets = [wb.sheets[s.strip()] for s in sheet_name.split(",")]
            else:
                target_sheets = list(wb.sheets)

            sheets_info: list[dict[str, Any]] = []

            for sheet in target_sheets:
                # Charts
                charts: list[dict[str, Any]] = []
                for chart in sheet.charts:
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
                        info["chart_type"] = str(chart.chart_type)
                    except Exception:
                        pass
                    charts.append(info)

                # Shapes
                shapes: list[dict[str, Any]] = []
                for shape in sheet.shapes:
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
                for pic in sheet.pictures:
                    pic_info: dict[str, Any] = {
                        "name": pic.name,
                        "position": {
                            "left": pic.left,
                            "top": pic.top,
                            "width": pic.width,
                            "height": pic.height,
                        },
                    }
                    try:
                        pic_info["filename"] = pic.filename
                    except Exception:
                        pass
                    pictures.append(pic_info)

                sheets_info.append(
                    {
                        "name": sheet.name,
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

    except ExcelRequiredError:
        raise
    except AgentExcelError:
        raise
    except Exception as exc:
        raise ExcelRequiredError("get_objects") from exc


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
    filepath = Path(filepath).resolve()

    if output_path is None:
        output_dir = Path("/tmp/agent-xlsx")
    else:
        output_dir = Path(output_path)
        if output_dir.suffix:
            output_dir = output_dir.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        with _excel_session(filepath) as (_app, wb):
            if sheet_name:
                search_sheets = [wb.sheets[s.strip()] for s in sheet_name.split(",")]
            else:
                search_sheets = list(wb.sheets)

            target_chart = None
            for sheet in search_sheets:
                for chart in sheet.charts:
                    if chart.name == chart_name:
                        target_chart = chart
                        break
                if target_chart:
                    break

            if target_chart is None:
                raise AgentExcelError(
                    "CHART_NOT_FOUND",
                    f"Chart '{chart_name}' was not found in the workbook",
                    ["Use 'get_objects' to list available charts"],
                )

            out_path = output_dir / f"{chart_name}.png"
            target_chart.to_png(str(out_path))

            return {
                "status": "success",
                "chart": chart_name,
                "path": str(out_path),
                "format": "png",
                "size_bytes": out_path.stat().st_size,
            }

    except ExcelRequiredError:
        raise
    except AgentExcelError:
        raise
    except Exception as exc:
        raise ExcelRequiredError("export_chart") from exc

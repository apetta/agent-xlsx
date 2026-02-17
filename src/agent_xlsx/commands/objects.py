"""Embedded object inspection and export — charts, shapes, and pictures.

Uses xlwings to inspect and export embedded objects from Excel workbooks.
Requires Microsoft Excel to be installed.
"""

from typing import Optional

import typer

from agent_xlsx.cli import app
from agent_xlsx.formatters.json_formatter import output
from agent_xlsx.utils.errors import handle_error
from agent_xlsx.utils.validation import validate_file


@app.command()
@handle_error
def objects(
    file: str = typer.Argument(..., help="Path to the Excel file"),
    sheet: Optional[str] = typer.Option(None, "--sheet", "-s", help="Sheet name to inspect"),
    export: Optional[str] = typer.Option(None, "--export", "-e", help="Export a chart by name"),
    output_path: Optional[str] = typer.Option(
        None, "--output", "-o", help="Output file/directory path"
    ),
    engine: str = typer.Option(
        "auto",
        "--engine",
        help="Backend: auto (default), excel, or aspose",
    ),
) -> None:
    """List or export embedded objects (charts, shapes, pictures).

    Uses Excel (xlwings) when available, falls back to Aspose.Cells.
    Use without --export to list all objects, or with --export to save
    a specific chart as PNG.
    """
    path = validate_file(file)
    filepath = str(path)

    # Engine selection
    from agent_xlsx.adapters.xlwings_adapter import is_excel_available

    use_engine = None
    engine_lower = engine.lower()

    if engine_lower == "excel":
        if not is_excel_available():
            from agent_xlsx.utils.errors import ExcelRequiredError

            raise ExcelRequiredError("objects")
        use_engine = "excel"
    elif engine_lower == "aspose":
        from agent_xlsx.adapters.aspose_adapter import is_aspose_available
        from agent_xlsx.utils.errors import AsposeNotInstalledError

        if not is_aspose_available():
            raise AsposeNotInstalledError()
        use_engine = "aspose"
    elif engine_lower == "auto":
        if is_excel_available():
            use_engine = "excel"
        else:
            from agent_xlsx.adapters.aspose_adapter import is_aspose_available

            if is_aspose_available():
                use_engine = "aspose"
            else:
                from agent_xlsx.utils.errors import ExcelRequiredError

                raise ExcelRequiredError("objects")
    else:
        from agent_xlsx.utils.errors import ExcelRequiredError

        raise ExcelRequiredError("objects")

    if export:
        if use_engine == "aspose":
            from agent_xlsx.adapters.aspose_adapter import export_chart

            result = export_chart(
                filepath=filepath,
                chart_name=export,
                sheet_name=sheet,
                output_path=output_path,
            )
        else:
            from agent_xlsx.adapters.xlwings_adapter import export_chart

            result = export_chart(
                filepath=filepath,
                chart_name=export,
                sheet_name=sheet,
                output_path=output_path,
            )
    else:
        if use_engine == "aspose":
            from agent_xlsx.adapters.aspose_adapter import get_objects

            result = get_objects(filepath=filepath, sheet_name=sheet)
        else:
            from agent_xlsx.adapters.xlwings_adapter import get_objects

            result = get_objects(filepath=filepath, sheet_name=sheet)

    output(result)

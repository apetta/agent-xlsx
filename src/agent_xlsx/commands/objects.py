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

    Uses Aspose.Cells when available, falls back to Excel (xlwings).
    Use without --export to list all objects, or with --export to save
    a specific chart as PNG.
    """
    path = validate_file(file)
    filepath = str(path)

    # Engine selection
    from agent_xlsx.utils.engine import resolve_engine

    # libreoffice=False: no LibreOffice object-extraction adapter exists.
    use_engine = resolve_engine("objects", engine, libreoffice=False)

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

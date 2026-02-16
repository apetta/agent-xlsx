"""Sheet management operations: list, create, rename, delete, copy, hide/unhide."""

from typing import Optional

import typer

from agent_xlsx.cli import app
from agent_xlsx.formatters import json_formatter
from agent_xlsx.utils.errors import AgentExcelError, handle_error
from agent_xlsx.utils.validation import validate_file


@app.command("sheet")
@handle_error
def sheet_cmd(
    file: str = typer.Argument(..., help="Path to the Excel file"),
    list_: bool = typer.Option(False, "--list", "-l", help="List all sheets"),
    create: Optional[str] = typer.Option(
        None,
        "--create",
        help="Create a new sheet with the given name",
    ),
    rename: Optional[str] = typer.Option(
        None,
        "--rename",
        help="Rename a sheet (use with positional new name)",
    ),
    new_name: Optional[str] = typer.Option(
        None,
        "--new-name",
        help="New name for rename/copy operations",
    ),
    delete: Optional[str] = typer.Option(None, "--delete", help="Delete a sheet"),
    copy: Optional[str] = typer.Option(
        None,
        "--copy",
        help="Copy a sheet (use --new-name for the copy name)",
    ),
    hide: Optional[str] = typer.Option(None, "--hide", help="Hide a sheet"),
    unhide: Optional[str] = typer.Option(None, "--unhide", help="Unhide a sheet"),
    output: Optional[str] = typer.Option(
        None,
        "--output",
        "-o",
        help="Save to a new file instead of overwriting",
    ),
) -> None:
    """Manage sheets: list, create, rename, delete, copy, hide/unhide."""
    path = validate_file(file)

    from agent_xlsx.adapters import openpyxl_adapter as oxl

    if list_:
        result = oxl.manage_sheet(str(path), "list", sheet_name="")
        json_formatter.output(result)
        return

    if create:
        result = oxl.manage_sheet(str(path), "create", sheet_name=create, output_path=output)
        json_formatter.output(result)
        return

    if rename:
        if not new_name:
            raise AgentExcelError(
                "MISSING_ARGUMENT",
                "Rename requires --new-name",
                ["Usage: agent-xlsx sheet file.xlsx --rename OldName --new-name NewName"],
            )
        result = oxl.manage_sheet(
            str(path), "rename", sheet_name=rename, new_name=new_name, output_path=output
        )
        json_formatter.output(result)
        return

    if delete:
        result = oxl.manage_sheet(str(path), "delete", sheet_name=delete, output_path=output)
        json_formatter.output(result)
        return

    if copy:
        result = oxl.manage_sheet(
            str(path), "copy", sheet_name=copy, new_name=new_name, output_path=output
        )
        json_formatter.output(result)
        return

    if hide:
        result = oxl.manage_sheet(str(path), "hide", sheet_name=hide, output_path=output)
        json_formatter.output(result)
        return

    if unhide:
        result = oxl.manage_sheet(str(path), "unhide", sheet_name=unhide, output_path=output)
        json_formatter.output(result)
        return

    # No action specified — show help
    raise AgentExcelError(
        "MISSING_ACTION",
        "No sheet action specified",
        [
            "Use --list to list sheets",
            "Use --create <name> to create a sheet",
            "Use --rename <name> --new-name <new> to rename",
            "Use --delete <name> to delete",
            "Use --copy <name> --new-name <new> to copy",
            "Use --hide <name> or --unhide <name>",
        ],
    )

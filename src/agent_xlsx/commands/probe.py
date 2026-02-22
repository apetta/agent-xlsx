"""Ultra-fast workbook profiling via Polars + fastexcel."""

from typing import Optional

import typer

from agent_xlsx.adapters.polars_adapter import probe_workbook
from agent_xlsx.cli import app
from agent_xlsx.formatters.json_formatter import output, output_spreadsheet_data
from agent_xlsx.utils.errors import handle_error
from agent_xlsx.utils.validation import validate_file


@app.command()
@handle_error
def probe(
    file: str = typer.Argument(..., help="Path to the Excel file"),
    sheet: Optional[str] = typer.Option(None, "--sheet", "-s", help="Profile a specific sheet"),
    sample: int = typer.Option(0, "--sample", "-n", help="Sample rows (head + tail) per sheet"),
    stats: bool = typer.Option(
        False,
        "--stats",
        help="Include numeric/string summaries (implies --types)",
    ),
    types: bool = typer.Option(
        False,
        "--types",
        help="Include column types and null counts",
    ),
    full: bool = typer.Option(
        False,
        "--full",
        help="All detail: types + nulls + sample(3) + stats",
    ),
    no_header: bool = typer.Option(
        False,
        "--no-header",
        help="Treat row 1 as data, use column letters (A, B, C) as headers. "
        "Use for non-tabular sheets like P&L reports and dashboards.",
    ),
) -> None:
    """Ultra-fast workbook profiling — lean by default.

    Returns sheet names, dimensions, and headers with zero data parsing (<10ms).
    Use --types, --sample, --stats, or --full to opt into richer detail.
    """
    path = validate_file(file)

    if full:
        types = True
        stats = True
        sample = max(sample, 3)
    if stats:
        types = True  # stats implies types (need df for stats anyway)

    result = probe_workbook(
        filepath=path,
        sheet_name=sheet,
        sample_rows=sample,
        stats=stats,
        include_types=types,
        no_header=no_header,
    )

    output_spreadsheet_data(result)

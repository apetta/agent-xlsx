"""Main CLI entry point for agent-xlsx."""

from __future__ import annotations

import typer

app = typer.Typer(
    name="agent-xlsx",
    help="XLSX file CLI built with Agent Experience (AX) in mind..",
    no_args_is_help=True,
    pretty_exceptions_enable=False,
)


def _register_commands() -> None:
    """Import and register all command modules."""
    from agent_xlsx.commands import (
        export,  # noqa: F401
        license_cmd,  # noqa: F401
        objects,  # noqa: F401
        overview,  # noqa: F401
        probe,  # noqa: F401
        read,  # noqa: F401
        recalc,  # noqa: F401
        screenshot,  # noqa: F401
        search,  # noqa: F401
        sheet,  # noqa: F401
        vba,  # noqa: F401
        write,  # noqa: F401
    )
    from agent_xlsx.commands import format as _format  # noqa: F401
    from agent_xlsx.commands import inspect as _inspect  # noqa: F401


_register_commands()

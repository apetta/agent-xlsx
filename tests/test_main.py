"""Tests for __main__.py entry point: broad exception handler.

CliRunner invokes app (cli.py) directly, bypassing __main__.py.
These tests call main() directly to exercise the except-Exception handler.
"""

import json
import sys
from io import StringIO
from unittest.mock import patch

from agent_xlsx.__main__ import main


def test_internal_error_produces_json(sample_xlsx):
    """Unexpected exceptions produce structured JSON on stdout, traceback on stderr."""
    stdout_buf = StringIO()
    stderr_buf = StringIO()

    with (
        patch("sys.argv", ["agent-xlsx", "read", str(sample_xlsx), "--formulas"]),
        patch(
            "agent_xlsx.commands.read._read_with_formulas",
            side_effect=ValueError("simulated internal error"),
        ),
        patch.object(sys, "stdout", stdout_buf),
        patch.object(sys, "stderr", stderr_buf),
    ):
        try:
            main()
        except SystemExit as e:
            assert e.code == 1

    data = json.loads(stdout_buf.getvalue())
    assert data["error"] is True
    assert data["code"] == "INTERNAL_ERROR"
    assert data["exception_type"] == "ValueError"
    assert "simulated internal error" in data["message"]

    # Traceback preserved on stderr for developers
    assert "ValueError: simulated internal error" in stderr_buf.getvalue()


def test_unknown_option_returns_cli_usage_error():
    """Unknown CLI flags produce structured JSON with CLI_USAGE_ERROR code.

    Exercises the custom click.UsageError handler in __main__.py which converts
    Typer/Click usage errors into structured JSON on stdout (exit code 2).
    """
    stdout_buf = StringIO()

    with (
        patch("sys.argv", ["agent-xlsx", "--nonexistent-flag"]),
        patch.object(sys, "stdout", stdout_buf),
    ):
        try:
            main()
        except SystemExit as e:
            assert e.code == 2

    data = json.loads(stdout_buf.getvalue())
    assert data["error"] is True
    assert data["code"] == "CLI_USAGE_ERROR"
    assert "--nonexistent-flag" in data["message"]

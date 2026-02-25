"""Entry point wrapper ensuring all CLI errors produce structured JSON.

Intercepts Click/Typer UsageErrors (e.g. negative numbers parsed as flags)
and formats them as JSON with actionable suggestions, maintaining the
structured-output contract that LLM agents rely on.
"""

import json
import re
import sys

import click


def main() -> None:
    """Run the CLI app with structured error handling for Click errors."""
    from agent_xlsx.cli import app

    try:
        app(standalone_mode=False)
    except click.UsageError as e:
        _handle_usage_error(e)
    except SystemExit:
        raise
    except click.Abort:
        raise SystemExit(1)


def _handle_usage_error(error: click.UsageError) -> None:
    """Format a Click UsageError as structured JSON with helpful suggestions."""
    message = error.format_message()
    suggestions: list[str] = []

    # Detect negative number mistaken for a flag (e.g. "No such option: -4")
    neg_match = re.search(r"No such option: (-\d)", message)
    if neg_match:
        suggestions = [
            "Negative numbers are parsed as flags by the shell",
            "Use --value '-4.095' instead",
            "Or use -- sentinel: write file.xlsx A1 -- -4.095",
        ]

    # Detect close option name mismatches (e.g. --number instead of --number-format)
    opt_match = re.search(r"No such option: (--[\w-]+)", message)
    if opt_match and not neg_match:
        _corrections = {"--number": "Did you mean --number-format?"}
        corr = _corrections.get(opt_match.group(1))
        if corr:
            suggestions.append(corr)

    result = {"error": True, "code": "CLI_USAGE_ERROR", "message": message}
    if suggestions:
        result["suggestions"] = suggestions
    json.dump(result, sys.stdout, indent=2)
    sys.stdout.write("\n")
    raise SystemExit(2)


if __name__ == "__main__":
    main()

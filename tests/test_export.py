"""Tests for export command: JSON, CSV, Markdown formats with flags and envelopes."""

import json

from typer.testing import CliRunner

from agent_xlsx.cli import app

runner = CliRunner()


# ---------------------------------------------------------------------------
# Default format (JSON) — stdout
# ---------------------------------------------------------------------------


def test_export_json_default(rich_xlsx):
    """Default export produces JSON with expected envelope fields."""
    result = runner.invoke(app, ["export", str(rich_xlsx)])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["_data_origin"] == "untrusted_spreadsheet"
    assert "source" in data
    assert "row_count" in data
    assert "columns" in data
    assert "data" in data
    assert isinstance(data["data"], list)
    assert data["row_count"] == len(data["data"])


def test_export_json_explicit_format(rich_xlsx):
    """--format json produces the same structured output as the default."""
    result = runner.invoke(app, ["export", str(rich_xlsx), "--format", "json"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["_data_origin"] == "untrusted_spreadsheet"
    assert data["row_count"] > 0
    assert "export_time_ms" in data


# ---------------------------------------------------------------------------
# --sheet flag
# ---------------------------------------------------------------------------


def test_export_sheet_flag(multisheet_xlsx):
    """--sheet selects the specified sheet for export."""
    result = runner.invoke(
        app, ["export", str(multisheet_xlsx), "--sheet", "Beta", "--format", "json"]
    )
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["sheet"] == "Beta"
    assert data["columns"] == ["ID", "Value"]
    assert data["row_count"] == 10


def test_export_sheet_not_found(multisheet_xlsx):
    """--sheet with a non-existent name produces a SHEET_NOT_FOUND error."""
    result = runner.invoke(app, ["export", str(multisheet_xlsx), "--sheet", "Nope"])
    assert result.exit_code == 1
    data = json.loads(result.stdout)
    assert data["error"] is True
    assert data["code"] == "SHEET_NOT_FOUND"


# ---------------------------------------------------------------------------
# CSV format — raw stdout
# ---------------------------------------------------------------------------


def test_export_csv_raw_stdout(sample_xlsx):
    """--format csv writes raw CSV to stdout (no JSON wrapping)."""
    result = runner.invoke(app, ["export", str(sample_xlsx), "--format", "csv"])
    assert result.exit_code == 0, result.stdout
    lines = result.stdout.strip().split("\n")
    # First line should be the CSV header row
    assert "header" in lines[0]
    assert "amount" in lines[0]
    # At least header + 1 data row
    assert len(lines) >= 2


def test_export_csv_json_envelope(sample_xlsx):
    """--format csv --json-envelope wraps CSV in a JSON envelope."""
    result = runner.invoke(app, ["export", str(sample_xlsx), "--format", "csv", "--json-envelope"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["_data_origin"] == "untrusted_spreadsheet"
    assert data["format"] == "csv"
    assert "row_count" in data
    # data field contains the raw CSV string
    assert isinstance(data["data"], str)
    assert "header" in data["data"]


# ---------------------------------------------------------------------------
# Markdown format
# ---------------------------------------------------------------------------


def test_export_markdown_raw_stdout(sample_xlsx):
    """--format markdown writes a Markdown table to stdout."""
    result = runner.invoke(app, ["export", str(sample_xlsx), "--format", "markdown"])
    assert result.exit_code == 0, result.stdout
    lines = result.stdout.strip().split("\n")
    # Markdown table: header row, separator row, data rows
    assert lines[0].startswith("|")
    assert "---" in lines[1]
    assert len(lines) >= 3


def test_export_markdown_json_envelope(sample_xlsx):
    """--format markdown --json-envelope wraps Markdown in a JSON envelope."""
    result = runner.invoke(
        app, ["export", str(sample_xlsx), "--format", "markdown", "--json-envelope"]
    )
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert data["_data_origin"] == "untrusted_spreadsheet"
    assert data["format"] == "markdown"
    assert isinstance(data["data"], str)
    assert "|" in data["data"]


# ---------------------------------------------------------------------------
# --output (file write)
# ---------------------------------------------------------------------------


def test_export_output_json_file(rich_xlsx, tmp_path):
    """--output writes JSON to a file and prints a success message."""
    out = tmp_path / "export.json"
    result = runner.invoke(app, ["export", str(rich_xlsx), "--output", str(out)])
    assert result.exit_code == 0, result.stdout
    # Stdout is a success envelope
    meta = json.loads(result.stdout)
    assert meta["status"] == "success"
    assert meta["format"] == "json"
    assert meta["row_count"] > 0
    # File should contain the full export data
    file_data = json.loads(out.read_text())
    assert "data" in file_data
    assert file_data["row_count"] == meta["row_count"]


def test_export_output_csv_file(sample_xlsx, tmp_path):
    """--output with CSV writes the file and prints a success message."""
    out = tmp_path / "export.csv"
    result = runner.invoke(
        app, ["export", str(sample_xlsx), "--format", "csv", "--output", str(out)]
    )
    assert result.exit_code == 0, result.stdout
    meta = json.loads(result.stdout)
    assert meta["status"] == "success"
    assert meta["format"] == "csv"
    # File should contain raw CSV
    csv_text = out.read_text()
    assert "header" in csv_text


def test_export_output_markdown_file(sample_xlsx, tmp_path):
    """--output with Markdown writes the file and prints a success message."""
    out = tmp_path / "export.md"
    result = runner.invoke(
        app, ["export", str(sample_xlsx), "--format", "markdown", "--output", str(out)]
    )
    assert result.exit_code == 0, result.stdout
    meta = json.loads(result.stdout)
    assert meta["status"] == "success"
    assert meta["format"] == "markdown"
    md_text = out.read_text()
    assert "|" in md_text
    assert "---" in md_text


# ---------------------------------------------------------------------------
# --no-header flag
# ---------------------------------------------------------------------------


def test_export_no_header(sample_xlsx):
    """--no-header uses column letters (A, B, ...) instead of row-1 values."""
    result = runner.invoke(app, ["export", str(sample_xlsx), "--no-header", "--format", "json"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    # Column names should be letters, not header values
    for col in data["columns"]:
        assert col.isalpha(), f"Expected letter column name, got '{col}'"
    # Row count should include the original header row as data
    assert data["row_count"] >= 2


# ---------------------------------------------------------------------------
# --compact / --no-compact
# ---------------------------------------------------------------------------


def test_export_compact_drops_null_column(compact_xlsx):
    """Default compact mode drops the fully-null NullCol column."""
    result = runner.invoke(app, ["export", str(compact_xlsx), "--format", "json"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert "NullCol" not in data["columns"]
    assert "Name" in data["columns"]
    assert "Value" in data["columns"]
    assert "Category" in data["columns"]


def test_export_no_compact_keeps_null_column(compact_xlsx):
    """--no-compact preserves all columns including fully-null ones."""
    result = runner.invoke(app, ["export", str(compact_xlsx), "--no-compact", "--format", "json"])
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)
    assert "NullCol" in data["columns"]
    assert len(data["columns"]) == 4


# ---------------------------------------------------------------------------
# --json-envelope is ignored when --output is used
# ---------------------------------------------------------------------------


def test_export_json_envelope_ignored_with_output(sample_xlsx, tmp_path):
    """--json-envelope is ignored when --output is specified (CSV goes straight to file)."""
    out = tmp_path / "export.csv"
    result = runner.invoke(
        app,
        [
            "export",
            str(sample_xlsx),
            "--format",
            "csv",
            "--output",
            str(out),
            "--json-envelope",
        ],
    )
    assert result.exit_code == 0, result.stdout
    meta = json.loads(result.stdout)
    # Stdout is the file-write success message, not a JSON envelope
    assert meta["status"] == "success"
    assert meta["format"] == "csv"
    # File is raw CSV, not JSON-wrapped
    csv_text = out.read_text()
    assert csv_text.startswith("header")

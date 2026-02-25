# agent-xlsx

[![PyPI](https://img.shields.io/pypi/v/agent-xlsx)](https://pypi.org/project/agent-xlsx/)
[![Python](https://img.shields.io/pypi/pyversions/agent-xlsx)](https://pypi.org/project/agent-xlsx/)
[![License](https://img.shields.io/badge/licence-Apache--2.0-blue)](LICENSE)

**XLSX file CLI built with Agent Experience (AX) in mind.**

agent-xlsx gives LLM agents the same depth of understanding of Excel workbooks that a human gets by opening them in Excel — structure, data, formatting, charts, formulas, VBA, and visual layout — all accessible through a single CLI that returns token-efficient JSON.

```bash
agent-xlsx probe report.xlsx
agent-xlsx read report.xlsx "Sales!A1:F50"
agent-xlsx search report.xlsx "revenue" --ignore-case
agent-xlsx screenshot report.xlsx
```

---

## Why agent-xlsx?

LLM agents working with Excel files face a fundamental problem: existing libraries are designed for humans writing Python scripts, not for agents that need to build understanding of a workbook incrementally and efficiently.

**agent-xlsx solves this with three design principles:**

1. **Progressive Disclosure** — `probe` (structure) &rarr; `screenshot` (visual) &rarr; `read` (data) &rarr; `inspect` (metadata). Each layer adds detail only when needed. No wasted tokens.

2. **Speed First** — The primary data backend is Polars + fastexcel (Rust/Calamine), delivering 7-10x faster reads than openpyxl with zero-copy Arrow integration.

3. **Token Efficiency** — Every output is optimised for minimal token consumption. Aggregation over enumeration. Capped lists with counts. An agent builds comprehensive understanding of a workbook in 1-2 round-trips, not 10.

---

## Installation

### uvx (recommended — zero install)

```bash
uvx agent-xlsx probe report.xlsx
```

### Global install

```bash
uv tool install agent-xlsx
```

Also available via `pipx install agent-xlsx`.

### pip

```bash
pip install agent-xlsx
```

### Agent Skill

Give AI agents built-in knowledge of agent-xlsx commands and workflows:

```bash
npx skills add apetta/agent-xlsx
```

Compatible with [Claude Code](https://claude.ai/code), [Cursor](https://cursor.com), [Gemini CLI](https://geminicli.com), and [20+ other agents](https://agentskills.io).

### Aspose.Cells — third-party licence required

agent-xlsx includes [Aspose.Cells for Python](https://pypi.org/project/aspose-cells-python/) as a dependency for cross-platform screenshot, recalc, and objects support.

> **Important:** Aspose.Cells is a **proprietary, commercially licensed** library by [Aspose Pty Ltd](https://www.aspose.com/). It is **not** covered by this project's Apache-2.0 licence. By installing agent-xlsx you also install Aspose.Cells and agree to [Aspose's EULA](https://company.aspose.com/legal/eula). A [separate Aspose licence](https://purchase.aspose.com/pricing/cells/python-java) is required for production use without watermarks.

Without a licence, Aspose runs in **evaluation mode** (watermarks on rendered images, 100-file-per-session limit). To remove watermarks, purchase and set an Aspose licence:

```bash
agent-xlsx license --set /path/to/Aspose.Cells.lic
```

Or via environment variable:

```bash
export ASPOSE_LICENSE_PATH=/path/to/Aspose.Cells.lic
```

### Optional: LibreOffice (free fallback for `screenshot` and `recalc`)

**macOS:**
```bash
brew install --cask libreoffice
```

**Ubuntu / Debian / ECS:**
```bash
apt install libreoffice-calc
```

**Alpine:**
```bash
apk add libreoffice-calc
```

All other commands (probe, read, search, export, write, format, inspect, overview, sheet, vba) work with zero system dependencies.

---

## Quick Start

The recommended agent workflow is **probe first, then drill down**:

1. **Profile** the workbook — lean skeleton:
```bash
agent-xlsx probe workbook.xlsx
```

2. **Drill into types / samples** if needed:
```bash
agent-xlsx probe workbook.xlsx --types --sample 3
```

3. **Visual understanding** — see formatting, charts, layout:
```bash
agent-xlsx screenshot workbook.xlsx
```

4. **Read specific data:**
```bash
agent-xlsx read workbook.xlsx --sheet Sales "A1:F100"
```

5. **Inspect metadata** — formulas, charts, merged cells, conditional formatting:
```bash
agent-xlsx inspect workbook.xlsx --sheet Sales
```

---

## Commands

### `probe` — Ultra-Fast Workbook Profiling

The **first command an agent should run**. Lean by default — returns sheet names, dimensions, and headers with zero data parsing. Use flags to opt into richer detail.

```bash
agent-xlsx probe data.xlsx
agent-xlsx probe data.xlsx --types
agent-xlsx probe data.xlsx --sample 3
agent-xlsx probe data.xlsx --stats
agent-xlsx probe data.xlsx --full
agent-xlsx probe data.xlsx --sheet "Sales"
```

| Flag | Effect |
|------|--------|
| *(none)* | Sheet names, dims, headers only |
| `--types` | Add column types + null counts |
| `--sample N` | Add N head + N tail rows |
| `--stats` | Full stats (implies `--types`) |
| `--full` | Shorthand for `--types --sample 3 --stats` |
| `--sheet` | Target a single sheet |

**Default output** (~250 tokens for 6 sheets):
```json
{
  "file": "data.xlsx",
  "size_bytes": 107679,
  "format": "xlsx",
  "probe_time_ms": 7.9,
  "sheets": [
    {
      "name": "txns",
      "index": 0,
      "visible": true,
      "rows": 255,
      "cols": 34,
      "headers": ["user_id", "txn_day", "txn_month", "amount", "currency", "..."]
    }
  ]
}
```

**With `--full`** (types + sample + stats):
```json
{
  "file": "data.xlsx",
  "size_bytes": 107679,
  "format": "xlsx",
  "probe_time_ms": 18.5,
  "sheets": [
    {
      "name": "txns",
      "index": 0,
      "visible": true,
      "rows": 255,
      "cols": 34,
      "headers": ["user_id", "txn_day", "txn_month", "amount", "currency", "..."],
      "column_types": {
        "user_id": "string",
        "txn_day": "float64",
        "amount": "float64",
        "txn_date": "datetime",
        "category": "string"
      },
      "null_counts": {"user_id": 0, "amount": 0, "currency": 0},
      "sample": {
        "head": [["8bb055ad-...", 1, 12, -39.0, "GBP"]],
        "tail": [["8bb055ad-...", 1, 8, -150.0, "GBP"]]
      },
      "numeric_summary": {
        "amount": {"min": -4888.06, "max": 5000.0, "mean": -142.3, "std": 892.1}
      },
      "string_summary": {
        "category": {"unique": 12, "top_values": ["Software & Technology", "Sales", "Employees"]}
      }
    }
  ]
}
```

### `overview` — Structural Metadata

Focuses on elements that `probe` cannot detect: formulas, charts, tables, named ranges. Uses openpyxl for metadata that the Rust backend doesn't expose.

```bash
agent-xlsx overview data.xlsx
agent-xlsx overview data.xlsx --include-formulas
agent-xlsx overview data.xlsx --include-formatting
```

```json
{
  "file": "data.xlsx",
  "size_bytes": 107679,
  "overview_time_ms": 157.2,
  "sheets": [
    {
      "name": "txns",
      "index": 0,
      "dimensions": "A1:AZ324",
      "row_count": 324,
      "col_count": 52,
      "has_formulas": false,
      "has_charts": true,
      "chart_count": 1,
      "has_tables": false
    }
  ]
}
```

### `read` — Data Extraction

Read data from any range or sheet. Default path uses Polars + fastexcel for speed. Use `--formulas` to fall back to openpyxl for formula string extraction.

```bash
agent-xlsx read data.xlsx
agent-xlsx read data.xlsx "A1:F50"
agent-xlsx read data.xlsx --sheet Sales "B2:G100"
agent-xlsx read data.xlsx --limit 500 --offset 100
agent-xlsx read data.xlsx --formulas
agent-xlsx read data.xlsx --sort amount --descending
```

```json
{
  "range": "A1:E5",
  "dimensions": {"rows": 4, "cols": 5},
  "headers": ["user_id", "txn_day", "txn_month", "txn_year", "txn_hour"],
  "data": [
    ["8bb055ad-caa1-40b6-a577-832425b02408", 1, 12, 2024, 8],
    ["8bb055ad-caa1-40b6-a577-832425b02408", 1, 12, 2024, 4]
  ],
  "row_count": 4,
  "truncated": false,
  "backend": "polars+fastexcel",
  "read_time_ms": 8.9
}
```

### `search` — Cross-Workbook Search

Search for values across all sheets. Supports regex and case-insensitive matching.

```bash
agent-xlsx search data.xlsx "revenue"
agent-xlsx search data.xlsx "rev.*" --regex
agent-xlsx search data.xlsx "stripe" --ignore-case
agent-xlsx search data.xlsx "error" --sheet Summary
agent-xlsx search data.xlsx "SUM(" --in-formulas
```

```json
{
  "query": "Stripe",
  "match_count": 25,
  "matches": [
    {"sheet": "txns", "column": "txn_description", "row": 12, "value": "Stripe DemoCompany Ltd. Payout UK"},
    {"sheet": "txns", "column": "merchant_name", "row": 12, "value": "Stripe"}
  ],
  "truncated": true,
  "search_time_ms": 18.8
}
```

### `inspect` — Detailed Element Inspection

Deep inspection of workbook elements: formulas, charts, merged cells, named ranges, comments, conditional formatting, data validation, and hyperlinks.

```bash
agent-xlsx inspect data.xlsx --sheet Sales
agent-xlsx inspect data.xlsx --sheet Sales --range A1:C10
agent-xlsx inspect data.xlsx --names
agent-xlsx inspect data.xlsx --charts
agent-xlsx inspect data.xlsx --comments
agent-xlsx inspect data.xlsx --conditional Sales
agent-xlsx inspect data.xlsx --validation Sales
agent-xlsx inspect data.xlsx --hyperlinks Sales
```

### `screenshot` — Full-Fidelity HD Visual Capture

Export workbook sheets as HD PNG images. Three rendering engines auto-detected in order: **Aspose.Cells** (cross-platform, included) → **Excel** (xlwings, highest fidelity) → **LibreOffice** (free fallback). Use `--engine` to force a specific backend.

```bash
agent-xlsx screenshot data.xlsx
agent-xlsx screenshot data.xlsx --sheet Summary
agent-xlsx screenshot data.xlsx --sheet "Sales,Summary"
agent-xlsx screenshot data.xlsx "Sales!A1:F20"
agent-xlsx screenshot data.xlsx --engine aspose
agent-xlsx screenshot data.xlsx --dpi 300
agent-xlsx screenshot data.xlsx --output ./shots/
agent-xlsx screenshot data.xlsx --timeout 60
```

**Single sheet/range output:**
```json
{
  "status": "success",
  "format": "png",
  "path": "/tmp/agent-xlsx/data_Summary.png",
  "sheet": "Summary",
  "size_bytes": 245000,
  "dpi": 200,
  "capture_time_ms": 3200.0,
  "engine": "libreoffice+pymupdf"
}
```

**Multi-sheet output:**
```json
{
  "status": "success",
  "format": "png",
  "dpi": 200,
  "sheets": [
    {"name": "Sales", "path": "/tmp/agent-xlsx/data_Sales.png", "size_bytes": 245000},
    {"name": "Summary", "path": "/tmp/agent-xlsx/data_Summary.png", "size_bytes": 89000}
  ],
  "capture_time_ms": 4100.0,
  "engine": "libreoffice+pymupdf"
}
```

### `export` — Bulk Data Export

Export entire sheets to JSON, CSV, or Markdown.

```bash
agent-xlsx export data.xlsx --format csv
agent-xlsx export data.xlsx --format markdown
agent-xlsx export data.xlsx --format json
agent-xlsx export data.xlsx --format csv --output out.csv
agent-xlsx export data.xlsx --format csv --sheet Sales
```

### `write` — Write Values and Formulas

Write values or formulas to cells. Supports single cells, ranges (via JSON), and CSV file imports.

```bash
agent-xlsx write data.xlsx "A1" "Hello"
agent-xlsx write data.xlsx "A1" "=SUM(B1:B100)" --formula
agent-xlsx write data.xlsx "A1:C3" --json '[[1,2,3],[4,5,6],[7,8,9]]'
agent-xlsx write data.xlsx "A1" --from-csv import.csv
agent-xlsx write data.xlsx "A1" "42" --number-format "0.00%"
agent-xlsx write data.xlsx "A1" "Hello" --sheet Summary
agent-xlsx write data.xlsx "A1" "Hello" --output new_file.xlsx
```

Use `--output` to write to a new file and preserve the original.

### `format` — Read and Apply Cell Formatting

Read or modify cell formatting: fonts, fills, borders, number formats.

Read formatting:

```bash
agent-xlsx format data.xlsx "A1" --read --sheet Sales
```

Apply formatting:

```bash
agent-xlsx format data.xlsx "A1:D1" --font '{"bold": true, "size": 14}'
agent-xlsx format data.xlsx "B2:B100" --number-format "#,##0.00"
agent-xlsx format data.xlsx "A1:D10" --fill '{"color": "FFFF00"}'
agent-xlsx format data.xlsx "A1:D10" --border '{"style": "thin"}'
agent-xlsx format data.xlsx "A1:D10" --copy-from "G1"
```

```json
{
  "cell": "A1",
  "value": "user_id",
  "font": {"name": "Aptos Narrow", "size": 12.0, "bold": false, "italic": false},
  "fill": {"type": "solid", "color": "indexed:9"},
  "border": {
    "top": {"style": "thin", "color": "indexed:10"},
    "bottom": {"style": "thin", "color": "indexed:10"}
  },
  "alignment": {"horizontal": null, "vertical": null, "wrap_text": null},
  "number_format": "@"
}
```

### `sheet` — Sheet Management

List, create, rename, delete, copy, hide, and unhide sheets.

```bash
agent-xlsx sheet data.xlsx --list
agent-xlsx sheet data.xlsx --create "New Sheet"
agent-xlsx sheet data.xlsx --rename "Old Name" --new-name "New Name"
agent-xlsx sheet data.xlsx --delete "Temp"
agent-xlsx sheet data.xlsx --copy "Template" --new-name "Q1 Report"
agent-xlsx sheet data.xlsx --hide "Internal"
agent-xlsx sheet data.xlsx --unhide "Internal"
```

### `vba` — VBA Macro Analysis

Extract and analyse VBA macros using oletools. Works headless on all platforms without Microsoft Excel.

```bash
agent-xlsx vba macros.xlsm --list
agent-xlsx vba macros.xlsm --read Main
agent-xlsx vba macros.xlsm --read-all
agent-xlsx vba macros.xlsm --security
```

### `recalc` — Formula Recalculation

Scan for formula errors or trigger a full recalculation. Auto-detects engine: Aspose → Excel → LibreOffice.

Scan for errors (no engine needed):

```bash
agent-xlsx recalc data.xlsx --check-only
```

Full recalculation (requires Excel, Aspose, or LibreOffice):

```bash
agent-xlsx recalc data.xlsx
agent-xlsx recalc data.xlsx --engine aspose
agent-xlsx recalc data.xlsx --timeout 120
```

```json
{
  "status": "success",
  "mode": "check_only",
  "total_formulas": 847,
  "total_errors": 3,
  "check_time_ms": 184.1,
  "error_summary": {
    "#REF!": {"count": 2, "locations": ["Sales!F12", "Sales!F15"]},
    "#DIV/0!": {"count": 1, "locations": ["Summary!C8"]}
  }
}
```

---

## Architecture

agent-xlsx uses a **multi-backend architecture**, choosing the fastest backend capable of satisfying each request:

```
                            agent-xlsx CLI
                                  |
      +---------------+-----------+-----------+---------------+
      |               |           |           |               |
Polars+fastexcel   openpyxl   Aspose.Cells   xlwings      LibreOffice
 (Rust/Calamine)  (Pure Py)  (Cross-plat)    (Excel)       (Headless)

  Data reads      Metadata   Screenshots    Screenshots   Screenshots
  Profiling       Formulas   Recalc         Recalc        Recalc
  Search          Formatting Objects        Objects
  Export          Writes

  + oletools (VBA extraction & analysis)
```

**Rendering engine auto-detection**: Aspose.Cells → Excel (xlwings) → LibreOffice. Use `--engine` to force a specific backend.

| Backend | Role | Speed | Used by |
|---------|------|-------|---------|
| **Polars + fastexcel** | Primary data engine | 7-10x faster than openpyxl | `probe`, `read`, `search`, `export` |
| **openpyxl** | Metadata + writes | Baseline | `overview`, `inspect`, `write`, `format`, `sheet` |
| **Aspose.Cells** ([separately licensed](https://company.aspose.com/legal/eula)) | Cross-platform rendering (default) | Fast (rendering) | `screenshot`, `recalc`, `objects` |
| **xlwings** (Excel) | Highest-fidelity rendering | Fast (rendering) | `screenshot`, `recalc`, `objects`, `vba --run` |
| **LibreOffice + PyMuPDF** | Free rendering fallback | Moderate (rendering) | `screenshot`, `recalc` |
| **oletools** | VBA extraction | Fast | `vba` |

### Why not just openpyxl?

openpyxl creates a Python object for every cell. For a 100K-row workbook, that's millions of allocations and ~50x the file size in RAM. Polars + fastexcel reads the same data through Rust with zero-copy Arrow transfer — the data never touches Python's heap until the agent needs it.

---

## File Format Support

| Format | Extension | Read | Write | Screenshot | VBA |
|--------|-----------|------|-------|------------|-----|
| Excel (Open XML) | `.xlsx` | Yes | Yes | Yes | N/A |
| Excel (Macro-enabled) | `.xlsm` | Yes | Yes | Yes | Yes |
| Excel (Binary) | `.xlsb` | Yes | - | Yes | Yes |
| Excel (Legacy) | `.xls` | Yes | - | Yes | - |
| OpenDocument | `.ods` | Yes | - | Yes | - |

---

## Deployment

agent-xlsx is designed for headless deployment in agentic infrastructure — no GUI, no Excel installation, no Docker requirement.

### AWS ECS / Container

Aspose.Cells is included as a dependency (see [licence note](#aspose-cells--third-party-licence-required)). Add system fonts for rendered output on Linux:

```dockerfile
FROM python:3.12-slim

RUN apt-get update && \
    apt-get install -y --no-install-recommends libgdiplus libfontconfig1 fonts-liberation && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

RUN pip install agent-xlsx
RUN agent-xlsx --help
```

Optionally add LibreOffice as a fallback rendering engine:

```dockerfile
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice-calc && \
    apt-get clean && rm -rf /var/lib/apt/lists/*
```

---

## Error Handling

All errors return structured JSON with an error code, message, and actionable suggestions:

```json
{
  "error": true,
  "code": "FILE_NOT_FOUND",
  "message": "File not found: missing.xlsx",
  "suggestions": [
    "Check the file path is correct",
    "Ensure the file exists and is readable"
  ]
}
```

Error codes include: `FILE_NOT_FOUND`, `INVALID_FORMAT`, `SHEET_NOT_FOUND`, `INVALID_RANGE`, `EXCEL_REQUIRED`, `ASPOSE_NOT_INSTALLED`, `LIBREOFFICE_REQUIRED`, `NO_RENDERING_BACKEND`, and more.

---

## Performance

agent-xlsx uses a tiered backend strategy that matches the fastest engine to each task:

| Operation | Speed | Backend |
|-----------|-------|---------|
| `probe` (default, lean) | Fastest | Polars + fastexcel |
| `probe --full` (types + sample + stats) | Fast | Polars + fastexcel |
| `read` (range) | Fastest | Polars + fastexcel |
| `search` (cross-workbook) | Fast | Polars + fastexcel |
| `overview` | Moderate | openpyxl |
| `inspect` | Moderate | openpyxl |
| `recalc --check-only` | Moderate | openpyxl |
| `screenshot` (PNG, per-sheet) | Slower (rendering) | LibreOffice + PyMuPDF |
| `recalc` (full) | Slower (rendering) | LibreOffice |

The Polars + fastexcel backend is 7-10x faster than openpyxl for equivalent operations. Performance scales linearly with file size. Every command output includes `file_size_human` so agents can calibrate timeout expectations accordingly.

---

## Development

Clone, install, and set up hooks:

```bash
git clone https://github.com/apetta/agent-xlsx.git
cd agent-xlsx
uv sync --group dev
uv run pre-commit install --install-hooks -t pre-commit -t pre-push
```

This installs two git hooks automatically:

- **pre-commit** — `ruff check --fix` + `ruff format` (runs on every commit)
- **pre-push** — lint + format + `pytest` (full suite before pushing)

Run commands locally:

```bash
uv run agent-xlsx probe sample_data.xlsx
```

Run checks manually:

```bash
uv run ruff check src/ tests/
uv run ruff format src/ tests/
uv run pytest
```

### Project Structure

```
src/agent_xlsx/
  cli.py                    # Typer CLI entry point
  commands/                 # 14 command implementations
    probe.py                  # Ultra-fast profiling (Polars)
    overview.py               # Structural metadata (openpyxl)
    read.py                   # Data extraction (Polars)
    search.py                 # Cross-workbook search (Polars)
    export.py                 # Bulk export (Polars)
    inspect.py                # Deep inspection (openpyxl)
    write.py                  # Write operations (openpyxl)
    format.py                 # Formatting read/write (openpyxl)
    sheet.py                  # Sheet management (openpyxl)
    screenshot.py             # Visual capture (Excel/Aspose/LO)
    objects.py                # Embedded objects (Excel/Aspose)
    vba.py                    # VBA analysis (oletools)
    recalc.py                 # Recalculation (Excel/Aspose/LO)
    license_cmd.py            # Aspose licence management
  adapters/                 # Backend adapters
    polars_adapter.py         # Polars + fastexcel (primary data)
    openpyxl_adapter.py       # openpyxl (metadata + writes)
    xlwings_adapter.py        # xlwings/Excel (rendering + objects)
    aspose_adapter.py         # Aspose.Cells (cross-platform rendering)
    libreoffice_adapter.py    # LibreOffice headless (fallback rendering)
    oletools_adapter.py       # oletools (VBA extraction)
  formatters/               # Output formatting
    json_formatter.py         # Token-efficient JSON output
    token_optimizer.py        # Output capping and aggregation
  utils/                    # Shared utilities
    errors.py                 # Error types and handler
    validation.py             # File and range validation
    constants.py              # Caps and limits
    memory.py                 # Memory budget checking
    dates.py                  # Date detection and serial→ISO conversion
    config.py                 # Persistent config (~/.agent-xlsx/)
```

---

## Licence

This project is licensed under **Apache-2.0** — see [LICENSE](LICENSE) for details.

**Third-party notice:** agent-xlsx depends on [Aspose.Cells for Python](https://pypi.org/project/aspose-cells-python/), which is proprietary software by Aspose Pty Ltd, distributed under its own [EULA](https://company.aspose.com/legal/eula). Aspose.Cells is **not** covered by this project's Apache-2.0 licence. A [separate commercial licence](https://purchase.aspose.com/pricing/cells/python-java) from Aspose is required for production use without evaluation watermarks. By installing agent-xlsx you accept responsibility for complying with Aspose's licensing terms.

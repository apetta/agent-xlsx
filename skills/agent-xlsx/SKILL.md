---
name: agent-xlsx
description: "Interact with Excel files (.xlsx, .xlsm, .xlsb, .xls, .ods) using the agent-xlsx CLI for data extraction, analysis, writing, formatting, visual capture, VBA analysis, and sheet management. Use when the user asks to: (1) Read, analyse, or search data in spreadsheets, (2) Write values or formulas to cells, (3) Inspect formatting, formulas, charts, or metadata, (4) Take screenshots or visual captures of sheets, (5) Export sheets to CSV/JSON/Markdown, (6) Manage sheets (create, rename, delete, copy, hide), (7) Analyse or execute VBA macros, (8) List/export embedded objects (charts, shapes, pictures), (9) Check for formula errors, or (10) Any task involving Excel file interaction. Prefer over openpyxl/pandas scripts — faster, structured JSON optimised for AI."
---

# agent-xlsx

XLSX CLI for AI agents. Structured JSON to stdout. Polars+fastexcel for data reads (7-10x faster than openpyxl), openpyxl for metadata/writes, three rendering engines for visual capture (Excel → Aspose → LibreOffice), oletools for VBA.

## Workflow: Progressive Disclosure

Start lean, opt into detail:

```
probe (<10ms)  →  screenshot (~3s)  →  read (data)  →  inspect (metadata)
```

**Always start with `probe`:**

```bash
agent-xlsx probe <file>                        # Sheet names, dims, headers, column_map
agent-xlsx probe <file> --types                # + column types, null counts
agent-xlsx probe <file> --full                 # + types, sample(3), stats, date_summary
agent-xlsx probe <file> -s "Sales" --full      # Single-sheet deep-dive
```

Probe returns `column_map` — map headers to column letters for building ranges:

```json
{"column_map": {"user_id": "A", "amount": "E"}, "last_col": "W"}
```

## Essential Commands

### Data (Polars — fast)

```bash
# Read
agent-xlsx read <file> "A1:F50"                    # Range (positional arg)
agent-xlsx read <file> -s Sales "B2:G100"          # Sheet + range
agent-xlsx read <file> --limit 500 --offset 100    # Pagination
agent-xlsx read <file> --sort amount --descending  # Sorted
agent-xlsx read <file> --formulas                  # Formula strings (slower, openpyxl)

# Search
agent-xlsx search <file> "revenue"                 # Exact match, all sheets
agent-xlsx search <file> "rev.*" --regex           # Regex
agent-xlsx search <file> "stripe" --ignore-case    # Case-insensitive
agent-xlsx search <file> "SUM(" --in-formulas      # Inside formula strings

# Export
agent-xlsx export <file> --format csv              # CSV to stdout
agent-xlsx export <file> --format markdown          # Markdown table
agent-xlsx export <file> --format csv -o out.csv -s Sales
```

### Metadata (openpyxl)

```bash
# Overview — structural summary
agent-xlsx overview <file>
agent-xlsx overview <file> --include-formulas --include-formatting

# Inspect — comprehensive single-pass metadata
agent-xlsx inspect <file> -s Sales                 # Everything: formulas, merges, tables, charts, comments, cond. formatting, validation, hyperlinks, freeze panes
agent-xlsx inspect <file> -s Sales --range A1:C10  # Scoped
agent-xlsx inspect <file> --names                  # Named ranges
agent-xlsx inspect <file> --charts                 # Chart metadata
agent-xlsx inspect <file> --comments               # Cell comments

# Format — read/write cell formatting
agent-xlsx format <file> "A1" --read -s Sales      # Read formatting
agent-xlsx format <file> "A1:D1" --font '{"bold": true, "size": 14}'
agent-xlsx format <file> "B2:B100" --number-format "#,##0.00"
agent-xlsx format <file> "A1:D10" --copy-from "G1" # Copy all formatting
```

### Write (openpyxl)

```bash
agent-xlsx write <file> "A1" "Hello"                               # Single value
agent-xlsx write <file> "A1" "=SUM(B1:B100)" --formula             # Formula
agent-xlsx write <file> "A1:C3" --json '[[1,2,3],[4,5,6],[7,8,9]]' # 2D array
agent-xlsx write <file> "A1" --from-csv data.csv                   # CSV import
agent-xlsx write <file> "A1" "Hello" -o new.xlsx -s Sales          # New file

# Sheet management
agent-xlsx sheet <file> --list
agent-xlsx sheet <file> --create "New Sheet"
agent-xlsx sheet <file> --rename "Old" --new-name "New"
agent-xlsx sheet <file> --delete "Temp"
agent-xlsx sheet <file> --copy "Template" --new-name "Q1"
agent-xlsx sheet <file> --hide "Internal"
```

### Visual & Analysis (3 engines: Excel → Aspose → LibreOffice)

```bash
# Screenshot — HD PNG capture (auto-fits columns)
agent-xlsx screenshot <file>                       # All sheets
agent-xlsx screenshot <file> -s Sales              # Specific sheet
agent-xlsx screenshot <file> -s "Sales,Summary"    # Multiple sheets
agent-xlsx screenshot <file> "Sales!A1:F20"        # Range capture
agent-xlsx screenshot <file> -o ./shots/           # Output directory
agent-xlsx screenshot <file> --engine aspose       # Force engine
agent-xlsx screenshot <file> --dpi 300             # DPI (Aspose/LibreOffice)

# Objects — embedded charts, shapes, pictures
agent-xlsx objects <file>                          # List all
agent-xlsx objects <file> --export "Chart 1"       # Export chart as PNG

# Recalc — formula error checking
agent-xlsx recalc <file> --check-only              # Scan for #REF!, #DIV/0! (no engine needed)
agent-xlsx recalc <file>                           # Full recalculation (needs engine)
```

### VBA (oletools + xlwings)

```bash
agent-xlsx vba <file> --list                       # List modules + security summary
agent-xlsx vba <file> --read ModuleName            # Read module code
agent-xlsx vba <file> --read-all                   # All module code
agent-xlsx vba <file> --security                   # Full security analysis (risk level, IOCs)
agent-xlsx vba <file> --run "Module1.MyMacro"      # Execute (requires Excel)
agent-xlsx vba <file> --run "MyMacro" --args '[1]' # With arguments
```

### Config

```bash
agent-xlsx license --status                        # Check Aspose install + licence status
agent-xlsx license --set /path/to/Aspose.Cells.lic # Save licence path
agent-xlsx license --clear                         # Remove saved licence
```

## Common Patterns

### Profile a new spreadsheet

```bash
agent-xlsx probe file.xlsx --full             # Structure + types + samples + stats
agent-xlsx screenshot file.xlsx               # Visual understanding
```

### Find and extract specific data

```bash
agent-xlsx probe file.xlsx                    # Get column_map
agent-xlsx search file.xlsx "overdue" -i      # Find matching cells
agent-xlsx read file.xlsx "A1:G50" -s Invoices  # Extract the range
```

### Audit formulas

```bash
agent-xlsx recalc file.xlsx --check-only      # Scan for errors (#REF!, #DIV/0!)
agent-xlsx read file.xlsx --formulas          # See formula strings
agent-xlsx search file.xlsx "VLOOKUP" --in-formulas  # Find specific formulas
```

### Write results back

```bash
agent-xlsx write file.xlsx "H1" "Status" -o updated.xlsx
agent-xlsx write updated.xlsx "H2" --json '["Done","Pending","Done"]'
```

### Export for downstream use

```bash
agent-xlsx export file.xlsx --format csv -s Sales -o sales.csv
agent-xlsx export file.xlsx --format markdown  # Stdout
```

### Analyse VBA for security

```bash
agent-xlsx vba suspect.xlsm --security        # Risk assessment
agent-xlsx vba suspect.xlsm --read-all        # Read all code
```

## Critical Rules

1. **Always `probe` first** — instant (<10ms), returns sheet names and column_map
2. **`--formulas` for formula strings** — default read returns computed values only (Polars, fast). Add `--formulas` for formula text (openpyxl, slower)
3. **`--in-formulas` for formula search** — default search checks cell values. Add `--in-formulas` to search formula strings
4. **Dates auto-convert** — Excel serial numbers (44927) become ISO strings ("2023-01-15") automatically
5. **Check `truncated` field** — results are capped (search: 25, formulas: 50, comments: 20). Narrow query if truncated
6. **Range is positional** — `"A1:F50"` or `"Sheet1!A1:F50"` is a positional argument, not a flag
7. **`-o` preserves original** — write/format save to a new file when `--output` specified
8. **Screenshot needs an engine** — requires Excel, Aspose, or LibreOffice. See [backends.md](references/backends.md)
9. **VBA read vs run** — oletools for read/analysis (cross-platform), xlwings for execution (Excel required)
10. **500MB memory limit** — large files auto-chunk. Use `--limit` for big reads
11. **Writable: .xlsx and .xlsm only** — .xlsb, .xls, .ods are read-only

## Output Format

JSON to stdout. Errors:

```json
{"error": true, "code": "SHEET_NOT_FOUND", "message": "...", "suggestions": ["..."]}
```

Codes: `FILE_NOT_FOUND`, `INVALID_FORMAT`, `SHEET_NOT_FOUND`, `INVALID_RANGE`, `EXCEL_REQUIRED`, `LIBREOFFICE_REQUIRED`, `ASPOSE_NOT_INSTALLED`, `NO_RENDERING_BACKEND`, `MEMORY_EXCEEDED`, `VBA_NOT_FOUND`, `CHART_NOT_FOUND`.

## Deep-Dive Reference

| Reference | When to Read |
|-----------|-------------|
| [commands.md](references/commands.md) | Full flag reference for all 14 commands with types and defaults |
| [backends.md](references/backends.md) | Rendering engine details, platform quirks, Aspose licensing, format support |

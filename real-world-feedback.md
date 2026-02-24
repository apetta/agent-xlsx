# agent-xlsx: Real-World Agent Feedback

Field testing of agent-xlsx **v0.4.0** in a multi-agent pipeline that downloads the World Bank WDI dataset (~76MB, 401K rows, 6 sheets), extracts economic indicators for specific countries, and builds a formatted comparison dashboard. All calls via `uvx agent-xlsx`.

The pipeline has three agent roles:
- **Miner** (read-only): probe, search, read, export
- **Modeler** (write): write, format, sheet, recalc, read (for verification)
- **Orchestrator**: delegates to the above, QA-checks the output

---

## P0 — Blockers

### 1. `write` cannot create a new file

`write` always calls `load_workbook()` on the target file, so writing to a non-existent file fails with `BadZipFile`. Agents building new spreadsheets from extracted data (the most common write use case) have no way to bootstrap a blank workbook without an external tool.

```bash
# Fails — "File is not a zip file"
uvx agent-xlsx write Dashboard.xlsx "A1" --json '[[1,2,3]]'
```

**Current workaround:** Create a blank xlsx externally via `uv run --with openpyxl python3 -c "from openpyxl import Workbook; Workbook().save('Dashboard.xlsx')"` before any agent-xlsx write calls. This requires the agent to have access to Python/openpyxl outside of agent-xlsx, which defeats the purpose of a self-contained CLI.

**Suggestion:** Auto-create a new workbook when the target file doesn't exist. This is the single highest-impact improvement for agent workflows.

### 2. `write --output` on large source files hangs silently

Using a large file as the source with `--output` to create a derived file causes openpyxl to load the entire source into memory before writing. On the 76MB WDI file, this hung indefinitely with no progress indicator or timeout.

```bash
# Hangs — openpyxl loads all 76MB into memory to copy the workbook
uvx agent-xlsx write WDIEXCEL.xlsx "A1" --json '[[1,2]]' --output Dashboard.xlsx
```

An agent that doesn't know the source file is huge will try this and burn its timeout. The error is silent — no warning, no memory guard, no progress output.

**Suggestion:** Either (a) emit a warning when the source file exceeds a size threshold (e.g. 10MB), or (b) add a `--max-source-size` guard that fails fast with a clear error, or (c) if `write` supports auto-creating new files (P0 #1), agents can just write to fresh files and this becomes a non-issue.

---

## P1 — Significant Token/Turn Waste

### 3. Search needs a `--columns` filter

Searching the WDI Series sheet for `"GDP growth"` returned 6 matches — but only 1 was the indicator name (column C). The other 5 were from description columns (K, E, O) containing thousands of characters of methodology text. Each match was 500-2,000 tokens of noise.

```bash
# Returns 6 matches, 5 of which are multi-paragraph descriptions
uvx agent-xlsx search WDIEXCEL.xlsx "GDP growth" --sheet Series

# What I actually needed:
uvx agent-xlsx search WDIEXCEL.xlsx "GDP growth" --sheet Series --columns C
# or
uvx agent-xlsx search WDIEXCEL.xlsx "GDP growth" --sheet Series --columns "Indicator Name"
```

For "inflation", 21 matches were returned — most from description text. That's easily 10K+ tokens of waste in a single search call.

**Suggestion:** Add `--columns` (accepting column letters or header names) to restrict which columns are searched. This is the single biggest token-efficiency improvement for agents working with wide or text-heavy sheets.

### 4. Search needs a `--limit` / `--max-results` flag

Searching for a country code like `^ARG$` in the 401K-row Data sheet returned 25 matches (the max cap) — all identical `"ARG"` values from consecutive rows. Only the first match was useful (to locate the start of Argentina's row block).

```bash
# Returns 25 identical matches, wastes 24x tokens
uvx agent-xlsx search WDIEXCEL.xlsx "^ARG$" --sheet Data --regex

# What I actually needed:
uvx agent-xlsx search WDIEXCEL.xlsx "^ARG$" --sheet Data --regex --limit 1
```

**Suggestion:** Add `--limit N` to cap results below the default 25. For agents doing positional lookups (finding where a value first appears), `--limit 1` would save significant context.

### 5. Search needs row-range scoping

After finding that Argentina starts at row 84,506 and spans ~1,505 rows, I needed to find which row within that block contains indicator code `NY.GDP.MKTP.KD.ZG`. But search scans the entire sheet — there's no way to restrict it to a row range.

The workaround was reading all ~1,500 indicator codes into context and scanning programmatically (via a Python pipe). An LLM agent can't pipe to Python — it would need to ingest 1,500 rows into its context window and scan them, or do multiple paginated reads.

```bash
# Scans all 401K rows, returns 25 matches across all countries
uvx agent-xlsx search WDIEXCEL.xlsx "NY.GDP.MKTP.KD.ZG" --sheet Data

# What I actually needed:
uvx agent-xlsx search WDIEXCEL.xlsx "NY.GDP.MKTP.KD.ZG" --sheet Data --range A84506:D86010
```

**Suggestion:** Add `--range` to restrict search to a cell range. Combined with `--columns`, this would let agents do surgical lookups: "find indicator X within country Y's row block, in the indicator code column only."

### 6. `--json` in `write` should support formula strings

The Modeler agent wrote 4 AVERAGE formulas as 4 separate `write --formula` calls. Each call is a full openpyxl load/save cycle. Formulas can't be embedded in `--json` arrays because `--formula` is a boolean flag that applies to the entire write.

```bash
# 4 separate calls, 4 load/save cycles:
uvx agent-xlsx write Dashboard.xlsx "C8" "=AVERAGE(C2:M2)" --formula
uvx agent-xlsx write Dashboard.xlsx "C9" "=AVERAGE(C3:M3)" --formula
uvx agent-xlsx write Dashboard.xlsx "C10" "=AVERAGE(C4:M4)" --formula
uvx agent-xlsx write Dashboard.xlsx "C11" "=AVERAGE(C5:M5)" --formula

# What I actually needed — one call:
uvx agent-xlsx write Dashboard.xlsx "C8:C11" --json '[["=AVERAGE(C2:M2)"],["=AVERAGE(C3:M3)"],["=AVERAGE(C4:M4)"],["=AVERAGE(C5:M5)"]]' --formula
# or auto-detect "=" prefix in --json values as formulas
```

**Suggestion:** Either (a) allow `--formula` to apply to `--json` arrays (treating all string values starting with `=` as formulas), or (b) auto-detect formula strings in `--json` values by the `=` prefix. This would cut the Modeler's write turns in half for formula-heavy dashboards.

---

## P2 — Minor / Documentation

### 7. Performance claims should be caveated for large files

The README and tool descriptions claim sub-50ms for probe/read/search. On the 76MB WDI file:

| Command | Claimed | Actual |
|---------|---------|--------|
| `probe` | <10ms | ~3,800ms |
| `read` (single range) | ~9ms | ~6,800ms |
| `search` | ~19ms | ~4,000ms |

These are dominated by file I/O on the 76MB xlsx, not by processing. The benchmarks are presumably from smaller files. Agents (and their timeout configurations) built around the sub-50ms claim will be surprised.

**Suggestion:** Caveat performance numbers with file size context, e.g. "sub-50ms on files under 10MB; scales linearly with file size."

### 8. Multi-range reads: column letters as headers

When reading from non-row-1 ranges (e.g. `A85014:D85014`), headers come back as column letters `["A", "B", "C", "D"]` instead of the actual header names from row 1. This is technically correct, but it means the agent has to mentally map column letters back to header names using the `column_map` from `probe`.

Not necessarily a bug — it may be intentional. But worth noting that agents will often need to cross-reference with the probe output to interpret results.

---

## Summary: Impact on Agent Turn Count

| Workflow Step | Actual Turns | With Improvements | Savings |
|--------------|-------------|-------------------|---------|
| Find indicator codes (Series search) | 3 (search + read + read) | 1 (search --columns C) | 2 turns |
| Find country row ranges | 2 (search ARG + search BRA) | 2 (search --limit 1 each) | 0 turns, ~75% token reduction |
| Find indicator rows within country blocks | 2 (read 1500 rows + Python pipe, per country) | 2 (search --range per country) | 0 turns, ~95% token reduction |
| Create new dashboard file | 2 (external Python bootstrap + first write) | 1 (write auto-creates) | 1 turn |
| Write formulas | 4 (one per formula) | 1 (--json with formulas) | 3 turns |
| **Total** | **~15+ turns** | **~8 turns** | **~7 turns saved** |

The token savings from `--columns` and `--limit` on search are arguably more impactful than the turn savings — preventing thousands of tokens of description text and duplicate matches from entering the agent's context window.

"""Output capping and aggregation helpers for token efficiency."""

from __future__ import annotations

from typing import Any


def cap_list(items: list[Any], max_count: int) -> dict[str, Any]:
    """Cap a list and add truncation metadata."""
    return {
        "items": items[:max_count],
        "total": len(items),
        "truncated": len(items) > max_count,
    }


def summarise_formulas(formula_cells: list[dict[str, Any]], max_count: int = 50) -> dict[str, Any]:
    """Aggregate formula cells into a token-efficient summary."""
    columns_with_formulas: set[str] = set()
    for cell in formula_cells:
        col = "".join(c for c in cell.get("cell", "") if c.isalpha())
        if col:
            columns_with_formulas.add(col)

    return {
        "formula_count": len(formula_cells),
        "formula_columns": sorted(columns_with_formulas),
        "sample_formulas": formula_cells[:max_count],
        "truncated": len(formula_cells) > max_count,
    }

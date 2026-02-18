"""Shared DataFrame utilities used across commands."""

from __future__ import annotations

import polars as pl


def apply_compact(df: pl.DataFrame, compact: bool) -> pl.DataFrame:
    """Drop fully-null columns when compact mode is enabled."""
    if not compact or len(df) == 0:
        return df
    non_null_cols = [col for col in df.columns if df[col].null_count() < len(df)]
    return df.select(non_null_cols) if non_null_cols else df

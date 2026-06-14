"""Aggregate tools: group_by_aggregate, value_counts, sample_rows, compute_percentiles."""

from __future__ import annotations

from typing import Optional

import polars as pl

from excel_mcp.tools.load import _load, _lazy, _df_to_markdown

_SUPPORTED_FUNCS = {"sum", "mean", "count", "min", "max", "median", "std", "var", "first", "last", "n_unique"}


def group_by_aggregate(
    file_path: str,
    sheet_name: str,
    group_by: list[str],
    aggregations: list[dict],
) -> str:
    """Group data by one or more columns and compute aggregate functions.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        group_by: List of column names to group by.
        aggregations: List of aggregation specs, each with "col" and "func".
            Supported functions: sum, mean, count, min, max, median, std, var,
            first, last, n_unique.
            Optionally include "alias" to rename the result column.
            Example: [{"col": "Revenue", "func": "sum", "alias": "total_revenue"},
                      {"col": "OrderId", "func": "count"}]
    """
    exprs = []
    for agg in aggregations:
        col = agg.get("col")
        func = agg.get("func", "").lower()
        alias = agg.get("alias", f"{col}_{func}")
        if func not in _SUPPORTED_FUNCS:
            raise ValueError(f"Unsupported function '{func}'. Supported: {sorted(_SUPPORTED_FUNCS)}")
        exprs.append(getattr(pl.col(col), func)().alias(alias))

    result = (
        _lazy(file_path, sheet_name)
        .group_by(group_by)
        .agg(exprs)
        .sort(group_by)
        .collect()
    )
    return f"Group-by result: {result.height} groups.\n\n" + _df_to_markdown(result)


def value_counts(
    file_path: str,
    sheet_name: str,
    column: str,
    sort: bool = True,
    normalize: bool = False,
    top_n: Optional[int] = None,
) -> str:
    """Count occurrences of each unique value in a column (frequency table).

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        column: Column name to compute value counts for.
        sort: Sort by count descending (default True).
        normalize: Return proportions instead of raw counts (default False).
        top_n: Only return the top N most frequent values.
    """
    df = _load(file_path, sheet_name)
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found. Available: {df.columns}")

    # value_counts isn't available on LazyFrame — use group_by + count instead
    lf = _lazy(file_path, sheet_name)
    counts_lf = lf.group_by(column).agg(pl.len().alias("count"))
    if sort:
        counts_lf = counts_lf.sort("count", descending=True)
    if normalize:
        total = df.height
        counts_lf = counts_lf.with_columns((pl.col("count") / total).alias("proportion"))
    if top_n is not None:
        counts_lf = counts_lf.head(top_n)

    counts = counts_lf.collect()
    return f"Value counts for '{column}' ({df.height} total rows).\n\n" + _df_to_markdown(counts)


def sample_rows(
    file_path: str,
    sheet_name: str,
    n: Optional[int] = None,
    fraction: Optional[float] = None,
    seed: int = 42,
    shuffle: bool = True,
) -> str:
    """Return a random sample of rows for exploratory analysis.

    Provide either `n` (absolute count) or `fraction` (proportion, e.g. 0.1 for 10%).

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        n: Exact number of rows to sample.
        fraction: Fraction of total rows to sample (0.0–1.0).
        seed: Random seed for reproducibility (default 42).
        shuffle: Shuffle the sample (default True).
    """
    if n is None and fraction is None:
        raise ValueError("Provide either 'n' or 'fraction'.")
    if n is not None and fraction is not None:
        raise ValueError("Provide either 'n' or 'fraction', not both.")

    # sample() is not available on LazyFrame — collect first, then sample
    df = _load(file_path, sheet_name)
    sample = df.sample(fraction=fraction, n=n, seed=seed, shuffle=shuffle)
    return f"Sample of {sample.height} rows (from {df.height} total).\n\n" + _df_to_markdown(sample)


def compute_percentiles(
    file_path: str,
    sheet_name: str,
    column: str,
    quantiles: list[float],
    group_by: Optional[list[str]] = None,
) -> str:
    """Compute percentiles for a numeric column, optionally grouped.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        column: Numeric column to compute percentiles for.
        quantiles: List of quantile values between 0.0 and 1.0.
            Examples: [0.5] for median, [0.25, 0.5, 0.75] for quartiles,
            [0.9, 0.95, 0.99] for tail percentiles.
        group_by: Optional list of columns to group by before computing.
            When provided, percentiles are computed per group.
    """
    bad = [q for q in quantiles if not (0.0 <= q <= 1.0)]
    if bad:
        raise ValueError(f"Quantiles must be between 0.0 and 1.0, got: {bad}")

    lf = _lazy(file_path, sheet_name)

    if group_by:
        exprs = [
            pl.col(column).quantile(q).alias(f"p{int(q * 100)}")
            for q in quantiles
        ]
        result = lf.group_by(group_by).agg(exprs).sort(group_by).collect()
    else:
        exprs = [
            pl.col(column).quantile(q).alias(f"p{int(q * 100)}")
            for q in quantiles
        ]
        result = lf.select(exprs).collect()

    label = f"Percentiles for '{column}'"
    if group_by:
        label += f" grouped by {group_by}"
    return f"{label}.\n\n" + _df_to_markdown(result)

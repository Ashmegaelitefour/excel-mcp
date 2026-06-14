"""In-memory DataFrame cache keyed by (absolute_path, sheet_name).

Stores collected DataFrames so metadata (height, columns, dtypes) is always
available cheaply. Lazy computation happens inside each tool, not here.
"""

from __future__ import annotations

import polars as pl

_cache: dict[tuple[str, str], pl.DataFrame] = {}


def get_df(file_path: str, sheet_name: str) -> pl.DataFrame | None:
    return _cache.get((file_path, sheet_name))


def set_df(file_path: str, sheet_name: str, df: pl.DataFrame) -> None:
    _cache[(file_path, sheet_name)] = df

def get_lazy(file_path: str, sheet_name: str) -> pl.LazyFrame | None:
    """Return a LazyFrame view of the cached DataFrame, or None if not cached."""
    df = _cache.get((file_path, sheet_name))
    return df.lazy() if df is not None else None


def clear_sheet(file_path: str, sheet_name: str) -> None:
    _cache.pop((file_path, sheet_name), None)


def clear(file_path: str | None = None) -> None:
    if file_path is None:
        _cache.clear()
    else:
        keys = [k for k in _cache if k[0] == file_path]
        for k in keys:
            del _cache[k]

"""Transform tools: filter_rows, select_columns, drop_columns, sort_data,
drop_duplicates, cast_columns, fill_nulls, add_column, reset_sheet."""

from __future__ import annotations

from typing import Literal, Optional

import polars as pl

from excel_mcp import state
from excel_mcp.security import validate_path
from excel_mcp.tools.load import _load, _lazy, _df_to_markdown


def filter_rows(
    file_path: str,
    sheet_name: str,
    expression: str,
) -> dict:
    """Filter rows using a SQL WHERE-style expression and update the session cache.

    The filtered result replaces the cached DataFrame so subsequent tool calls
    on this sheet operate on the filtered data.

    Accepts standard SQL conditions, e.g.:
      - "Age > 30"
      - "Country = 'US' AND Revenue > 1000"
      - "Status IN ('Active', 'Pending')"
      - "Name LIKE '%Corp%'"

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        expression: SQL WHERE clause (without the WHERE keyword).
    """
    original_rows = _load(file_path, sheet_name).height
    lf = _lazy(file_path, sheet_name)
    # Use lazy SQL — filter pushes down into scan for CSV/Parquet
    ctx = pl.SQLContext(data=lf, eager=False)
    filtered = ctx.execute(f"SELECT * FROM data WHERE {expression}").collect()

    abs_path = validate_path(file_path)
    state.set_df(abs_path, sheet_name, filtered)
    null_counts = filtered.null_count().row(0)
    schema = [
        {"name": col, "dtype": str(dtype), "null_count": nulls}
        for col, dtype, nulls in zip(filtered.columns, filtered.dtypes, null_counts)
    ]
    return {
        "filtered_rows": filtered.height,
        "original_rows": original_rows,
        "columns": schema,
        "tip": "Use preview(offset=N) to inspect rows, or run_query with a LIMIT clause.",
    }


def select_columns(
    file_path: str,
    sheet_name: str,
    columns: list[str],
    rename_map: Optional[dict[str, str]] = None,
) -> dict:
    """Select and optionally rename columns, updating the session cache.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        columns: List of column names to keep.
        rename_map: Optional dict mapping old names to new names, e.g. {"OldName": "NewName"}.
    """
    lf = _lazy(file_path, sheet_name).select(columns)
    if rename_map:
        lf = lf.rename(rename_map)
    df = lf.collect()
    abs_path = validate_path(file_path)
    state.set_df(abs_path, sheet_name, df)
    return {
        "selected_columns": df.columns,
        "rows": df.height,
        "preview": _df_to_markdown(df.head(5)),
    }


def drop_columns(
    file_path: str,
    sheet_name: str,
    columns: list[str],
) -> dict:
    """Drop specific columns from the sheet, updating the session cache.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        columns: List of column names to remove.
    """
    df = _load(file_path, sheet_name)
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ValueError(f"Columns not found: {missing}. Available: {df.columns}")
    df = _lazy(file_path, sheet_name).drop(columns).collect()
    abs_path = validate_path(file_path)
    state.set_df(abs_path, sheet_name, df)
    return {
        "dropped": columns,
        "remaining_columns": df.columns,
        "rows": df.height,
    }


def sort_data(
    file_path: str,
    sheet_name: str,
    by: list[str],
    descending: bool | list[bool] = False,
    nulls_last: bool = True,
) -> str:
    """Sort the sheet data by one or more columns and update the session cache.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        by: List of column names to sort by.
        descending: True to sort descending. Pass a list to control direction per column.
        nulls_last: Place null values at the end (default True).
    """
    if isinstance(descending, bool):
        descending = [descending] * len(by)
    df = _lazy(file_path, sheet_name).sort(by, descending=descending, nulls_last=nulls_last).collect()
    abs_path = validate_path(file_path)
    state.set_df(abs_path, sheet_name, df)
    return _df_to_markdown(df.head(20))


def drop_duplicates(
    file_path: str,
    sheet_name: str,
    subset: Optional[list[str]] = None,
    keep: Literal["first", "last", "any", "none"] = "first",
) -> str:
    """Remove duplicate rows from the sheet and update the session cache.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        subset: Columns to consider for deduplication. Defaults to all columns.
        keep: Which duplicate to retain — "first", "last", "any", or "none" (drop all).
    """
    original_rows = _load(file_path, sheet_name).height
    df = _lazy(file_path, sheet_name).unique(subset=subset, keep=keep, maintain_order=True).collect()
    abs_path = validate_path(file_path)
    state.set_df(abs_path, sheet_name, df)
    removed = original_rows - df.height
    return (
        f"Removed {removed} duplicate rows. {df.height} rows remain.\n\n"
        + _df_to_markdown(df.head(10))
    )


def cast_columns(
    file_path: str,
    sheet_name: str,
    cast_map: dict[str, str],
) -> dict:
    """Cast one or more columns to a different data type and update the session cache.

    Supported type strings: Int8, Int16, Int32, Int64, UInt8, UInt16, UInt32, UInt64,
    Float32, Float64, Boolean, String, Utf8, Date, Datetime, Time, Duration.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        cast_map: Dict mapping column name to target type, e.g. {"Age": "Int32", "Price": "Float64"}.
    """
    existing_cols = _load(file_path, sheet_name).columns
    exprs = []
    for col, type_str in cast_map.items():
        if col not in existing_cols:
            raise ValueError(f"Column '{col}' not found. Available: {existing_cols}")
        dtype = getattr(pl, type_str, None)
        if dtype is None or not isinstance(dtype, type):
            raise ValueError(
                f"Unknown type '{type_str}'. Use Polars type names like Int64, Float64, String, Date."
            )
        exprs.append(pl.col(col).cast(dtype))
    df = _lazy(file_path, sheet_name).with_columns(exprs).collect()
    abs_path = validate_path(file_path)
    state.set_df(abs_path, sheet_name, df)
    return {
        "cast": cast_map,
        "schema": {c: str(t) for c, t in zip(df.columns, df.dtypes)},
    }


def fill_nulls(
    file_path: str,
    sheet_name: str,
    strategy: Literal["forward", "backward", "mean", "median", "min", "max"] | None = None,
    fill_value: str | int | float | None = None,
    columns: Optional[list[str]] = None,
) -> str:
    """Fill null values using a strategy or a literal value and update the session cache.

    Provide either `strategy` or `fill_value`, not both.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        strategy: Fill strategy — "forward", "backward", "mean", "median", "min", "max".
        fill_value: A literal value to fill nulls with (e.g. 0, "Unknown").
        columns: Columns to apply filling to. Defaults to all columns.
    """
    if strategy and fill_value is not None:
        raise ValueError("Provide either 'strategy' or 'fill_value', not both.")
    if not strategy and fill_value is None:
        raise ValueError("Provide either 'strategy' or 'fill_value'.")

    lf = _lazy(file_path, sheet_name)
    schema = lf.collect_schema()
    target_cols = columns if columns else list(schema.names())

    if fill_value is not None:
        lf = lf.with_columns([pl.col(c).fill_null(fill_value) for c in target_cols])
    elif strategy in ("forward", "backward"):
        lf = lf.with_columns([pl.col(c).fill_null(strategy=strategy) for c in target_cols])
    else:
        # stat-based: numeric cols only; must collect stat values first then re-apply
        df_tmp = lf.select(target_cols).collect()
        exprs = []
        for c in target_cols:
            if df_tmp[c].dtype.is_numeric():
                stat = getattr(df_tmp[c], strategy)()
                exprs.append(pl.col(c).fill_null(stat))
        if exprs:
            lf = lf.with_columns(exprs)

    df = lf.collect()
    abs_path = validate_path(file_path)
    state.set_df(abs_path, sheet_name, df)
    remaining_nulls = sum(df.select(target_cols).null_count().row(0))
    return f"Nulls filled. Remaining nulls across target columns: {remaining_nulls}.\n\n" + _df_to_markdown(df.head(10))


def add_column(
    file_path: str,
    sheet_name: str,
    column_name: str,
    expression: str,
) -> dict:
    """Add a new computed column using a SQL expression and update the session cache.

    The expression is evaluated via Polars SQLContext. Reference existing columns
    by name and use standard SQL functions.

    Examples:
      - expression: "Revenue * 0.1"
      - expression: "UPPER(Country)"
      - expression: "Revenue - Cost"

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        column_name: Name for the new column.
        expression: SQL expression for the new column value.
    """
    lf = _lazy(file_path, sheet_name)
    ctx = pl.SQLContext(data=lf, eager=False)
    result = ctx.execute(f'SELECT *, ({expression}) AS "{column_name}" FROM data').collect()
    abs_path = validate_path(file_path)
    state.set_df(abs_path, sheet_name, result)
    return {
        "added_column": column_name,
        "dtype": str(result[column_name].dtype),
        "columns": result.columns,
        "preview": _df_to_markdown(result.head(5)),
    }


def reset_sheet(
    file_path: str,
    sheet_name: str,
) -> str:
    """Reload the original sheet from disk, discarding all in-session transformations.

    Use this to undo filters, column drops, or any other cached changes.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name to reset.
    """
    abs_path = validate_path(file_path)
    state.clear_sheet(abs_path, sheet_name)
    df = _load(abs_path, sheet_name)
    return f"Sheet '{sheet_name}' reset from disk: {df.height} rows × {df.width} columns."

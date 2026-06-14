"""Inspect tools: list_sheets, load_sheet, peek, get_schema, describe, preview."""

from __future__ import annotations

import os
from typing import Optional

import polars as pl

from excel_mcp import state
from excel_mcp.security import validate_path

# Maximum rows rendered in any markdown table response to keep LLM context manageable.
_MAX_DISPLAY_ROWS = 50


def _scan_file(abs_path: str, skip_rows: int = 0) -> pl.LazyFrame | None:
    """Return a LazyFrame for CSV/TSV/Parquet, or None for Excel (must be eager)."""
    ext = os.path.splitext(abs_path)[1].lower()
    if ext == ".csv":
        return pl.scan_csv(abs_path, skip_rows=skip_rows, infer_schema_length=10_000)
    if ext == ".tsv":
        return pl.scan_csv(abs_path, separator="\t", skip_rows=skip_rows, infer_schema_length=10_000)
    if ext == ".parquet":
        return pl.scan_parquet(abs_path)
    return None


def _read_file(abs_path: str, sheet_name: str, skip_rows: int = 0) -> pl.DataFrame:
    """Collect a full DataFrame from disk. Used for initial cache population."""
    lf = _scan_file(abs_path, skip_rows)
    if lf is not None:
        return lf.collect()
    ext = os.path.splitext(abs_path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xlsb"):
        return pl.read_excel(abs_path, sheet_name=sheet_name, read_options={"skip_rows": skip_rows})
    if ext == ".xls":
        return pl.read_excel(abs_path, sheet_name=sheet_name, engine="xlrd", read_options={"skip_rows": skip_rows})
    raise ValueError(f"Unsupported file type: {ext}. Supported: .xlsx, .xls, .csv, .tsv, .parquet")


def _load(file_path: str, sheet_name: str = "Sheet1") -> pl.DataFrame:
    """Return the cached DataFrame, loading from disk on first access."""
    abs_path = validate_path(file_path)
    cached = state.get_df(abs_path, sheet_name)
    if cached is not None:
        return cached
    df = _read_file(abs_path, sheet_name)
    state.set_df(abs_path, sheet_name, df)
    return df


def _lazy(file_path: str, sheet_name: str) -> pl.LazyFrame:
    """Return a LazyFrame for the sheet.

    - For CSV/TSV/Parquet: returns a fresh scan so Polars can push down
      filters and projections into the file read.
    - For Excel (or anything already cached): wraps the cached DataFrame
      in a LazyFrame so the same lazy API works everywhere.
    """
    abs_path = validate_path(file_path)
    ext = os.path.splitext(abs_path)[1].lower()

    # For flat files prefer a fresh scan even if already cached — predicate
    # and projection pushdown only work from a LazyFrame rooted at scan_*.
    if ext in (".csv", ".tsv", ".parquet"):
        cached = state.get_df(abs_path, sheet_name)
        if cached is not None:
            # Use cached schema / skip_rows already applied; wrap as lazy
            return cached.lazy()
        lf = _scan_file(abs_path)
        # Populate the cache so metadata tools (get_schema, list_sheets) work
        df = lf.collect()
        state.set_df(abs_path, sheet_name, df)
        return df.lazy()

    # Excel: always use the collected cache
    return _load(abs_path, sheet_name).lazy()


def _df_to_markdown(df: pl.DataFrame, max_rows: int = _MAX_DISPLAY_ROWS) -> str:
    display = df.head(max_rows)
    header = "| " + " | ".join(display.columns) + " |"
    separator = "| " + " | ".join("---" for _ in display.columns) + " |"
    rows = []
    for row in display.iter_rows():
        rows.append("| " + " | ".join("" if v is None else str(v) for v in row) + " |")
    table = "\n".join([header, separator] + rows)
    if df.height > max_rows:
        table += f"\n\n_Showing {max_rows} of {df.height} rows._"
    return table


def peek(
    file_path: str,
    sheet_name: str,
    n_rows: int = 10,
) -> str:
    """Read the first N raw rows of a file directly from disk, bypassing the session cache.

    Use this BEFORE load_sheet to inspect the raw structure of the file — title rows,
    subtitle rows, merged headers — so you can determine the correct skip_rows value
    to pass to load_sheet.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        n_rows: Number of raw rows to return (default 10).
    """
    abs_path = validate_path(file_path)
    ext = os.path.splitext(abs_path)[1].lower()

    if ext in (".xlsx", ".xlsm", ".xlsb"):
        df = pl.read_excel(abs_path, sheet_name=sheet_name, has_header=False, read_options={"n_rows": n_rows})
    elif ext == ".xls":
        df = pl.read_excel(abs_path, sheet_name=sheet_name, engine="xlrd", has_header=False, read_options={"n_rows": n_rows})
    elif ext in (".csv", ".tsv"):
        sep = "\t" if ext == ".tsv" else ","
        df = pl.scan_csv(abs_path, separator=sep, has_header=False, infer_schema_length=0).head(n_rows).collect()
    else:
        raise ValueError(f"peek is not supported for {ext}")

    return (
        f"Raw first {df.height} rows (no header parsing, row numbers are 0-based).\n"
        f"Identify which row contains the real column headers, then call load_sheet with skip_rows=<that row index>.\n\n"
        + _df_to_markdown(df)
    )


def load_sheet(
    file_path: str,
    sheet_name: str,
    skip_rows: int = 0,
) -> dict:
    """Explicitly load a sheet into the session cache with control over header rows.

    Use this when the Excel file has extra title/subtitle rows above the real
    column headers — pass skip_rows=N to skip those rows so they are not counted
    as data. All subsequent tool calls on this sheet will use the correctly loaded
    version.

    Example: a sheet with 1 title row + 1 subtitle row above the headers → skip_rows=2

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name to load.
        skip_rows: Number of rows to skip before reading the header (default 0).
    """
    abs_path = validate_path(file_path)
    state.clear_sheet(abs_path, sheet_name)
    df = _read_file(abs_path, sheet_name, skip_rows=skip_rows)
    state.set_df(abs_path, sheet_name, df)
    return {
        "sheet": sheet_name,
        "skipped_rows": skip_rows,
        "data_rows": df.height,
        "columns": df.columns,
    }


def find_dirty_rows(
    file_path: str,
    sheet_name: str,
    columns: Optional[list[str]] = None,
    include_empty_strings: bool = True,
) -> dict:
    """Find rows that contain null, NaN, or empty string values.

    Returns a per-column summary of dirty row counts plus a preview of the
    affected rows so you can decide whether to fill, drop, or flag them.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        columns: Columns to check. Defaults to all columns.
        include_empty_strings: Treat empty/whitespace-only strings as dirty
                               in addition to nulls (default True).
    """
    lf = _lazy(file_path, sheet_name)
    schema = lf.collect_schema()
    target_cols = columns if columns else list(schema.names())

    missing_cols = [c for c in target_cols if c not in schema.names()]
    if missing_cols:
        raise ValueError(f"Columns not found: {missing_cols}. Available: {list(schema.names())}")

    # Build a dirty-flag expression per column
    dirty_exprs = []
    for col in target_cols:
        expr = pl.col(col).is_null()
        if include_empty_strings and schema[col] == pl.String:
            expr = expr | (pl.col(col).str.strip_chars() == "")
        dirty_exprs.append(expr.alias(f"_dirty_{col}"))

    # Single lazy pass: add dirty flags, filter to rows with at least one dirty col
    flagged_lf = lf.with_columns(dirty_exprs)
    any_dirty = pl.any_horizontal([pl.col(f"_dirty_{c}") for c in target_cols])
    dirty_rows = flagged_lf.filter(any_dirty).collect()

    # Build per-column summary from the dirty flags
    summary = []
    for col in target_cols:
        flag_col = f"_dirty_{col}"
        count = int(dirty_rows[flag_col].sum())
        if count > 0:
            summary.append({"column": col, "dirty_rows": count})

    # Drop the helper flag columns before displaying
    flag_cols = [f"_dirty_{c}" for c in target_cols]
    display_df = dirty_rows.drop(flag_cols)

    if display_df.height == 0:
        return {
            "total_dirty_rows": 0,
            "columns_checked": target_cols,
            "message": "No dirty rows found.",
        }

    return {
        "total_dirty_rows": display_df.height,
        "columns_checked": target_cols,
        "per_column": summary,
        "preview": _df_to_markdown(display_df),
        "tip": (
            "Use fill_nulls to fill these values, filter_rows to remove them, "
            "or drop_duplicates after fixing."
        ),
    }


def list_sheets(file_path: str) -> dict:
    """List all sheet names in a workbook or describe a flat file.

    Works with .xlsx, .xls, .csv, .tsv, and .parquet files.
    For Excel files: returns all sheet names with their row/column counts.
    For flat files (CSV/TSV/Parquet): returns a single virtual sheet named after the file.

    Args:
        file_path: Absolute or relative path to the file.
    """
    abs_path = validate_path(file_path)
    ext = os.path.splitext(abs_path)[1].lower()

    if ext in (".xlsx", ".xlsm", ".xlsb"):
        import fastexcel
        wb = fastexcel.read_excel(abs_path)
        sheet_names = wb.sheet_names
    elif ext == ".xls":
        import xlrd
        wb = xlrd.open_workbook(abs_path)
        sheet_names = wb.sheet_names()
    elif ext in (".csv", ".tsv", ".parquet"):
        df = _load(abs_path, os.path.basename(abs_path))
        return {
            "file": abs_path,
            "sheets": [{"sheet": os.path.basename(abs_path), "rows": df.height, "columns": df.width}],
        }
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    sheets = []
    for name in sheet_names:
        df = _load(abs_path, name)
        sheets.append({"sheet": name, "rows": df.height, "columns": df.width})
    return {"file": abs_path, "sheets": sheets}


def get_schema(file_path: str, sheet_name: str) -> dict:
    """Return column names, data types, null counts, and sample values for a sheet.

    Args:
        file_path: Path to the file (.xlsx, .xls, .csv, .tsv, .parquet).
        sheet_name: Sheet name (use filename for CSV/TSV/Parquet).
    """
    df = _load(file_path, sheet_name)
    null_counts = df.null_count().row(0)
    sample_row = df.head(1).row(0) if df.height > 0 else [None] * df.width
    columns = []
    for col, dtype, nulls, sample in zip(df.columns, df.dtypes, null_counts, sample_row):
        columns.append({
            "name": col,
            "dtype": str(dtype),
            "null_count": nulls,
            "null_pct": round(nulls / df.height * 100, 2) if df.height else 0,
            "sample": str(sample),
        })
    return {
        "sheet": sheet_name,
        "total_rows": df.height,
        "total_columns": df.width,
        "columns": columns,
    }


def describe(
    file_path: str,
    sheet_name: str,
    columns: Optional[list[str]] = None,
) -> str:
    """Return descriptive statistics (min, max, mean, std, percentiles) for numeric columns.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        columns: Optional list of column names to limit the output. Defaults to all numeric columns.
    """
    lf = _lazy(file_path, sheet_name)
    if columns:
        lf = lf.select(columns)

    # Collect schema without full materialise to know numeric cols
    schema = lf.collect_schema()
    numeric_cols = [name for name, dtype in schema.items() if dtype.is_numeric()]
    if not numeric_cols:
        return "No numeric columns found in the selection."

    # Only collect the columns we need
    stats = lf.select(numeric_cols).collect().describe()
    return _df_to_markdown(stats)


def preview(
    file_path: str,
    sheet_name: str,
    n_rows: int = 10,
    offset: int = 0,
    columns: Optional[list[str]] = None,
) -> str:
    """Preview rows of a sheet as a markdown table.

    Args:
        file_path: Path to the file.
        sheet_name: Sheet name.
        n_rows: Number of rows to return (default 10, max 50).
        offset: Row offset for pagination (0-based, default 0).
        columns: Optional list of column names to include.
    """
    n_rows = min(n_rows, _MAX_DISPLAY_ROWS)
    lf = _lazy(file_path, sheet_name)
    if columns:
        lf = lf.select(columns)
    # slice pushes down into the scan for CSV/Parquet
    slice_df = lf.slice(offset, n_rows).collect()
    total_rows = _load(file_path, sheet_name).height
    header = f"Rows {offset}–{offset + slice_df.height - 1} of {total_rows} total.\n\n"
    return header + _df_to_markdown(slice_df, max_rows=_MAX_DISPLAY_ROWS)

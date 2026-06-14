"""MCP server entry point — registers all Excel ETL tools via FastMCP."""

from mcp.server.fastmcp import FastMCP

from excel_mcp.security import warn_if_unrestricted
from excel_mcp.tools import (
    # inspect
    list_sheets,
    load_sheet,
    peek,
    get_schema,
    describe,
    preview,
    find_dirty_rows,
    # transform
    filter_rows,
    select_columns,
    drop_columns,
    sort_data,
    drop_duplicates,
    cast_columns,
    fill_nulls,
    add_column,
    reset_sheet,
    # aggregate
    group_by_aggregate,
    value_counts,
    sample_rows,
    compute_percentiles,
    # advanced
    join_sheets,
    pivot_data,
    unpivot_data,
    concat_sheets,
    window_function,
    save_sheet,
    clone_sheet,
    to_pandas,
    export_data,
    run_query,
)

app = FastMCP(
    "excel-mcp",
    instructions=(
        "ETL tools for Excel, CSV, TSV and Parquet files powered by Polars. "
        "Inspect, filter, transform, aggregate, join, pivot, and export tabular data "
        "at scale. All operations are cached in-session; use reset_sheet to reload from disk."
    ),
)

# --- Inspect ---
app.tool()(list_sheets)
app.tool()(peek)
app.tool()(load_sheet)
app.tool()(get_schema)
app.tool()(describe)
app.tool()(preview)
app.tool()(find_dirty_rows)

# --- Transform ---
app.tool()(filter_rows)
app.tool()(select_columns)
app.tool()(drop_columns)
app.tool()(sort_data)
app.tool()(drop_duplicates)
app.tool()(cast_columns)
app.tool()(fill_nulls)
app.tool()(add_column)
app.tool()(reset_sheet)

# --- Aggregate ---
app.tool()(group_by_aggregate)
app.tool()(value_counts)
app.tool()(sample_rows)
app.tool()(compute_percentiles)

# --- Advanced ---
app.tool()(join_sheets)
app.tool()(pivot_data)
app.tool()(unpivot_data)
app.tool()(concat_sheets)
app.tool()(window_function)
app.tool()(save_sheet)
app.tool()(clone_sheet)
app.tool()(to_pandas)
app.tool()(export_data)
app.tool()(run_query)


def main() -> None:
    warn_if_unrestricted()
    app.run(transport="stdio")


if __name__ == "__main__":
    main()

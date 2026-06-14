# excel-mcp

An MCP server that lets LLMs inspect, query, and transform tabular files using [Polars](https://pola.rs/) — the fast, Arrow-native DataFrame library.

Supports **Excel** (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`), **CSV**, **TSV**, and **Parquet** files.

---

## Quick start

### With `uvx` (no install needed)

The fastest way — `uvx` downloads and runs the package on demand:

```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["excel-mcp"]
    }
  }
}
```

### With `pip`

```bash
pip install excel-mcp
```

Then use `excel-mcp` as the command in your client config (see setup sections below).

---

## Client setup

### Claude Desktop

Edit the config file:

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

**With `uvx`** (no prior install needed):

```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["excel-mcp"]
    }
  }
}
```

**With `pip install excel-mcp`**:

```json
{
  "mcpServers": {
    "excel": {
      "type": "stdio",
      "command": "excel-mcp"
    }
  }
}
```

Restart Claude Desktop. You can now ask:

> "Load sales.xlsx, filter to Q4, group by region, and show me total revenue."

### Claude Code (CLI)

```bash
claude mcp add excel-mcp -- uvx excel-mcp
```

### GitHub Copilot / VS Code

**Option A — Workspace config** (`.vscode/mcp.json`, committed to the repo):

```json
{
  "servers": {
    "excel": {
      "type": "stdio",
      "command": "uvx",
      "args": ["excel-mcp"]
    }
  }
}
```

**Option B — User settings** (`settings.json`):

```json
{
  "mcp.servers": {
    "excel": {
      "type": "stdio",
      "command": "uvx",
      "args": ["excel-mcp"]
    }
  }
}
```

After saving, run **MCP: Start Server** from the VS Code command palette.

> **Note on `pip install`**: If you installed via `pip` instead of `uvx`, replace `"uvx"` / `"args": ["excel-mcp"]` with the full path to the binary (e.g. `/usr/local/bin/excel-mcp`). Claude Desktop and VS Code launch the subprocess in a minimal environment and may not inherit your shell's `PATH`. Run `which excel-mcp` to find the path.

---

## Security: restricting file access

By default the server can access any file path and prints a warning at startup. To restrict it to specific directories, set the `EXCEL_MCP_ALLOWED_DIRS` environment variable:

```bash
export EXCEL_MCP_ALLOWED_DIRS="/home/alice/reports:/home/alice/exports"
```

In your MCP client config:

```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["excel-mcp"],
      "env": {
        "EXCEL_MCP_ALLOWED_DIRS": "/home/alice/reports:/home/alice/exports"
      }
    }
  }
}
```

Any attempt to read or write outside the listed directories will be rejected with a `PermissionError`.

---

## Tools

### Inspect

| Tool | Description |
|---|---|
| `list_sheets` | List all sheets in a workbook with row/column counts |
| `load_sheet` | Load a sheet into session cache with optional `skip_rows` for title rows above headers |
| `peek` | Read the first N raw rows directly from disk (bypasses cache) — use before `load_sheet` to inspect header layout |
| `get_schema` | Column names, data types, null counts, and a sample value per column |
| `describe` | Descriptive statistics (min, max, mean, std, percentiles) for numeric columns |
| `preview` | First N rows as a markdown table, with optional column filter and row offset |
| `find_dirty_rows` | Identify rows with nulls, type mismatches, or other data quality issues |

### Transform *(operations update the session cache)*

| Tool | Description |
|---|---|
| `filter_rows` | Filter rows using a SQL WHERE expression |
| `select_columns` | Keep and optionally rename a subset of columns |
| `drop_columns` | Remove specific columns |
| `sort_data` | Sort by one or more columns, with per-column direction control |
| `drop_duplicates` | Remove duplicate rows, optionally scoped to a column subset |
| `cast_columns` | Change column data types (e.g. `String` → `Int64`, `String` → `Date`) |
| `fill_nulls` | Fill null values using a strategy (`forward`, `mean`, `median`, …) or a literal value |
| `add_column` | Add a computed column using a SQL expression (e.g. `"Revenue * 0.1"`) |
| `reset_sheet` | Reload the original file from disk, discarding all in-session transformations |

### Aggregate

| Tool | Description |
|---|---|
| `group_by_aggregate` | Group by columns and compute `sum`, `mean`, `count`, `min`, `max`, `median`, `std`, `first`, `last`, `n_unique` |
| `value_counts` | Frequency table for a column, with optional normalization and top-N filter |
| `sample_rows` | Random sample by row count or fraction |
| `compute_percentiles` | Compute arbitrary percentiles (e.g. p50, p90, p95, p99) for a column, optionally grouped |

### Advanced

| Tool | Description |
|---|---|
| `join_sheets` | Join two sheets/files on key columns (`inner`, `left`, `right`, `full`, `cross`) |
| `pivot_data` | Pivot: unique values of one column become new column headers |
| `unpivot_data` | Melt wide-format data into long format |
| `concat_sheets` | Stack or side-by-side multiple sheets/files (`vertical`, `horizontal`, `diagonal`) |
| `window_function` | Apply a window/analytic function: `rank`, `cumsum`, `diff`, `pct_change`, `rolling_mean`, `rolling_sum`, … |
| `run_query` | Run a full SQL SELECT statement via Polars SQLContext (`data` is the primary table) |
| `save_sheet` | Write the current in-session state back to disk (Excel or CSV) |
| `clone_sheet` | Copy a sheet within the session under a new name |
| `export_data` | Export to `csv`, `xlsx`, `parquet`, or `json` |

---

## Why stdio (local-only)

excel-mcp runs as a local process on your machine, not as a hosted service. This is intentional.

The tools operate on file paths that point to **your filesystem**. A remote server would require uploading your files to third-party infrastructure before any query could run — spreadsheets that often contain payroll data, customer records, financial projections, or other sensitive information.

Running locally means:

- Your files never leave your machine
- No account, API key, or internet connection required
- No per-query cost or rate limits
- Works on air-gapped or restricted networks

The tradeoff is that users install the package themselves, but that is a one-time step (`pip install excel-mcp` or `uvx`) and is the right default for a tool whose entire job is reading your private data.

---



Transform and aggregate operations update an **in-memory cache** for the current server session. This lets you chain operations naturally:

```
load → filter rows → add column → group by → export
```

Each tool call operates on the result of the previous one. Call `reset_sheet` to discard changes and reload from disk.

---

## SQL reference (`filter_rows`, `run_query`, `add_column`)

These tools use [Polars SQLContext](https://docs.pola.rs/api/python/stable/reference/sql.html). Standard SQL syntax is supported:

```sql
-- filter_rows expression (no WHERE keyword)
"Country = 'US' AND Revenue > 1000"
"Status IN ('Active', 'Pending')"
"Name LIKE '%Corp%'"

-- run_query (full SELECT)
SELECT Region, SUM(Revenue) AS total
FROM data
GROUP BY Region
ORDER BY total DESC
LIMIT 10

-- add_column expression
"Revenue * 0.1"
"UPPER(Country)"
"Revenue - Cost"
```

---

## Development

```bash
git clone https://github.com/your-org/excel-mcp
cd excel-mcp
pip install -e .
excel-mcp  # starts the server on stdio
```

Test interactively with the MCP Inspector:

```bash
npx @modelcontextprotocol/inspector excel-mcp
```

---

## License

MIT

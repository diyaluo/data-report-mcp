# data-report-mcp

A local-first **Model Context Protocol (MCP)** server that turns a CSV/XLSX
table into a real data analysis report — summary statistics, duplicate /
missing diagnostics, top category frequencies, pairwise correlations, a
Chinese Markdown insights writeup, and an Excel workbook with **native
interactive charts**.

- 🔒 **100% local, no API key, no network.** Business / financial spreadsheets
  never leave the machine. There is no telemetry upload — only an optional
  local usage counter under `~/.cache/data-report-mcp`.
- 🧮 **No pandas.** Runs on the standard library + `openpyxl`, so it bundles
  small and starts fast inside an MCP host.
- 🔌 Two transports: **stdio** (Claude Desktop / Cursor / any MCP client) and
  **streamable-http** for a hosted endpoint.
- 🆓 Anonymous free tier (100 tables / 30 days per client); over-quota returns
  an in-band `QUOTA_EXCEEDED` signal, it never silently bills.

## Tools

| Tool | Purpose |
|------|---------|
| `analyze_table(file_base64 \| file_url, filename, charts=3, return_xlsx=false)` | Analyze a `.csv`/`.xlsx` fully locally; returns shape, column profiles, significant correlations, a Markdown report, and optionally a base64 `.xlsx` workbook with native charts. |
| `health()` | Service status + free-tier limits. |

`analyze_table` accepts the raw file as **base64** (preferred) or a direct
**https URL** to a `.csv`/`.xlsx`. Inputs larger than 30 MB are rejected.

## Install & run

```bash
pip install -e .                 # or: pip install data-report-mcp
# stdio (default MCP transport)
python -m data_report_mcp.server
# streamable-http
REMOTE_TRANSPORT=http MCP_PORT=5070 python -m data_report_mcp.server
```

### Claude Desktop / MCP client config

```json
{
  "mcpServers": {
    "data-report": {
      "command": "python",
      "args": ["-m", "data_report_mcp.server"]
    }
  }
}
```

### Environment variables

| Var | Default | Meaning |
|-----|---------|---------|
| `FREE_TABLES_PER_30D` | `100` | anonymous monthly table limit |
| `DATA_REPORT_API_KEY` | _(unset)_ | paid key that bypasses the quota |
| `DATA_REPORT_STATE_DIR` | `~/.cache/data-report-mcp` | local usage-counter directory |
| `REMOTE_TRANSPORT` / `MCP_HOST` / `MCP_PORT` | `stdio` / `127.0.0.1` / `5070` | http transport options |

## What the report contains

- Overall: row/column count, duplicate rows, columns with missing rate ≥ 20%.
- Numeric columns: min / max / mean / median / sum / stdev / missing rate.
- Categorical columns: unique count + top-5 value frequencies.
- Significant pairwise correlations (Pearson |r| ≥ 0.5).
- An Excel workbook (`return_xlsx=true`) with sheets 概览 / 数值统计 / 明细 /
  分布_* (each with a native bar chart) / 显著相关.

> Output is a starting point for analysis, not a conclusion: correlation is
> not causation — combine it with business judgment.

## License

MIT. Core analysis adapted from the open `data-report` local toolkit.

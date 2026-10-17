"""
Data Report MCP Server
======================
Model Context Protocol server that turns a CSV/XLSX table into a data
analysis report (statistics + native Excel charts) and a Chinese Markdown
insights summary. 100% local, no API key, no network — sensitive business
or financial spreadsheets never leave the machine.

Two transports:
  - stdio:  default for Claude Desktop / Cursor / local MCP clients
            (python -m data_report_mcp.server)
  - streamable-http: REMOTE_TRANSPORT=http for a hosted endpoint

Tools:
  analyze_table(file_base64 | file_url, filename, charts=3, return_md=true)
  health()

Input is accepted as base64 (preferred) or a direct URL to a .csv/.xlsx.
The analysis itself runs fully locally; output JSON carries row/column
metadata + Markdown insights, and optionally an xlsx workbook (base64).
"""
from __future__ import annotations

import base64
import io
import json
import os
import sys
import time
import uuid
import hashlib
import logging
import tempfile
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

# ---- state / config ----------------------------------------------------------
def _state_dir() -> Path:
    override = os.environ.get("DATA_REPORT_STATE_DIR")
    candidate = Path(override) if override else (Path.home() / ".cache" / "data-report-mcp")
    try:
        candidate.mkdir(parents=True, exist_ok=True)
        return candidate
    except Exception:
        fallback = Path(tempfile.gettempdir()) / "data-report-mcp"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback

STATE_DIR = _state_dir()
USAGE_LOG = STATE_DIR / "mcp_usage.jsonl"
QUOTA_FILE = STATE_DIR / "mcp_quota.json"

FREE_TABLES_PER_30D = int(os.environ.get("FREE_TABLES_PER_30D", "100"))
MAX_UPLOAD_BYTES = 30 * 1024 * 1024
FETCH_TIMEOUT_SEC = 30
ALLOWED_EXT = (".csv", ".xlsx")

UPGRADE_INFO = {
    "plan": "data-report-pro",
    "free_limit_tables_30d": FREE_TABLES_PER_30D,
    "message": (
        "Free quota used. Set a paid api_key (DATA_REPORT_API_KEY env) to "
        "bypass the anonymous monthly limit."
    ),
}

logging.basicConfig(
    level=os.environ.get("MCP_LOG_LEVEL", "INFO"),
    format="%(asctime)s [%(levelname)s] %(message)s",
    stream=sys.stderr,
)
log = logging.getLogger("data-report-mcp")

# ---- MCP SDK (1.x / 2.x compatible) -----------------------------------------
try:  # mcp >= 2.0
    from mcp.server.mcpserver import MCPServer  # noqa: E402
except ImportError:  # mcp 1.x
    from mcp.server.fastmcp import FastMCP as MCPServer  # noqa: E402

mcp = MCPServer("data-report")

# ---- bundled core ------------------------------------------------------------
try:
    from .report_core import read_table, analyze, to_excel, to_markdown  # noqa: E402
except Exception:  # script-style fallback
    try:
        from report_core import read_table, analyze, to_excel, to_markdown  # noqa: E402
    except Exception:
        read_table = analyze = to_excel = to_markdown = None


# ---- telemetry + quota -------------------------------------------------------
def _client_id(api_key: str | None) -> str:
    raw = api_key or f"anon-{uuid.getnode()}"
    return hashlib.sha256(raw.encode()).hexdigest()[:16]


def _record_usage(rec: dict) -> None:
    rec["ts"] = datetime.now(timezone.utc).isoformat()
    try:
        with USAGE_LOG.open("a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except Exception as e:
        log.warning("telemetry write failed: %s", e)


def tables_used_30d(client: str) -> int:
    cutoff = time.time() - 30 * 86400
    if not USAGE_LOG.exists():
        return 0
    total = 0
    with USAGE_LOG.open(encoding="utf-8") as f:
        for line in f:
            try:
                r = json.loads(line)
                if r.get("client") == client and r.get("ts_epoch", 0) >= cutoff:
                    total += int(r.get("tables", 0))
            except Exception:
                continue
    return total


# ---- input -------------------------------------------------------------------
def _load_bytes(file_base64: str | None, file_url: str | None) -> bytes:
    if file_base64:
        return base64.b64decode(file_base64)
    if file_url:
        req = urllib.request.Request(
            file_url,
            headers={"User-Agent": "data-report-mcp/1.0"},
        )
        with urllib.request.urlopen(req, timeout=FETCH_TIMEOUT_SEC) as resp:
            data = resp.read()
            if len(data) > MAX_UPLOAD_BYTES:
                raise ValueError("file too large (30 MB max)")
            return data
    raise ValueError("Provide either file_base64 or file_url")


# ---- tools -------------------------------------------------------------------
@mcp.tool()
def health() -> dict:
    """Service health + free-tier limits. No auth required."""
    return {
        "status": "ok",
        "service": "data-report-mcp",
        "free_tier": {"tables_per_30d": FREE_TABLES_PER_30D},
        "inputs": ["csv", "xlsx"],
        "local_only": True,
        "server_time_utc": datetime.now(timezone.utc).isoformat(),
    }


@mcp.tool()
def analyze_table(
    file_base64: str | None = None,
    file_url: str | None = None,
    filename: str = "upload.csv",
    charts: int = 3,
    return_xlsx: bool = False,
    api_key: str | None = None,
) -> dict:
    """Analyze a CSV/XLSX table fully locally: infer numeric/categorical
    columns, compute shape, missing rate, duplicates, min/max/mean/median/sum/
    stddev, top frequencies for text columns and pairwise correlations
    (|r|>=0.5), and return a Chinese Markdown insights report. Optionally a
    base64 Excel workbook with native interactive bar charts.

    Args:
        file_base64: raw CSV/XLSX content as base64 (preferred).
        file_url:    direct HTTPS URL to a .csv/.xlsx (alternative).
        filename:    original filename (used only to decide csv vs xlsx).
        charts:      max category columns to chart in the xlsx workbook.
        return_xlsx: when true, include the generated workbook as base64.
        api_key:     paid key to bypass the anonymous monthly quota.
    """
    started = time.time()
    client = _client_id(api_key or os.environ.get("MCP_CLIENT_ID"))
    is_paid = bool(api_key and api_key == os.environ.get("DATA_REPORT_API_KEY", ""))
    rec: dict = {"client": client, "ts_epoch": time.time(), "tool": "analyze_table"}

    if read_table is None or analyze is None:
        return {"ok": False, "error": "analysis deps unavailable on server"}

    name = (filename or "").lower().split("?")[0]
    if not name.endswith(ALLOWED_EXT):
        return {"ok": False, "error": "only .csv and .xlsx inputs are supported"}

    try:
        data = _load_bytes(file_base64, file_url)
    except Exception as e:
        rec.update(ok=False, tables=0, error=f"fetch:{e}")
        _record_usage(rec)
        return {"ok": False, "error": f"could not load file: {e}"}

    if len(data) > MAX_UPLOAD_BYTES:
        return {"ok": False, "error": "file too large (30 MB max)"}

    suffix = ".xlsx" if name.endswith(".xlsx") else ".csv"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(data)
        in_path = tmp.name
    out_xlsx = in_path + ".out.xlsx"

    try:
        if not is_paid:
            used = tables_used_30d(client)
            if used + 1 > FREE_TABLES_PER_30D:
                rec.update(ok=False, blocked="quota")
                _record_usage(rec)
                return {
                    "ok": False,
                    "error_code": "QUOTA_EXCEEDED",
                    "tables_used_30d": used,
                    "free_limit": FREE_TABLES_PER_30D,
                    "upgrade": UPGRADE_INFO,
                }

        rows = read_table(in_path)
        if not rows:
            return {"ok": True, "rows": 0, "columns": [], "markdown": "(空表，无数据行)"}

        a = analyze(rows)
        to_excel(a, out_xlsx, chart_n=max(0, int(charts)))
        md = to_markdown(a, "数据分析报告")

        col_profiles = a.get("columns", [])
        num_cols = [str(c.get("name")) for c in col_profiles if c.get("is_numeric")]
        cat_cols = [str(c.get("name")) for c in col_profiles if not c.get("is_numeric")]
        result = {
            "ok": True,
            "rows": int(a.get("rows", len(rows))),
            "columns": [str(h) for h in a.get("headers", [])],
            "format": "json+markdown",
            "markdown": md,
            "summary": {
                "n_rows": a.get("rows"),
                "n_cols": a.get("cols"),
                "dup_rows": a.get("dup_rows"),
                "numeric_cols": num_cols,
                "categorical_cols": cat_cols,
                "significant_correlations": [
                    {"a": x, "b": y, "r": round(float(r), 3)}
                    for x, y, r in (a.get("corr") or [])
                ],
            },
            "elapsed_ms": int((time.time() - started) * 1000),
        }
        if return_xlsx:
            buf = Path(out_xlsx).read_bytes()
            result["xlsx_base64"] = base64.b64encode(buf).decode()
            result["xlsx_bytes"] = len(buf)

        rec.update(ok=True, tables=1, rows=result["rows"])
        _record_usage(rec)
        return result
    finally:
        for p in (in_path, out_xlsx):
            try:
                os.unlink(p)
            except OSError:
                pass


# ---- entrypoint --------------------------------------------------------------
def main() -> None:
    transport = os.environ.get("REMOTE_TRANSPORT", "stdio").lower()
    if transport in ("http", "streamable-http", "sse"):
        host = os.environ.get("MCP_HOST", "127.0.0.1")
        port = int(os.environ.get("MCP_PORT", "5070"))
        log.info("data-report MCP over streamable-http on %s:%s", host, port)
        mcp.run(transport="streamable-http", host=host, port=port)
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()

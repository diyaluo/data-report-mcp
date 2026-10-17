"""End-to-end stdio acceptance test using the official MCP Client SDK.

Spawns the server over stdio, initializes, lists tools, runs analyze_table
on an in-memory CSV, and asserts the analysis round-trips. Run:

    python test_stdio.py
"""
import asyncio
import base64
import json
import os
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable

CSV = (
    "城市,品类,销售额,数量\n"
    "郑州,电子,1200.5,3\n"
    "郑州,电子,1200.5,3\n"
    "北京,服装,800,5\n"
    "上海,电子,2500,2\n"
    "北京,食品,,4\n"
).encode("utf-8")


def _content_to_obj(result):
    for block in result.content:
        if getattr(block, "type", None) == "text":
            return json.loads(block.text)
    raise AssertionError("no text content in tool result")


async def main():
    child_env = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
    child_env["PYTHONPATH"] = os.path.join(ROOT, "src")
    child_env["DATA_REPORT_STATE_DIR"] = "/tmp/dr-mcp-test-state"
    params = StdioServerParameters(
        command=PY,
        args=["-m", "data_report_mcp.server"],
        cwd=ROOT,
        env=child_env,
    )
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as s:
            init = await s.initialize()
            si = getattr(init, "server_info", None) or getattr(init, "serverInfo")
            print("server:", si.name, getattr(si,"version",""))
            assert si.name == "data-report", si

            tools = await s.list_tools()
            names = {t.name for t in tools.tools}
            print("tools:", sorted(names))
            assert {"health", "analyze_table"} <= names, names

            h = _content_to_obj(await s.call_tool("health", {}))
            print("health:", h["status"], "inputs", h["inputs"])
            assert h["status"] == "ok" and h["local_only"] is True

            res = _content_to_obj(await s.call_tool("analyze_table", {
                "file_base64": base64.b64encode(CSV).decode(),
                "filename": "sales.csv",
                "charts": 3,
                "return_xlsx": True,
            }))
            print("ok:", res["ok"], "rows:", res["rows"], "cols:", res["columns"])
            assert res["ok"] is True
            assert res["rows"] == 5, res["rows"]
            assert res["columns"] == ["城市", "品类", "销售额", "数量"]
            assert res["summary"]["dup_rows"] == 1, res["summary"]
            assert "销售额" in res["summary"]["numeric_cols"]
            assert "数据分析报告" in res["markdown"]
            xb = base64.b64decode(res["xlsx_base64"])
            print("xlsx bytes:", res["xlsx_bytes"], "decoded:", len(xb))
            assert xb[:2] == b"PK", "not a real xlsx (zip)"
            assert len(xb) > 3000

            # error path: unsupported extension
            bad = _content_to_obj(await s.call_tool("analyze_table", {
                "file_base64": base64.b64encode(b"hello").decode(),
                "filename": "notes.txt",
            }))
            print("bad ext ok=False:", not bad["ok"], "-", bad.get("error", "")[:40])
            assert bad["ok"] is False

            # missing input
            none = _content_to_obj(await s.call_tool("analyze_table", {}))
            assert none["ok"] is False

    print("\nALL ASSERTIONS PASSED")


if __name__ == "__main__":
    asyncio.run(main())

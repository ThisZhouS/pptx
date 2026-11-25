"""MCP 协议端到端测试：以 stdio 子进程启动 server，走真实 JSON-RPC。

覆盖 initialize → tools/list → tools/call（catalog / build / validate）。
"""
import asyncio
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples", "stdio-demo.pptx")

PARAMS = StdioServerParameters(
    command=sys.executable,
    args=["-m", "mcp_ppt.server"],
    cwd=ROOT,
    env={**os.environ, "PYTHONIOENCODING": "utf-8"},
)


async def main() -> int:
    async with stdio_client(PARAMS) as (read, write):
        async with ClientSession(read, write) as session:
            init = await session.initialize()
            print("[init] server:", init.serverInfo.name, init.serverInfo.version)

            tools = await session.list_tools()
            print(f"[tools] {len(tools.tools)} 个：")
            for t in tools.tools:
                print("   -", t.name, "|", (t.description or "").splitlines()[0][:52])

            res = await session.call_tool("capability_report", {})
            print("[capability]", res.content[0].text if res.content else None)

            res = await session.call_tool("catalog_transitions", {"query": "变形"})
            print("[catalog 变形]", res.content[0].text[:300])

            model = {
                "theme": "tech" if False else "dark",
                "footer": "stdio test",
                "defaults": {"transition": {"effect": "fade", "duration": 0.6}},
                "slides": [
                    {"layout": "title", "title": "MCP 端到端",
                     "subtitle": "stdio 传输 · 生成 · 校验",
                     "transition": {"effect": "morph", "duration": 1.0}},
                    {"layout": "bullets", "kicker": "验证", "title": "一切走协议",
                     "bullets": ["tools/list", "tools/call", "原生 pptx 输出"],
                     "animations": [
                         {"target": "title", "effect": "fade", "trigger": "click"},
                         {"target": "bullets", "effect": "wipe", "trigger": "after_previous",
                          "build": "paragraph"}]},
                ],
            }
            res = await session.call_tool("build_deck", {"model": model, "out_path": OUT,
                                                         "motion_style": "tech"})
            payload = json.loads(res.content[0].text)
            print("[build] slides:", payload.get("slide_count"),
                  "warnings:", payload.get("warnings"),
                  "validate:", payload.get("validate"))

            res = await session.call_tool("validate", {"path": OUT})
            print("[validate]", res.content[0].text[:400])

            res = await session.call_tool("inspect", {"path": OUT})
            data = json.loads(res.content[0].text)
            for s in data["slides"]:
                tr = s["transition"] or {}
                print(f"   slide {s['index']}: transition={tr.get('effects')} "
                      f"animations={[(a['presetClass'], a['nodeType']) for a in s['animations']]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
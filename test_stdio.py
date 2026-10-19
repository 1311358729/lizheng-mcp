# -*- coding: utf-8 -*-
"""MCP 协议级握手测试：stdio 拉起 lizheng-mcp，list_tools + 实调 4 个工具。

用法：
    python test_stdio.py <路径/到/工程.lz> [输出目录]

不带参数时只跑不依赖工程库的部分（backend_info / get_spec / 规范加载）。
"""
import asyncio
import json
import os
import shutil
import sys
import tempfile

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
LZ = sys.argv[1] if len(sys.argv) > 1 else ""
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(tempfile.gettempdir(), "lz_mcp_test")
if LZ:
    shutil.rmtree(OUT, ignore_errors=True)


async def main():
    params = StdioServerParameters(
        command=sys.executable,
        args=[os.path.join(HERE, "server.py")],
    )
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools = await session.list_tools()
            print("工具数:", len(tools.tools))
            for t in tools.tools:
                print("  -", t.name)

            r = await session.call_tool("lz_backend_info", {})
            d = json.loads(r.content[0].text)
            print()
            print("backend_info: pyodbc=%s ace=%s 规范节=%s" % (
                d["pyodbc_installed"], d["ace_access_driver"], d["spec_sections"]))

            r = await session.call_tool("lz_get_spec", {"section": "ZK"})
            d = json.loads(r.content[0].text)
            print("get_spec(#ZK): 字段=%d 必填=%s" % (len(d["fields"]), d["required"]))

            if not LZ:
                print()
                print("(未提供 .lz 路径，跳过工程库相关调用)")
                print("用法: python test_stdio.py <工程.lz> [输出目录]")
                return

            r = await session.call_tool("lz_project_summary", {"lz_path": LZ})
            d = json.loads(r.content[0].text)
            print()
            print("project_summary: 工程=%s %s" % (d["project"].get("GCBH"), d["project"].get("GCMC")))
            print("  钻孔数:", d["borehole_count"], "| 有数据的节:", list(d["sections_with_data"].keys()))
            print("  首孔:", json.dumps(d["boreholes"][0], ensure_ascii=False)[:150])

            r = await session.call_tool("lz_db_to_interface", {
                "lz_path": LZ, "output_dir": OUT, "boreholes": ["ZK1"]})
            d = json.loads(r.content[0].text)
            print()
            print("db_to_interface: 文件数=%d" % d["file_count"])
            for f in d["files"]:
                print("  ->", f["output_path"], "| 节:", f["sections"], "| 记录:", f["records"])

            r = await session.call_tool("lz_validate_interface_file", {
                "path": d["files"][0]["output_path"]})
            v = json.loads(r.content[0].text)
            print()
            print("validate: valid=%s errors=%d warnings=%d" % (v["valid"], len(v["errors"]), len(v["warnings"])))

asyncio.run(main())

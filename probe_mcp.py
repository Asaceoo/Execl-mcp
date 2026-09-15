"""MCP stdio probe: verify the deployed server advertises the slicer-linkage surface.

Sends initialize + tools/list over stdio JSON-RPC and asserts:
  - the `slicer` tool's action enum contains connect-pivots / disconnect-pivots
  - the `slicer` tool's schema declares `pivot_table_names` and `slicer_name`
  - the `pivottable` tool's schema declares `share_cache_from`
  - the advertised tool count
"""
import json
import os
import subprocess
import sys

# Derived from this file's location so the probe runs in any checkout: release.sh deploys the
# server to excel-mcp-bin/ beside this script. Override with EXCEL_MCP_BIN / EXCEL_MCP_DOTNET.
ROOT = os.path.dirname(os.path.abspath(__file__))
BIN = os.environ.get("EXCEL_MCP_BIN") or os.path.join(ROOT, "excel-mcp-bin")
EXE = os.path.join(BIN, "Sbroenne.ExcelMcp.McpServer.exe")
ENV = dict(os.environ)
ENV["DOTNET_ROOT"] = os.environ.get("EXCEL_MCP_DOTNET") or os.path.join(ROOT, ".dotnet10")
ENV["DOTNET_CLI_TELEMETRY_OPTOUT"] = "1"


def send(proc, payload):
    proc.stdin.write((json.dumps(payload) + "\n").encode())
    proc.stdin.flush()


def read_response(proc, want_id, timeout_lines=4000):
    for _ in range(timeout_lines):
        line = proc.stdout.readline()
        if not line:
            return None
        text = line.decode("utf-8", "replace").strip()
        if not text.startswith("{"):
            continue
        try:
            msg = json.loads(text)
        except json.JSONDecodeError:
            continue
        if msg.get("id") == want_id:
            return msg
    return None


proc = subprocess.Popen(
    [EXE],
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    stderr=subprocess.DEVNULL,
    env=ENV,
)

try:
    send(proc, {
        "jsonrpc": "2.0", "id": 1, "method": "initialize",
        "params": {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "probe", "version": "1.0"},
        },
    })
    init = read_response(proc, 1)
    if init is None:
        print("FAIL: no initialize response")
        sys.exit(1)

    server = init.get("result", {}).get("serverInfo", {})
    print(f"serverInfo: name={server.get('name')} version={server.get('version')}")

    send(proc, {"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}})
    send(proc, {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}})
    tools_msg = read_response(proc, 2)
    if tools_msg is None:
        print("FAIL: no tools/list response")
        sys.exit(1)

    tools = tools_msg.get("result", {}).get("tools", [])
    print(f"tools advertised: {len(tools)}")

    by_name = {t["name"]: t for t in tools}
    fails = []

    slicer = by_name.get("slicer")
    if slicer is None:
        fails.append("slicer tool missing")
    else:
        blob = json.dumps(slicer)
        for action in ("connect-pivots", "disconnect-pivots"):
            if action in blob:
                print(f"  OK  slicer action: {action}")
            else:
                fails.append(f"slicer action missing: {action}")
        # The generated schema uses snake_case names, so assert on the actual property keys
        # instead of substring-matching the serialized tool. Missing keys must FAIL the probe,
        # not merely print a MISS - a probe that reports PASS while printing MISS is worthless.
        props = slicer.get("inputSchema", {}).get("properties", {})
        for prop in ("pivot_table_names", "slicer_name"):
            if prop in props:
                print(f"  OK  slicer param: {prop}")
            else:
                fails.append(f"slicer param missing: {prop}")

    pivottable = by_name.get("pivottable")
    if pivottable is None:
        fails.append("pivottable tool missing")
    else:
        props = pivottable.get("inputSchema", {}).get("properties", {})
        if "share_cache_from" in props:
            print("  OK  pivottable param: share_cache_from")
        else:
            fails.append("pivottable param missing: share_cache_from")

    print()
    if fails:
        print("FAIL:")
        for f in fails:
            print("  -", f)
        sys.exit(1)
    print("PASS - deployed server advertises the slicer-linkage surface")
finally:
    proc.terminate()

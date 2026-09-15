"""Reproduce F5-9: real parameter-validation errors swallowed into a generic message.

Prints the raw tools/call result so the exact shape (isError content vs JSON-RPC error)
is visible, instead of the parsed-and-rethrown text the other probes use.
"""

from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from demo_slicer_link import Server  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "_demo" / f"f59-repro-{datetime.now().strftime('%H%M%S')}.xlsx"


def show(label: str, server: Server, tool: str, arguments: dict) -> None:
    message = server.request("tools/call", {"name": tool, "arguments": arguments})
    print(f"--- {label}")
    if "error" in message:
        print(f"    JSON-RPC error: {json.dumps(message['error'], ensure_ascii=False)[:300]}")
    result = message.get("result", {})
    print(f"    isError={result.get('isError')}")
    for part in result.get("content", []):
        if part.get("type") == "text":
            print(f"    text: {part.get('text', '')[:400]}")


def main() -> int:
    server = Server()
    server.start()
    try:
        session = server.request("tools/call", {
            "name": "file", "arguments": {"action": "create", "path": str(OUTPUT), "show": False},
        })["result"]["content"][0]["text"]
        session_id = json.loads(session)["session_id"]
        print(f"session {session_id}")

        show("1. 未知参数名 (bogus_param)", server, "range_edit", {
            "action": "set-values", "session_id": session_id, "sheet_name": "Sheet1",
            "range_address": "A1", "values": [[1]], "bogus_param": "x",
        })
        show("2. 缺必填参数 (range_address)", server, "range_edit", {
            "action": "set-values", "session_id": session_id, "sheet_name": "Sheet1",
            "values": [[1]],
        })
        show("3. action 不在枚举内", server, "range_edit", {
            "action": "definitely-not-an-action", "session_id": session_id,
        })
        show("4. range_format 未知参数", server, "range_format", {
            "action": "format-range", "session_id": session_id, "sheet_name": "Sheet1",
            "range_address": "A1", "nope": 1,
        })
        show("5. 参数类型错误 (values 传字符串)", server, "range_edit", {
            "action": "set-values", "session_id": session_id, "sheet_name": "Sheet1",
            "range_address": "A1", "values": "not-an-array",
        })
        show("6. 工具内部真实错误 (非法 sheet)", server, "range_edit", {
            "action": "set-values", "session_id": session_id, "sheet_name": "NoSuchSheet",
            "range_address": "A1", "values": [[1]],
        })
    finally:
        server.proc.terminate()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

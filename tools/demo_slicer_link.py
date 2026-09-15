#!/usr/bin/env python
"""End-to-end proof that one slicer drives several PivotTables on the deployed server.

Drives the real Excel COM through the MCP stdio interface - no mocks, no fixtures:

    create workbook -> write source data -> PivotTable A
                                        -> PivotTable B (share_cache_from=A)
                                        -> one slicer on Region
                                        -> connect-pivots A,B
                                        -> set-slicer-selection and read both tables back

Every step asserts on the response instead of printing and hoping: a step that reports
Success=false, or a slicer that ends up connected to fewer tables than requested, fails
the run with a non-zero exit code.

Usage:  python tools/demo_slicer_link.py [--show] [output.xlsx]

--show  runs Excel visibly and leaves it in the foreground for a while so the slicer can be
        clicked by hand; without it Excel stays hidden (the default for background automation,
        which is what MCP clients normally want).
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
EXE = ROOT / "excel-mcp-bin" / "Sbroenne.ExcelMcp.McpServer.exe"
DEFAULT_OUTPUT = ROOT / "_demo" / "slicer-link-demo.xlsx"
# Override with DEMO_WATCH_SECONDS when you just want a quick verification run.
WATCH_SECONDS = int(os.environ.get("DEMO_WATCH_SECONDS", "180"))

DATA = [
    ["Region", "Product", "Sales"],
    ["华东", "A", 100],
    ["华东", "B", 150],
    ["华南", "A", 200],
    ["华南", "B", 250],
    ["华北", "A", 300],
    ["华北", "B", 350],
]


class DemoFailure(RuntimeError):
    """Raised when a step does not produce the outcome the demo promises."""


class Server:
    def __init__(self) -> None:
        env = dict(os.environ)
        env["DOTNET_ROOT"] = str(ROOT / ".dotnet10")
        self.proc = subprocess.Popen(
            [str(EXE)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            env=env,
        )
        self.next_id = 1

    def request(self, method: str, params: dict) -> dict:
        ident = self.next_id
        self.next_id += 1
        self.proc.stdin.write(
            (json.dumps({"jsonrpc": "2.0", "id": ident, "method": method, "params": params}) + "\n").encode()
        )
        self.proc.stdin.flush()
        for _ in range(20000):
            line = self.proc.stdout.readline()
            if not line:
                raise DemoFailure(f"server closed the stream while waiting for {method}")
            text = line.decode("utf-8", "replace").strip()
            if not text.startswith("{"):
                continue
            try:
                message = json.loads(text)
            except json.JSONDecodeError:
                continue
            if message.get("id") == ident:
                return message
        raise DemoFailure(f"no response to {method}")

    def start(self) -> None:
        self.request(
            "initialize",
            {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "demo-slicer-link", "version": "1"},
            },
        )
        self.proc.stdin.write(
            (json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}}) + "\n").encode()
        )
        self.proc.stdin.flush()

    def call(self, tool: str, arguments: dict) -> dict:
        result = self.request("tools/call", {"name": tool, "arguments": arguments})
        if "error" in result:
            raise DemoFailure(f"{tool} returned an error: {result['error']}")
        content = result.get("result", {}).get("content", [])
        payload = "".join(part.get("text", "") for part in content)
        # Image content blocks (e.g. `screenshot`) are not text: keep them so callers can
        # save the bytes. Without this they vanish silently and the tool looks broken.
        images = [part for part in content if part.get("type") == "image"]
        try:
            parsed = json.loads(payload)
        except json.JSONDecodeError:
            parsed = {"raw": payload}
        if images:
            parsed["_images"] = images
        if result.get("result", {}).get("isError"):
            raise DemoFailure(f"{tool} reported an error: {payload[:400]}")
        # The tool's own payload also carries success/error; the MCP-level isError flag alone is
        # not enough, because COM failures come back as a normal result whose body says otherwise.
        if isinstance(parsed, dict) and parsed.get("success") is False:
            detail = parsed.get("errorMessage") or parsed.get("error") or payload
            raise DemoFailure(f"{tool} returned success=false: {str(detail)[:400]}")
        return parsed

    def stop(self) -> None:
        self.proc.terminate()


def step(name: str) -> None:
    print(f"\n-- {name}")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise DemoFailure(message)


def main(argv: list[str]) -> int:
    args = argv[1:]
    show = "--show" in args
    if show:
        args.remove("--show")
    output = Path(args[0]) if args else DEFAULT_OUTPUT
    output.parent.mkdir(parents=True, exist_ok=True)
    # Never delete an existing file: this sandbox refuses removal (safe-delete fails closed),
    # so the demo writes to a fresh timestamped name instead of cleaning up first.
    if output.exists():
        stamp = datetime.now().strftime("%H%M%S")
        output = output.with_name(f"{output.stem}-{stamp}{output.suffix}")

    server = Server()
    try:
        server.start()

        step("create workbook")
        created = server.call("file", {"action": "create", "path": str(output), "show": show})
        session = created.get("session_id") or created.get("sessionId")
        require(bool(session), f"no session id in response: {json.dumps(created)[:300]}")
        print(f"   session: {session}")

        step("write source data")
        server.call(
            "range",
            {
                "action": "set-values",
                "session_id": session,
                "sheet_name": "Sheet1",
                "range_address": "A1:C7",
                "values": DATA,
            },
        )

        # Destination sheets have to exist up front: the lookup is Worksheets[name], and a missing
        # name surfaces as COMException 0x8002000B (DISP_E_BADINDEX), not as a friendly error.
        for sheet in ("Sheet2", "Sheet3"):
            server.call("worksheet", {"action": "create", "session_id": session, "sheet_name": sheet})

        step("PivotTable A (own cache)")
        server.call(
            "pivottable",
            {
                "action": "create-from-range",
                "session_id": session,
                "pivot_table_name": "PT_A",
                "source_sheet": "Sheet1",
                "source_range": "A1:C7",
                "destination_sheet": "Sheet2",
                "destination_cell": "A1",
            },
        )

        step("PivotTable B (share_cache_from = PT_A)")
        shared = server.call(
            "pivottable",
            {
                "action": "create-from-range",
                "session_id": session,
                "pivot_table_name": "PT_B",
                "source_sheet": "Sheet1",
                "source_range": "A1:C7",
                "destination_sheet": "Sheet3",
                "destination_cell": "A1",
                "share_cache_from": "PT_A",
            },
        )
        print(f"   {str(shared.get('workflowHint') or shared.get('message') or '')[:160]}")

        for table in ("PT_A", "PT_B"):
            server.call(
                "pivottable_field",
                {"action": "add-row-field", "session_id": session, "pivot_table_name": table, "field_name": "Region"},
            )
            server.call(
                "pivottable_field",
                {"action": "add-value-field", "session_id": session, "pivot_table_name": table, "field_name": "Sales"},
            )

        step("create one slicer on Region")
        server.call(
            "slicer",
            {
                "action": "create-slicer",
                "session_id": session,
                "pivot_table_name": "PT_A",
                "field_name": "Region",
                "slicer_name": "RegionSlicer",
                "destination_sheet": "Sheet2",
                "position": "E1",
            },
        )

        step("connect-pivots PT_A,PT_B")
        connected = server.call(
            "slicer",
            {
                "action": "connect-pivots",
                "session_id": session,
                "slicer_name": "RegionSlicer",
                # plural parameters take a JSON array as a string, not a comma-separated list
                "pivot_table_names": '["PT_A","PT_B"]',
            },
        )
        print(f"   {str(connected.get('message') or '')[:300]}")

        step("verify connection list")
        listed = server.call("slicer", {"action": "list-slicers", "session_id": session})
        slicers = listed.get("slicers") or listed.get("items") or []
        target = next((s for s in slicers if str(s.get("name", "")).lower() == "regionslicer"), None)
        require(target is not None, f"RegionSlicer not found in list-slicers: {json.dumps(listed)[:400]}")
        linked = target.get("pivotTables") or target.get("pivot_tables") or target.get("connectedPivotTables") or []
        print(f"   connected PivotTables: {linked}")
        require(len(linked) >= 2, f"slicer is linked to {len(linked)} table(s), expected at least 2")

        step("select one region and read both PivotTables")
        server.call(
            "slicer",
            {
                "action": "set-slicer-selection",
                "session_id": session,
                "slicer_name": "RegionSlicer",
                "selected_items": '["华东"]',
                "clear_first": True,
            },
        )
        for sheet in ("Sheet2", "Sheet3"):
            values = server.call(
                "range",
                {
                    "action": "get-values",
                    "session_id": session,
                    "sheet_name": sheet,
                    "range_address": "A1:B10",
                },
            )
            rows = values.get("values") or []
            flat = [str(cell) for row in rows for cell in row]
            print(f"   {sheet}: {[r for r in rows if r][:6]}")
            require(any("华东" in cell for cell in flat), f"{sheet} does not show the selected region")
            require(
                not any("华北" in cell for cell in flat),
                f"{sheet} still shows a filtered-out region - the slicer did not filter it",
            )

        # Persist BEFORE any interactive time. If Excel dies while it is on screen (the user
        # closes the window, or COM drops), a close(save=True) afterwards loses everything that
        # was built - the first --show run ended with an empty one-sheet shell on disk.
        saved = False
        if show:
            step("save workbook before opening the window")
            try:
                server.call(
                    "workbook",
                    {
                        "action": "save-as",
                        "session_id": session,
                        "target_path": str(output),
                        "overwrite": True,
                    },
                )
                saved = True
            except DemoFailure as error:
                print(f"   early save failed ({error}); relying on close(save=True)")

        if show:
            step("bring Excel to the foreground")
            server.call("window", {"action": "show", "session_id": session})
            server.call("window", {"action": "bring-to-front", "session_id": session})
            print(f"   Excel is visible for {WATCH_SECONDS}s - click the Region slicer on Sheet2.")
            try:
                time.sleep(WATCH_SECONDS)
            except KeyboardInterrupt:
                print("   interrupted, wrapping up")

        step("close and save")
        try:
            server.call("file", {"action": "close", "session_id": session, "save": True})
        except DemoFailure as error:
            if not saved:
                raise
            # The copy written before the window is intact, so the demo still produced a usable
            # workbook - only the edits made while Excel was on screen are missing.
            print(f"\nWARNING: final save failed ({error})", file=sys.stderr)
            print(
                "The workbook saved before the interactive window is intact; "
                "only edits made on screen were lost.",
                file=sys.stderr,
            )
        print(f"\nOK - one slicer filtered both PivotTables. Workbook saved to:\n  {output}")
        return 0
    except DemoFailure as error:
        print(f"\nFAILED: {error}", file=sys.stderr)
        return 1
    finally:
        server.stop()


if __name__ == "__main__":
    sys.exit(main(sys.argv))

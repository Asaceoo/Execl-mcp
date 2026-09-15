"""Perspective 3 audit: dump the whole MCP surface and report usability defects.

Answers, with evidence rather than impressions:
  1. How many tools / actions does the deployed server really advertise?
  2. Are parameter names uniform (snake_case) or mixed?  (an LLM guessing `sheetName`
     where the schema says `sheet_name` produces an opaque failure)
  3. Does every tool carry a description, and does every parameter carry one?
  4. How much context does the schema cost before the model even starts working?
  5. Are there actions advertised by a description but missing from the enum (or vice versa)?
  6. Do error paths carry actionable text, or a generic sentence?

Usage:
    python tools/audit_schema.py                 # human report
    python tools/audit_schema.py --json out.json # machine-readable dump
    python tools/audit_schema.py --full          # include per-tool parameter tables
"""
from __future__ import annotations

import argparse
import collections
import json
import os
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
EXE = ROOT / "excel-mcp-bin" / "Sbroenne.ExcelMcp.McpServer.exe"

SNAKE_RE = re.compile(r"^[a-z0-9]+(_[a-z0-9]+)*$")
CAMEL_RE = re.compile(r"^[a-z]+[A-Za-z0-9]*$")


class McpClient:
    """Minimal stdio JSON-RPC client (one line per message)."""

    def __init__(self, exe: Path) -> None:
        env = dict(os.environ)
        env["DOTNET_ROOT"] = str(ROOT / ".dotnet10")
        env["DOTNET_CLI_TELEMETRY_OPTOUT"] = "1"
        self.proc = subprocess.Popen(
            [str(exe)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=env,
        )
        self._id = 0

    def _send(self, payload: dict) -> None:
        assert self.proc.stdin is not None
        self.proc.stdin.write((json.dumps(payload) + "\n").encode())
        self.proc.stdin.flush()

    def _read(self, want_id: int | None, limit: int = 20000):
        assert self.proc.stdout is not None
        for _ in range(limit):
            line = self.proc.stdout.readline()
            if not line:
                return None
            text = line.decode("utf-8", "replace").strip()
            if not text.startswith("{"):
                continue
            try:
                msg = json.loads(text)
            except json.JSONDecodeError:
                continue
            if want_id is None or msg.get("id") == want_id:
                return msg
        return None

    def request(self, method: str, params: dict | None = None):
        self._id += 1
        self._send({"jsonrpc": "2.0", "id": self._id, "method": method,
                    "params": params or {}})
        return self._read(self._id)

    def initialize(self) -> dict:
        reply = self.request("initialize", {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "audit-schema", "version": "1.0"},
        })
        self._send({"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}})
        return reply

    def tools(self) -> list[dict]:
        reply = self.request("tools/list")
        return (reply or {}).get("result", {}).get("tools", [])

    def call(self, tool: str, args: dict) -> dict:
        self._id += 1
        self._send({"jsonrpc": "2.0", "id": self._id, "method": "tools/call",
                    "params": {"name": tool, "arguments": args}})
        return self._read(self._id) or {}

    def close(self) -> None:
        try:
            self.proc.terminate()
        except Exception:  # noqa: BLE001
            pass


def text_of(reply: dict) -> str:
    parts = (reply.get("result") or {}).get("content") or []
    return "".join(str(p.get("text", "")) for p in parts if p.get("type") == "text")


FINDINGS: list[tuple[str, str, str]] = []


def finding(severity: str, code: str, detail: str) -> None:
    FINDINGS.append((severity, code, detail))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", help="write the raw schema dump here")
    ap.add_argument("--full", action="store_true", help="print per-tool parameter tables")
    args = ap.parse_args()

    if not EXE.exists():
        print(f"FATAL: deployed server not found: {EXE}")
        print("       run `bash release.sh --no-bump --no-build` first")
        return 2

    client = McpClient(EXE)
    try:
        init = client.initialize()
        server = ((init or {}).get("result") or {}).get("serverInfo") or {}
        tools = client.tools()

        print("=" * 78)
        print("MCP SURFACE AUDIT")
        print("=" * 78)
        print(f"serverInfo.name    : {server.get('name')}")
        print(f"serverInfo.version : {server.get('version')}")
        print(f"tool count         : {len(tools)}")

        schema_bytes = len(json.dumps({"tools": tools}, ensure_ascii=False))
        print(f"schema size        : {schema_bytes:,} bytes (~{schema_bytes // 4:,} tokens)")
        if schema_bytes > 120_000:
            finding("MED", "SCHEMA-SIZE",
                    f"tools/list is {schema_bytes:,} bytes; every client pays this "
                    f"before the first call. Look for fat descriptions to trim.")

        if args.json:
            Path(args.json).write_text(
                json.dumps({"serverInfo": server, "tools": tools},
                           ensure_ascii=False, indent=2),
                encoding="utf-8")
            print(f"raw dump           : {args.json}")

        # ---- per-tool walk -------------------------------------------------
        action_names: dict[str, list[str]] = {}
        total_actions = 0
        naming = collections.Counter()
        no_tool_desc: list[str] = []
        no_param_desc: list[str] = []
        camel_params: list[tuple[str, str]] = []
        tool_desc_len: list[tuple[int, str]] = []
        per_tool_actions: list[tuple[str, int]] = []

        for tool in tools:
            name = tool["name"]
            desc = (tool.get("description") or "").strip()
            if not desc:
                no_tool_desc.append(name)
            tool_desc_len.append((len(desc), name))

            schema = tool.get("inputSchema") or {}
            props = schema.get("properties") or {}
            for pname, pspec in props.items():
                if SNAKE_RE.match(pname):
                    naming["snake"] += 1
                elif CAMEL_RE.match(pname):
                    naming["camel"] += 1
                    camel_params.append((name, pname))
                else:
                    naming["other"] += 1
                if not (pspec.get("description") or "").strip():
                    no_param_desc.append(f"{name}.{pname}")

            action = props.get("action") or {}
            enum = action.get("enum") or []
            action_names[name] = list(enum)
            total_actions += len(enum)
            per_tool_actions.append((name, len(enum)))

        print(f"declared actions   : {total_actions}")
        print(f"param naming       : {dict(naming)}")

        # ---- checks --------------------------------------------------------
        if camel_params:
            sample = ", ".join(f"{t}.{p}" for t, p in camel_params[:8])
            finding("HIGH", "PARAM-CASING",
                    f"{len(camel_params)} parameter(s) are camelCase while the rest are "
                    f"snake_case -> {sample}. An LLM that guesses the wrong case gets an "
                    f"opaque binding failure.")
        if no_tool_desc:
            finding("HIGH", "TOOL-NO-DESC", f"{len(no_tool_desc)} tool(s) without description: "
                                            f"{', '.join(no_tool_desc)}")
        if no_param_desc:
            finding("MED", "PARAM-NO-DESC",
                    f"{len(no_param_desc)} parameter(s) without description, e.g. "
                    f"{', '.join(no_param_desc[:10])}")

        # actions that appear in the free-text description but not in the enum
        for tool in tools:
            name = tool["name"]
            desc = (tool.get("description") or "")
            enum = set(action_names.get(name) or [])
            quoted = set(re.findall(r"`([a-z][a-z0-9-]{2,})`", desc))
            # only consider tokens that look like actions of this tool
            suspicious = {q for q in quoted if "-" in q or q in {
                "create", "delete", "list", "read", "update", "set", "get",
            }}
            missing = suspicious - enum
            if missing and enum:
                finding("LOW", "DESC-ENUM-DRIFT",
                        f"`{name}` description mentions {sorted(missing)} which is not in the "
                        f"action enum {sorted(enum)[:6]}...")

        # ---- actionability of a deliberately broken call -------------------
        print()
        print("-" * 78)
        print("ERROR-ACTIONABILITY PROBES (a good message tells the caller how to fix it)")
        print("-" * 78)
        probes = [
            ("missing required action", "worksheet", {"session_id": "nope"}),
            ("bad enum value", "worksheet", {"action": "no-such-action", "session_id": "x"}),
            ("bad param name", "worksheet", {"action": "list", "sessionId": "x"}),
            ("unknown tool", "__no_such_tool__", {}),
        ]
        for label, tool, payload in probes:
            reply = client.call(tool, payload)
            body = text_of(reply)
            if not body:
                body = json.dumps(reply.get("error") or reply)[:300]
            body = body.strip()
            # score: does it name the offending parameter / list valid values?
            actionable = bool(re.search(
                r"no-such-action|合法|valid|allowed|expected|enum|"
                r"rejected the arguments|unknown|not found|not recognised|not recognized",
                body, re.I))
            print(f"  [{'OK ' if actionable else 'WEAK'}] {label}")
            print(f"         -> {body.splitlines()[0][:150] if body else '(empty)'}")
            if not actionable:
                finding("HIGH", "ERROR-NOT-ACTIONABLE",
                        f"{label}: server replied without naming the problem or the valid "
                        f"set -> {(body or '(empty)')[:120]}")

        # ---- report --------------------------------------------------------
        print()
        print("=" * 78)
        print("FINDINGS")
        print("=" * 78)
        if not FINDINGS:
            print("  none")
        for severity, code, detail in sorted(FINDINGS, key=lambda f: "HIGH MED LOW".index(f[0])):
            print(f"  [{severity:4s}] {code}: {detail}")

        if args.full:
            print()
            print("=" * 78)
            print("PER-TOOL TABLE")
            print("=" * 78)
            for name, count in sorted(per_tool_actions, key=lambda x: -x[1]):
                print(f"  {name:<22} {count:3d} actions")
            print()
            print("FATTEST DESCRIPTIONS")
            for length, name in sorted(tool_desc_len, reverse=True)[:10]:
                print(f"  {name:<22} {length:6,d} chars")

        return 1 if any(s == "HIGH" for s, _, _ in FINDINGS) else 0
    finally:
        client.close()


if __name__ == "__main__":
    sys.exit(main())

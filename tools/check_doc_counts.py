#!/usr/bin/env python
"""Faithful Python replica of mcp-server-excel/scripts/check-doc-counts.ps1.

The sandbox cannot run the PowerShell script (its output never reaches stdout here), so this
mirrors its logic step for step: derive the canonical counts from the generated skill manifest,
then assert every headline claim in the user-facing docs matches. Run from the repo root.

  python tools/check_doc_counts.py
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent / "mcp-server-excel"
MANIFEST = (
    REPO
    / "src/ExcelMcp.Core/obj/GeneratedFiles/ExcelMcp.Generators"
    / "Sbroenne.ExcelMcp.Generators.ServiceRegistryGenerator/_SkillManifest.g.cs"
)
TOOL_ACTIONS = REPO / "src/ExcelMcp.Core/Models/Actions/ToolActions.cs"
MCP_SERVER_DIR = REPO / "src/ExcelMcp.McpServer"
PROGRAM = REPO / "src/ExcelMcp.McpServer/Program.cs"

# file + regex with optional named groups t (tools) and o (operations)
CHECKS = (
    ("README.md", r"(?P<t>\d+) tools with (?P<o>\d+) operations"),
    ("README.md", r"all (?P<o>\d+) operations"),
    ("FEATURES.md", r"(?P<t>\d+) specialized tools with (?P<o>\d+) operations"),
    ("src/ExcelMcp.McpServer/README.md", r"(?P<t>\d+) specialized tools with (?P<o>\d+) operations"),
    ("src/ExcelMcp.McpServer/README.md", r"all (?P<o>\d+) operations"),
    ("src/ExcelMcp.CLI/README.md", r"with (?P<o>\d+) operations matching"),
    ("src/ExcelMcp.CLI/README.md", r"\*\*(?P<o>\d+) operations\*\* across"),
    ("vscode-extension/README.md", r"(?P<t>\d+) specialized tools with (?P<o>\d+) operations"),
    ("mcpb/README.md", r"(?P<t>\d+) tools with (?P<o>\d+) operations"),
    ("mcpb/manifest.json", r"(?P<t>\d+) specialized tools with (?P<o>\d+) operations"),
    ("src/ExcelMcp.CLI/ExcelMcp.CLI.csproj", r"(?P<o>\d+) operations across"),
    ("gh-pages/docs/index.md", r"(?P<t>\d+) tools and (?P<o>\d+) operations"),
    (".github/plugins/excel-mcp/README.md", r"(?P<t>\d+) specialized tools with (?P<o>\d+) operations"),
    (".github/plugins/excel-cli/README.md", r"command categories with (?P<o>\d+) operations"),
    ("skills/excel-mcp/SKILL.md", r"Provides (?P<o>\d+) Excel operations"),
)

failures: list[str] = []


def add_failure(message: str) -> None:
    failures.append(message)


def canonical_counts() -> tuple[int, int, int, int, int]:
    content = MANIFEST.read_text(encoding="utf-8-sig")
    marker = 'public const string Json = @"'
    start = content.index(marker) + len(marker)
    end = content.rindex('";')
    manifest = json.loads(content[start:end].replace('""', '"'))

    # The manifest emits camelCase keys (totalCommands / totalOperations), matching the C# object.
    manifest_tools = int(manifest["totalCommands"])
    manifest_ops = int(manifest["totalOperations"])

    diag = next((c for c in manifest["commands"] if c["name"] == "diag"), None)
    if diag is None:
        add_failure("Expected a 'diag' command in the manifest; the user-facing count maths is stale.")
        diag_ops = 0
    else:
        diag_ops = len(diag.get("actions") or [])

    actions_text = TOOL_ACTIONS.read_text(encoding="utf-8-sig")
    enum_match = re.search(r"enum\s+FileAction\s*\{(?P<body>[^}]*)\}", actions_text)
    if enum_match is None:
        print("ERROR: could not locate the FileAction enum in ToolActions.cs", file=sys.stderr)
        sys.exit(1)
    file_ops = len(re.findall(r"JsonStringEnumMemberName", enum_match.group("body")))
    if file_ops == 0:
        print("ERROR: FileAction enum parsed to 0 operations - parsing bug.", file=sys.stderr)
        sys.exit(1)

    canonical_tools = manifest_tools - 1 + 1  # - diag + file
    canonical_ops = manifest_ops - diag_ops + file_ops
    return manifest_tools, manifest_ops, diag_ops, file_ops, (canonical_tools, canonical_ops)


def main() -> int:
    manifest_tools, manifest_ops, diag_ops, file_ops, (tools, ops) = canonical_counts()

    # 3. cross-check the real MCP tool surface
    mcp_names: set[str] = set()
    for path in MCP_SERVER_DIR.rglob("*.cs"):
        text = path.read_text(encoding="utf-8-sig", errors="ignore")
        mcp_names.update(re.findall(r'McpServerTool\s*\(\s*Name\s*=\s*"([^"]+)"', text))

    if len(mcp_names) != tools:
        add_failure(
            f"MCP tool surface has {len(mcp_names)} tools but canonical is {tools}."
        )
    if "diag" in mcp_names:
        add_failure("A 'diag' MCP tool exists - the diag-is-CLI-only assumption is broken.")
    if "file" not in mcp_names:
        add_failure("No 'file' MCP tool found - the file-adds-ops assumption is broken.")

    # 4. csproj GenerateSkillFile parameters
    for project in ("src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj", "src/ExcelMcp.CLI/ExcelMcp.CLI.csproj"):
        text = (REPO / project).read_text(encoding="utf-8-sig")
        match = re.search(r'ExtraOperationCount\s*=\s*"(\d+)"', text)
        if match and int(match.group(1)) != file_ops:
            add_failure(
                f"{project} sets ExtraOperationCount={match.group(1)} but FileAction has {file_ops}."
            )

    # 5. headline claims
    for relative, pattern in CHECKS:
        path = REPO / relative
        if not path.exists():
            add_failure(f"Expected doc not found: {relative}")
            continue
        content = path.read_text(encoding="utf-8-sig")
        matches = list(re.finditer(pattern, content))
        if not matches:
            add_failure(f"{relative}: headline pattern not found: /{pattern}/")
            continue
        for match in matches:
            groups = match.groupdict()
            if groups.get("t") and int(groups["t"]) != tools:
                add_failure(f"{relative}: tool count is {groups['t']} but should be {tools} -> {match.group(0).strip()!r}")
            if groups.get("o") and int(groups["o"]) != ops:
                add_failure(f"{relative}: operation count is {groups['o']} but should be {ops} -> {match.group(0).strip()!r}")

    # 5b. --help banner must stay derived
    program = PROGRAM.read_text(encoding="utf-8-sig")
    for match in re.finditer(r"Provides\s+\d+\s+tools\s+with\s+\d+\+?\s+operations", program):
        add_failure(f"Program.cs: --help banner restates counts as a literal -> {match.group(0)!r}")
    if "Provides {McpToolSurface.ToolCount} tools with {McpToolSurface.OperationCount} operations" not in program:
        add_failure("Program.cs: --help banner no longer derives its counts from McpToolSurface.")

    print(f"Canonical (from code): {tools} tools, {ops} operations")
    print(
        f"  manifest: {manifest_tools} tools / {manifest_ops} ops; "
        f"- diag({diag_ops}) + file({file_ops}); MCP tool surface: {len(mcp_names)} tools"
    )

    if failures:
        print(f"\nFAILED - {len(failures)} problem(s):")
        for message in failures:
            print(f"  - {message}")
        return 1

    print(f"\nPASS - all {len(CHECKS)} headline checks + surface cross-checks agree.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env bash
# One-key release for the local ExcelMcp build.
#
#   bump version -> build -> deploy -> emit versioned artefacts -> verify
#
# The version in mcp-server-excel/Directory.Build.props is the single source of truth. Every
# artefact written to dist/ carries it as a filename suffix, so nothing is published unversioned.
#
# Usage:
#   bash release.sh                 bump the build suffix, rebuild, deploy, package
#   bash release.sh --no-bump       repackage the current version as-is
#   bash release.sh --no-build      redeploy and repackage from the existing Release output
#   bash release.sh --test          additionally run the Excel-driven core test suites
#
# Builds go through build.sh, which restores the Windows environment variables the sandbox shell
# starts without; calling dotnet directly fails with "Value cannot be null (Parameter 'path1')".
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# Allow the tree to live anywhere; build.sh derives its own paths from the same variable.
export EXCEL_MCP_ROOT="$ROOT"

# Native tools need Windows paths; the user profile may contain spaces, so quote it.
to_win() {
    case "$1" in
        [A-Za-z]:[\\/]*) printf '%s' "$1" ;;
        *) cygpath -w "$1" 2>/dev/null || printf '%s' "$1" ;;
    esac
}
export DOTNET_ROOT="$(to_win "$ROOT/.dotnet10")"

BUMP=1
BUILD=1
RUN_TESTS=0

for arg in "$@"; do
    case "$arg" in
        --no-bump)  BUMP=0 ;;
        --no-build) BUILD=0 ;;
        --test)     RUN_TESTS=1 ;;
        *) echo "Unknown option: $arg" >&2; exit 2 ;;
    esac
done

# Prefer an explicit override, then a python on PATH, then the usual Windows install location.
PY=""
for candidate in "${EXCEL_MCP_PYTHON:-python}" python3 python.exe "/c/Program Files/Python312/python.exe"; do
    if command -v "$candidate" >/dev/null 2>&1; then
        PY="$candidate"
        break
    fi
done
if [ -z "$PY" ]; then
    echo "No python interpreter found; tools/pack.py needs one." >&2
    exit 1
fi

PACK="tools/pack.py"

# Run from the workspace root and hand python a RELATIVE path. MSYS rewrites POSIX paths when it
# launches a native executable, and an absolute one like /d/execl-mcp/tools/pack.py arrives at
# python.exe as D:\d\execl-mcp\tools\pack.py - the drive prefix gets duplicated. A relative path
# sidesteps the rewriting entirely. pack.py resolves the repo root from its own __file__.
cd "$ROOT"

pack() {
    "$PY" "$PACK" "$@"
}

if [ "$BUMP" -eq 1 ]; then
    VERSION="$(pack bump)"
    echo "== version bumped to $VERSION"
else
    VERSION="$(pack show)"
    echo "== keeping version $VERSION"
fi

if [ "$BUILD" -eq 1 ]; then
    echo "== building Release"
    # build.sh cds into the repo itself, so the project paths must be RELATIVE. An absolute
    # POSIX path (/d/execl-mcp/...) travels through `env` unrewritten and reaches MSBuild as a
    # path it reads as an unknown switch (MSB1001: 未知开关).
    bash "$ROOT/build.sh" build -c Release "src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj"
    bash "$ROOT/build.sh" build -c Release "src/ExcelMcp.CLI/ExcelMcp.CLI.csproj"
fi

echo "== deploying and packaging"
pack collect "$VERSION"

if [ "$RUN_TESTS" -eq 1 ]; then
    echo "== running core test suites"
    bash "$ROOT/build.sh" test -c Release "tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj" \
        --filter "FullyQualifiedName~Slicer|FullyQualifiedName~PivotTable"
    bash "$ROOT/build.sh" test -c Release "tests/ExcelMcp.McpServer.Tests/ExcelMcp.McpServer.Tests.csproj"
fi

echo "== done: $VERSION"

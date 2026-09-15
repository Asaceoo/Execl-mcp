#!/usr/bin/env bash
# Build wrapper for the pinned ExcelMcp SDK on Windows / Git Bash.
#
# Why this exists: a sandboxed Git Bash shell can start with none of the core Windows environment
# variables set (SystemRoot, windir, COMSPEC, APPDATA, PROGRAMDATA, ProgramFiles...), which makes
# `dotnet restore` / `dotnet build` fail with "Value cannot be null (Parameter 'path1')".
# This restores the minimum set, including the two variables that bash cannot export by name
# (ProgramFiles(x86), ProgramFiles) via `env`, then delegates to the pinned SDK.
#
# Layout assumed (override any of it with the env vars below):
#   <workspace>/build.sh              <- this file
#   <workspace>/mcp-server-excel/     <- the source tree that gets built
#   <workspace>/.dotnet10/dotnet.exe  <- pinned SDK
#   <workspace>/.nuget-packages/      <- redirected NuGet cache (never writes the user profile)
#
# Env overrides:
#   EXCEL_MCP_ROOT      workspace root            (default: the directory holding this script)
#   EXCEL_MCP_REPO      source tree to build      (default: $EXCEL_MCP_ROOT/mcp-server-excel)
#   EXCEL_MCP_DOTNET    dotnet executable         (default: $EXCEL_MCP_ROOT/.dotnet10/dotnet.exe)
#   EXCEL_MCP_NUGET     NuGet package cache dir   (default: $EXCEL_MCP_ROOT/.nuget-packages)
#
# Usage:
#   bash build.sh build -c Release src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj
#   bash build.sh test  -c Release tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj --filter ...
#
# NOTE: project paths must be RELATIVE to the source tree. An absolute POSIX path travels through
# `env` unrewritten and reaches MSBuild as a path it reads as an unknown switch (MSB1001).
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

WORKSPACE="${EXCEL_MCP_ROOT:-$SCRIPT_DIR}"
REPO="${EXCEL_MCP_REPO:-$WORKSPACE/mcp-server-excel}"
DOTNET="${EXCEL_MCP_DOTNET:-$WORKSPACE/.dotnet10/dotnet.exe}"
NUGET_DIR="${EXCEL_MCP_NUGET:-$WORKSPACE/.nuget-packages}"
DOTNET_HOME="$WORKSPACE/.dotnet-home"

if [ ! -d "$REPO" ]; then
    echo "Source tree not found at $REPO (set EXCEL_MCP_REPO to point at it)." >&2
    exit 1
fi

if [ ! -x "$DOTNET" ]; then
    echo "Pinned SDK not found at $DOTNET (set EXCEL_MCP_DOTNET to point at dotnet)." >&2
    exit 1
fi

# Native tools want Windows paths; Git Bash hands us POSIX ones. Values already in Windows form
# (C:\..., D:/...) pass through untouched, so an override may use either convention.
to_win() {
    case "$1" in
        [A-Za-z]:[\\/]*) printf '%s' "$1" ;;
        *) cygpath -w "$1" 2>/dev/null || printf '%s' "$1" ;;
    esac
}

WIN_DOTNET_ROOT="$(to_win "$WORKSPACE/.dotnet10")"
WIN_NUGET="$(to_win "$NUGET_DIR")"
WIN_DOTNET_HOME="$(to_win "$DOTNET_HOME")"

# Reuse the caller's Windows user dirs when Git Bash already exports them; derive them from $HOME
# otherwise (a bare sandbox shell has none of them).
WIN_USERPROFILE="${USERPROFILE:-$(to_win "$HOME")}"
WIN_DRIVE="${WIN_USERPROFILE%%:*}"
WIN_HOMEPATH="${WIN_USERPROFILE#*:}"
if [ "$WIN_DRIVE" = "$WIN_USERPROFILE" ]; then
    # No drive prefix at all - fall back to a plain split so the env block stays well formed.
    WIN_DRIVE="C"
    WIN_HOMEPATH="$WIN_USERPROFILE"
fi
WIN_APPDATA="${APPDATA:-${WIN_USERPROFILE}\\AppData\\Roaming}"
WIN_LOCALAPPDATA="${LOCALAPPDATA:-${WIN_USERPROFILE}\\AppData\\Local}"
WIN_SYSTEMROOT="${SystemRoot:-${SYSTEMROOT:-C:\\WINDOWS}}"

mkdir -p "$NUGET_DIR" "$DOTNET_HOME"

# Keep the MSYS path (so helper binaries stay reachable) and add the Windows system dirs.
MSYS_PATH="$PATH"
cd "$REPO"

exec env \
    "SystemRoot=$WIN_SYSTEMROOT" \
    "windir=$WIN_SYSTEMROOT" \
    "COMSPEC=$WIN_SYSTEMROOT\\system32\\cmd.exe" \
    "PATHEXT=.COM;.EXE;.BAT;.CMD" \
    "ProgramFiles=${ProgramFiles:-C:\\Program Files}" \
    "ProgramFiles(x86)=${ProgramFiles_x86:-C:\\Program Files (x86)}" \
    "APPDATA=$WIN_APPDATA" \
    "LOCALAPPDATA=$WIN_LOCALAPPDATA" \
    "PROGRAMDATA=${PROGRAMDATA:-C:\\ProgramData}" \
    "ALLUSERSPROFILE=${ALLUSERSPROFILE:-C:\\ProgramData}" \
    "USERPROFILE=$WIN_USERPROFILE" \
    "HOMEDRIVE=${WIN_DRIVE}:" \
    "HOMEPATH=$WIN_HOMEPATH" \
    "TEMP=${TEMP:-${WIN_LOCALAPPDATA}\\Temp}" \
    "TMP=${TMP:-${WIN_LOCALAPPDATA}\\Temp}" \
    "DOTNET_ROOT=$WIN_DOTNET_ROOT" \
    "NUGET_PACKAGES=$WIN_NUGET" \
    "DOTNET_CLI_HOME=$WIN_DOTNET_HOME" \
    "DOTNET_CLI_TELEMETRY_OPTOUT=1" \
    "DOTNET_NOLOGO=1" \
    "DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1" \
    "PATH=${MSYS_PATH}:/c/WINDOWS/system32:/c/WINDOWS" \
    "$DOTNET" "$@"

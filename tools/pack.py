#!/usr/bin/env python
"""Release helper for the local ExcelMcp build.

Single source of truth for the version is <Version> in mcp-server-excel/Directory.Build.props.
Every published artefact carries that version as a filename suffix - no unversioned output.

Subcommands
    show                print the current version
    bump                increment the trailing build suffix, write it back, print the new version
    collect <version>   deploy the Release output to excel-mcp-bin and emit versioned artefacts

Called by release.sh, which owns the build step because the sandbox shell needs its environment
vars restored first. Run this file directly only for show / bump.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REPO = ROOT / "mcp-server-excel"
PROPS = REPO / "Directory.Build.props"
DEPLOY_DIR = ROOT / "excel-mcp-bin"
DIST_DIR = ROOT / "dist"

TFM = "net10.0-windows"
BUILD_ROOT = REPO / "src"
# (project directory, output file name) - the output file name is the AssemblyName from each csproj.
PRODUCTS = (
    ("ExcelMcp.McpServer", "Sbroenne.ExcelMcp.McpServer.exe"),
    ("ExcelMcp.CLI", "excelcli.exe"),
)
# A .NET apphost is only a loader: it needs its managed dll, .deps.json and .runtimeconfig.json
# beside it. Copying the .exe alone produces a file that always fails with
# "The application to execute does not exist", so entry points are shipped as a set, never solo.
APPHOST_SIBLINGS = (".exe", ".dll", ".deps.json", ".runtimeconfig.json")

VERSION_RE = re.compile(r"<Version>([^<]+)</Version>")
BUILD_SUFFIX_RE = re.compile(r"^(?P<base>.+?)\.(?P<build>\d+)$")


class ReleaseError(RuntimeError):
    """Raised for conditions the operator has to fix, rather than bugs in this script."""


def read_version() -> str:
    text = PROPS.read_text(encoding="utf-8-sig")
    match = VERSION_RE.search(text)
    if match is None:
        raise ReleaseError(f"No <Version> element found in {PROPS}")
    return match.group(1).strip()


def write_version(new_version: str) -> None:
    text = PROPS.read_text(encoding="utf-8-sig")
    replaced, count = VERSION_RE.subn(f"<Version>{new_version}</Version>", text, count=1)
    if count != 1:
        raise ReleaseError(f"Expected exactly one <Version> element in {PROPS}")
    PROPS.write_text(replaced, encoding="utf-8-sig", newline="")


def next_version(current: str) -> str:
    """Increment the trailing build suffix: 2.0.8-slicerlink.1 -> 2.0.8-slicerlink.2.

    A version without a trailing number gets '.1' so the first bump of a bare release is defined.
    """
    match = BUILD_SUFFIX_RE.match(current)
    if match is None:
        return f"{current}.1"
    return f"{match.group('base')}.{int(match.group('build')) + 1}"


def output_dir(product: str) -> Path:
    return BUILD_ROOT / product / "bin" / "Release" / TFM


def kill_processes_locking_output() -> None:
    """Stop processes that hold the deployment directory open.

    Overwriting a running executable's directory fails with 'Device or resource busy', so the
    deployed server is stopped first. taskkill is invoked directly (not through a shell) so no
    MSYS path rewriting can mangle its arguments. Absent processes are not an error.
    """
    for _, exe_name in PRODUCTS:
        result = subprocess.run(
            ["taskkill", "/F", "/T", "/IM", exe_name],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        if result.returncode == 0:
            print(f"stopped a running {exe_name} so its files could be replaced")


def copy_tree_reporting_stale(source: Path, destination: Path) -> list[str]:
    """Mirror source into destination, returning files that exist only in destination.

    Stale entries are reported rather than deleted: this sandbox fails closed on recursive
    deletion, and silently leaving an old assembly behind is exactly what the --version check
    below is for.
    """
    destination.mkdir(parents=True, exist_ok=True)
    source_files = {p.relative_to(source) for p in source.rglob("*") if p.is_file()}

    for relative in source_files:
        target = destination / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source / relative, target)

    return sorted(
        str(p.relative_to(destination))
        for p in destination.rglob("*")
        if p.is_file() and p.relative_to(destination) not in source_files
    )


def run_version_check(exe_path: Path) -> str:
    """Run `<exe> --version` and return its combined output.

    Captured as BYTES and decoded defensively. The banner goes through the Windows console code
    page, so any non-ASCII character in it arrives as GBK/ANSI, not UTF-8 - and `text=True` would
    then raise UnicodeDecodeError inside the reader thread, leaving the release self-check silently
    empty (it reported "deployed binary reports ''" instead of a real mismatch).
    """
    result = subprocess.run(
        [str(exe_path), "--version"],
        capture_output=True,
        check=False,
        cwd=str(exe_path.parent),
    )
    raw = (result.stdout or b"") + (result.stderr or b"")
    for encoding in ("utf-8", "cp936", "mbcs", "latin-1"):
        try:
            return raw.decode(encoding).strip()
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", errors="replace").strip()


def build_zip(destination: Path, files_root: Path) -> None:
    """Zip with zipfile rather than PowerShell Compress-Archive: the sandbox kills that helper's
    process tree, which leaves behind a zero-byte archive."""
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(files_root.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(files_root.parent))


def copy_entry_point(source_dir: Path, destination: Path, exe_name: str) -> list[str]:
    """Copy an apphost together with the files it loads, returning the copied file names."""
    stem = Path(exe_name).stem
    copied: list[str] = []
    for suffix in APPHOST_SIBLINGS:
        source = source_dir / f"{stem}{suffix}"
        if not source.is_file():
            raise ReleaseError(f"Expected build output missing: {source}")
        shutil.copy2(source, destination / source.name)
        copied.append(source.name)
    return copied


def is_entry_point_file(name: str) -> bool:
    """True for the apphost sets that are intentionally added on top of the server output."""
    return any(name.startswith(Path(exe).stem + ".") for _, exe in PRODUCTS)


def assert_bundle_runnable(bundle: Path) -> None:
    """Fail the release if a shipped entry point cannot start, instead of shipping a dead file."""
    with zipfile.ZipFile(bundle) as archive:
        names = {Path(n).name for n in archive.namelist()}
    missing = [
        f"{Path(exe).stem}{suffix}"
        for _, exe in PRODUCTS
        for suffix in APPHOST_SIBLINGS
        if f"{Path(exe).stem}{suffix}" not in names
    ]
    if missing:
        raise ReleaseError(
            f"{bundle.name} would ship an entry point that cannot start. Missing: "
            + ", ".join(missing)
        )


def collect(version: str) -> int:
    actual_version = read_version()
    if actual_version != version:
        raise ReleaseError(
            f"Version mismatch: Directory.Build.props says '{actual_version}', collect was given "
            f"'{version}'. Build and package must use the same version."
        )

    missing = [name for name, _ in PRODUCTS if not output_dir(name).is_dir()]
    if missing:
        raise ReleaseError(
            "Release output missing for: " + ", ".join(missing) + ". Run the build first."
        )

    kill_processes_locking_output()

    build_time = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%SZ")

    # The MCP server is the deployable unit, so its output is what gets mirrored. The CLI shares
    # that dependency set, so only its own apphost set is added on top - it then runs in place,
    # which is what makes a single self-contained bundle possible.
    server_dir = output_dir("ExcelMcp.McpServer")
    DEPLOY_DIR.mkdir(parents=True, exist_ok=True)
    copy_entry_point(output_dir(PRODUCTS[1][0]), DEPLOY_DIR, PRODUCTS[1][1])
    stale = [
        name
        for name in copy_tree_reporting_stale(server_dir, DEPLOY_DIR)
        if not is_entry_point_file(Path(name).name) and Path(name).name != "VERSION.txt"
    ]

    (DEPLOY_DIR / "VERSION.txt").write_text(
        "\n".join(
            [
                f"version: {version}",
                f"built: {build_time}",
                f"source: {server_dir}",
                "",
                "Built from the local patch to upstream 2.0.8. Regenerated by release.sh; edits here",
                "are overwritten on the next release.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    deployed_files = sorted(p for p in DEPLOY_DIR.rglob("*") if p.is_file())

    DIST_DIR.mkdir(parents=True, exist_ok=True)
    bundle = DIST_DIR / f"jyyj-mcp-{version}-win-x64.zip"
    build_zip(bundle, DEPLOY_DIR)

    # One versioned bundle is the published unit. Solo .exe copies are deliberately not produced:
    # an apphost without its managed dll next to it can never start, so shipping one would put a
    # dead file in the hands of the operator. assert_bundle_runnable guards that same invariant.
    published = [bundle]
    assert_bundle_runnable(bundle)

    server_exe = DEPLOY_DIR / PRODUCTS[0][1]
    cli_exe = DEPLOY_DIR / PRODUCTS[1][1]
    manifest = {
        "version": version,
        "built": build_time,
        "deploy_directory": str(DEPLOY_DIR),
        "artefacts": [p.name for p in published],
        "deployed_file_count": len(deployed_files),
        "stale_files_tolerated": stale,
        "version_check": run_version_check(server_exe),
        "cli_check": run_version_check(cli_exe),
    }
    manifest_path = DIST_DIR / f"jyyj-mcp-{version}-manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    published.append(manifest_path)

    print(f"version: {version}")
    print(f"deployed: {DEPLOY_DIR} ({len(deployed_files)} files)")
    if stale:
        print(f"WARNING - {len(stale)} file(s) present in the deploy directory but not in the build:")
        for name in stale:
            print(f"  stale: {name}")
    for path in published:
        size = path.stat().st_size
        print(f"artefact: {path.name} ({size:,} bytes)")
    print(f"version check: {manifest['version_check']}")
    print(f"cli check: {(manifest['cli_check'] or '(no output)').splitlines()[0]}")

    if version not in manifest["version_check"]:
        print(
            f"FAILED - deployed binary reports '{manifest['version_check']}' "
            f"but expected version '{version}'",
            file=sys.stderr,
        )
        return 1

    if not manifest["cli_check"].strip():
        print("FAILED - deployed CLI produced no output; the bundle would ship a dead entry point")
        return 1

    return 0


def selftest() -> int:
    """Assert the version arithmetic, including the four-part form used from 3.0.0.0 onwards.

    The packaging step rewrites <Version> on every release, so a silent regression here would
    publish a mislabelled artefact - a failure that no build error would catch.
    """
    cases = [
        ("2.0.8", "2.0.9"),
        ("2.0.8-jyyj.1", "2.0.8-jyyj.2"),
        ("3.0.0.0", "3.0.0.1"),
        ("3.0.0.9", "3.0.0.10"),
        ("3.0.1.0", "3.0.1.1"),
    ]
    failures = [
        f"  next_version({current!r}) = {next_version(current)!r}, expected {expected!r}"
        for current, expected in cases
        if next_version(current) != expected
    ]
    if failures:
        print("FAILED - version arithmetic is wrong:", file=sys.stderr)
        print("\n".join(failures), file=sys.stderr)
        return 1

    print(f"OK - {len(cases)} version-arithmetic cases; source version is {read_version()}")
    return 0


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print(__doc__, file=sys.stderr)
        return 2

    command = argv[1]

    if command == "selftest":
        return selftest()

    if command == "show":
        print(read_version())
        return 0

    if command == "bump":
        new_version = next_version(read_version())
        write_version(new_version)
        print(new_version)
        return 0

    if command == "collect":
        if len(argv) < 3:
            print("collect requires the version to package", file=sys.stderr)
            return 2
        return collect(argv[2])

    print(f"Unknown subcommand: {command}", file=sys.stderr)
    return 2


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv))
    except ReleaseError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        sys.exit(1)

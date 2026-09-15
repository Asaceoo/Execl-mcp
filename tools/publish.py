#!/usr/bin/env python
"""Stage this workspace into a clean standalone git repository and publish it.

Why a separate tree, instead of `git init` at the workspace root
-----------------------------------------------------------------
The source tree under `mcp-server-excel/` is itself a git checkout of the upstream project. Git
refuses to record the *contents* of a nested repository: `git add` stores a 40-character gitlink
instead, so committing from the workspace root would silently replace ~1100 source files with a
single pointer and produce a repository that cannot be built. Staging into a tree that contains no
nested `.git` sidesteps that completely, and leaves the working checkout (and its ability to pull
upstream) untouched.

The published file set is enumerated explicitly rather than copied wholesale
----------------------------------------------------------------------------
* every file the nested repository tracks (`git ls-files`) - this is how the local patches, which
  live in the working tree, travel along;
* untracked-but-not-ignored files inside the nested tree (the new sources added by those patches);
* this workspace's own root files and `tools/`.

`mcp-server-excel/.github/workflows/` is excluded on purpose - see `.gitignore` for the reasoning.

Subcommands
    stage                refresh the publish tree, print the expected file count
    check                stage, then assert git's index contains exactly the expected files
    push [--message M]   check, commit, and push to the GitHub remote
    status               show what the publish tree currently differs by
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

WORKSPACE = Path(__file__).resolve().parent.parent
NESTED = WORKSPACE / "mcp-server-excel"
DEFAULT_DEST = WORKSPACE / "_publish"
REMOTE_URL = "git@github.com:Asaceoo/Execl-mcp.git"
REMOTE_NAME = "origin"
BRANCH = "main"

# Root-level entries that are local toolchain state, build output, or other publish trees.
ROOT_EXCLUDE_DIRS = {
    ".dotnet10",
    ".dotnet-home",
    ".nuget-packages",
    ".workbuddy",
    "excel-mcp-bin",
    "dist",
    "mcp-server-excel",  # handled separately, from the nested repo's own tracking
    "_publish",
    "__pycache__",
}
ROOT_EXCLUDE_NAMES = {"nul", "desktop.ini"}

# Upstream CI cannot run here (self-hosted runners, sealed secrets, deploy targets we do not have).
NESTED_EXCLUDE_PREFIXES = (".github/workflows/",)

NESTED_EXTRA_EXCLUDE_SUFFIXES = (".log",)


class PublishError(RuntimeError):
    """A condition the operator has to fix, rather than a bug in this script."""


def run(args: list[str], cwd: Path | None = None, env: dict | None = None, check: bool = True):
    result = subprocess.run(args, cwd=str(cwd) if cwd else None, env=env, capture_output=True, text=True)
    if check and result.returncode != 0:
        raise PublishError(
            f"command failed ({result.returncode}): {' '.join(args)}\n{result.stdout}\n{result.stderr}"
        )
    return result


def git(args: list[str], cwd: Path, env: dict) -> str:
    return run(["git", *args], cwd=cwd, env=env).stdout


def clean_git_env() -> dict:
    """A GIT_CONFIG_GLOBAL without the machine's `insteadOf` mirrors.

    The developer machine rewrites https://github.com/... to a mirror that answers 502, so any
    HTTPS GitHub operation fails. SSH remotes are not affected by that particular rule, but a clean
    config makes the publish deterministic regardless of what the machine's global config holds.
    """
    handle, path = tempfile.mkstemp(prefix="publish-gitconfig-", suffix=".ini")
    os.close(handle)
    Path(path).write_text(
        "[user]\n\tname = asaceoo\n\temail = 280346890@qq.com\n[safe]\n\tdirectory = *\n",
        encoding="utf-8",
    )
    env = dict(os.environ)
    env["GIT_CONFIG_GLOBAL"] = path
    env["GIT_CONFIG_SYSTEM"] = os.devnull
    return env


def nested_file_set(env: dict) -> list[str]:
    """Tracked files plus untracked-but-not-ignored files, from the nested repository."""
    tracked = git(["ls-files", "-z"], NESTED, env).split("\0")
    untracked = git(["ls-files", "-z", "--others", "--exclude-standard"], NESTED, env).split("\0")
    paths = [p for p in {*tracked, *untracked} if p]
    kept, dropped = [], []
    for path in paths:
        if path.endswith(NESTED_EXTRA_EXCLUDE_SUFFIXES) or path.startswith(NESTED_EXCLUDE_PREFIXES):
            dropped.append(path)
        else:
            kept.append(path)
    if dropped:
        print(f"  excluded from the nested tree: {len(dropped)} file(s)")
        for path in sorted(dropped)[:5]:
            print(f"    - {path}")
        if len(dropped) > 5:
            print(f"    ... and {len(dropped) - 5} more")
    return sorted(kept)


def root_file_set() -> list[str]:
    """Workspace-root files and tools/, minus local-only state. Paths are relative to the root."""
    collected: list[str] = []
    for entry in sorted(WORKSPACE.iterdir()):
        name = entry.name
        if name in ROOT_EXCLUDE_DIRS or name in ROOT_EXCLUDE_NAMES:
            continue
        if name.startswith("_"):  # _demo, _regshards, _tlist_v5.log, _tests_list.txt, ...
            continue
        if entry.is_dir():
            for path in sorted(entry.rglob("*")):
                if not path.is_file():
                    continue
                if "__pycache__" in path.parts:
                    continue
                relative = path.relative_to(WORKSPACE).as_posix()
                if relative.endswith(".log"):
                    continue
                collected.append(relative)
        elif entry.is_file():
            if name.endswith(".log"):
                continue
            collected.append(name)
    return collected


def collect_expected() -> tuple[list[str], list[str], list[str]]:
    env = clean_git_env()
    nested = nested_file_set(env)
    root = root_file_set()
    missing = [p for p in nested if not (NESTED / p).is_file()]
    return root, nested, missing


def stage(dest: Path) -> int:
    print(f"staging into {dest}")
    root_files, nested_files, missing = collect_expected()
    if missing:
        print(f"  WARNING: {len(missing)} tracked file(s) are absent from the working tree:")
        for path in missing[:5]:
            print(f"    ! {path}")

    expected: list[str] = []
    copies: list[tuple[Path, str]] = []
    for relative in root_files:
        copies.append((WORKSPACE / relative, relative))
        expected.append(relative)
    for relative in nested_files:
        target = f"mcp-server-excel/{relative}"
        copies.append((NESTED / relative, target))
        expected.append(target)

    # Remove anything in the destination that the current expectation does not cover, so that a
    # file deleted from the sources actually disappears instead of lingering as a stale commit.
    if dest.exists():
        keep = set(expected) | {".git"}
        stale: list[Path] = []
        for path in sorted(dest.rglob("*")):
            relative = path.relative_to(dest).as_posix()
            if relative.split("/", 1)[0] in keep or relative in keep:
                continue
            stale.append(path)
        removed = 0
        for path in sorted(stale, key=lambda p: len(p.parts), reverse=True):
            if path.is_file() or path.is_symlink():
                path.unlink()
                removed += 1
            elif path.is_dir():
                try:
                    path.rmdir()
                except OSError:
                    pass
        if removed:
            print(f"  removed {removed} stale file(s) from the publish tree")

    copied = 0
    for source, relative in copies:
        if not source.is_file():
            continue
        target = dest / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
        copied += 1

    print(f"  copied {copied} file(s)  (root {len(root_files)} + nested {len(nested_files)})")
    return len(expected)


def ensure_repo(dest: Path, env: dict) -> None:
    if not (dest / ".git").exists():
        print(f"  initialising a fresh repository at {dest}")
        dest.mkdir(parents=True, exist_ok=True)
        run(["git", "init", "-b", BRANCH], cwd=dest, env=env)
        run(["git", "remote", "add", REMOTE_NAME, REMOTE_URL], cwd=dest, env=env)
    remotes = git(["remote"], dest, env).split()
    if REMOTE_NAME not in remotes:
        run(["git", "remote", "add", REMOTE_NAME, REMOTE_URL], cwd=dest, env=env)


def check(dest: Path, expected_count: int) -> None:
    """Assert git's index holds exactly the files we meant to publish.

    Without this, a rule in a .gitignore (the workspace's, or the nested one) can silently swallow
    a file and the publish still reports success - the repository just quietly ships less than the
    source tree. An always-green guard is worse than no guard, so a mismatch fails the run.
    """
    env = clean_git_env()
    # -f matters: upstream tracks a few files that its own .gitignore also matches
    # (llm-tests/aitest-reports/results.json is one). Plain `git add -A` drops them and the
    # published repository silently ships fewer files than the source tree. `check()` compares the
    # index against the enumerated set afterwards, so forcing cannot let extras slip through.
    run(["git", "add", "-A", "-f"], cwd=dest, env=env)
    staged = [p for p in git(["ls-files", "-z"], dest, env).split("\0") if p]
    print(f"  index holds {len(staged)} file(s); expected {expected_count}")
    if len(staged) != expected_count:
        expected_root, expected_nested, _ = collect_expected()
        want = set(expected_root) | {f"mcp-server-excel/{p}" for p in expected_nested}
        got = set(staged)
        missing = sorted(want - got)
        extra = sorted(got - want)
        detail = []
        if missing:
            detail.append(f"missing from the index ({len(missing)}): " + ", ".join(missing[:8]))
        if extra:
            detail.append(f"unexpected in the index ({len(extra)}): " + ", ".join(extra[:8]))
        raise PublishError("staged file set does not match the expected set\n  " + "\n  ".join(detail))


def commit(dest: Path, message: str) -> str:
    env = clean_git_env()
    # -f matters: upstream tracks a few files that its own .gitignore also matches
    # (llm-tests/aitest-reports/results.json is one). Plain `git add -A` drops them and the
    # published repository silently ships fewer files than the source tree. `check()` compares the
    # index against the enumerated set afterwards, so forcing cannot let extras slip through.
    run(["git", "add", "-A", "-f"], cwd=dest, env=env)

    # The root-commit requirement belongs to the FIRST publish only. A fresh staging tree must not
    # inherit a history it never had - that is what the check was written for. Every later publish
    # legitimately continues that root commit, and demanding rootness again made the second publish
    # impossible: `git commit` always writes a parent, so the guard could only ever fire.
    first_publish = (
        run(["git", "rev-parse", "--verify", "--quiet", "HEAD"], cwd=dest, env=env, check=False)
        .returncode != 0
    )

    # An unchanged tree is not an error: re-publishing the same content should be idempotent, not
    # fail with git's "nothing to commit". Refresh the message on the existing commit instead.
    staged_changes = run(
        ["git", "diff", "--cached", "--quiet", "--exit-code"], cwd=dest, env=env, check=False
    ).returncode != 0

    if not staged_changes and not first_publish:
        run(["git", "commit", "--amend", "-m", message], cwd=dest, env=env)
        sha = git(["rev-parse", "HEAD"], dest, env).strip()
        print(f"  tree unchanged; amended the message on {sha[:12]}")
        return sha

    run(["git", "commit", "-m", message], cwd=dest, env=env)
    sha = git(["rev-parse", "HEAD"], dest, env).strip()

    if first_publish:
        parents = git(["cat-file", "-p", "HEAD"], dest, env).count("\nparent ")
        if parents:
            raise PublishError("the first publish commit is not a root commit")
        print(f"  committed root commit {sha[:12]}")
    else:
        print(f"  committed {sha[:12]} on top of {git(['rev-parse', '--short', 'HEAD^'], dest, env).strip()}")

    return sha


def push(dest: Path, message: str | None, force: bool) -> None:
    env = clean_git_env()
    if message:
        ensure_repo(dest, env)
        check(dest, stage(dest))
        # commit() already reports what it did (root commit vs. child commit); saying "root commit"
        # again here was a lie for every publish after the first one.
        sha = commit(dest, message)
    else:
        ensure_repo(dest, env)
        # `.git/HEAD` exists from `git init` onwards, so it cannot tell "no commit yet" apart from
        # "nothing new to push". Ask git for a commit instead.
        has_commit = run(
            ["git", "rev-parse", "--verify", "--quiet", "HEAD"], cwd=dest, env=env, check=False
        ).returncode == 0
        if not has_commit:
            raise PublishError("nothing to push and no commit given; pass --message")
        sha = git(["rev-parse", "HEAD"], dest, env).strip()
        print(f"  pushing existing {sha[:12]}")

    args = ["push", REMOTE_NAME, f"{BRANCH}:{BRANCH}"]
    if force:
        args.append("--force")
    print(f"  git {' '.join(args)} -> {REMOTE_URL}")
    result = run(["git", *args], cwd=dest, env=env, check=False)
    sys.stdout.write(result.stdout)
    sys.stderr.write(result.stderr)
    if result.returncode != 0:
        raise PublishError(f"push failed with exit code {result.returncode}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("command", choices=["stage", "check", "push", "status"])
    parser.add_argument("--dest", default=os.environ.get("EXCEL_MCP_PUBLISH_DIR", str(DEFAULT_DEST)))
    parser.add_argument("--message", "-m", default=None, help="commit message; omit to push the existing HEAD")
    parser.add_argument("--force", action="store_true", help="force-push (rewrites the remote branch)")
    options = parser.parse_args()

    dest = Path(options.dest).resolve()
    try:
        if options.command == "stage":
            stage(dest)
        elif options.command == "check":
            ensure_repo(dest, clean_git_env())
            check(dest, stage(dest))
            print("OK - the index matches the expected file set")
        elif options.command == "status":
            env = clean_git_env()
            ensure_repo(dest, env)
            print(git(["status", "--short"], dest, env) or "  (clean)")
        elif options.command == "push":
            push(dest, options.message, options.force)
    except PublishError as error:
        print(f"FAILED - {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

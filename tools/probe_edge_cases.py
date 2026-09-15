"""Edge-case harvester: drive the deployed server into the corners and record the REAL reply.

Every case carries an expectation, so this doubles as a regression guard. A case that
returns something other than expected is a finding, and the raw reply is printed verbatim
so the fix can be driven by evidence instead of by reading the source.

Usage:
    python tools/probe_edge_cases.py            # all cases
    python tools/probe_edge_cases.py --list     # show case ids
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from demo_slicer_link import Server  # noqa: E402  (shared stdio client)

ROOT = Path(__file__).resolve().parent.parent


class Case:
    def __init__(self, cid, group, what, tool, args, expect_ok,
                 expect_regex=None, note=""):
        self.cid = cid
        self.group = group
        self.what = what
        self.tool = tool
        self.args = args
        self.expect_ok = expect_ok
        self.expect_regex = expect_regex
        self.note = note


RESULTS: list[tuple[Case, bool, str]] = []


def body_of(reply: dict) -> str:
    """Server replies with a JSON envelope; surface success/error as one readable string."""
    if "_images" in reply:
        reply = {k: v for k, v in reply.items() if k != "_images"}
    if "raw" in reply:
        return str(reply["raw"])
    if "error" in reply and "success" not in reply:
        return str(reply["error"])
    ok = reply.get("success")
    if ok is False:
        return str(reply.get("error") or reply.get("errorMessage") or reply)
    return json.dumps(
        {k: v for k, v in reply.items()
         if k in ("success", "message", "chartName", "chartType", "error", "errorMessage")},
        ensure_ascii=False)


def run(server: Server, session: str, case: Case) -> None:
    args = dict(case.args)
    args["session_id"] = session
    ok = False
    try:
        reply = server.call(case.tool, args)
        text = body_of(reply)
        ok = reply.get("success") is not False and not reply.get("error")
    except Exception as error:  # noqa: BLE001
        # The shared client raises when the server answers success=false. A raised message is still
        # a failure observation, so grade it against the expectation instead of assuming the worst -
        # an earlier version failed every expected-failure case, which read as nine phantom defects.
        text = str(error)

    verdict = ok == case.expect_ok
    # Keep enough of the reply to show WHY it did not match. At 400 characters the actionable
    # hint the server appends was cut off, and the truncation read as "the hint is missing".
    detail = text[:1200]
    if verdict and case.expect_regex and not re.search(case.expect_regex, text, re.I | re.S):
        verdict = False
        detail += f"\n        !! expected /{case.expect_regex}/ in the reply but it is absent"
    RESULTS.append((case, verdict, detail))
    mark = "ok  " if verdict else "FAIL"
    print(f"  [{mark}] {case.cid:<28} {'success' if ok else 'failure'}")
    print(f"         {detail.splitlines()[0][:170]}")


def build_cases() -> list[Case]:
    chart = "create-from-range"
    return [
        # --- chart source-address quoting -----------------------------------
        Case("CH-QUOTE-APOSTROPHE", "chart", "sheet name contains an apostrophe, bare address",
             "chart", {"action": chart, "sheet_name": "O'Brien",
                       "source_range_address": "A1:C7", "chart_type": "ColumnClustered",
                       "chart_name": "ch_apos", "target_range": "E2:K16"},
             True, None,
             "Excel sheet names may contain apostrophes; the reference must double them."),
        Case("CH-QUOTE-SPACE", "chart", "sheet name contains a space, bare address (control)",
             "chart", {"action": chart, "sheet_name": "With Space",
                       "source_range_address": "A1:C7", "chart_type": "ColumnClustered",
                       "chart_name": "ch_space", "target_range": "E2:K16"},
             True, None),
        Case("CH-QUOTE-PLAIN", "chart", "plain sheet name, bare address (control)",
             "chart", {"action": chart, "sheet_name": "Plain",
                       "source_range_address": "A1:C7", "chart_type": "ColumnClustered",
                       "chart_name": "ch_plain", "target_range": "E2:K16"},
             True, None),

        # --- chart type discovery / downgrade -------------------------------
        Case("CH-TYPE-INVALID", "chart", "unknown chart type must list the legal values",
             "chart", {"action": chart, "sheet_name": "Plain",
                       "source_range_address": "A1:C7", "chart_type": "Windows",
                       "chart_name": "ch_bad", "target_range": "E18:K32"},
             False, r"Valid values:.*ColumnClustered"),
        Case("CH-TYPE-SUNBURST", "chart", "sunburst is refused when Excel downgrades it",
             "chart", {"action": chart, "sheet_name": "Plain",
                       "source_range_address": "A1:C7", "chart_type": "Sunburst",
                       "chart_name": "ch_sun", "target_range": "E34:K48"},
             False, r"fell back|fallback|not available"),

        # --- modern / stock data-shape guards -------------------------------
        Case("CH-MODERN-ALLTEXT", "chart", "modern chart over an all-text range",
             "chart", {"action": chart, "sheet_name": "TextOnly",
                       "source_range_address": "A1:B7", "chart_type": "Treemap",
                       "chart_name": "ch_text", "target_range": "D2:J16"},
             False, r"numeric"),
        Case("CH-STOCK-WRONGCOLS", "chart", "HLC stock chart fed four value columns",
             "chart", {"action": chart, "sheet_name": "FourNum",
                       "source_range_address": "A1:E7", "chart_type": "StockHLC",
                       "chart_name": "ch_stock", "target_range": "M2:S16"},
             False, r"stock|series|fall back|fell back|not available"),

        # --- range error surface (F5-9 regression) --------------------------
        Case("RANGE-EDIT-BADACTION", "error-surface", "range_edit with an unknown action",
             "range_edit", {"action": "no-such-action", "sheet_name": "Plain",
                            "range_address": "A1"},
             False, r"rejected the arguments|Valid|no-such-action"),
        Case("RANGE-FORMAT-BADPARAM", "error-surface", "range_format with an undeclared argument",
             "range_format", {"action": "set-number-format", "sheet_name": "Plain",
                              "range_address": "A1", "not_a_real_parameter": 1},
             False, r"Unexpected argument|Accepted argument|rejected the arguments"),
        Case("RANGE-VALUES-BADSHAPE", "error-surface", "set-values with a ragged 2-D array",
             "range", {"action": "set-values", "sheet_name": "Plain",
                       "range_address": "A1:B2", "values": [[1, 2], [3]]},
             False, r"ragged|dimension|shape|length|count|rejected"),

        # --- list wire format: the flagship action sits on the quoted-JSON-array side --------
        Case("PARAM-LIST-NATIVE-ARRAY", "error-surface",
             "native array sent where the schema declares a JSON string",
             "slicer", {"action": "connect-pivots", "slicer_name": "Nope",
                        "pivot_table_names": ["A", "B"]},
             False, r"was sent as a JSON array"),

        # --- slicer guards (also covered by the Core-level SlicerLink tests) ----------------
        Case("SLICER-CONNECT-BLANK", "slicer", "connect-pivots with only blank names",
             "slicer", {"action": "connect-pivots", "slicer_name": "Nope",
                        "pivot_table_names": "[\"\", \"   \"]"},
             False, r"non-empty"),
        Case("SLICER-CONNECT-NOPIVOT", "slicer", "connect-pivots to a missing slicer",
             "slicer", {"action": "connect-pivots", "slicer_name": "NoSuchSlicer",
                        "pivot_table_names": "[\"NoSuchPivot\"]"},
             False, r"Slicer 'NoSuchSlicer' not found"),
        Case("SLICER-DISCONNECT-BLANK", "slicer", "disconnect-pivots with only blank names",
             "slicer", {"action": "disconnect-pivots", "slicer_name": "Nope",
                        "pivot_table_names": "[\"  \"]"},
             False, r"non-empty"),
        Case("SLICER-DISCONNECT-NOSLICER", "slicer", "disconnect-pivots on a missing slicer",
             "slicer", {"action": "disconnect-pivots", "slicer_name": "NoSuchSlicer",
                        "pivot_table_names": "[\"P1\"]"},
             False, r"not found|NoSuchSlicer"),
    ]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--list", action="store_true")
    args = ap.parse_args()

    cases = build_cases()
    if args.list:
        for c in cases:
            print(f"{c.cid:<28} {c.group:<14} {c.what}")
        return 0

    stamp = time.strftime("%H%M%S")
    out = ROOT / "_demo" / f"edge-{stamp}.xlsx"
    server = Server()
    print(f"edge-case harvest -> {out.name}")
    print("=" * 78)
    try:
        server.start()
        session = server.call("file", {"action": "create", "path": str(out), "show": False})["session_id"]

        # Fixture shapes: TextOnly has no numeric column; FourNum has four, which is one more than
        # StockHLC accepts (HLC needs exactly three).
        fixtures = {
            "Plain": ("A1:D7", [["Region", "Q1", "Q2", "Q3"],
                                ["East", 100, 130, 120], ["South", 200, 170, 180],
                                ["North", 300, 240, 260], ["West", 180, 260, 250],
                                ["Centre", 90, 150, 140], ["Northeast", 140, 110, 130]]),
            "With Space": ("A1:D7", [["Region", "Q1", "Q2", "Q3"],
                                     ["East", 100, 130, 120], ["South", 200, 170, 180],
                                     ["North", 300, 240, 260], ["West", 180, 260, 250],
                                     ["Centre", 90, 150, 140], ["Northeast", 140, 110, 130]]),
            "O'Brien": ("A1:D7", [["Region", "Q1", "Q2", "Q3"],
                                  ["East", 100, 130, 120], ["South", 200, 170, 180],
                                  ["North", 300, 240, 260], ["West", 180, 260, 250],
                                  ["Centre", 90, 150, 140], ["Northeast", 140, 110, 130]]),
            "TextOnly": ("A1:C7", [["Label", "A", "B"], ["a", "x", "y"], ["b", "x", "y"],
                                   ["c", "z", "y"], ["d", "x", "z"], ["e", "y", "x"],
                                   ["f", "z", "z"]]),
            "FourNum": ("A1:E7", [["Region", "Q1", "Q2", "Q3", "Q4"],
                                  ["East", 100, 130, 120, 110], ["South", 200, 170, 180, 190],
                                  ["North", 300, 240, 260, 250], ["West", 180, 260, 250, 240],
                                  ["Centre", 90, 150, 140, 130], ["Northeast", 140, 110, 130, 120]]),
        }
        for name, (address, values) in fixtures.items():
            server.call("worksheet", {"action": "create", "session_id": session,
                                      "sheet_name": name})
            server.call("range", {"action": "set-values", "session_id": session,
                                  "sheet_name": name, "range_address": address,
                                  "values": values})

        by_group: dict[str, list[Case]] = {}
        for case in cases:
            by_group.setdefault(case.group, []).append(case)

        for group, items in by_group.items():
            print()
            print(f"--- {group} " + "-" * (74 - len(group)))
            for case in items:
                run(server, session, case)

        server.call("file", {"action": "close", "session_id": session, "save": True})
    finally:
        server.proc.terminate()

    print()
    print("=" * 78)
    ok = [c for c, good, _ in RESULTS if good]
    bad = [(c, d) for c, good, d in RESULTS if not good]
    print(f"SUMMARY: {len(ok)}/{len(RESULTS)} matched expectation")
    if bad:
        print()
        print("MISMATCHES (each one is either a real defect or a wrong expectation):")
        for case, detail in bad:
            print(f"  {case.cid}  [{case.group}]")
            print(f"    expected: {'success' if case.expect_ok else 'failure'}"
                  + (f" matching /{case.expect_regex}/" if case.expect_regex else ""))
            if case.note:
                print(f"    note    : {case.note}")
            print(f"    observed: {detail!r}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())

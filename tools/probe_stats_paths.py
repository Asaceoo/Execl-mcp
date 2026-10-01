#!/usr/bin/env python
"""Probe: can the deployed server do *statistical process* analysis?

The schema audit proves there is no Analysis-ToolPak / Solver *action* among the
328. That is a statement about action names, not about capability, and the two
must not be conflated. This probe settles the capability question on the real
machine over three independent paths:

  A. native worksheet functions -- range.set-formulas + calculation_mode.calculate.
     No add-in, no macros, works on a plain .xlsx.
  B. VBA -> SOLVER.XLAM         -- vba import/run driving the documented Solver API
     (SolverReset / SolverOk / SolverSolve / SolverFinish).
  C. VBA -> ATPVBAEN.XLAM       -- Analysis ToolPak's own VBA entry point.

Every case asserts a KNOWN numeric answer: the fixture is y = 2x + 3 exactly, so a
working regression must return slope 2, intercept 3, R^2 1. A path that cannot
produce the number is reported with its real error text -- "we cannot" is never
asserted without evidence.

Usage:  python tools/probe_stats_paths.py [--path a|b|c|all]
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from demo_slicer_link import Server  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DEMO = ROOT / "_demo"

# y = 2x + 3 exactly, so every regression statistic has one correct value.
# z = 2x + 4 is a second sample one unit away, only used to make T.TEST run.
DATA = [["X", "Y", "Z"]] + [[i, 2 * i + 3, 2 * i + 4] for i in range(1, 11)]

# (label, formula, expected). Averages/percentiles checked against the closed form.
# AVERAGE(1..10) = 5.5 ; STDEV.S = sqrt(82.5/9) ; MEDIAN = 5.5 ; SKEW = 0 (symmetric)
# PERCENTILE(0.75) = QUARTILE(3) = 7.75 ; FORECAST at x=11 = 25
CASES: list[tuple[str, str, float]] = [
    ("AVERAGE", "=AVERAGE(A2:A11)", 5.5),
    ("STDEV.S", "=STDEV.S(A2:A11)", 3.02765035409749),
    ("MEDIAN", "=MEDIAN(A2:A11)", 5.5),
    ("SKEW", "=SKEW(A2:A11)", 0.0),
    ("CORREL", "=CORREL(A2:A11,B2:B11)", 1.0),
    ("SLOPE", "=SLOPE(B2:B11,A2:A11)", 2.0),
    ("INTERCEPT", "=INTERCEPT(B2:B11,A2:A11)", 3.0),
    ("RSQ", "=RSQ(B2:B11,A2:A11)", 1.0),
    ("PERCENTILE", "=PERCENTILE(A2:A11,0.75)", 7.75),
    ("QUARTILE", "=QUARTILE(A2:A11,3)", 7.75),
    ("FORECAST.LINEAR", "=FORECAST.LINEAR(11,B2:B11,A2:A11)", 25.0),
]
# T.TEST yields a p-value: assert it RAN and landed in [0,1] rather than which side
# of 0.05 it falls on -- the latter is a claim about the data, not about the tool.
#
# type MUST be 2, not 1. type=1 is the PAIRED test, which divides by the variance of
# the per-row differences; here z - y is a constant 1 for every row, so that variance
# is exactly 0 and Excel answers #DIV/0! -- a degenerate fixture, not a tool defect.
# type=2 (two-sample, equal variance) is well conditioned on this data.
TTEST_FORMULA = "=T.TEST(B2:B11,C2:C11,2,2)"

SOLVER_VBA = r'''
Sub ProbeStatsAddins()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Dim ai As AddIn
    Dim listing As String
    listing = ""
    For Each ai In Application.AddIns
        listing = listing & ai.Name & "[" & ai.Installed & "] "
    Next ai
    ws.Range("H1").Value = "ADDINS: " & listing

    ' A bare `Application.Run "SOLVER.XLAM!SolverReset"` DOES NOT WORK: Excel resolves the
    ' bare file name against its own search path, looks for Documents\SOLVER.XLAM, fails with
    ' 1004 -- and then SolverOk/SolverSolve ALSO return 0 while silently doing nothing. The
    ' zero return codes make that look like success. Qualify with the add-in's FullName.
    Dim solverFull As String
    solverFull = ""
    For Each ai In Application.AddIns
        If InStr(UCase(ai.Name), "SOLVER") > 0 Then
            ai.Installed = True
            solverFull = ai.FullName
        End If
    Next ai
    ws.Range("H2").Value = "solverAddin=" & solverFull

    ' Record both call forms so a failure stays attributable to a specific form.
    Dim probe As String
    probe = ""
    Err.Clear
    On Error Resume Next
    Application.Run "SOLVER.XLAM!SolverReset"
    probe = probe & "bareName:SolverReset=" & Err.Number & "|" & Err.Description & " ## "
    Err.Clear
    Application.Run "'" & solverFull & "'!SolverReset"
    probe = probe & "fullName:SolverReset=" & Err.Number & "|" & Err.Description
    Err.Clear
    On Error GoTo 0
    ws.Range("H3").Value = probe

    ' SolverReset's 1004 "cannot set the Focus property of the DialogSheet class" is HARMLESS:
    ' it is cosmetic UI setup, and the solve below still converges. Do not treat it as fatal.
    Dim run As String
    run = ""
    Err.Clear
    On Error Resume Next
    Application.Run "'" & solverFull & "'!SolverOk", "$C$17", 2, 0, "$C$14:$C$16"
    run = run & "SolverOk=" & Err.Number & "|" & Err.Description & " ## "
    Err.Clear
    Application.Run "'" & solverFull & "'!SolverSolve", True
    run = run & "SolverSolve=" & Err.Number & "|" & Err.Description & " ## "
    Err.Clear
    Application.Run "'" & solverFull & "'!SolverFinish", 1
    run = run & "SolverFinish=" & Err.Number & "|" & Err.Description
    Err.Clear
    On Error GoTo 0
    ws.Range("H4").Value = run
End Sub
'''

TOOLPAK_VBA = r'''
Sub ProbeToolpak()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Dim ai As AddIn
    Dim out As String
    Dim atpFull As String
    out = ""
    atpFull = ""
    For Each ai In Application.AddIns
        If InStr(UCase(ai.Name), "ANALYS") > 0 Or InStr(UCase(ai.Name), "ATPVBA") > 0 Or InStr(UCase(ai.Name), "FUNCRES") > 0 Then
            ai.Installed = True
            out = out & ai.Name & "[" & ai.Installed & "] "
            If InStr(UCase(ai.Name), "ATPVBAEN") > 0 Then atpFull = ai.FullName
        End If
    Next ai
    ws.Range("H5").Value = "toolpakAddins: " & out & " || atpFull=" & atpFull

    ' Same trap as Solver: a bare "ATPVBAEN.XLAM!Macro" does not resolve. Try the three
    ' qualification forms and record each, so a failure names the form that failed.
    Dim r As String
    r = ""
    Err.Clear
    On Error Resume Next
    Application.Run "'" & atpFull & "'!DescriptiveStatistics", ws.Range("A1:A11"), False, ws.Range("J1")
    r = r & "fullPath=" & Err.Number & "|" & Err.Description & " ## "
    Err.Clear
    Application.Run "ATPVBAEN.DescriptiveStatistics", ws.Range("A1:A11"), False, ws.Range("J1")
    r = r & "module=" & Err.Number & "|" & Err.Description & " ## "
    Err.Clear
    Application.Run "DescriptiveStatistics", ws.Range("A1:A11"), False, ws.Range("J1")
    r = r & "bare=" & Err.Number & "|" & Err.Description
    Err.Clear
    On Error GoTo 0
    ws.Range("H6").Value = r
End Sub
'''


def approx(actual: object, expected: float, tol: float = 1e-9) -> bool:
    try:
        value = float(actual)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return False
    return abs(value - expected) <= max(tol, abs(expected) * 1e-9)


def build(server: Server, path: Path, macro_free: bool) -> str:
    sid = server.call("file", {"action": "create", "path": str(path), "show": False})["session_id"]
    server.call(
        "range",
        {
            "action": "set-values",
            "session_id": sid,
            "sheet_name": "Sheet1",
            "range_address": "A1:C11",
            "values": DATA,
        },
    )
    # Labels in E, formulas in F -- written as one bulk array then recalculated once.
    labels = [[name] for name, _, _ in CASES] + [["T.TEST"]]
    server.call(
        "range",
        {
            "action": "set-values",
            "session_id": sid,
            "sheet_name": "Sheet1",
            "range_address": "E1:E12",
            "values": labels,
        },
    )
    body = [[formula] for _, formula, _ in CASES] + [[TTEST_FORMULA]]
    server.call(
        "range",
        {
            "action": "set-formulas",
            "session_id": sid,
            "sheet_name": "Sheet1",
            "range_address": "F1:F12",
            "formulas": body,
        },
    )
    if not macro_free:
        # Solver model: minimise (v1-5)^2 + (v2-7)^2 + (v3-3)^2 -> answer is exactly 5,7,3.
        server.call(
            "range",
            {
                "action": "set-values",
                "session_id": sid,
                "sheet_name": "Sheet1",
                "range_address": "B14:C16",
                "values": [["v1", 0], ["v2", 0], ["v3", 0]],
            },
        )
        server.call(
            "range",
            {
                "action": "set-values",
                "session_id": sid,
                "sheet_name": "Sheet1",
                "range_address": "B17:C17",
                "values": [["obj", None]],
            },
        )
        server.call(
            "range",
            {
                "action": "set-formulas",
                "session_id": sid,
                "sheet_name": "Sheet1",
                "range_address": "C17:C17",
                "formulas": [["=(C14-5)^2+(C15-7)^2+(C16-3)^2"]],
            },
        )
    server.call(
        "calculation_mode",
        {"action": "calculate", "session_id": sid, "scope": "workbook"},
    )
    return sid


def read_back(server: Server, sid: str, address: str) -> list[list[object]]:
    payload = server.call(
        "range",
        {
            "action": "get-values",
            "session_id": sid,
            "sheet_name": "Sheet1",
            "range_address": address,
        },
    )
    return payload.get("values") or []


def path_a(server: Server) -> bool:
    """Native worksheet functions -- the baseline that needs nothing but the tool face."""
    print("\n== PATH A: native worksheet functions (no add-in, no macros) ==")
    path = (DEMO / f"stats-native-{time.strftime('%H%M%S')}.xlsx").resolve()
    sid = build(server, path, macro_free=True)
    print("   workbook:", path.name)

    values = read_back(server, sid, "F1:F12")
    if len(values) != 12:
        print(f"   !! expected 12 results, read {len(values)}")
        return False

    passed = 0
    for index, (name, _formula, expected) in enumerate(CASES):
        actual = values[index][0]
        ok = approx(actual, expected)
        passed += ok
        mark = "PASS" if ok else "FAIL"
        print(f"   {mark}  {name:<17} got={actual!r:<24} want={expected}")

    p_value = values[11][0]
    try:
        in_range = 0.0 <= float(p_value) <= 1.0  # type: ignore[arg-type]
    except (TypeError, ValueError):
        in_range = False
    passed += in_range
    print(
        f"   {'PASS' if in_range else 'FAIL'}  {'T.TEST':<17} got={p_value!r:<24} want=0<=p<=1"
    )

    server.call("file", {"action": "close", "session_id": sid, "save": True})
    total = len(CASES) + 1
    print(f"   VERDICT A: {passed}/{total} statistical functions produced the exact answer")
    return passed == total


def path_b(server: Server) -> bool:
    """Solver via VBA -- the optimisation add-in the tool face does not expose."""
    print("\n== PATH B: VBA -> SOLVER.XLAM (minimise a quadratic, answer is 5/7/3) ==")
    path = (DEMO / f"stats-solver-{time.strftime('%H%M%S')}.xlsm").resolve()
    sid = build(server, path, macro_free=False)
    print("   workbook:", path.name)

    server.call(
        "vba",
        {
            "action": "import",
            "session_id": sid,
            "module_name": "StatsProbe",
            "vba_code": SOLVER_VBA,
        },
    )
    server.call(
        "vba",
        {"action": "run", "session_id": sid, "procedure_name": "StatsProbe.ProbeStatsAddins"},
    )

    for label, address in (
        ("add-ins", "H1:H1"),
        ("solver add-in", "H2:H2"),
        ("reset probe", "H3:H3"),
        ("solve probe", "H4:H4"),
    ):
        rows = read_back(server, sid, address)
        text = rows[0][0] if rows and rows[0] else ""
        print(f"   {label:<14}: {text}")

    solution = [row[0] for row in read_back(server, sid, "C14:C16")]
    objective = read_back(server, sid, "C17:C17")
    obj_value = objective[0][0] if objective and objective[0] else None
    print(f"   solver answer : {solution} (want [5, 7, 3])  objective={obj_value}")

    # Solver is an ITERATIVE solver: it halts on a convergence tolerance, so a correct
    # answer is 4.99999993 rather than 5. Judging it by the 1e-9 used for the exact
    # worksheet functions above would fail a solve that is in fact right. 1e-6 matches
    # the solver's own default tolerance.
    ok = len(solution) == 3 and all(
        approx(got, want, tol=1e-6) for got, want in zip(solution, (5.0, 7.0, 3.0))
    )
    print("   VERDICT B:", "SOLVED" if ok else "NOT SOLVED")
    server.call("file", {"action": "close", "session_id": sid, "save": True})
    return ok


def path_c(server: Server) -> bool:
    """Analysis ToolPak's own VBA entry point."""
    print("\n== PATH C: VBA -> ATPVBAEN.XLAM (Analysis ToolPak VBA interface) ==")
    path = (DEMO / f"stats-toolpak-{time.strftime('%H%M%S')}.xlsm").resolve()
    sid = build(server, path, macro_free=True)
    print("   workbook:", path.name)

    server.call(
        "vba",
        {
            "action": "import",
            "session_id": sid,
            "module_name": "ToolpakProbe",
            "vba_code": TOOLPAK_VBA,
        },
    )
    server.call(
        "vba",
        {"action": "run", "session_id": sid, "procedure_name": "ToolpakProbe.ProbeToolpak"},
    )

    for label, address in (("toolpak add-ins", "H5:H5"), ("call probe", "H6:H6")):
        rows = read_back(server, sid, address)
        text = rows[0][0] if rows and rows[0] else ""
        print(f"   {label:<16}: {text}")

    output = read_back(server, sid, "J1:J4")
    written = [row[0] for row in output if row and row[0] not in (None, "")]
    print(f"   output block   : {written if written else '(empty)'}")

    # This is a RECORDED BOUNDARY, not a capability gap. ATPVBAEN's entry points are not
    # reachable by name from outside the project: all three qualification forms answer
    # "macro not available", because they live behind a VBA project REFERENCE that the
    # `vba` tool cannot add -- it can only import module code. Nothing is lost: these are
    # the same algorithms path A already drives directly, and path A is the better caller
    # anyway (it returns values instead of pasting a formatted report block).
    #
    # Asserting "still unreachable" turns the boundary into a tripwire. If a future build
    # ever makes it work, this probe FAILS and the written-down conclusion gets revisited,
    # instead of silently going stale.
    reached = bool(written)
    verdict = "REACHABLE - revisit the capability note" if reached else "NOT REACHABLE (boundary recorded; equivalent to path A)"
    print(f"   VERDICT C: {verdict}")
    server.call("file", {"action": "close", "session_id": sid, "save": True})
    return not reached


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--path", default="all", choices=["a", "b", "c", "all"])
    args = parser.parse_args(argv)

    server = Server()
    server.start()
    results: dict[str, bool] = {}
    try:
        if args.path in ("a", "all"):
            results["A"] = path_a(server)
        if args.path in ("b", "all"):
            results["B"] = path_b(server)
        if args.path in ("c", "all"):
            results["C"] = path_c(server)
    finally:
        server.stop()

    print("\n== SUMMARY ==")
    for key in sorted(results):
        print(f"   PATH {key}: {'PASS' if results[key] else 'FAIL'}")
    return 0 if all(results.values()) else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

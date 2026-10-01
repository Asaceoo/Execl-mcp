#!/usr/bin/env python
"""Probe: do the `analysis` tool's documented parameter names actually work?

USER-GUIDE recipe 6 spelled them target_cell / target_value / result_cell, none of
which exist in the schema, and omitted the required table_range -- so the recipe as
written could not run. This probe drives the CORRECTED calls against real Excel and
asserts goals that have a closed-form answer, so a wrong parameter name shows up as a
real binder error rather than as prose nobody tested.

    model:  C1 = A1 * (1 + B1)  with  A1 = 100, B1 = 0.15   ->   C1 = 115
    goal-seek C1 to 200 by varying B1                       ->   B1 = 1.0 exactly
    data table: rows vary B1, columns vary A1

Usage:  python tools/probe_analysis.py
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from demo_slicer_link import Server  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DEMO = ROOT / "_demo"


def number(rows: list[list[object]], row: int = 0, col: int = 0) -> float | None:
    try:
        return float(rows[row][col])  # type: ignore[arg-type]
    except (IndexError, TypeError, ValueError):
        return None


def main() -> int:
    path = (DEMO / f"analysis-{time.strftime('%H%M%S')}.xlsx").resolve()
    server = Server()
    server.start()
    checks: list[tuple[str, bool, str]] = []
    try:
        sid = server.call(
            "file", {"action": "create", "path": str(path), "show": False}
        )["session_id"]
        print("workbook:", path.name)

        def read(address: str) -> list[list[object]]:
            return server.call("range", {
                "action": "get-values", "session_id": sid,
                "sheet_name": "Sheet1", "range_address": address,
            }).get("values") or []

        # --- model: A1 price, B1 rate, C1 = A1*(1+B1)
        server.call("range", {
            "action": "set-values", "session_id": sid, "sheet_name": "Sheet1",
            "range_address": "A1:B1", "values": [[100, 0.15]],
        })
        server.call("range", {
            "action": "set-formulas", "session_id": sid, "sheet_name": "Sheet1",
            "range_address": "C1:C1", "formulas": [["=A1*(1+B1)"]],
        })
        server.call("calculation_mode", {
            "action": "calculate", "session_id": sid, "scope": "workbook",
        })
        baseline = number(read("C1:C1"))
        # 100 * (1 + 0.15) evaluates to 114.99999999999999: 0.15 is not representable in
        # binary floating point, so an exact == would fail a model that is working fine.
        checks.append(("baseline C1 = 115", baseline is not None and abs(baseline - 115.0) < 1e-9,
                       f"got {baseline}"))

        # --- goal-seek: the schema says formula_cell / goal / changing_cell.
        print("\n-- goal-seek (formula_cell / goal / changing_cell)")
        server.call("analysis", {
            "action": "goal-seek", "session_id": sid, "sheet_name": "Sheet1",
            "formula_cell": "C1", "goal": 200, "changing_cell": "B1",
        })
        server.call("calculation_mode", {
            "action": "calculate", "session_id": sid, "scope": "workbook",
        })
        rate = number(read("B1:B1"))
        result = number(read("C1:C1"))
        print(f"   B1 -> {rate} (want 1.0)   C1 -> {result} (want 200)")
        checks.append(("goal-seek drove B1 to 1.0", rate is not None and abs(rate - 1.0) < 1e-9,
                       f"got {rate}"))
        checks.append(("goal-seek drove C1 to 200", result is not None and abs(result - 200.0) < 1e-6,
                       f"got {result}"))

        # --- scenario manager
        print("\n-- scenarios (changing_cells as a string, values as a native array)")
        server.call("analysis", {
            "action": "create-scenario", "session_id": sid, "sheet_name": "Sheet1",
            "scenario_name": "Optimistic", "changing_cells": "B1", "values": [0.3],
        })
        listed = server.call("analysis", {
            "action": "list-scenarios", "session_id": sid, "sheet_name": "Sheet1",
        })
        text = str(listed)
        found = "Optimistic" in text
        print(f"   list-scenarios contains the new name: {found}")
        checks.append(("scenario created and listed", found, text[:160]))

        server.call("analysis", {
            "action": "create-scenario-summary", "session_id": sid,
            "sheet_name": "Sheet1", "result_cells": "C1",
        })
        sheets = server.call("worksheet", {
            "action": "list", "session_id": sid,
        })
        sheet_names = [item.get("name") for item in (sheets.get("worksheets") or sheets.get("sheets") or [])
                       if isinstance(item, dict)]
        print(f"   worksheets after summary: {sheet_names}")
        checks.append(("summary produced a report sheet", len(sheet_names) > 1, str(sheet_names)))

        # --- data table: needs the top-left formula cell, row values down column D,
        # column values across row 1. This is the parameter the old recipe dropped.
        print("\n-- data table (table_range + row_input_cell + column_input_cell)")
        server.call("range", {
            "action": "set-values", "session_id": sid, "sheet_name": "Sheet1",
            "range_address": "D1:E4",
            "values": [[None, 100], [0.1, None], [0.2, None], [0.3, None]],
        })
        server.call("range", {
            "action": "set-formulas", "session_id": sid, "sheet_name": "Sheet1",
            "range_address": "D1:D1", "formulas": [["=C1"]],
        })
        server.call("analysis", {
            "action": "create-data-table", "session_id": sid, "sheet_name": "Sheet1",
            "table_range": "D1:E4", "row_input_cell": "B1", "column_input_cell": "A1",
        })
        server.call("calculation_mode", {
            "action": "calculate", "session_id": sid, "scope": "workbook",
        })
        grid = read("D1:E4")
        print(f"   table after recalc: {grid}")
        # The first row carries the row-input values (they replace B1) and the first
        # column carries the column-input ones (they replace A1), so E2 = 0.1*(1+100).
        # Asserting the actual numbers, rather than "a grid appeared", is what makes
        # this a check instead of a decoration.
        got = [number(grid, row, 1) for row in (1, 2, 3)]
        expected = [10.1, 20.2, 30.3]
        ok = all(value is not None and abs(value - want) < 1e-9
                 for value, want in zip(got, expected))
        checks.append(("data table computed 10.1 / 20.2 / 30.3", ok, f"got {got}"))

        server.call("file", {"action": "close", "session_id": sid, "save": True})
    finally:
        server.stop()

    print("\n== SUMMARY ==")
    failed = 0
    for label, ok, detail in checks:
        print(f"   {'PASS' if ok else 'FAIL'}  {label:<38} {detail}")
        failed += not ok
    print(f"   {len(checks) - failed}/{len(checks)} checks passed")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())

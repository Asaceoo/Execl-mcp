"""Measure how far slicer cross-filtering scales on a real Excel instance.

Builds one shared PivotCache, clones it into N PivotTables, creates M slicers on separate
fields, connects every slicer to every PivotTable, then proves the fan-out actually filters
by selecting one item and reading every PivotTable's sheet.

    PIVOTS=4 SLICERS=12 python tools/bench_slicer_scale.py

Exit code 0 means the requested M x N fan-out worked end to end; anything else reports the
first size that failed, so the number in the output is measured, not assumed.
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from demo_slicer_link import Server  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
PIVOTS = int(os.environ.get("PIVOTS", "4"))
SLICERS = int(os.environ.get("SLICERS", "12"))
ROWS = 12

def col_letter(index: int) -> str:
    """1 -> A, 27 -> AA. Needed because the source range grows with the slicer count."""
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# Every slicer needs its own source field, so the data carries SLICERS dimension columns.
FIELDS = [f"F{i:02d}" for i in range(1, SLICERS + 1)]
HEADER = FIELDS + ["Sales"]
LAST_COL = col_letter(len(HEADER))
DATA = [HEADER]
for r in range(ROWS):
    DATA.append([("华东" if (r + i) % 3 == 0 else "华南" if (r + i) % 3 == 1 else "华北")
                 if i == 0 else f"V{(r + i) % 3}" for i in range(SLICERS)] + [100 + r * 10])

OUTPUT = ROOT / "_demo" / f"slicer-scale-{PIVOTS}x{SLICERS}.xlsx"


def main() -> int:
    server = Server()
    server.start()
    session = None
    try:
        OUTPUT.parent.mkdir(parents=True, exist_ok=True)
        session = server.call("file", {"action": "create", "path": str(OUTPUT), "show": False})["session_id"]

        server.call("range", {
            "action": "set-values", "session_id": session,
            "sheet_name": "Sheet1", "range_address": f"A1:{LAST_COL}{ROWS + 1}",
            "values": DATA,
        })
        source = f"A1:{LAST_COL}{ROWS + 1}"

        # Destination sheets must exist - the code indexes Worksheets[name] by name.
        for name in [f"PT{i}" for i in range(1, PIVOTS + 1)] + ["Dash"]:
            server.call("worksheet", {"action": "create", "session_id": session, "sheet_name": name})

        names = []
        for i in range(1, PIVOTS + 1):
            args = {
                "action": "create-from-range", "session_id": session,
                "pivot_table_name": f"PT_{i}",
                "source_sheet": "Sheet1", "source_range": source,
                "destination_sheet": f"PT{i}", "destination_cell": "A1",
            }
            if i > 1:
                args["share_cache_from"] = "PT_1"
            server.call("pivottable", args)
            names.append(f"PT_{i}")
        print(f"PivotTables sharing one cache: {len(names)}")

        # A freshly created PivotTable has no fields laid out yet.
        for name in names:
            server.call("pivottable_field", {
                "action": "add-row-field", "session_id": session,
                "pivot_table_name": name, "field_name": FIELDS[0],
            })
            server.call("pivottable_field", {
                "action": "add-value-field", "session_id": session,
                "pivot_table_name": name, "field_name": "Sales",
            })

        slicers = []
        for idx, field in enumerate(FIELDS):
            name = f"SL_{field}"
            server.call("slicer", {
                "action": "create-slicer", "session_id": session,
                "pivot_table_name": "PT_1", "field_name": field,
                "slicer_name": name, "destination_sheet": "Dash",
                "position": f"E{1 + idx * 6}",
            })
            slicers.append(name)
        print(f"Slicers created: {len(slicers)}")

        linked = 0
        for name in slicers:
            server.call("slicer", {
                "action": "connect-pivots", "session_id": session,
                "slicer_name": name, "pivot_table_names": json.dumps(names),
            })
            linked += 1
        print(f"Slicers connected to all {PIVOTS} PivotTables: {linked}")

        listed = server.call("slicer", {"action": "list-slicers", "session_id": session})
        count = len(listed.get("slicers") or listed.get("items") or [])
        print(f"list-slicers reports: {count}")

        # Prove the fan-out filters: one selection must reach every PivotTable.
        server.call("slicer", {
            "action": "set-slicer-selection", "session_id": session,
            "slicer_name": slicers[0], "selected_items": '["华东"]', "clear_first": True,
        })
        ok = 0
        for i in range(1, PIVOTS + 1):
            values = server.call("range", {
                "action": "get-values", "session_id": session,
                "sheet_name": f"PT{i}", "range_address": "A1:B12",
            }).get("values") or []
            flat = [str(c) for row in values for c in row]
            if any("华东" in c for c in flat) and not any("华北" in c for c in flat):
                ok += 1
        print(f"PivotTables filtered by the single selection: {ok}/{PIVOTS}")
        if ok != PIVOTS:
            print("FAILED: not every PivotTable reacted to the slicer", file=sys.stderr)
            return 1

        server.call("file", {"action": "close", "session_id": session, "save": True})
        session = None
        print(f"\nOK - {SLICERS} slicers x {PIVOTS} PivotTables all linked. Saved to:\n  {OUTPUT}")
        return 0
    except Exception as error:  # noqa: BLE001 - the point is to report the real failure
        print(f"\nFAILED at {PIVOTS}x{SLICERS}: {type(error).__name__}: {error}", file=sys.stderr)
        return 1
    finally:
        if session:
            try:
                server.call("file", {"action": "close", "session_id": session, "save": False})
            except Exception:  # noqa: BLE001
                pass
        server.stop()


if __name__ == "__main__":
    sys.exit(main())

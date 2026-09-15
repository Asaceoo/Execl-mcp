"""Probe which analysis capabilities actually run on this machine.

The MCP surface advertises 31 tools / 328 actions, but "advertised" is not "works". This
script drives the ones that matter for data analysis against real Excel and reports the
outcome of each, including the real error text when it fails.

Sections:
  BASIC    - sorting, filtering, aggregation, conditional formatting, charting
  ADVANCED - what-if goal seek, two-variable data table, scenarios, DAX measures in the
             Power Pivot data model, Power Query M, Python in Excel
"""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from demo_slicer_link import Server  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
STAMP = datetime.now().strftime("%H%M%S")
OUTPUT = ROOT / "_demo" / f"analysis-probe-{STAMP}.xlsx"

SALES = [
    ["Region", "Product", "Sales"],
    ["华东", "A", 100], ["华东", "B", 150],
    ["华南", "A", 200], ["华南", "B", 250],
    ["华北", "A", 300], ["华北", "B", 350],
]


def main() -> int:
    server = Server()
    server.start()
    session = None
    results: list[tuple[str, bool, str]] = []

    def probe_any(name: str, tool: str, variants: list[dict]) -> None:
        """Try several parameter shapes and report the first that works."""
        last = ""
        for index, arguments in enumerate(variants, 1):
            try:
                server.call(tool, arguments)
                results.append((name, True, ""))
                print(f"  OK    {name} (variant {index})")
                return
            except Exception as error:  # noqa: BLE001
                last = str(error)[:200]
        results.append((name, False, last))
        print(f"  FAIL  {name}: {last}")

    def probe(name: str, tool: str, arguments: dict) -> None:
        try:
            server.call(tool, arguments)
            results.append((name, True, ""))
            print(f"  OK    {name}")
        except Exception as error:  # noqa: BLE001 - the real error text is the point
            results.append((name, False, str(error)[:220]))
            print(f"  FAIL  {name}: {str(error)[:220]}")

    try:
        session = server.call("file", {"action": "create", "path": str(OUTPUT), "show": False})["session_id"]
        server.call("worksheet", {"action": "create", "session_id": session, "sheet_name": "Data"})
        server.call("range", {"action": "set-values", "session_id": session,
                              "sheet_name": "Data", "range_address": "A1:C7", "values": SALES})
        server.call("worksheet", {"action": "create", "session_id": session, "sheet_name": "Model"})
        # Simple model: revenue = price * quantity
        server.call("range", {"action": "set-values", "session_id": session, "sheet_name": "Model",
                              "range_address": "A1:B3",
                              "values": [["单价", 100], ["数量", 5], ["收入", None]]})
        server.call("range", {"action": "set-formulas", "session_id": session, "sheet_name": "Model",
                              "range_address": "B3", "formulas": [["=B1*B2"]]})

        print("BASIC")
        # sort_columns is a real JSON array of objects, NOT a JSON string like pivot_table_names.
        probe_any("sort a range", "range_edit", [
            {"action": "sort", "session_id": session, "sheet_name": "Data",
             "range_address": "A1:C7", "sort_columns": [{"columnIndex": 3, "ascending": False}],
             "has_headers": True},
            {"action": "sort", "session_id": session, "sheet_name": "Data",
             "range_address": "A1:C7", "sort_columns": '[{"columnIndex":3,"ascending":false}]',
             "has_headers": True},
        ])
        probe("create a table", "table",
              {"action": "create", "session_id": session, "sheet_name": "Data",
               "range_address": "A1:C7", "table_name": "SalesTbl", "has_headers": True})
        probe("filter the table", "table_column",
              {"action": "apply-filter-values", "session_id": session, "sheet_name": "Data",
               "table_name": "SalesTbl", "column_name": "Region", "values": '["华东"]'})
        # A PivotTable owns its rectangle, so give every later feature its own sheet.
        for name in ("PT", "WhatIf"):
            server.call("worksheet", {"action": "create", "session_id": session, "sheet_name": name})
        probe("aggregate with a PivotTable", "pivottable",
              {"action": "create-from-range", "session_id": session, "pivot_table_name": "PT1",
               "source_sheet": "Data", "source_range": "A1:C7",
               "destination_sheet": "PT", "destination_cell": "A1"})
        probe("conditional format a column", "conditionalformat",
              {"action": "add-rule", "session_id": session, "sheet_name": "Data",
               "range_address": "C2:C7", "rule_type": "data-bar", "data_bar_color": "#1F4E79"})
        probe("chart from the table", "chart",
              {"action": "create-from-table", "session_id": session, "table_name": "SalesTbl",
               "chart_name": "C1", "sheet_name": "Model", "chart_type": "ColumnClustered"})

        print("\nADVANCED - what-if")
        probe("goal seek (solve price for revenue=1000)", "analysis",
              {"action": "goal-seek", "session_id": session, "sheet_name": "Model",
               "formula_cell": "B3", "goal": 1000, "changing_cell": "B1"})
        # Excel requires the data table's input cells on the same sheet as the table itself.
        server.call("range", {"action": "set-values", "session_id": session, "sheet_name": "WhatIf",
                              "range_address": "A1:B3", "values": [["单价", 100], ["数量", 5], ["收入", None]]})
        server.call("range", {"action": "set-formulas", "session_id": session, "sheet_name": "WhatIf",
                              "range_address": "B3", "formulas": [["=B1*B2"]]})
        server.call("range", {"action": "set-values", "session_id": session, "sheet_name": "WhatIf",
                              "range_address": "A5:E9",
                              "values": [["=B3", 4, 5, 6, 7], [80, None, None, None, None],
                                         [100, None, None, None, None], [120, None, None, None, None],
                                         [140, None, None, None, None]]})
        probe("two-variable data table (price x quantity)", "analysis",
              {"action": "create-data-table", "session_id": session, "sheet_name": "WhatIf",
               "table_range": "A5:E9", "row_input_cell": "B1", "column_input_cell": "B2"})
        probe("create a scenario", "analysis",
              {"action": "create-scenario", "session_id": session, "sheet_name": "Model",
               "scenario_name": "乐观", "changing_cells": "B1:B2", "values": [120, 8]})
        probe("list scenarios", "analysis",
              {"action": "list-scenarios", "session_id": session, "sheet_name": "Model"})

        print("\nADVANCED - Power Pivot / DAX")
        probe("load the table into the data model", "table",
              {"action": "add-to-data-model", "session_id": session, "table_name": "SalesTbl"})
        probe("list data-model tables", "datamodel",
              {"action": "list-tables", "session_id": session})
        probe("create a DAX measure", "datamodel",
              {"action": "create-measure", "session_id": session, "table_name": "SalesTbl",
               "measure_name": "总销售额", "dax_formula": "SUM([Sales])"})
        probe("evaluate a DAX query", "datamodel",
              {"action": "evaluate", "session_id": session,
               "dax_query": "EVALUATE ROW(\"总额\", SUM('SalesTbl'[Sales]))"})

        print("\nADVANCED - Power Query / Python")
        probe("evaluate M code", "powerquery",
              {"action": "evaluate", "session_id": session, "m_code": "let x = 1 + 1 in x"})
        probe("Python in Excel", "pythoninexcel",
              {"action": "set-formula", "session_id": session, "sheet_name": "Model",
               "range_address": "K1", "code": "1 + 1"})

        server.call("file", {"action": "close", "session_id": session, "save": True})
        session = None

        passed = sum(1 for _, ok, _ in results if ok)
        print(f"\n=== {passed}/{len(results)} probes succeeded ===")
        for name, ok, err in results:
            if not ok:
                print(f"  FAILED: {name}\n          {err}")
        print(f"workbook -> {OUTPUT}")
        return 0
    finally:
        if session:
            try:
                server.call("file", {"action": "close", "session_id": session, "save": False})
            except Exception:  # noqa: BLE001
                pass
        server.stop()


if __name__ == "__main__":
    sys.exit(main())

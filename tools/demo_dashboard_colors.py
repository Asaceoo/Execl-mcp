"""Build a small styled dashboard on real Excel and prove the colour pipeline works.

Answers two practical questions with evidence instead of guesses:
  1. can the board's colours be driven from the MCP tool surface?
  2. what integer encoding do the *_color parameters expect?

Excel's Interior.Color is a VBA-style RGB long (red in the low byte), so the script writes
pure red (255), green (65280) and blue (16711680) swatches alongside the real styling and
captures a screenshot. If the swatches render as red/green/blue the encoding is confirmed.
"""

from __future__ import annotations

import base64
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from demo_slicer_link import Server  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
# Never delete an existing artefact - this sandbox blocks removal, so roll to a new name.
OUTPUT = ROOT / "_demo" / "dashboard-colors.xlsx"
SHOT = ROOT / "_demo" / "dashboard-colors.png"
_STAMP = datetime.now().strftime("%H%M%S")
if OUTPUT.exists():
    OUTPUT = OUTPUT.with_name(f"dashboard-colors-{_STAMP}.xlsx")
    SHOT = SHOT.with_name(f"dashboard-colors-{_STAMP}.png")

DATA = [
    ["Region", "Month", "Sales"],
    ["华东", "1月", 120], ["华东", "2月", 150], ["华东", "3月", 180],
    ["华南", "1月", 200], ["华南", "2月", 170], ["华南", "3月", 210],
    ["华北", "1月", 90], ["华北", "2月", 130], ["华北", "3月", 160],
]

# Colour parameters are STRINGS: "#RRGGBB" (or a numeric colour index as text). FormattingHelpers
# turns #RRGGBB into Excel's RGB long, where red is the low byte.
NAVY = "#1F4E79"
LIGHT = "#DDEBF7"
WHITE = "#FFFFFF"
ACCENT = "#C00000"


def step(name: str, fn) -> None:
    try:
        fn()
        print(f"  OK    {name}")
    except Exception as error:  # noqa: BLE001 - report and keep building the board
        print(f"  FAIL  {name}: {str(error)[:160]}")


def main() -> int:
    server = Server()
    server.start()
    session = None
    try:
        OUTPUT.parent.mkdir(parents=True, exist_ok=True)
        session = server.call("file", {"action": "create", "path": str(OUTPUT), "show": False})["session_id"]

        server.call("range", {"action": "set-values", "session_id": session,
                              "sheet_name": "Sheet1", "range_address": "A1:C10", "values": DATA})
        for name in ("PT", "Dash"):
            server.call("worksheet", {"action": "create", "session_id": session, "sheet_name": name})

        server.call("pivottable", {"action": "create-from-range", "session_id": session,
                                   "pivot_table_name": "PT_Sales", "source_sheet": "Sheet1",
                                   "source_range": "A1:C10", "destination_sheet": "PT",
                                   "destination_cell": "A1"})
        for action, field in (("add-row-field", "Region"), ("add-column-field", "Month"),
                              ("add-value-field", "Sales")):
            server.call("pivottable_field", {"action": action, "session_id": session,
                                             "pivot_table_name": "PT_Sales", "field_name": field})
        server.call("slicer", {"action": "create-slicer", "session_id": session,
                               "pivot_table_name": "PT_Sales", "field_name": "Region",
                               "slicer_name": "RegionSlicer", "destination_sheet": "Dash",
                               "position": "B2"})

        print("styling:")
        step("worksheet tab colour (navy)", lambda: server.call(
            "worksheet_style", {"action": "set-tab-color", "session_id": session,
                                "sheet_name": "Dash", "red": 31, "green": 78, "blue": 121}))

        step("title band fill + white bold font", lambda: server.call(
            "range_format", {"action": "format-range", "session_id": session, "sheet_name": "Dash",
                             "range_address": "A1:H1", "fill_color": NAVY, "font_color": WHITE,
                             "font_size": 16, "bold": True}))

        step("KPI row fill + accent font", lambda: server.call(
            "range_format", {"action": "format-range", "session_id": session, "sheet_name": "Dash",
                             "range_address": "A20:H20", "fill_color": LIGHT, "font_color": ACCENT,
                             "bold": True}))

        step("conditional data bars on the source column", lambda: server.call(
            "conditionalformat", {"action": "add-rule", "session_id": session, "sheet_name": "Sheet1",
                                  "range_address": "C2:C10", "rule_type": "data-bar",
                                  "data_bar_color": NAVY}))

        step("chart from the PivotTable", lambda: server.call(
            "chart", {"action": "create-from-pivottable", "session_id": session,
                      "pivot_table_name": "PT_Sales", "chart_name": "SalesChart",
                      "sheet_name": "Dash", "chart_type": "ColumnClustered"}))

        step("chart series fill", lambda: server.call(
            "chart_config", {"action": "set-series-format", "session_id": session,
                              "sheet_name": "Dash", "chart_name": "SalesChart",
                              "series_index": 1, "fill_color": NAVY}))

        # Encoding probe: these three must render as red / green / blue.
        for col, value in (("A", "#FF0000"), ("B", "#00FF00"), ("C", "#0000FF")):
            step(f"swatch {col}24 = {value}", lambda c=col, v=value: server.call(
                "range_format", {"action": "format-range", "session_id": session,
                                 "sheet_name": "Dash", "range_address": f"{c}24", "fill_color": v}))

        # The screenshot comes back as an image content part; the text part is only a caption.
        response = server.request("tools/call", {"name": "screenshot", "arguments": {
            "action": "capture", "session_id": session,
            "sheet_name": "Dash", "range_address": "A1:H24"}})
        image = next((part.get("data") for part in response.get("result", {}).get("content", [])
                      if part.get("type") == "image" and part.get("data")), None)
        if image:
            Path(SHOT).write_bytes(base64.b64decode(image))
            print(f"\nscreenshot -> {SHOT}")
        else:
            print(f"\nscreenshot returned no image payload: {json.dumps(response)[:200]}")

        server.call("file", {"action": "close", "session_id": session, "save": True})
        session = None
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

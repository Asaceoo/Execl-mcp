"""Root-cause check: is the modern-chart failure Excel's fault or ours?

The MCP server creates every chart with Shapes.AddChart(). That entry point predates
Excel 2016, so the modern chart types (treemap, waterfall, funnel, histogram, pareto,
box-whisker, region map, combo) and stock charts may simply be unsupported there.

This script talks to Excel directly and tries BOTH entry points per type, with the data
layout each type actually requires:
  A) Shapes.AddChart(type, ...)        - what ExcelMcp uses today
  B) Shapes.AddChart2(-1, type, ...)   - Excel 2013+ entry point (already used for
                                         PivotCharts in ChartCommands.Lifecycle.cs)
"""

from __future__ import annotations

import sys

import pythoncom
import win32com.client

# name, code, source range (data layout matters: stock charts need numeric columns only,
# histogram/pareto/box-whisker need a single numeric block)
CASES = [
    ("Treemap", 117, "A1:C7"), ("Sunburst", 116, "A1:C7"),
    ("Histogram", 118, "B1:C7"), ("Pareto", 122, "B1:C7"),
    ("BoxWhisker", 121, "B1:C7"), ("Waterfall", 119, "A1:C7"),
    ("Funnel", 123, "A1:C7"), ("RegionMap", 140, "A1:C7"),
    ("ColumnLineCombo", 120, "A1:C7"),
    ("StockHLC", 88, "E1:G7"), ("StockOHLC", 89, "E1:H7"),
    ("StockVHLC", 90, "E1:H7"), ("StockVOHLC", 91, "E1:I7"),
    ("ColumnClustered", 51, "A1:C7"),
]

DATA = [
    ["Region", "Q1", "Q2"],
    ["华东", 100, 130], ["华南", 200, 170], ["华北", 300, 240],
    ["西南", 180, 260], ["东北", 90, 150], ["西北", 140, 110],
]
NUMERIC = [
    [320, 100, 130, 90, 100], [240, 200, 170, 180, 200], [180, 300, 240, 250, 300],
    [420, 180, 260, 160, 180], [150, 90, 150, 80, 90], [260, 140, 110, 120, 140],
]


def try_create(ws, source: str, method: str, type_code: int) -> str:
    try:
        if method == "AddChart":
            shape = ws.Shapes.AddChart(type_code, 10, 10, 300, 200)
        else:
            shape = ws.Shapes.AddChart2(-1, type_code, 10, 10, 300, 200)
    except Exception as error:  # noqa: BLE001
        return f"FAIL create: {str(error)[:44]}"
    try:
        shape.Chart.SetSourceData(ws.Range(source), 2)  # 2 = xlColumns
        actual = int(shape.Chart.ChartType)
        shape.Delete()
        return "OK" if actual == type_code else f"OK but type={actual}"
    except Exception as error:  # noqa: BLE001
        try:
            shape.Delete()
        except Exception:  # noqa: BLE001
            pass
        return f"FAIL source: {str(error)[:44]}"


def main() -> int:
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = None
    try:
        print(f"Excel {excel.Version}")
        book = excel.Workbooks.Add()
        ws = book.Worksheets(1)
        ws.Name = "Probe"
        ws.Range("A1:C7").Value = DATA
        ws.Range("E1:I7").Value = NUMERIC

        print(f"{'type':<17}{'code':<6}{'AddChart (MCP 现状)':<36}AddChart2")
        for name, code, source in CASES:
            a = try_create(ws, source, "AddChart", code)
            b = try_create(ws, source, "AddChart2", code)
            print(f"{name:<17}{code:<6}{a:<36}{b}")
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()
    return 0


if __name__ == "__main__":
    sys.exit(main())

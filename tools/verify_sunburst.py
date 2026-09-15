"""Is Sunburst (116) reachable at all on this Excel build?

AddChart2(-1, 116) creates a chart but reports ChartType -4111, so the type request is refused.
Two alternatives worth trying: convert a working treemap to sunburst, and pass Style/NewLayout.
"""

from __future__ import annotations

import pythoncom
import win32com.client

DATA = [
    ["大区", "城市", "销售额"],
    ["华东", "上海", 320], ["华东", "杭州", 210], ["华东", "南京", 180],
    ["华南", "广州", 280], ["华南", "深圳", 350],
]


def attach(chart, ws, text_cols: int) -> None:
    series = chart.SeriesCollection().NewSeries()
    series.Values = ws.Range("C2:C6")
    series.XValues = ws.Range("A2:A6")


def report(name: str, shape, requested: int) -> None:
    try:
        actual = int(shape.Chart.ChartType)
        print(f"  {name}: type={actual} {'OK' if actual == requested else '回落'}")
    except Exception as error:  # noqa: BLE001
        print(f"  {name}: cannot read type ({str(error)[:40]})")


def main() -> int:
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = None
    try:
        book = excel.Workbooks.Add()
        ws = book.Worksheets(1)
        ws.Range("A1:C6").Value = DATA

        shape = ws.Shapes.AddChart2(-1, 117, 10, 10, 300, 200)          # treemap first
        attach(shape.Chart, ws, 1)
        print("H. 树图 -> 改类型 116")
        try:
            shape.Chart.ChartType = 116
        except Exception as error:  # noqa: BLE001
            print(f"  setting type failed: {str(error)[:60]}")
        report("H", shape, 116)
        shape.Delete()

        shape = ws.Shapes.AddChart2(201, 116, 10, 10, 300, 200, True)   # explicit style + new layout
        attach(shape.Chart, ws, 1)
        print("I. AddChart2(style=201, 116, NewLayout)")
        report("I", shape, 116)
        shape.Delete()

        shape = ws.Shapes.AddChart2(-1, 116, 10, 10, 300, 200)
        attach(shape.Chart, ws, 2)
        print("J. AddChart2(-1, 116) + 层级 XValues(两列)")
        report("J", shape, 116)
        shape.Delete()

        print("reference: AddChart2(-1, 117) treemap")
        shape = ws.Shapes.AddChart2(-1, 117, 10, 10, 300, 200)
        attach(shape.Chart, ws, 1)
        report("ref", shape, 117)
        shape.Delete()
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

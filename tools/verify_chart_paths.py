"""Which creation path can Excel actually use for the modern chart types?

Path C: create a plain column chart, switch ChartType, then attach the source.
Path D: create a plain column chart, attach the source, then switch ChartType.

If either path works, ExcelMcp can support these types without touching Shapes.AddChart.
"""

from __future__ import annotations

import pythoncom
import win32com.client

DATA = [
    ["Region", "Q1", "Q2"],
    ["华东", 100, 130], ["华南", 200, 170], ["华北", 300, 240],
    ["西南", 180, 260], ["东北", 90, 150], ["西北", 140, 110],
]
NUM = [
    [320, 100, 130, 90, 100], [240, 200, 170, 180, 200], [180, 300, 240, 250, 300],
    [420, 180, 260, 160, 180], [150, 90, 150, 80, 90], [260, 140, 110, 120, 140],
]
CASES = [
    ("Treemap", 117, "A1:C7"), ("Sunburst", 116, "A1:C7"),
    ("Histogram", 118, "B1:C7"), ("Pareto", 122, "B1:C7"),
    ("BoxWhisker", 121, "B1:C7"), ("Waterfall", 119, "A1:C7"),
    ("Funnel", 123, "A1:C7"), ("RegionMap", 140, "A1:C7"),
    ("ColumnLineCombo", 120, "A1:C7"),
    ("StockHLC", 88, "E1:G7"), ("StockOHLC", 89, "E1:H7"),
]


def run(ws, source: str, code: int, mode: str) -> str:
    try:
        shape = ws.Shapes.AddChart(51, 10, 10, 300, 200)
        chart = shape.Chart
        if mode == "C":
            chart.ChartType = code
            chart.SetSourceData(ws.Range(source), 2)
        else:
            chart.SetSourceData(ws.Range(source), 2)
            chart.ChartType = code
        actual = int(chart.ChartType)
        shape.Delete()
        return "OK" if actual == code else f"回落{actual}"
    except Exception as error:  # noqa: BLE001
        return "FAIL " + str(error).replace("\n", " ")[:30]


def main() -> int:
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = None
    try:
        print(f"Excel {excel.Version} build {excel.Build}")
        book = excel.Workbooks.Add()
        ws = book.Worksheets(1)
        ws.Range("A1:C7").Value = DATA
        ws.Range("E1:I7").Value = NUM
        print(f"{'type':<17}{'C: 先改类型再设数据源':<34}D: 先设数据源再改类型")
        for name, code, source in CASES:
            print(f"{name:<17}{run(ws, source, code, 'C'):<34}{run(ws, source, code, 'D')}")
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

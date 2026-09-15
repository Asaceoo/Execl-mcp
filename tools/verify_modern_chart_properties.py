"""Which COM member fails on a modern chart?

ChartCommands.Read walks a chart's properties behind `catch (COMException)` guards, but the
late binder reports a *missing* member as NotImplementedException (E_NOTIMPL, 0x80004001),
which slips through every one of those guards and kills the whole call.

This script visits each property the read path touches, on both a classic chart and a modern
one, and prints the HRESULT that comes back so the exact member can be identified.
"""

from __future__ import annotations

import pythoncom
import win32com.client

DATA = [
    ["Region", "Q1", "Q2"],
    ["华东", 100, 130], ["华南", 200, 170], ["华北", 300, 240],
    ["西南", 180, 260], ["东北", 90, 150], ["西北", 140, 110],
]

PROBES = [
    ("shape.Type", lambda s, c: s.Type),
    ("shape.Name", lambda s, c: s.Name),
    ("shape.Left", lambda s, c: s.Left),
    ("shape.Top", lambda s, c: s.Top),
    ("shape.Width", lambda s, c: s.Width),
    ("shape.Height", lambda s, c: s.Height),
    ("shape.TopLeftCell", lambda s, c: s.TopLeftCell.Address()),
    ("shape.BottomRightCell", lambda s, c: s.BottomRightCell.Address()),
    ("shape.Placement", lambda s, c: s.Placement),
    ("chart.ChartType", lambda s, c: c.ChartType),
    ("chart.PivotLayout", lambda s, c: c.PivotLayout),
    ("chart.HasTitle", lambda s, c: c.HasTitle),
    ("chart.ChartTitle", lambda s, c: c.ChartTitle.Text),
    ("chart.HasLegend", lambda s, c: c.HasLegend),
    ("chart.SeriesCollection().Count", lambda s, c: c.SeriesCollection().Count),
    ("chart.ChartArea", lambda s, c: c.ChartArea.Width),
    ("chart.ChartArea.Parent", lambda s, c: c.ChartArea.Parent.Name),
    ("chart.Axes(xlValue)", lambda s, c: c.Axes(2).MinimumScale),
    ("series.Name", lambda s, c: c.SeriesCollection(1).Name),
    ("series.Values", lambda s, c: c.SeriesCollection(1).Values),
    ("series.XValues", lambda s, c: c.SeriesCollection(1).XValues),
    ("series.Formula", lambda s, c: c.SeriesCollection(1).Formula),
    ("chart.ChartArea.Parent.SeriesCollection(1).Formula",
     lambda s, c: c.ChartArea.Parent.SeriesCollection(1).Formula),
    ("chart.PlotArea", lambda s, c: c.PlotArea.Width),
]


def probe(label: str, ws, chart_type: int) -> None:
    shape = ws.Shapes.AddChart2(-1, chart_type, 10, 10, 300, 200)
    chart = shape.Chart
    if chart_type in (117, 119, 123, 118, 121, 122, 140, 120):
        series = chart.SeriesCollection().NewSeries()
        series.Values = ws.Range("B2:B7")
        series.XValues = ws.Range("A2:A7")
    else:
        chart.SetSourceData(ws.Range("A1:C7"), 2)

    print(f"--- chart type {chart_type}")
    for name, accessor in PROBES:
        try:
            accessor(shape, chart)
            print(f"    ok    {name}")
        except Exception as error:  # noqa: BLE001 - the HRESULT is the finding
            detail = str(error)
            hresult = getattr(error, "hresult", None)
            code = f"0x{hresult & 0xFFFFFFFF:08X}" if isinstance(hresult, int) else "?"
            marker = "E_NOTIMPL" if hresult == -2147467263 else ""
            print(f"    FAIL  {name}: {code} {marker} {detail[:60]}")
    shape.Delete()


def main() -> int:
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    book = None
    try:
        book = excel.Workbooks.Add()
        ws = book.Worksheets(1)
        ws.Range("A1:C7").Value = DATA
        probe("classic", ws, 51)
        probe("waterfall", ws, 119)
        probe("treemap", ws, 117)
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

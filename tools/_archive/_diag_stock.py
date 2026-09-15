"""Replicate ChartTestsFixture layout to find which stock-chart call throws."""
import pythoncom
import win32com.client

# 复刻 ChartTestsFixture 布局：A1:E6，5 列数值
FIX = [
    ["X", "Y", "Series2", "Series3", "Series4"],
    [1, 100, 15, 50, 80], [2, 150, 25, 75, 90], [3, 200, 35, 100, 110],
    [4, 250, 45, 125, 120], [5, 300, 55, 150, 130],
]
CASES = [
    ("StockHLC", 88, "A1:C6"),
    ("StockOHLC", 89, "A1:D6"),
    ("StockVHLC", 90, "A1:D6"),
    ("StockVOHLC", 91, "A1:E6"),
]

pythoncom.CoInitialize()
xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
bk = None
try:
    bk = xl.Workbooks.Add()
    ws = bk.Worksheets(1)
    ws.Range("A1:E6").Value = FIX
    for name, code, src in CASES:
        steps = []
        sh = None
        try:
            sh = ws.Shapes.AddChart(51, 10, 10, 300, 200)
            steps.append("AddChart51:OK")
            ch = sh.Chart
            ch.SetSourceData(ws.Range(src))  # 不传 PlotBy，与 C# 一致
            steps.append("SetSourceData:OK")
            ch.ChartType = code
            steps.append(f"ChartType={code}:OK")
            actual = int(ch.ChartType)
            sh.Delete()
            steps.append(f"readback={actual} {'PASS' if actual == code else 'FELL BACK'}")
        except Exception as e:  # noqa: BLE001
            hr = hex(e.hresult) if hasattr(e, "hresult") else "?"
            last = steps[-1] if steps else "?"
            steps.append(f"FAIL@{last} hr={hr} {str(e)[:70]}")
            try:
                if sh is not None:
                    sh.Delete()
            except Exception:  # noqa: BLE001
                pass
        print(f"{name:<12}{src:<10}", " | ".join(steps))
finally:
    if bk is not None:
        bk.Close(SaveChanges=False)
    xl.Quit()
    pythoncom.CoUninitialize()

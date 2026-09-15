import pythoncom, win32com.client
PATH = r"D:\execl-mcp\_demo\slicer-skin-164547.xlsm"
pythoncom.CoInitialize()
xl = win32com.client.Dispatch("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
bk=None
try:
    bk = xl.Workbooks.Open(PATH)
    for sc in bk.SlicerCaches:
        for i in range(1, sc.Slicers.Count + 1):
            sl = sc.Slicers(i)
            print(f"slicer {sl.Name}: Style={sl.Style!r}")
    # 标记落在哪张表
    for s in bk.Worksheets:
        z1 = s.Range("Z1").Value
        z3 = s.Range("Z3").Value
        if z1 or z3:
            print(f"trace on {s.Name}: Z1={z1!r} Z3={z3!r}")
finally:
    if bk is not None: bk.Close(SaveChanges=False)
    xl.Quit(); pythoncom.CoUninitialize()

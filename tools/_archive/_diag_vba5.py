import pythoncom, win32com.client
PATH = None
import glob
cands = sorted(glob.glob(r"D:\execl-mcp\_demo\_vba_bisect2_*.xlsm"))
PATH = cands[-1]
pythoncom.CoInitialize()
xl = win32com.client.Dispatch("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
bk=None
try:
    bk = xl.Workbooks.Open(PATH)
    print("sheets order:", [s.Name for s in bk.Worksheets])
    for sname in bk.Worksheets:
        ws = bk.Worksheets(sname.Name)
        print(f"  {sname.Name}: A1={ws.Range('A1').Value!r}")
finally:
    if bk is not None: bk.Close(SaveChanges=False)
    xl.Quit(); pythoncom.CoUninitialize()

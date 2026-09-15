"""Isolate the silent `vba run`: does Application.Run work at all on this machine?

Uses the workbook the MCP probe just saved (SkinModule + slicer already inside).
"""
import pythoncom
import win32com.client

PATH = r"D:\execl-mcp\_demo\slicer-skin-164547.xlsm"

pythoncom.CoInitialize()
xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
bk = None
try:
    xl.AutomationSecurity = 1  # msoAutomationSecurityLow BEFORE open
    print("security set to:", xl.AutomationSecurity)
    bk = xl.Workbooks.Open(PATH)
    print("opened:", bk.Name)
    print("modules:", [c.Name for c in bk.VBProject.VBComponents])
    print("slicer caches:", bk.SlicerCaches.Count)

    for label, name in [
        ("Module.Sub", "SkinModule.ApplySlicerSkin"),
        ("Sub only", "ApplySlicerSkin"),
        ("Book!Module.Sub", f"'{bk.Name}'!SkinModule.ApplySlicerSkin"),
    ]:
        try:
            xl.Run(name)
            z1 = bk.Worksheets(1).Range("Z1").Value
            z2 = bk.Worksheets(1).Range("Z2").Value
            z3 = bk.Worksheets(1).Range("Z3").Value
            z9 = bk.Worksheets(1).Range("Z9").Value
            print(f"Run({label!r}): OK  Z1={z1!r} Z2={z2!r} Z3={z3!r} Z9={z9!r}")
            if z1 == "START":
                break
        except Exception as e:  # noqa: BLE001
            print(f"Run({label!r}): FAIL {str(e)[:120]}")
finally:
    if bk is not None:
        bk.Close(SaveChanges=False)
    xl.Quit()
    pythoncom.CoUninitialize()

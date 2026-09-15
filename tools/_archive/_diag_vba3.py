import sys, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from demo_slicer_link import Server

DATA = [["Region","Q1","Q2"],["华东",100,130],["华南",200,170],["华北",300,240],
        ["西南",180,260],["东北",90,150],["西北",140,110]]

SKIN = '''Public Sub ApplySlicerSkin()
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim i As Long
    Dim report As String
    ThisWorkbook.Worksheets(1).Range("Z1").Value = "START"
    For Each sc In ThisWorkbook.SlicerCaches
        For i = 1 To sc.Slicers.Count
            Set sl = sc.Slicers(i)
            sl.Style = "SlicerStyleLight2"
            report = report & sl.Name & "=" & sl.Style & ";"
        Next i
    Next sc
    ThisWorkbook.Worksheets(1).Range("Z3").Value = report
End Sub
'''
MARK = 'Public Sub MarkIt()\n    ThisWorkbook.Worksheets(1).Range("A1").Value = "RAN"\nEnd Sub\n'

path = Path(f"_demo/_vba_bisect_{time.strftime('%H%M%S')}.xlsm").resolve()
s = Server(); s.start()
try:
    sid = s.call("file", {"action": "create", "path": str(path), "show": False})["session_id"]
    s.call("range", {"action": "set-values", "session_id": sid, "sheet_name": "Sheet1",
                     "range_address": "A1:C7", "values": DATA})
    s.call("worksheet", {"action": "create", "session_id": sid, "sheet_name": "Pivot"})
    s.call("pivottable", {"action": "create-from-range", "session_id": sid,
                          "pivot_table_name": "PT_S", "source_sheet": "Sheet1",
                          "source_range": "A1:C7", "destination_sheet": "Pivot",
                          "destination_cell": "A1"})
    s.call("pivottable_field", {"action": "add-row-field", "session_id": sid,
                                "pivot_table_name": "PT_S", "field_name": "Region"})
    s.call("pivottable_field", {"action": "add-value-field", "session_id": sid,
                                "pivot_table_name": "PT_S", "field_name": "Q1"})
    s.call("slicer", {"action": "create-slicer", "session_id": sid,
                      "destination_sheet": "Pivot", "pivot_table_name": "PT_S",
                      "field_name": "Region", "slicer_name": "sl_x", "position": "F3"})
    print("slicer OK")

    s.call("vba", {"action": "import", "session_id": sid, "module_name": "M1", "vba_code": MARK})
    try:
        s.call("vba", {"action": "run", "session_id": sid, "procedure_name": "M1.MarkIt"})
        v = s.call("range", {"action": "get-values", "session_id": sid,
                             "sheet_name": "Sheet1", "range_address": "A1"})
        print("M1.MarkIt ->", (v.get("values") or [[None]])[0][0])
    except Exception as e:
        print("M1.MarkIt FAIL:", str(e)[:150])

    s.call("vba", {"action": "import", "session_id": sid, "module_name": "SkinModule", "vba_code": SKIN})
    try:
        s.call("vba", {"action": "run", "session_id": sid, "procedure_name": "SkinModule.ApplySlicerSkin"})
        v = s.call("range", {"action": "get-values", "session_id": sid,
                             "sheet_name": "Sheet1", "range_address": "Z1:Z3"})
        print("Skin Z1:Z3 ->", v.get("values"))
    except Exception as e:
        print("Skin FAIL:", str(e)[:200])
    s.call("file", {"action": "close", "session_id": sid, "save": True})
except Exception as e:
    print("ERR:", str(e)[:250])
finally:
    s.proc.terminate()

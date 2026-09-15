"""Probe: can the MCP `vba` tool re-skin a slicer? (the documented fallback path)

Flow: create .xlsm -> data + pivottable + slicer -> `vba import` a module whose
Sub restyles every slicer and writes a report cell -> `vba run` -> read the
report cell back through `range get-values`.

Usage:
    python probe_slicer_skin.py                # default: SlicerStyleLight2
    python probe_slicer_skin.py --style SlicerStyleDark3
    python probe_slicer_skin.py --all          # sweep all 12 built-in styles
"""
import argparse
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from demo_slicer_link import Server

DATA = [
    ["Region", "Q1", "Q2"],
    ["华东", 100, 130], ["华南", 200, 170], ["华北", 300, 240],
    ["西南", 180, 260], ["东北", 90, 150], ["西北", 140, 110],
]

ALL_STYLES = [f"SlicerStyle{fam}{n}" for fam in ("Light", "Dark") for n in range(1, 7)]


def vba_apply_one(style: str) -> str:
    """VBA that applies one style to every slicer, with report cells."""
    return f"""Public Sub ApplySlicerSkin()
    On Error GoTo Fail
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    ws.Range("Z1").Value = "START"
    ws.Range("Z2").Value = "caches=" & ThisWorkbook.SlicerCaches.Count
    Dim report As String
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim i As Long
    For Each sc In ThisWorkbook.SlicerCaches
        For i = 1 To sc.Slicers.Count
            Set sl = sc.Slicers(i)
            sl.Style = "{style}"
            report = report & sl.Name & "=" & sl.Style & ";"
        Next i
    Next sc
    ws.Range("Z3").Value = report
    Exit Sub
Fail:
    ws.Range("Z9").Value = "ERR " & Err.Number & ": " & Err.Description
End Sub"""


def vba_sweep_all(styles) -> str:
    """VBA that tries every style on slicer #1 and reports OK/ERR per style."""
    names = ", ".join(f'"{s}"' for s in styles)
    return f"""Public Sub SweepSlicerSkins()
    On Error GoTo Fail
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Dim allStyles As Variant
    allStyles = Array({names})
    Dim sc As SlicerCache
    Set sc = ThisWorkbook.SlicerCaches(1)
    Dim sl As Slicer
    Set sl = sc.Slicers(1)
    Dim i As Long, row As Long
    row = 1
    ws.Range("Z1").Value = "style"
    ws.Range("AA1").Value = "result"
    For i = LBound(allStyles) To UBound(allStyles)
        On Error Resume Next
        sl.Style = CStr(allStyles(i))
        If Err.Number <> 0 Then
            ws.Cells(row + 1, 26).Value = CStr(allStyles(i))
            ws.Cells(row + 1, 27).Value = "ERR " & Err.Number
            Err.Clear
        Else
            ws.Cells(row + 1, 26).Value = CStr(allStyles(i))
            ws.Cells(row + 1, 27).Value = "OK ->" & sl.Style
        End If
        row = row + 1
        On Error GoTo Fail
    Next i
    ws.Range("AC1").Value = "SWEEP DONE"
    Exit Sub
Fail:
    ws.Range("AC2").Value = "ERR " & Err.Number & ": " & Err.Description
End Sub"""


def build_workbook(s: Server, path: Path) -> str:
    """Create .xlsm with one pivottable + one slicer; return session id."""
    sid = s.call("file", {"action": "create", "path": str(path), "show": False})["session_id"]
    s.call("range", {"action": "set-values", "session_id": sid,
                     "sheet_name": "Sheet1", "range_address": "A1:C7", "values": DATA})
    s.call("worksheet", {"action": "create", "session_id": sid, "sheet_name": "Pivot"})
    s.call("pivottable", {"action": "create-from-range", "session_id": sid,
                          "pivot_table_name": "PT_Skin",
                          "source_sheet": "Sheet1", "source_range": "A1:C7",
                          "destination_sheet": "Pivot", "destination_cell": "A1"})
    s.call("pivottable_field", {"action": "add-row-field", "session_id": sid,
                                "pivot_table_name": "PT_Skin", "field_name": "Region"})
    s.call("pivottable_field", {"action": "add-value-field", "session_id": sid,
                                "pivot_table_name": "PT_Skin", "field_name": "Q1"})
    s.call("slicer", {"action": "create-slicer", "session_id": sid,
                      "destination_sheet": "Pivot", "pivot_table_name": "PT_Skin",
                      "field_name": "Region", "slicer_name": "sl_skin", "position": "F3"})
    return sid


def run_one_style(style: str) -> None:
    stamp = time.strftime("%H%M%S")
    path = Path(f"_demo/slicer-skin-{style.replace('SlicerStyle', '').lower()}-{stamp}.xlsm").resolve()
    s = Server()
    s.start()
    try:
        sid = build_workbook(s, path)
        print("1 slicer created:", path.name)
        s.call("vba", {"action": "import", "session_id": sid,
                       "module_name": "SkinModule", "vba_code": vba_apply_one(style)})
        s.call("vba", {"action": "run", "session_id": sid,
                       "procedure_name": "SkinModule.ApplySlicerSkin"})
        trace = s.call("range", {"action": "get-values", "session_id": sid,
                                 "sheet_name": "Sheet1", "range_address": "Z1:Z9"})
        rows = trace.get("values") or []
        for row in rows:
            cell = row[0]
            if cell not in (None, ""):
                print("   cell:", cell)
        report = "".join(str(r[0] or "") for r in rows)
        ok = style in report
        print("VERDICT:", f"{style} APPLIED" if ok else f"{style} NOT REPORTED")
        s.call("file", {"action": "close", "session_id": sid, "save": True})
        print("saved:", path)
    except Exception as error:  # noqa: BLE001
        print("ERR:", str(error)[:400])
    finally:
        s.proc.terminate()


def run_sweep() -> None:
    stamp = time.strftime("%H%M%S")
    path = Path(f"_demo/slicer-skin-sweep-{stamp}.xlsm").resolve()
    s = Server()
    s.start()
    try:
        sid = build_workbook(s, path)
        print("1 slicer created:", path.name)
        s.call("vba", {"action": "import", "session_id": sid,
                       "module_name": "SkinModule", "vba_code": vba_sweep_all(ALL_STYLES)})
        s.call("vba", {"action": "run", "session_id": sid,
                       "procedure_name": "SkinModule.SweepSlicerSkins"})
        trace = s.call("range", {"action": "get-values", "session_id": sid,
                                 "sheet_name": "Sheet1", "range_address": "Z1:AC13"})
        rows = trace.get("values") or []
        ok_list, err_list = [], []
        for row in rows:
            style_cell = str(row[0] or "")
            result_cell = str(row[1] or "")
            if not style_cell or style_cell == "style":
                continue
            print(f"   {style_cell:<22} {result_cell}")
            (ok_list if result_cell.startswith("OK") else err_list).append(style_cell)
        done = any("SWEEP DONE" in str(cell) for r in rows for cell in r)
        print(f"SUMMARY: {len(ok_list)}/{len(ALL_STYLES)} styles applied"
              + ("" if done else "  (SWEEP marker missing - check AC2)"))
        if err_list:
            print("FAILED:", ", ".join(err_list))
        s.call("file", {"action": "close", "session_id": sid, "save": True})
        print("saved:", path)
    except Exception as error:  # noqa: BLE001
        print("ERR:", str(error)[:400])
    finally:
        s.proc.terminate()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="VBA slicer re-skin probe")
    parser.add_argument("--style", default="SlicerStyleLight2",
                        help="built-in style name, e.g. SlicerStyleDark3")
    parser.add_argument("--all", action="store_true",
                        help="sweep all 12 built-in styles in one session")
    args = parser.parse_args()
    if args.all:
        run_sweep()
    else:
        run_one_style(args.style)

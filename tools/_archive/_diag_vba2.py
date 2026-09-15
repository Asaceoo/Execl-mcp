import sys, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from demo_slicer_link import Server

path = Path(f"_demo/_vba_diag_{time.strftime('%H%M%S')}.xlsm").resolve()
s = Server(); s.start()
try:
    sid = s.call("file", {"action": "create", "path": str(path), "show": False})["session_id"]
    code = 'Public Sub MarkIt()\n    ThisWorkbook.Worksheets(1).Range("A1").Value = "RAN"\nEnd Sub\n'
    s.call("vba", {"action": "import", "session_id": sid, "module_name": "M1", "vba_code": code})
    print("imported")
    for label, proc in [("existing", "M1.MarkIt"), ("missing", "M1.Nope"), ("missing module", "Zzz.Nope")]:
        try:
            r = s.call("vba", {"action": "run", "session_id": sid, "procedure_name": proc})
            print(f"run {label:<14}: OK   ->", str(r)[:80])
        except Exception as e:
            print(f"run {label:<14}: FAIL ->", str(e)[:160])
    v = s.call("range", {"action": "get-values", "session_id": sid, "sheet_name": "Sheet1", "range_address": "A1"})
    print("A1 after:", (v.get("values") or [[None]])[0][0])
    s.call("file", {"action": "close", "session_id": sid, "save": True})
except Exception as e:
    print("ERR:", str(e)[:200])
finally:
    s.proc.terminate()

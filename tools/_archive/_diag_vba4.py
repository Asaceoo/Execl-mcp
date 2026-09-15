import sys, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from demo_slicer_link import Server

DATA = [["Region","Q1","Q2"],["华东",100,130],["华南",200,170],["华北",300,240],
        ["西南",180,260],["东北",90,150],["西北",140,110]]
MARK = 'Public Sub MarkIt()\n    ThisWorkbook.Worksheets(1).Range("A1").Value = "RAN"\nEnd Sub\n'

path = Path(f"_demo/_vba_bisect2_{time.strftime('%H%M%S')}.xlsm").resolve()
s = Server(); s.start()
try:
    sid = s.call("file", {"action": "create", "path": str(path), "show": False})["session_id"]

    def checkpoint(label):
        s.call("vba", {"action": "run", "session_id": sid, "procedure_name": "M1.MarkIt"})
        v = s.call("range", {"action": "get-values", "session_id": sid,
                             "sheet_name": "Sheet1", "range_address": "A1"})
        got = (v.get("values") or [[None]])[0][0]
        print(f"[{label}] run -> A1={got!r}")
        # 恢复现场，避免影响下一检查点
        s.call("range", {"action": "set-values", "session_id": sid,
                         "sheet_name": "Sheet1", "range_address": "A1", "values": [["Region"]]})

    s.call("vba", {"action": "import", "session_id": sid, "module_name": "M1", "vba_code": MARK})
    checkpoint("0. 刚建好空簿")

    s.call("range", {"action": "set-values", "session_id": sid, "sheet_name": "Sheet1",
                     "range_address": "A1:C7", "values": DATA})
    checkpoint("1. 写入数据后")

    s.call("worksheet", {"action": "create", "session_id": sid, "sheet_name": "Pivot"})
    s.call("pivottable", {"action": "create-from-range", "session_id": sid,
                          "pivot_table_name": "PT_S", "source_sheet": "Sheet1",
                          "source_range": "A1:C7", "destination_sheet": "Pivot",
                          "destination_cell": "A1"})
    s.call("pivottable_field", {"action": "add-row-field", "session_id": sid,
                                "pivot_table_name": "PT_S", "field_name": "Region"})
    s.call("pivottable_field", {"action": "add-value-field", "session_id": sid,
                                "pivot_table_name": "PT_S", "field_name": "Q1"})
    checkpoint("2. 建透视表后")

    s.call("slicer", {"action": "create-slicer", "session_id": sid,
                      "destination_sheet": "Pivot", "pivot_table_name": "PT_S",
                      "field_name": "Region", "slicer_name": "sl_x", "position": "F3"})
    checkpoint("3. 建切片器后")
    s.call("file", {"action": "close", "session_id": sid, "save": True})
except Exception as e:
    print("ERR:", str(e)[:250])
finally:
    s.proc.terminate()

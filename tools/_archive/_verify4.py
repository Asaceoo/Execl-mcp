import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from demo_slicer_link import Server
DATA=[["Region","Q1","Q2"],["华东",100,130],["华南",200,170],["华北",300,240],["西南",180,260],["东北",90,150],["西北",140,110]]
s=Server(); s.start()
try:
    sid=s.call("file",{"action":"create","path":str(Path("_demo/_addseries.xlsx").resolve()),"show":False})["session_id"]
    s.call("range",{"action":"set-values","session_id":sid,"sheet_name":"Sheet1","range_address":"A1:C7","values":DATA})
    s.call("chart",{"action":"create-from-range","session_id":sid,"sheet_name":"Sheet1","source_range_address":"A1:C7","chart_type":"ColumnClustered","chart_name":"c1","target_range":"E1:J12"})
    for label,args in [
        ("series_name+values_range", {"series_name":"Q1dup","values_range":"B2:B7"}),
        ("values_range only", {"values_range":"B2:B7"}),
        ("values+category", {"series_name":"Q1dup","values_range":"B2:B7","category_range":"A2:A7"}),
    ]:
        try:
            r=s.call("chart_config", dict({"action":"add-series","session_id":sid,"chart_name":"c1"}, **args))
            print("OK  ", label, str(r)[:100])
        except Exception as e:
            print("FAIL", label, str(e)[:150])
    s.call("file",{"action":"close","session_id":sid})
except Exception as e:
    print("ERR", str(e)[:200])
finally:
    s.proc.terminate()

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from demo_slicer_link import Server
DATA=[["Region","Q1","Q2"],["华东",100,130],["华南",200,170],["华北",300,240],["西南",180,260],["东北",90,150],["西北",140,110]]
s=Server(); s.start()
try:
    sid=s.call("file",{"action":"create","path":str(Path("_demo/_diag_read.xlsx").resolve()),"show":False})["session_id"]
    s.call("range",{"action":"set-values","session_id":sid,"sheet_name":"Sheet1","range_address":"A1:C7","values":DATA})
    for label, ctype in [("柱形图","ColumnClustered"),("瀑布图","Waterfall"),("树状图","Treemap")]:
        try:
            s.call("chart",{"action":"create-from-range","session_id":sid,"sheet_name":"Sheet1",
                            "source_range_address":"A1:C7","chart_type":ctype,"chart_name":f"c_{ctype}","target_range":"E1:J12"})
            print(f"create {label}: OK")
        except Exception as e:
            print(f"create {label}: FAIL {str(e)[:160]}")
    try:
        r=s.call("chart",{"action":"list","session_id":sid})
        print("list:", str(r)[:200])
    except Exception as e:
        print("list: FAIL", str(e)[:300])
    for name in ["c_ColumnClustered","c_Waterfall","c_Treemap"]:
        try:
            r=s.call("chart",{"action":"read","session_id":sid,"chart_name":name})
            print(f"read {name}: type={r.get('chart_type')} series={r.get('series_count')}")
        except Exception as e:
            print(f"read {name}: FAIL {str(e)[:300]}")
    s.call("file",{"action":"close","session_id":sid})
except Exception as e:
    print("ERR", str(e)[:200])
finally:
    s.proc.terminate()

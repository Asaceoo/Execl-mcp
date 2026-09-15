import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from demo_slicer_link import Server
s = Server(); s.start()
try:
    r = s.request("tools/list", {})
    names = sorted(t["name"] for t in r["result"]["tools"])
    print(len(names), "tools:")
    print(" ".join(names))
finally:
    s.proc.terminate()

"""Focused probe: MODERN charts + STOCK charts + screenshot, on the deployed v.5 build.

Why a separate probe: probe_chart_types.py runs the full 64-case matrix (~25 min).
This one answers exactly three questions:
  1. Do the 8 modern chart types (F6-1) build and read back correctly?
  2. Do the 4 stock chart variants (F6-1/F6-4) match their column counts?
  3. Does `screenshot` work in a VISIBLE session (show=true)? The headless session
     could never capture, which is why the full probe reported a screenshot failure.
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from demo_slicer_link import Server

CATEGORY = [["Region", "Q1", "Q2"],
            ["华东", 100, 130], ["华南", 200, 170], ["华北", 300, 240],
            ["西南", 180, 260], ["东北", 90, 150], ["西北", 140, 110]]
NUMERIC = [["A", "B", "C", "D", "E"],
           [10, 24, 31, 12, 5], [22, 18, 14, 30, 9], [35, 12, 28, 17, 21],
           [14, 27, 19, 25, 33], [28, 31, 11, 22, 16], [19, 15, 36, 28, 12]]
HIERARCHY = [["大区", "城市", "销售额"],
             ["华东", "上海", 320], ["华东", "杭州", 210], ["华东", "南京", 180],
             ["华南", "广州", 280], ["华南", "深圳", 350]]

MODERN = [
    ("树状图", "Treemap", "A1:C7"),
    ("旭日图", "Sunburst", "K1:M7"),
    ("直方图", "Histogram", "B1:C7"),
    ("帕累托图", "Pareto", "B1:C7"),
    ("箱形图", "BoxWhisker", "B1:C7"),
    ("瀑布图", "Waterfall", "A1:C7"),
    ("漏斗图", "Funnel", "A1:C7"),
    ("区域地图", "RegionMap", "A1:C7"),
    ("柱线组合图", "ColumnLineCombo", "A1:C7"),
]
STOCK = [
    ("盘高-盘低-收盘 HLC (3 列)", "StockHLC", "E1:G7"),
    ("开盘-盘高-盘低-收盘 OHLC (4 列)", "StockOHLC", "E1:H7"),
    ("成交量-盘高-盘低-收盘 VHLC (4 列)", "StockVHLC", "E1:H7"),
    ("成交量-开盘-盘高-盘低-收盘 VOHLC (5 列)", "StockVOHLC", "E1:I7"),
]

stamp = time.strftime("%H%M%S")
path = Path(f"_demo/mshot-probe-{stamp}.xlsx").resolve()
server = Server()
server.start()
results = []


def check(section, name, ok, note=""):
    results.append((section, name, ok, note))
    print(f"  {'OK  ' if ok else 'FAIL'} {name}" + (f"  <- {note}" if note and not ok else ""))


try:
    # show=true: the screenshot action needs a visible window (headless capture always failed).
    session = server.call("file", {"action": "create", "path": str(path), "show": True})["session_id"]
    print(f"1 session (visible) -> {path.name}")
    server.call("worksheet", {"action": "create", "session_id": session, "sheet_name": "Charts"})
    server.call("range", {"action": "set-values", "session_id": session,
                          "sheet_name": "Charts", "range_address": "A1:C7", "values": CATEGORY})
    server.call("range", {"action": "set-values", "session_id": session,
                          "sheet_name": "Charts", "range_address": "E1:I7", "values": NUMERIC})
    server.call("range", {"action": "set-values", "session_id": session,
                          "sheet_name": "Charts", "range_address": "K1:M7", "values": HIERARCHY})

    row = 10
    created = []
    print("MODERN")
    for name, ctype, source in MODERN:
        chart_name = f"ch_{ctype}"
        target = f"A{row}:F{row + 11}"
        row += 13
        try:
            server.call("chart", {"action": "create-from-range", "session_id": session,
                                  "sheet_name": "Charts", "source_range_address": source,
                                  "chart_type": ctype, "chart_name": chart_name,
                                  "target_range": target})
            created.append((name, ctype, chart_name))
            check("MODERN", f"{name} ({ctype})", True)
        except Exception as error:  # noqa: BLE001
            detail = str(error)[:140]
            # Sunburst is an Excel-side limit: Excel itself falls back to another type, and the
            # tool now refuses instead of reporting a fake success. Kept visible, never hidden.
            if ctype == "Sunburst" and "fell back" in detail:
                check("MODERN-LIMIT", f"{name} ({ctype}) 已知 Excel 限制", False, detail)
            else:
                check("MODERN", f"{name} ({ctype})", False, detail)

    print("STOCK")
    for name, ctype, source in STOCK:
        chart_name = f"ch_{ctype}"
        target = f"A{row}:F{row + 11}"
        row += 13
        try:
            server.call("chart", {"action": "create-from-range", "session_id": session,
                                  "sheet_name": "Charts", "source_range_address": source,
                                  "chart_type": ctype, "chart_name": chart_name,
                                  "target_range": target})
            created.append((name, ctype, chart_name))
            check("STOCK", f"{name} ({ctype})", True)
        except Exception as error:  # noqa: BLE001
            check("STOCK", f"{name} ({ctype})", False, str(error)[:140])

    # Read back what Excel actually built - a silent fallback would show a different type.
    print("READBACK")
    reported = {}
    try:
        listing = server.call("chart", {"action": "list", "session_id": session})
        for item in listing.get("charts") or []:
            # The list payload is camelCase (`chartType`), not snake_case.
            reported[item.get("name")] = str(item.get("chartType") or item.get("type") or "")
        print(f"  Excel reports {len(reported)} charts")
    except Exception as error:  # noqa: BLE001
        check("READBACK", "chart list", False, str(error)[:160])
    for name, ctype, chart_name in created:
        got = reported.get(chart_name, "")
        check("READBACK", f"{chart_name} type={got or '?'}", got == ctype,
              f"expected {ctype}, got {got or 'nothing'}")

    # F6-5: modern charts used to blow up `chart read` (E_NOTIMPL on PivotLayout / series).
    # Reading back the type also catches a silent fallback to another chart type.
    for _, ctype, chart_name in [c for c in created if c[1] in ("Treemap", "StockVOHLC")]:
        try:
            detail = server.call("chart", {"action": "read", "session_id": session,
                                           "chart_name": chart_name})
            check("READBACK", f"chart read {chart_name} -> {detail.get('chartType')}",
                  str(detail.get("chartType")) == ctype, f"expected {ctype}")
        except Exception as error:  # noqa: BLE001
            check("READBACK", f"chart read {chart_name}", False, str(error)[:160])

    print("SCREENSHOT")
    try:
        # NOTE: the MCP action takes only sheet_name + quality. `--output` is a CLI-only flag;
        # the server answers with an image content block (base64) plus a text summary.
        shot = server.call("screenshot", {"action": "capture-sheet", "session_id": session,
                                          "sheet_name": "Charts", "quality": "High"})
        blocks = shot.get("_images") or []
        out_png = Path(f"_demo/mshot-{stamp}.png").resolve()
        written = 0
        mime = ""
        if blocks:
            import base64
            mime = blocks[0].get("mimeType", "")
            data = blocks[0].get("data") or ""
            out_png.write_bytes(base64.b64decode(data))
            written = out_png.stat().st_size
        note = str(shot.get("raw", ""))[:80] or f"blocks={len(blocks)}"
        check("SCREENSHOT", f"截图 {note} -> {out_png.name} ({written} B, {mime or 'no image'})",
              written > 0, f"image blocks={len(blocks)}")
    except Exception as error:  # noqa: BLE001
        check("SCREENSHOT", "截图", False, str(error)[:160])

    server.call("file", {"action": "close", "session_id": session, "save": True})
    print(f"2 saved: {path}")
except Exception as error:  # noqa: BLE001
    print("ERR:", str(error)[:400])
finally:
    server.proc.terminate()

passed = sum(1 for _, _, ok, _ in results if ok)
print(f"\n=== {passed} passed / {len(results) - passed} failed ===")
for section in ["MODERN", "MODERN-LIMIT", "STOCK", "READBACK", "SCREENSHOT"]:
    items = [r for r in results if r[0] == section]
    if not items:
        continue
    good = sum(1 for _, _, ok, _ in items if ok)
    bad = [f"{r[1]}({r[3][:60]})" for r in items if not r[2]]
    print(f"{section}: {good}/{len(items)}" + (f"  失败: {'; '.join(bad)}" if bad else ""))

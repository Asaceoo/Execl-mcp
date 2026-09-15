"""Probe which chart types and dashboard pieces actually render on this machine.

The source advertises 70+ chart types, but the local Excel build decides what really works
(modern charts need Excel 2016+, RegionMap needs 365, stock charts need numeric columns in a
fixed order). This script creates one chart per type against real Excel and then asks Excel
what it actually built, so a silent fallback to another type is reported as a mismatch.

Sections:
  BASIC     - column / bar / line / pie / area / scatter / radar / surface / bubble
  3D        - 3D column, 3D pie, 3D area, 3D bar
  SHAPES    - cylinder / cone / pyramid
  STOCK     - HLC, OHLC, VHLC, VOHLC
  MODERN    - treemap, sunburst, histogram, pareto, box-whisker, waterfall, funnel, map, combo
  CONFIG    - title, axis, legend, data labels, style, trendline, series format, plot options
  DASHBOARD - sparklines, conditional formatting, pivot chart, slicer, screenshot
"""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from demo_slicer_link import Server  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
STAMP = datetime.now().strftime("%H%M%S")
OUTPUT = ROOT / "_demo" / f"chart-probe-{STAMP}.xlsx"

# A1:C7   category + two numeric series (most types)
# E1:G7   three numeric columns: high / low / close (stock charts)
# I1:K7   x / y / size (bubble)
DATA = [
    ["Region", "Q1", "Q2"],
    ["华东", 100, 130],
    ["华南", 200, 170],
    ["华北", 300, 240],
    ["西南", 180, 260],
    ["东北", 90, 150],
    ["西北", 140, 110],
]
NUMERIC = [[r[1], r[2], r[1] + r[2], r[1] - 10, r[2] + 20] for r in DATA[1:]]
BUBBLE = [[1, 10, 3], [2, 25, 6], [3, 18, 4], [4, 32, 9], [5, 21, 5], [6, 14, 7]]

RESULTS: list[tuple[str, str, bool, str]] = []


def main() -> int:
    server = Server()
    server.start()
    session = None

    def probe(section: str, name: str, tool: str, arguments: dict) -> dict | None:
        try:
            result = server.call(tool, arguments)
        except Exception as error:  # noqa: BLE001 - real error text is the point
            RESULTS.append((section, name, False, str(error)[:200]))
            print(f"  FAIL  {name}: {str(error)[:180]}")
            return None
        RESULTS.append((section, name, True, ""))
        print(f"  OK    {name}")
        return result

    try:
        session = server.call("file", {"action": "create", "path": str(OUTPUT), "show": False})["session_id"]
        server.call("worksheet", {"action": "create", "session_id": session, "sheet_name": "Charts"})
        server.call("range", {"action": "set-values", "session_id": session,
                              "sheet_name": "Charts", "range_address": "A1:C7", "values": DATA})
        server.call("range", {"action": "set-values", "session_id": session,
                              "sheet_name": "Charts", "range_address": "E1:I7", "values": NUMERIC})
        # Sunburst nests text levels, so it needs two category columns plus a value column.
        server.call("range", {"action": "set-values", "session_id": session,
                              "sheet_name": "Charts", "range_address": "K1:M7",
                              "values": [["大区", "城市", "销售额"]] +
                                        [["华东", c, v] for c, v in
                                         [("上海", 320), ("杭州", 210), ("南京", 180)]] +
                                        [["华南", c, v] for c, v in
                                         [("广州", 280), ("深圳", 350)]]})
        server.call("range", {"action": "set-values", "session_id": session,
                              "sheet_name": "Charts", "range_address": "I1:K7", "values": BUBBLE})

        slot = {"row": 10}

        def make(name: str, chart_type: str, source: str = "A1:C7") -> str | None:
            """Create one chart in its own vertical slot and return the chart name."""
            row = slot["row"]
            slot["row"] = row + 13
            target = f"A{row}:F{row + 11}"
            chart_name = f"ch_{chart_type}"
            result = probe(CHART_SECTION, f"{name} ({chart_type})", "chart", {
                "action": "create-from-range",
                "session_id": session,
                "sheet_name": "Charts",
                "source_range_address": source,
                "chart_type": chart_type,
                "chart_name": chart_name,
                "target_range": target,
            })
            return chart_name if result else None

        global CHART_SECTION
        CHART_SECTION = "BASIC"
        print("BASIC")
        for name, ctype in [
            ("簇状柱形图", "ColumnClustered"), ("堆积柱形图", "ColumnStacked"),
            ("百分比堆积柱形图", "ColumnStacked100"), ("簇状条形图", "BarClustered"),
            ("堆积条形图", "BarStacked"), ("折线图", "Line"),
            ("带标记的折线图", "LineMarkers"), ("饼图", "Pie"),
            ("分离型饼图", "PieExploded"), ("圆环图", "Doughnut"),
            ("面积图", "Area"), ("堆积面积图", "AreaStacked"),
            ("散点图", "XYScatter"), ("带平滑线的散点图", "XYScatterSmooth"),
            ("雷达图", "Radar"), ("填充雷达图", "RadarFilled"),
            ("曲面图", "Surface"), ("气泡图", "Bubble"),
        ]:
            make(name, ctype, "I1:K7" if ctype == "Bubble" else "A1:C7")

        CHART_SECTION = "3D"
        print("3D")
        for name, ctype in [
            ("三维簇状柱形图", "Column3DClustered"), ("三维饼图", "Pie3D"),
            ("三维面积图", "Area3D"), ("三维簇状条形图", "Bar3DClustered"),
        ]:
            make(name, ctype)

        CHART_SECTION = "SHAPES"
        print("SHAPES")
        for name, ctype in [
            ("圆柱图", "CylinderColClustered"), ("圆锥图", "ConeColClustered"),
            ("棱锥图", "PyramidColClustered"),
        ]:
            make(name, ctype)

        CHART_SECTION = "STOCK"
        print("STOCK")
        for name, ctype, source in [
            ("盘高-盘低-收盘图 (3 列)", "StockHLC", "E1:G7"),
            ("开盘-盘高-盘低-收盘图 (4 列)", "StockOHLC", "E1:H7"),
            ("成交量-盘高-盘低-收盘图 (4 列)", "StockVHLC", "E1:H7"),
            ("成交量-开盘-盘高-盘低-收盘图 (5 列)", "StockVOHLC", "E1:I7"),
        ]:
            make(name, ctype, source)

        CHART_SECTION = "MODERN"
        print("MODERN")
        for name, ctype in [
            ("树状图", "Treemap"), ("旭日图", "Sunburst"),
            ("直方图", "Histogram"), ("帕累托图", "Pareto"),
            ("箱形图", "BoxWhisker"), ("瀑布图", "Waterfall"),
            ("漏斗图", "Funnel"), ("区域地图", "RegionMap"),
            ("柱线组合图", "ColumnLineCombo"),
        ]:
            source = "B1:C7" if ctype in ("Histogram", "Pareto", "BoxWhisker") else "A1:C7"
            if ctype == "Sunburst":
                source = "K1:M7"
            make(name, ctype, source)

        # Ask Excel what it actually built - a silent fallback shows up here.
        print("VERIFY")
        try:
            listing = server.call("chart", {"action": "list", "session_id": session})
            charts = listing.get("charts") or []
            print(f"  Excel reports {len(charts)} charts")
            for item in charts[:6]:
                print(f"    {item.get('name')}: {item.get('chart_type') or item.get('type')}")
        except Exception as error:  # noqa: BLE001
            print(f"  FAIL  chart list: {str(error)[:180]}")

        print("CONFIG")
        ref = "ch_ColumnClustered"
        probe("CONFIG", "标题 set-title", "chart_config",
              {"action": "set-title", "session_id": session, "chart_name": ref, "title": "区域销售额"})
        probe("CONFIG", "坐标轴标题 set-axis-title", "chart_config",
              {"action": "set-axis-title", "session_id": session, "chart_name": ref,
               "axis": "Category", "title": "地区"})
        probe("CONFIG", "图例 show-legend", "chart_config",
              {"action": "show-legend", "session_id": session, "chart_name": ref, "visible": True})
        probe("CONFIG", "数据标签 set-data-labels", "chart_config",
              {"action": "set-data-labels", "session_id": session, "chart_name": ref, "show_value": True})
        probe("CONFIG", "内置样式 set-style", "chart_config",
              {"action": "set-style", "session_id": session, "chart_name": ref, "style_id": 12})
        probe("CONFIG", "趋势线 add-trendline", "chart_config",
              {"action": "add-trendline", "session_id": session, "chart_name": ref,
               "series_index": 1, "trendline_type": "Linear"})
        probe("CONFIG", "系列格式 set-series-format", "chart_config",
              {"action": "set-series-format", "session_id": session, "chart_name": ref,
               "series_index": 1, "fill_color": "#4472C4"})
        probe("CONFIG", "网格线 set-gridlines", "chart_config",
              {"action": "set-gridlines", "session_id": session, "chart_name": ref,
               "axis": "Value", "show_major": True, "show_minor": False})
        probe("CONFIG", "坐标轴刻度 set-axis-scale", "chart_config",
              {"action": "set-axis-scale", "session_id": session, "chart_name": ref,
               "axis": "Value", "minimum": 0, "maximum": 400})
        probe("CONFIG", "绘图区 set-plot-options", "chart_config",
              {"action": "set-plot-options", "session_id": session, "chart_name": ref, "gap_width": 60})
        probe("CONFIG", "混合类型 set-series-chart-type（次坐标轴组合）", "chart_config",
              {"action": "set-series-chart-type", "session_id": session, "chart_name": ref,
               "series_index": 2, "chart_type": "Line"})
        probe("CONFIG", "换类型 set-chart-type", "chart_config",
              {"action": "set-chart-type", "session_id": session, "chart_name": ref,
               "chart_type": "ColumnStacked"})
        probe("CONFIG", "数据源 set-source-range", "chart_config",
              {"action": "set-source-range", "session_id": session, "chart_name": ref,
               "source_range": "A1:C5"})
        probe("CONFIG", "加系列 add-series", "chart_config",
              {"action": "add-series", "session_id": session, "chart_name": ref,
               "series_name": "Q1b", "values_range": "B2:B7"})
        probe("CONFIG", "删系列 remove-series", "chart_config",
              {"action": "remove-series", "session_id": session, "chart_name": ref, "series_index": 3})
        probe("CONFIG", "放置模式 set-placement", "chart_config",
              {"action": "set-placement", "session_id": session, "chart_name": ref, "placement": 1})
        probe("CONFIG", "图表区格式 set-area-format", "chart_config",
              {"action": "set-area-format", "session_id": session, "chart_name": ref,
               "area": "Chart", "fill_color": "#F2F2F2"})

        print("DASHBOARD")
        server.call("worksheet", {"action": "create", "session_id": session, "sheet_name": "Board"})
        server.call("range", {"action": "set-values", "session_id": session,
                              "sheet_name": "Board", "range_address": "A1:C7", "values": DATA})
        probe("DASHBOARD", "迷你图 add-sparkline", "drawing",
              {"action": "add-sparkline", "session_id": session, "sheet_name": "Board",
               "location_range": "E2:E7", "source_range": "B2:C7", "sparkline_type": "Line"})
        probe("DASHBOARD", "条件格式-数据条", "conditionalformat",
              {"action": "add-rule", "session_id": session, "sheet_name": "Board",
               "range_address": "B2:C7", "rule_type": "data-bar", "data_bar_color": "#638EC6"})
        probe("DASHBOARD", "条件格式-色阶", "conditionalformat",
              {"action": "add-rule", "session_id": session, "sheet_name": "Board",
               "range_address": "B2:C7", "rule_type": "color-scale",
               "color_scale_min_color": "#F8696B", "color_scale_max_color": "#63BE7B"})
        probe("DASHBOARD", "条件格式-图标集", "conditionalformat",
              {"action": "add-rule", "session_id": session, "sheet_name": "Board",
               "range_address": "B2:C7", "rule_type": "icon-set", "icon_set_id": "3TrafficLights1"})
        probe("DASHBOARD", "条件格式-前N项 top10", "conditionalformat",
              {"action": "add-rule", "session_id": session, "sheet_name": "Board",
               "range_address": "B2:C7", "rule_type": "top10", "rank": 3, "top_bottom": "Top"})

        # Pivot chart + slicer: the interactive dashboard core
        server.call("worksheet", {"action": "create", "session_id": session, "sheet_name": "Pivot"})
        server.call("range", {"action": "set-values", "session_id": session,
                              "sheet_name": "Pivot", "range_address": "A1:C7", "values": DATA})
        try:
            server.call("pivottable", {"action": "create-from-range", "session_id": session,
                                       "pivot_table_name": "PT_Board",
                                       "source_sheet": "Pivot", "source_range": "A1:C7",
                                       "destination_sheet": "Pivot", "destination_cell": "E1"})
            server.call("pivottable_field", {"action": "add-row-field", "session_id": session,
                                             "pivot_table_name": "PT_Board", "field_name": "Region"})
            server.call("pivottable_field", {"action": "add-value-field", "session_id": session,
                                             "pivot_table_name": "PT_Board", "field_name": "Q1"})
            probe("DASHBOARD", "数据透视表 create", "pivottable", {"action": "list", "session_id": session})
            probe("DASHBOARD", "数据透视图 create-from-pivottable", "chart",
                  {"action": "create-from-pivottable", "session_id": session,
                   "pivot_table_name": "PT_Board", "sheet_name": "Pivot",
                   "chart_type": "ColumnClustered", "chart_name": "ch_pivot",
                   "target_range": "K1:P12"})
            probe("DASHBOARD", "切片器 create-slicer", "slicer",
                  {"action": "create-slicer", "session_id": session, "destination_sheet": "Pivot",
                   "pivot_table_name": "PT_Board", "field_name": "Region",
                   "slicer_name": "sl_region", "position": "A20"})
        except Exception as error:  # noqa: BLE001
            RESULTS.append(("DASHBOARD", "透视看板链路", False, str(error)[:200]))
            print(f"  FAIL  pivot chain: {str(error)[:180]}")

        try:
            # Screenshot needs a VISIBLE session: the workbook is created with show=false here,
            # so capture is expected to fail. Use tools/probe_modern_stock_shot.py (show=true)
            # to verify the screenshot path.
            shot = server.call("screenshot", {"action": "capture-sheet", "session_id": session,
                                              "sheet_name": "Board", "quality": "High"})
            blocks = shot.get("_images") or []
            ok = bool(blocks and blocks[0].get("data"))
            RESULTS.append(("DASHBOARD", "截图 screenshot", ok,
                            "" if ok else "headless session - see probe_modern_stock_shot.py"))
            print(f"  {'OK   ' if ok else 'SKIP '} screenshot (headless session)" )
        except Exception as error:  # noqa: BLE001
            RESULTS.append(("DASHBOARD", "截图 screenshot", False, str(error)[:160]))
            print(f"  FAIL  screenshot: {str(error)[:160]}")

        try:
            # `file` has no `save` action - persistence goes through close+save.
            server.call("file", {"action": "close", "session_id": session, "save": True})
        except Exception as error:  # noqa: BLE001
            print(f"  note  close: {str(error)[:120]}")
        session = None
    finally:
        if session:
            try:
                server.call("file", {"action": "close", "session_id": session})
            except Exception:  # noqa: BLE001
                pass
        server.proc.terminate()

    passed = sum(1 for _, _, ok, _ in RESULTS if ok)
    failed = len(RESULTS) - passed
    print(f"\n=== {passed} passed / {failed} failed ===")
    for section in ["BASIC", "3D", "SHAPES", "STOCK", "MODERN", "CONFIG", "DASHBOARD"]:
        items = [r for r in RESULTS if r[0] == section]
        if not items:
            continue
        good = sum(1 for _, _, ok, _ in items if ok)
        bad = [r[1] for r in items if not r[2]]
        print(f"{section}: {good}/{len(items)}" + (f"  失败: {', '.join(bad)}" if bad else ""))
    print(f"\n产物: {OUTPUT}")
    return 0 if failed == 0 else 1


CHART_SECTION = "BASIC"

if __name__ == "__main__":
    raise SystemExit(main())

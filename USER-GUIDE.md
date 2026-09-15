# jyyj-mcp 助手 · 用户手册

面向**使用者**的完整操作指南：装好 → 接上 → 用起来 → 出错自己查。

- 版本：`3.0.0.0`
- 工具面：**31 个工具 / 328 个动作**
- 上游项目：[`sbroenne/mcp-server-excel`](https://github.com/sbroenne/mcp-server-excel)（MIT，v2.0.8）

> 想看内部实现、补丁原理与实测事实，请读 [技术手册](TECHNICAL-MANUAL.md)。
> 想看部署细节、验证证据与重建步骤，请读 [DEPLOYMENT.md](DEPLOYMENT.md)。

---

## 目录

1. [它是什么](#1-它是什么)
2. [环境要求](#2-环境要求)
3. [安装](#3-安装)
4. [接入 MCP 客户端](#4-接入-mcp-客户端)
5. [核心心智模型：一切都围绕 session](#5-核心心智模型一切都围绕-session)
6. [工具速查表（31 个）](#6-工具速查表31-个)
7. [两种列表线格式（最容易踩的坑）](#7-两种列表线格式最容易踩的坑)
8. [实战配方](#8-实战配方)
9. [错误排查表](#9-错误排查表)
10. [已知限制](#10-已知限制)

---

## 1. 它是什么

jyyj-mcp 助手是一个 **MCP 服务器**：把「驱动真实的 Microsoft Excel」这件能力，通过
[Model Context Protocol](https://modelcontextprotocol.io) 暴露给 AI 助手（Claude Desktop、Cursor、WorkBuddy 等）。

它和「解析 xlsx 的库」**不是一类东西**：

| | 解析型库（openpyxl / pandas） | jyyj-mcp 助手 |
|---|---|---|
| 驱动方式 | 直接读写文件字节 | 调用 **Excel 官方 COM 接口**，驱动真实 Excel 进程 |
| 透视表 | 基本只能读，写几乎不可用 | 建、改、刷新、钻取、计算字段/成员、布局、日期分组 |
| 切片器 | 不支持 | 建、选、删，**并支持一个切片器驱动多张透视表** |
| 图表 | 只能生成静态图 | 真·Excel 图表对象，可后续被 Excel 编辑 |
| VBA / 宏 | 不支持 | 导入、运行、读取（`vba run`） |
| Power Query (M) | 不支持 | 建、改、刷新、装载 |
| 数据模型 / DAX | 不支持 | 建度量值、建关系、执行 DMV 查询 |
| 文件打开时 | 可以 | **必须关闭**（COM 需要独占访问） |

一句话：**凡是"只有 Excel 自己才懂"的东西，它都能创建出来并原样保留**，而不是在读写过程中被降级成静态值。

---

## 2. 环境要求

| 项 | 要求 | 说明 |
|---|---|---|
| 操作系统 | **Windows x64** | COM 自动化只存在于 Windows |
| Excel | **Microsoft Excel 2016 或更高版本（桌面版）** | 必须是真装的 Excel；**WPS 不适用** |
| 目标工作簿 | 操作期间**在 Excel 里关闭** | COM 需要独占访问，被打开的文件会报占用错误 |
| .NET | **仅「从源码构建」时需要** .NET 10 SDK | 用预编译包不需要装 .NET |
| VBA 功能 | 需在 Excel 里手工开启「信任对 VBA 工程对象模型的访问」 | 否则 `vba` 工具会失败 |
| 截图功能 | 需要 `show: true` 的可见会话 | 无头会话下 `capture-sheet` 必然失败 |

---

## 3. 安装

### 方式 A：用预编译包（推荐）

1. 从 Releases 下载 `jyyj-mcp-3.0.0.0-win-x64.zip`。
2. 解压到任意目录，例如 `D:\jyyj-mcp`。

> ⚠️ **不要只把 `.exe` 单独拷出来用。** .NET 的可执行文件必须和同目录的 `*.dll`、
> `*.deps.json`、`*.runtimeconfig.json` 待在一起，单独一个 exe 会报
> `The application to execute does not exist`。所以发布只提供**完整 zip**，不提供裸 exe。

3. 确认解压目录里有 `Sbroenne.ExcelMcp.McpServer.exe`（约 139 个文件）。

### 方式 B：从源码构建

```bash
git clone https://github.com/Asaceoo/Execl-mcp.git
cd Execl-mcp

# 构建（项目路径传相对路径）
bash build.sh build -c Release src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj

# 一键发布：版本递增 → 构建 → 部署到 excel-mcp-bin/ → 产物落 dist/ → 自校验
bash release.sh --no-bump
```

`Directory.Build.props` 里的 `<Version>` 是版本号的**唯一来源**，所有产物都带该版本号后缀。
构建脚本可用的环境变量覆盖项（默认取脚本所在目录）：

```bash
export EXCEL_MCP_ROOT=/d/my/jyyj-mcp              # 工作区根
export EXCEL_MCP_REPO=$EXCEL_MCP_ROOT/mcp-server-excel
export EXCEL_MCP_DOTNET=$EXCEL_MCP_ROOT/.dotnet10/dotnet.exe
export EXCEL_MCP_NUGET=$EXCEL_MCP_ROOT/.nuget-packages
```

---

## 4. 接入 MCP 客户端

### 通用 JSON 配置

在客户端的 MCP 配置里加一段（路径改成你的解压位置）：

```json
{
  "mcpServers": {
    "jyyj-mcp": {
      "command": "D:\\jyyj-mcp\\Sbroenne.ExcelMcp.McpServer.exe",
      "args": [],
      "timeout": 600
    }
  }
}
```

`timeout` 建议给足：Excel COM 启动会话本身就有秒级开销，长任务（刷新 Power Query、
跑大范围计算）可能需要数分钟。

### WorkBuddy 连接器

把上面的配置写入 `~/.workbuddy/mcp.json` 的 `mcpServers` 之后，**新服务器不会自动生效**——
需要打开「连接器管理」页，在右上角的「自定义连接器」入口里对新服务器点一次 **「信任」**。

### 验证接好了

重启客户端，在工具列表里应看到 **31 个工具**，名字如下：

```
analysis  calculation_mode  chart  chart_config  conditionalformat
connection  datamodel  datamodel_relationship  drawing  file
namedrange  pivottable  pivottable_calc  pivottable_field  powerquery
pythoninexcel  querytable  range  range_edit  range_format
range_link  screenshot  slicer  table  table_column
vba  window  workbook  worksheet  worksheet_style  xmlmap
```

也可以用仓库自带探针离线校验（不需要 AI 客户端）：

```bash
python probe_mcp.py          # 校验已部署 server 的工具 schema
```

---

## 5. 核心心智模型：一切都围绕 session

**这是最容易踩坑的地方。** 除了 `file` 工具本身，**所有工具都需要 `session_id`**。

```
file(create|open)  ──►  返回 session_id  ──►  拿它调其它 29 个工具  ──►  file(close, save)
     ▲                                                                        │
     └──────────────────── 会话期间 Excel 在后台独占该文件 ◄──────────────────┘
```

### 一次完整的调用序列

```python
# 1) 建会话 —— 这一步是"打开 Excel"
r = file(action="create", path="D:\\work\\report.xlsx")
sid = r["session_id"]

# 2) 干活 —— 所有操作都带上同一个 session_id
worksheet(action="create",  session_id=sid, sheet_name="Data")
range(action="set-values",  session_id=sid, sheet_name="Data",
      range_address="A1:C4", values=[["Region","Q1","Q2"],
                                    ["East",100,130],
                                    ["South",200,170],
                                    ["North",300,240]])

# 3) 收尾 —— save=true 才会写盘
file(action="close", session_id=sid, save=True)
```

### `file` 工具：唯一入口

| 动作 | 作用 | 关键参数 |
|---|---|---|
| `create` | 新建一个空工作簿并开会话 | `path`（**必须由用户提供，不要猜路径**） |
| `open` | 打开已有工作簿并开会话 | `path` |
| `close` | 关闭会话 | `session_id`、`save`（默认 `false`＝**丢弃改动**） |
| `list` | 列出当前活跃会话 | — |
| `test` | 检查文件是否可被打开（可用性预检） | `path` |

### 三个必须知道的参数

| 参数 | 默认 | 说明 |
|---|---|---|
| `save`（`file close`） | **`false`** | 默认**丢弃**所有改动。想保留结果必须显式传 `save: true` |
| `show`（`file create`/`open`） | **`false`** | 默认后台无头运行，**看不到 Excel 窗口**。要在屏幕上看着它操作就传 `show: true` |
| `timeout_seconds` | `120` | 范围为 10–3600 秒。大工作簿打开、Power Query 刷新建议调大 |

> **看不到窗口是设计如此，不是坏了。** 无头是默认行为（源码注释：*"default: false for
> background automation"*）。已经开好的会话想让它显示出来，用
> `window(action="show")` 再 `window(action="bring-to-front")`。
> 但注意 `screenshot capture-sheet` **必须**在可见会话下用。

---

## 6. 工具速查表（31 个）

按「你要做什么」组织，不是按字母序。

### 会话与文件

| 工具 | 动作数 | 干什么 |
|---|---|---|
| `file` | 5 | **会话入口**：create / open / close / list / test |
| `workbook` | 15 | 文档属性、另存为、导出固定格式、外部链接、保护、视图选项 |
| `calculation_mode` | 3 | 切换自动/手动计算。**批量写入前切手动能快一个数量级** |

### 工作表与区域

| 工具 | 动作数 | 干什么 |
|---|---|---|
| `worksheet` | 8 | 建、改名、复制、删除、移动；跨文件复制/移动 |
| `worksheet_style` | 25 | 标签颜色、页面设置、**显示/隐藏/深度隐藏**、**分组与分级显示**、批注、保护 |
| `range` | 17 | 读写值与公式、清空、复制、数字格式、取已用区域 |
| `range_edit` | 9 | 插/删行列与单元格、查找替换、**排序** |
| `range_format` | 14 | 样式、条件性格式、数据验证、自动列宽、合并单元格、列宽行高 |
| `range_link` | 11 | 超链接、**现代批注（Threaded Comment）**、单元格锁定 |
| `namedrange` | 6 | 命名区域：建、读、写、改、删 |
| `conditionalformat` | 4 | 条件格式规则：加、清、列 |

### 表格 / 透视表 / 切片器

| 工具 | 动作数 | 干什么 |
|---|---|---|
| `table` | 16 | Excel 表格（ListObject）全生命周期；`preflight` 预检；接数据模型 |
| `table_column` | 12 | 表格列：增删改名、**筛选**、**多列排序**、结构化引用 |
| `pivottable` | 10 | 透视表：建（range / table / datamodel）、读、刷新、缓存选项、钻取 |
| `pivottable_field` | 15 | 字段：加行/列/值/筛选、改聚合方式、**日期分组**、数值分组、排序 |
| `pivottable_calc` | 10 | 计算字段、计算成员、布局、分类汇总、总计 |
| `slicer` | 10 | 切片器：建、选、删；**`connect-pivots` / `disconnect-pivots`（多表联动）** |

### 数据与查询

| 工具 | 动作数 | 干什么 |
|---|---|---|
| `powerquery` | 12 | Power Query M：建、改、刷新、装载、卸载、重命名、求值 |
| `querytable` | 9 | 本地 QueryTable：文本 / Web 导入、刷新、取消 |
| `connection` | 11 | 数据连接（OLEDB / ODBC / ODC）：建、刷新、装到表、属性 |
| `datamodel` | 15 | 数据模型（Power Pivot）：表、列、度量值、**执行 DMV 查询** |
| `datamodel_relationship` | 5 | 数据模型关系：建、改、读、删 |
| `xmlmap` | 6 | XML 映射：加映射、映区域、导入/导出 XML |

### 图表与视觉

| 工具 | 动作数 | 干什么 |
|---|---|---|
| `chart` | 8 | 图表生命周期：建（range / table / pivottable）、读、删、移动、贴合区域 |
| `chart_config` | 25 | 数据源、系列增删、类型、标题、坐标轴、图例、数据标签、趋势线、区域格式 |
| `drawing` | 14 | 图片、AutoShape、文本框、连接线、表单控件、**迷你图** |
| `screenshot` | 2 | 把工作表或区域导出成图片，**用于视觉校验**（需可见会话） |

### 高级 / 平台

| 工具 | 动作数 | 干什么 |
|---|---|---|
| `vba` | 6 | VBA 模块与过程：列、看、导入、改、**运行**、删（`.xlsm`） |
| `pythoninexcel` | 2 | Microsoft 365「Python in Excel」的 `=PY()` 公式：写、取结果 |
| `window` | 15 | Excel 窗口：显示/隐藏、置前、位置、状态栏、冻结窗格、缩放 |
| `analysis` | 8 | 目标求解、方案管理器（Scenario）、**模拟运算表**（Data Table） |

---

## 7. 两种列表线格式（最容易踩的坑）

MCP 面上「一串字符串」有**两种写法**，取决于工具源码里参数声明的类型。传错会被参数绑定直接拒绝。

| 类型 | 写法 | 参数 |
|---|---|---|
| **原生 JSON 数组**（9 个） | `["a","b"]` | `range.values`、`range.formulas`、`range.formats`、`range_format.range_addresses`、`table.rows`、`range_edit.sort_columns`、`table_column.sort_columns`、`vba.parameters`、`analysis.values` |
| **装 JSON 数组的字符串**（5 个） | `"[\"a\",\"b\"]"` | `slicer.pivot_table_names`、`slicer.selected_items`、`pivottable_field.selected_values`、`pivottable_field.item_names`、`table_column.values` |

### 例子：同一个「两个值」的两种写法

```python
# ✅ 原生数组：range.values 声明为 array
range(action="set-values", session_id=sid, sheet_name="S1",
      range_address="A1:B1", values=[["a", "b"]])

# ✅ 字符串装 JSON：slicer.pivot_table_names 声明为 string
slicer(action="connect-pivots", session_id=sid, slicer_name="区域",
       pivot_table_names='["Pivot1", "Pivot2"]')

# ❌ 会被拒绝：把数组传给了声明为 string 的参数
slicer(action="connect-pivots", session_id=sid, slicer_name="区域",
       pivot_table_names=["Pivot1", "Pivot2"])
```

传错时的报错是这样的（工具会主动告诉你怎么改）：

```
Argument 'pivot_table_names' was sent as a JSON array, but this tool declares it as a
string that carries a JSON array - send it quoted, for example
"pivot_table_names":"[\"value1\",\"value2\"]".
```

**记住这一条就够了**：名字里带 `_names` / `_items` / `selected_*` 的（切片器、透视表字段）
基本是**字符串装 JSON**；`values` / `rows` / `formulas` / `sort_columns` 这类是**原生数组**。
拿不准时，先按错的方式发一次——报错会直接告诉你正确写法，不用翻文档。

---

## 8. 实战配方

> 下面所有例子都省略了 `session_id` 的传递，实际调用每条都要带。

### 配方 1：从数据到带格式的报表

```python
file(action="create", path="D:\\work\\report.xlsx")           # → sid

# 写入数据
worksheet(action="create", sheet_name="Data")
range(action="set-values", sheet_name="Data", range_address="A1:D7",
      values=[["Region","Q1","Q2","Q3"],
              ["East",100,130,120], ["South",200,170,180],
              ["North",300,240,260], ["West",180,260,250],
              ["Centre",90,150,140], ["Northeast",140,110,130]])

# 变成 Excel 表格（后续透视表/图表更好引用）
table(action="create", sheet_name="Data", table_name="Sales",
      range_address="A1:D7", has_headers=True)

# 美化
range_format(action="format-range", sheet_name="Data",
             range_address="A1:D1", bold=True,
             fill_color="#1F4E78", font_color="#FFFFFF")
range_format(action="auto-fit-columns", sheet_name="Data", range_address="A1:D7")
table(action="set-style", sheet_name="Data", table_name="Sales",
      style_name="TableStyleMedium2")

workbook(action="set-document-property", property_name="Title", value="季度销售")
file(action="close", save=True)
```

### 配方 2：透视表 + 切片器驱动多张表（旗舰能力）

**前提知识**：Excel 的一个切片器只能驱动**共享同一 PivotCache** 的多张透视表。
本工具用 `share_cache_from` 建缓存共享，用 `connect-pivots` 把切片器接上去。

```python
# ① 第一张透视表（它自建缓存）
pivottable(action="create-from-table",  session_id=sid,
           pivot_table_name="Pivot1", table_name="Sales",
           destination_sheet="Dash", destination_cell="A3")

# ② 第二张透视表 —— 关键：share_cache_from 复用第一张的缓存
#    注意每个透视表要独占一张工作表（Excel 的硬限制）
worksheet(action="create", session_id=sid, sheet_name="Dash2")
pivottable(action="create-from-table",  session_id=sid,
           pivot_table_name="Pivot2", table_name="Sales",
           destination_sheet="Dash2", destination_cell="A3",
           share_cache_from="Pivot1")          # ← 这一行是"联动"的前提

# ③ 配字段
pivottable_field(action="add-row-field",   session_id=sid,
                 pivot_table_name="Pivot1", field_name="Region")
pivottable_field(action="add-value-field", session_id=sid,
                 pivot_table_name="Pivot1", field_name="Q1", function="Sum")

# ④ 建一个切片器
slicer(action="create-slicer", session_id=sid,
       slicer_name="区域", pivot_table_name="Pivot1",
       field_name="Region", destination_sheet="Dash", position="H2")

# ⑤ 把切片器接到第二张表上（列表参数 → 字符串装 JSON）
slicer(action="connect-pivots", session_id=sid, slicer_name="区域",
       pivot_table_names='["Pivot1", "Pivot2"]')

# ⑥ 想断开某张表（至少保留一个连接，重复名字会自动去重）
slicer(action="disconnect-pivots", session_id=sid, slicer_name="区域",
       pivot_table_names='["Pivot2"]')

# ⑦ 选中某些项
slicer(action="set-slicer-selection", session_id=sid, slicer_name="区域",
       selected_items='["East", "South"]', clear_first=True)
```

> 常见错误：`connect-pivots` 报 `COMException 0x800A03EC`。原因几乎总是
> **两张透视表的缓存不是同一个** —— 检查第 ② 步有没有传 `share_cache_from`。

### 配方 3：建图表（含现代图与股价图）

```python
# 普通柱形图
chart(action="create-from-range", session_id=sid,
      sheet_name="Data", chart_name="ch1", chart_type="ColumnClustered",
      source_range_address="A1:D7", target_range="F2:L18")

# 现代图（树状图 / 直方图 / 帕累托 / 箱形图 / 瀑布图 / 漏斗图 / 区域地图 / 柱线组合）
chart(action="create-from-range", session_id=sid,
      sheet_name="Data", chart_name="ch_treemap", chart_type="Treemap",
      source_range_address="A1:D7", target_range="F20:L36")

# 股价图（HLC / OHLC / VHLC / VOHLC）
# ⚠️ 列数必须精确匹配：HLC=3 列, OHLC/VHLC=4 列, VOHLC=5 列，且不要额外的类别列
chart(action="create-from-range", session_id=sid,
      sheet_name="StockData", chart_name="ch_stock", chart_type="StockHLC",
      source_range_address="A1:C30", target_range="E2:K18")

# 后续配置
chart_config(action="set-title", session_id=sid, chart_name="ch1", title="季度销售")
chart_config(action="set-chart-type", session_id=sid, chart_name="ch1", chart_type="BarClustered")
chart_config(action="add-series", session_id=sid, chart_name="ch1",
             series_name="Q4", source_range="Data!E2:E7")
chart_config(action="set-data-labels", session_id=sid, chart_name="ch1", show_value=True)

# 用截图校验版面有没有压住数据
screenshot(action="capture-sheet", session_id=sid, sheet_name="Data",
           range_address="F1:L40")
```

> **可用图表 37 种**（基础 25 + 现代 8 + 股价 4）。Excel 一共能枚举 84 种，其余需要
> 数据模型、地图服务或前台窗口等外部条件。传了不可用的类型，工具会**明确报错**
> 而不是静默降级；列数不对也会告诉你「这个变体需要几列、你给了几列」。

### 配方 4：Power Query / DAX / 数据模型

```python
# 建 Power Query 并装载到表
powerquery(action="create", session_id=sid, query_name="Sales",
           m_code='let Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content] in Source')
powerquery(action="load-to", session_id=sid, query_name="Sales",
           destination_sheet="PQ", destination_cell="A1", load_type="table")

# 把表加进数据模型，然后建度量值
table(action="add-to-data-model", session_id=sid, sheet_name="PQ", table_name="Sales")
datamodel(action="create-measure", session_id=sid, table_name="Sales",
          measure_name="TotalQ1", dax_expression="SUM(Sales[Q1])")

# 跨表计算要建关系
datamodel_relationship(action="create-relationship", session_id=sid,
                       from_table="Sales", from_column="RegionId",
                       to_table="Region", to_column="Id")

# 直接跑 DAX / DMV 查询
datamodel(action="evaluate", session_id=sid,
          dax_expression='EVALUATE SUMMARIZECOLUMNS(Sales[Region], "Q1", SUM(Sales[Q1]))')
```

### 配方 5：用 VBA 给切片器换皮肤

切片器外观在原生 API 里改不了，但可以用 VBA 兜底（需要 `.xlsm` 且已开启 VBA 访问信任）：

```python
vba(action="import", session_id=sid, module_name="Skin", code=(
    'Sub Paint()\n'
    '    Dim sc As SlicerCache\n'
    '    For Each sc In ActiveWorkbook.SlicerCaches\n'
    '        Dim sl As Slicer\n'
    '        For Each sl In sc.Slicers\n'
    '            sl.Style = "SlicerStyleDark2"\n'
    '        Next sl\n'
    '    Next sc\n'
    'End Sub\n'))
vba(action="run", session_id=sid, procedure="Skin.Paint")
```

内置样式共 12 种：`SlicerStyleLight1`–`SlicerStyleLight6`、`SlicerStyleDark1`–`SlicerStyleDark6`
（本机实测 **12/12 全部可用**，且随保存持久化）。

### 配方 6：分析工具

```python
# 目标求解：让 A10 等于 1000，调整 B2
analysis(action="goal-seek", session_id=sid, sheet_name="Model",
         target_cell="A10", target_value=1000, changing_cell="B2")

# 方案管理器
analysis(action="create-scenario", session_id=sid, sheet_name="Model",
         scenario_name="乐观", changing_cells=["B2","B3"], values=[120, 90])
analysis(action="create-scenario-summary", session_id=sid,
         sheet_name="Model", result_cell="A10", report_type="Summary")

# 模拟运算表（双变量）
analysis(action="create-data-table", session_id=sid, sheet_name="Model",
         row_input_cell="B1", column_input_cell="B2", formula="=A10")
```

---

## 9. 错误排查表

| 你看到的报错 | 原因 | 怎么办 |
|---|---|---|
| `The application to execute does not exist`（退出码 2147516570） | 只拷了 `.exe`，缺同目录的 dll/deps.json | 用完整解压目录，别单独拷 exe |
| `Argument 'xxx' was sent as a JSON array, but this tool declares it as a string...` | 列表参数写成了原生数组 | 按报错里的示例加引号：`"xxx":"[\"a\",\"b\"]"` |
| `Tool 'xxx' rejected the arguments before executing` | 参数名拼错，或传了该工具不接受的参数（如 `pivot_table_names` 传给 `pivottable`） | 看报错里的 `Accepted argument(s)` 列表；**补救提示在最前面** |
| `Invalid value 'X' for parameter 'chartType'. Valid values: ...` | 图表类型名不对 | 从报错列出的合法值里选 |
| `Excel fell back to a different chart type` | 该图表类型在这台 Excel 上不可用（如旭日图） | 换类型；这是 Excel 的限制，不是缺陷 |
| `Excel refused to convert the scaffold chart into StockHLC (COM 0xB0D7018E)... has 5 column(s)` | 股价图列数与变体不匹配 | 按报错提示的列数规则调整源区域 |
| `Modern charts need at least one numeric column` | 现代图的数据区全是文本 | 加一列数字 |
| `Slicer 'X' not found in workbook` | 切片器名字写错，或还没创建 | 先 `slicer list-slicers` 看真实名字 |
| `connect-pivots` 报 `COMException 0x800A03EC` | 两张透视表**没有共享 PivotCache** | 建第二张表时传 `share_cache_from` |
| `connect-pivots requires at least one non-empty PivotTable name` | 列表里全是空串或空格 | 传真实的表名 |
| `Value array row 2 column count (1) doesn't match range column count (2)` | 二维数组行列不齐 | 补齐成矩形 |
| `0x800706BA`（RPC 服务器不可用） | 连续自动化约 1 小时后 Excel COM 瞬态不可用 | 停下重试；不是项目缺陷 |
| 操作报文件被占用 | 目标工作簿正在 Excel 里打开 | 先在 Excel 里关闭该文件 |
| 看不到 Excel 窗口 | `show` 默认 `false`（设计如此） | 建会话时传 `show: true`，或 `window show` + `bring-to-front` |
| 改动没保存 | `file close` 的 `save` 默认 `false` | 传 `save: true` |
| 透视表建不出来 | 同一张工作表已经有别的透视表 | Excel 限制：**每张表只能有一个透视表**，换个 sheet |
| `vba` 相关失败 | 未开启「信任对 VBA 工程对象模型的访问」 | 在 Excel 信任中心里开启 |

---

## 10. 已知限制

| 限制 | 说明 |
|---|---|
| 旭日图（Sunburst） | Excel 自身会回退成其它类型，工具**明确报错**而不是假报成功 |
| `screenshot capture-sheet` | 必须 `show: true` 的可见会话；无头会话下必然失败 |
| 连续自动化 | 约 1 小时后 Excel COM 可能出现瞬态不可用（`0x800706BA`），停下重试即可 |
| 一个工作表一个透视表 | Excel 的硬限制，不是工具限制 |
| VBA | 需要在 Excel 里手工开启信任 |
| 平台 | 只能跑在 Windows + 桌面版 Excel；WPS 与 Excel 网页版不适用 |
| 并发 | 单个 Excel 进程一次服务一个会话；要并行处理多个文件需错开 |

---

## 相关文档

| 文档 | 内容 |
|---|---|
| [README.md](README.md) | 项目概览、相对上游做了什么 |
| [TECHNICAL-MANUAL.md](TECHNICAL-MANUAL.md) | 技术手册：架构、COM 层、补丁原理、实测事实集、构建与验证 |
| [DEPLOYMENT.md](DEPLOYMENT.md) | 部署手册：版本、补丁表、MCP 配置、验证证据、重建步骤 |
| [ADVERSARIAL-REVIEW.md](ADVERSARIAL-REVIEW.md) | 对抗性审查报告：问题清单 + 真机证据 + 处置 |

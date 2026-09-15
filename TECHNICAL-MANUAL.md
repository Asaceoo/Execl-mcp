# jyyj-mcp 助手 · 技术手册

面向**维护者与二次开发者**：架构、COM 层设计、补丁原理、实测事实、构建与验证链路。

- 版本：`3.0.0.0`（四段版本，唯一来源 `mcp-server-excel/Directory.Build.props` 的 `<Version>`）
- 工具面：31 个工具 / 328 个动作 / 534 个参数
- 基线：上游 `sbroenne/mcp-server-excel` v2.0.8（MIT）
- 运行时：.NET 10（`net10.0` / `net10.0-windows`）

> 面向使用者的内容见 [用户手册](USER-GUIDE.md)；部署与验证证据见 [DEPLOYMENT.md](DEPLOYMENT.md)；
> 问题清单与处置见 [ADVERSARIAL-REVIEW.md](ADVERSARIAL-REVIEW.md)。

---

## 目录

1. [分层架构](#1-分层架构)
2. [会话模型与 COM 层](#2-会话模型与-com-层)
3. [工具面的生成机制](#3-工具面的生成机制)
4. [两种列表线格式的根因](#4-两种列表线格式的根因)
5. [本地补丁面（五类修复）](#5-本地补丁面五类修复)
6. [实测事实集（Excel COM 行为）](#6-实测事实集excel-com-行为)
7. [构建与发布链路](#7-构建与发布链路)
8. [验证体系与反假绿守卫](#8-验证体系与反假绿守卫)
9. [扩展指南](#9-扩展指南)
10. [诊断工具速查](#10-诊断工具速查)

---

## 1. 分层架构

```
┌──────────────────────────────────────────────────────────────────────┐
│  MCP 客户端（Claude Desktop / Cursor / WorkBuddy …）                  │
└───────────────────────────┬──────────────────────────────────────────┘
                            │ stdio JSON-RPC
┌───────────────────────────▼──────────────────────────────────────────┐
│  ExcelMcp.McpServer      （可执行：Sbroenne.ExcelMcp.McpServer.exe）  │
│   ├─ SessionIdentityFilter   会话身份缺失/歧义的早报错                │
│   ├─ ToolErrorSurfaceFilter  参数绑定失败 → 可读文本 + 补救提示        │
│   ├─ Tools/Excel*Tool.cs     工具类（[McpServerToolType] + 手写文档） │
│   └─ ServiceBridge           对 Core 的调用桥                          │
└───────────────────────────┬──────────────────────────────────────────┘
                            │ 进程内调用
┌───────────────────────────▼──────────────────────────────────────────┐
│  ExcelMcp.Service        （会话调度、超时、错误分类）                  │
└───────────────────────────┬──────────────────────────────────────────┘
                            │
┌───────────────────────────▼──────────────────────────────────────────┐
│  ExcelMcp.Core           业务命令层（Commands/<Domain>/…）             │
│   每个域一个 I*Commands 接口 + 实现；方法名 → kebab-case 动作名        │
└───────────────────────────┬──────────────────────────────────────────┘
                            │
┌───────────────────────────▼──────────────────────────────────────────┐
│  ExcelMcp.ComInterop     会话、COM 释放、写保护、OLE 消息过滤、进度    │
└───────────────────────────┬──────────────────────────────────────────┘
                            │ COM（late binding）
┌───────────────────────────▼──────────────────────────────────────────┐
│  Microsoft Excel 桌面进程                                             │
└──────────────────────────────────────────────────────────────────────┘
```

配套项目：

| 项目 | 作用 |
|---|---|
| `ExcelMcp.Generators.Mcp` | **源生成器**：从 Core 的接口 + 属性生成 MCP 工具的 schema 与参数绑定 |
| `ExcelMcp.Generators` | 生成 ServiceRegistry / 技能清单（`_SkillManifest.g.cs`） |
| `ExcelMcp.Generators.Shared` | 生成器共用的字符串/命名工具 |
| `ExcelMcp.Build.Tasks` | 构建期任务（技能文档模板渲染等） |
| `ExcelMcp.CLI` | `excelcli.exe`：同一套 Core 能力的命令行入口 |
| `ExcelMcp.Cleanup` | 清理工具 |

### 依赖方向是单向的

```
McpServer ─► Service ─► Core ─► ComInterop
                              └► Generators（编译期，只读属性）
```

Core **不认识** MCP，只暴露接口与属性。这一条决定了：新增能力应该写在 Core，
MCP 面由生成器推导——**不要手写 schema**。

---

## 2. 会话模型与 COM 层

### 一切围绕 session

Excel COM 的自动化单元是「一个 Excel 进程 + 一个已打开的工作簿」。工具面把它建模成 **session**：

- `file create` / `file open` 建立一个会话，返回 `session_id`；
- 之后每个工具调用都带 `session_id`，Service 用它路由到对应的 Excel 实例；
- `file close` 结束会话；`save=false`（默认）丢弃改动。

这条设计的直接后果：**目标工作簿在操作期间必须在 Excel 里关闭**——COM 需要独占访问。

### 无头是默认，不是缺陷

建会话时 `tempExcel.Visible = _showExcel`，而 `show` 参数默认 `false`
（源码注释：*"default: false for background automation"*）。所以「看不到 Excel 窗口」是设计如此。
要可见就显式 `show: true`，或对已开会话用 `window show` + `window bring-to-front`。

唯一**必须**可见的功能是 `screenshot capture-sheet`（无头会话下捕获必然失败）。

### `ComUtilities.Release`：确定性的 COM 释放

迟到绑定（`dynamic`）的 COM 对象不会自动释放，泄漏会累积到 Excel 进程里。
代码里统一的模式是：

```csharp
dynamic? chart = null;
try
{
    chart = shape.Chart;
    // … 使用 chart
}
finally
{
    ComUtilities.Release(ref chart!);
}
```

`Release(ref T)` 在释放后把引用置空，因此**重复释放是安全的**，且调用方可以无条件写
`finally { ComUtilities.Release(ref x!); }`。这是本仓库所有 COM 代码的硬约定。

### 失败即清理：不留下半个对象

创建类操作（图表、形状）失败时，必须删掉已经建出来的半成品，否则工作表上会残留空对象：

```csharp
catch (COMException ex) when (IsStockShapeRefusal(ex))
{
    TryDeleteShape(stockShape);      // 先清理
    throw new InvalidOperationException(...);  // 再抛带解释的错
}
```

`TryDeleteShape` 自己吞掉清理阶段的异常——**原始失败才是要浮出来的那个**。

### 其他 COM 层设施

| 组件 | 作用 |
|---|---|
| `OleMessageFilter` / `IOleMessageFilter` | 抑制「服务器忙」弹窗导致的死锁（Excel 自动化经典坑） |
| `ExcelWriteGuard` | 批量写入期间抑制 `ScreenUpdating` 等，失败可恢复 |
| `FileAccessValidator` | 打开前校验文件可访问/未被占用 |
| `ExcelShutdownService` | 带重试的优雅退出（应对瞬态锁 `0x800706BA`） |
| `ComDiagnostics` | COM 层诊断埋点 |
| `OleCompoundFileReader` | 不启动 Excel 直接读 OLE 复合文档（轻量探测） |

---

## 3. 工具面的生成机制

### 一个动作的完整链路

```
Core 接口方法（带 [ServiceAction("connect-pivots")]）
      │
      ├─► ServiceRegistryGenerator  生成 ServiceRegistry（编译期常量）
      │
      └─► McpToolGenerator          生成 MCP 参数绑定 + JSON schema
                │
                └─► 运行时：tools/list 返回 31 个工具，每个工具的 action 是 enum
```

动作名默认由方法名经 **PascalCase → kebab-case** 推导
（`GetLoadConfig` → `get-load-config`），需要覆盖时用 `[ServiceAction("…")]`。

### schema 的关键规则（生成器决定，不手写）

| Core 侧声明 | MCP 面上 | 说明 |
|---|---|---|
| `action`（enum） | `enum`，**校验失败会列出全部合法值** | 这是错误可操作性的主要来源 |
| `int?` / `bool?` / `double?` | 同名可选参数 | 所有参数在 MCP 面都是可选；缺省由 Core 分发时补 |
| `TimeSpan` | `int?`（秒） | 生成器把名字改成 `*_seconds`，并追加 `Accepted range: …` |
| `List<string>` | **`string`**，内容是 JSON 数组 | 见下一节 |
| `string[]` | **`array`** | 原生 JSON 数组 |
| 其他 | 直通，追加 `?` 变可选 | |

参数名统一 `snake_case`（生成器从 PascalCase 转），实测 **534/534 全部 snake_case**。

### 参数描述里的"这个参数对哪些动作有效"

生成器会把每个参数的适用动作写进描述：

```
(required for: create-slicer) (valid for: create-slicer, list-slicers)
```

这让模型在**不读文档**的情况下也能知道某参数该不该传。改动 Core 接口后这些文本自动更新。

### 内置的技能提示词

`src/ExcelMcp.McpServer/Program.cs` 的 `ServerInstructions` 会在会话初始化时下发给模型，
里面有工作流约定、列表参数线格式说明等。`skills/` 下的 `excel-mcp/SKILL.md` 与
`excel-cli/SKILL.md` 由 `*.sbn` 模板在**构建期**渲染生成——**改模板才生效，改产物没用**。

---

## 4. 两种列表线格式的根因

这是本项目最容易被误判成"bug"的设计后果，值得单独一节。

### 根因

`McpToolGenerator.cs` 对 `List<string>` 的分支（约 432–447 行）：

```csharp
else if (ep.TypeName.StartsWith("System.Collections.Generic.List<string>") ||
         ep.TypeName.StartsWith("List<string>"))
{
    // List<string> → string (JSON array) in MCP, parse via ParseJsonList
    …
    preProcessingCode: $"var {localVarName} =
        Sbroenne.ExcelMcp.Core.Utilities.ParameterTransforms.ParseJsonList({snakeName}, nameof({snakeName}));"
}
```

即：**Core 里声明为 `List<string>` 的参数，在 MCP 面上被降级成 `string`**，
由 `ParameterTransforms.ParseJsonList` 负责把字符串解析回列表。
而声明为 `string[]` 的参数走 `else` 分支直通，保持原生 `array`。

结果是同一个「一组字符串」有两种线上写法，取决于上游作者当时用了哪种集合类型。

### 当前分布（实测，来自 deployed server 的 schema）

| 线格式 | 数量 | 参数 |
|---|---|---|
| **`string`（装 JSON 数组）** | 5 | `slicer.pivot_table_names`、`slicer.selected_items`、`pivottable_field.selected_values`、`pivottable_field.item_names`、`table_column.values` |
| **`array`（原生）** | 9 | `range.values`、`range.formulas`、`range.formats`、`range_format.range_addresses`、`table.rows`、`range_edit.sort_columns`、`table_column.sort_columns`、`vba.parameters`、`analysis.values` |

### 本地处置（不改上游契约）

不做破坏性统一（那会破坏 CLI 与既有调用），而是在**参数绑定失败时把补救写进错误文本**：

`ToolErrorSurfaceFilter.ShapeMismatches` 比较「实际送来的 JSON 形状」与
「schema 声明的类型」，命中就给出一句话修法：

```
Argument 'pivot_table_names' was sent as a JSON array, but this tool declares it as a
string that carries a JSON array - send it quoted, for example
"pivot_table_names":"[\"value1\",\"value2\"]".
```

### 一个由此暴露的**次级缺陷**（已修）

这段补救提示原先被**追加在 `Accepted argument(s): …` 长列表之后**，
导致它在字符串尾部；而客户端的输出上限（本仓库的边界探针当时是 400 字符）
恰好把它切掉——**最有用的一句话成了最先被丢弃的一句话**。

实测：补救位置从第 437 字符**前移到第 270 字符**（总长 674），
现在稳居任何合理截断之内。

> **一般化原则**：错误文本里，**可操作的补救必须排在最前**，诊断细节排后面。
> 任何"先列一堆上下文、最后才说要怎么改"的错误消息，在真实链路上都是不可用的。

---

## 5. 本地补丁面（五类修复）

相对上游 v2.0.8 的全部改动，按"修什么"归类。

### 补丁 1：切片器多表联动（新增能力，不是修 bug）

**背景**：Excel 的一个切片器只能驱动**共享同一 `PivotCache`** 的多个透视表。
上游没有暴露这个能力，文档里写的是"在 Excel UI 里手工做"。

**实现**：三个新动作/参数，全走原生 COM。

| 工具 | 新增 | 语义 |
|---|---|---|
| `pivottable` | 参数 `share_cache_from` | 新建透视表时复用已有透视表的 PivotCache |
| `slicer` | 动作 `connect-pivots` | `SlicerCache.PivotTables.AddPivotTable(...)` |
| `slicer` | 动作 `disconnect-pivots` | 安全断开：**至少保留一个连接**，重复名字去重 |

关键约束（真机实测）：

- 跨缓存 `AddPivotTable` → `COMException 0x800A03EC`，连接静默不生效。
  这是 `connect-pivots` 报错的首要原因。
- 对已移除条目重复 `RemovePivotTable` → `COMException 0x8002000B`（无效索引），
  **不静默忽略**。
- 请求里重复的表名必须**先去重**：既污染"至少保留一个"的算术，也会让删除执行两次。

### 补丁 2：图表创建路径分流（可用图表 25 → 37 种）

**背景**：上游走单一 `AddChart` + `SetSourceData` 路径。现代图与股价图在这条路径上必然失败。

**实现**：`Core/Commands/Chart/ChartCreationPath.cs` 按类型三分流。

| 路径 | 适用 | 做法 |
|---|---|---|
| 默认 | 全部经典图 | `Shapes.AddChart(code)` + `Chart.SetSourceData(range)` |
| 现代图 | Treemap/Sunburst/Histogram/Pareto/BoxWhisker/Waterfall/Funnel/RegionMap/ColumnLineCombo | `Shapes.AddChart2(-1, code)` + 逐列 `SeriesCollection.NewSeries()` |
| 股价图 | StockHLC/OHLC/VHLC/VOHLC | 先建柱形图 → `SetSourceData(range, PlotBy=2)` → 改 `ChartType` |

**现代图为什么必须手动挂系列**：这些类型对 `SetSourceData` 抛 `0x800A03EC`，
但图表**本身已经建出来了**——所以要在 `catch` 里先删掉半成品再抛错。

**股价图的两个坑**：

1. **系列数必须与变体精确匹配**：HLC=3、OHLC/VHLC=4、VOHLC=5 列，且**不接受额外的类别列**。
   多一列少一列都拒。
2. **`PlotBy` 必须显式传 `2`（xlColumns）**。否则 `AddChart(51)` 会先用整区 `CurrentRegion`
   自动填充，`SetSourceData` 改不动那个系列数，切 `ChartType` 时抛 scode `0xB0D7018E/190/191`。
   失败集合会随夹具列数漂移，恰好列数匹配的变体反而"假绿"——极难定位。

**本地新增的错误解释**（本轮）：

```csharp
catch (COMException ex) when (IsStockShapeRefusal(ex))
{
    TryDeleteShape(stockShape);
    throw new InvalidOperationException(
        $"Excel refused to convert the scaffold chart into {chartType} (COM 0x{ex.HResult:X8}). "
        + "A stock chart takes exactly one value column per series and no separate category column: "
        + "StockHLC needs 3 columns, StockOHLC and StockVHLC need 4, StockVOHLC needs 5. "
        + ColumnCountPhrase(sourceRange), ex);
}
```

`IsStockShapeRefusal` **刻意不包含**通用的 `0x800A03EC`：在这个分支上它也可能是"源区域不可用"，
把列数规则扣到它头上是错误断言。为兜住这种情况另有一个不带因果断言的 `catch (COMException)`。

**创建后回读校验**：`VerifyChartType` 比对回读的 `ChartType` 与请求值，
不一致就**删掉图表并报错**——拒绝静默降级。旭日图就是这样被识别为"Excel 不支持"的
（本机 COM 回读 `ChartType = -4111`）。

### 补丁 3：`chart_config add-series` 传 Range 对象

原实现把地址字符串直接赋给 `Series.Values`，恒定抛 `0x800A03EC`
（该属性要的是 `Range` 对象）。现改为按图表所在工作表把地址解析成 `Range` 再赋值。

### 补丁 4：参数绑定错误透传

**背景**：MCP SDK 在**参数绑定阶段**抛的异常不进入工具内部的 `catch`，
导致 6 种不同的调用错误被吞成同一句 "An error occurred invoking 'x'"。

**实现**：`ToolErrorSurfaceFilter.Wrap`（注册在 `SessionIdentityFilter` **之后**）
捕获并转成文本，内容依次是：

1. 失败句 + 绑定器原始消息；
2. **形状补救提示**（`ShapeMismatches`，位置已前移，见第 4 节）；
3. 未声明的参数名；
4. 工具接受的参数名全表；
5. 重发指引。

### 补丁 5：现代图读取容错

现代图的 `series.Values` / `XValues` / `Formula` 与 `chart.PivotLayout`
在**读取**时抛 `E_NOTIMPL`（迟到绑定器表现为 `NotImplementedException`），
会让 `chart list` / `chart read` 整体崩溃。

**实现**：`CanHandle` 同时捕 `COMException` 与 `NotImplementedException`；
读取路径降级为 `SafeText`。**写入能力不受影响**——这些属性写是好的，只是读不了。

### 附带的工程性修复（本轮）

| 项 | 修复 |
|---|---|
| 工作表名含单引号 | 图表数据源地址拼接未转义 → `'O''Brien'!A1:D7`。抽出 `Core/Utilities/SheetReference.cs` 统一 `QuoteSheetName` + `BuildRangeReference`，配 `SheetReferenceTests` 单测 |
| 死代码 | `ChartCreationPath.NeedsSeriesAttach` / `IsStockChart` 声明后零引用（`internal` 成员编译器不告警，只能靠人工/审查发现）→ 删除 |
| 改名残留 | `ToolErrorSurfaceFilter` 与 `ExcelToolsBase` 日志前缀 `[ExcelMcp]` → `[jyyj-mcp]` |
| 缺参数描述 | `screenshot` 的 `sheet_name` / `range_address` / `quality` 三个参数无 `[Description]` → 补齐 |
| 过时描述 | `chart` 工具描述 "70+ types" → "84 types are enumerable and 37 are verified to work" |
| 引导缺失 | `ServerInstructions` 补「LIST PARAMETERS - two wire formats」段落 |

---

## 6. 实测事实集（Excel COM 行为）

以下都是**真机验证过**的行为，不是文档抄录。踩坑时先查这张表。

### 透视表与切片器

| 事实 | 含义 |
|---|---|
| `PivotCache.Index` **不可**用于判断是否同一缓存（独立缓存都为 0） | 判定"共享缓存"只能靠 `SourceData` 比对 |
| `PivotCache.SourceData` 是 **R1C1** 记法（`'Sheet'!A1:D6` → `Sheet!R1C1:R6C4`） | 整列源两侧都是 `C1:C4`，不会误拒 |
| 跨缓存 `SlicerCache.PivotTables.AddPivotTable` → `0x800A03EC` | 连接不生效；先 `share_cache_from` |
| 重复 `RemovePivotTable` → `0x8002000B` | 无效索引，不静默忽略 |
| 一个切片器驱动多表 ⟺ 共享同一 PivotCache | 联动的充要条件 |
| **每张工作表只能有一个透视表** | Excel 硬限制；集成测试夹具必须一表一透视 |
| 切片器内置样式 **12/12 全部可用** | `SlicerStyleLight1-6` + `SlicerStyleDark1-6`，随保存持久化 |
| 切片器皮肤只能用 VBA 改 | `.xlsm` + `vba import` + `vba run "模块.过程"`，`Slicer.Style = "…"` |

### 图表

| 事实 | 含义 |
|---|---|
| 枚举 84 种，**真机可用 37 种** | BASIC 18/18、3D 4/4、SHAPES 3/3、STOCK 4/4、MODERN 8/9、CONFIG 17/17、DASHBOARD 8/9 |
| 仅剩 2 项失败均为非代码限制 | 旭日图（COM 回读 `ChartType=-4111` 假成功）、截图（需 `show:true`） |
| 现代图 `SetSourceData` → `0x800A03EC`，但图表已建出 | 必须 `catch` 里清理半成品 |
| 现代图属性**写可用、读抛 `E_NOTIMPL`** | `CanHandle` 要同时捕 `COMException` 与 `NotImplementedException` |
| 股价图 `AddChart(51)` 会先用 `CurrentRegion` 自动填充 | `SetSourceData` 必须显式 `PlotBy=2` |
| 股价图系列数必须精确匹配变体 | HLC=3 / OHLC·VHLC=4 / VOHLC=5，**不接受类别列** |

### 运行环境

| 事实 | 含义 |
|---|---|
| Excel 默认无头（`show` 默认 `false`） | 看不到窗口是设计如此 |
| 连续自动化约 **1 小时**后 COM 瞬态不可用（`0x800706BA`） | 停下重试；分片回归必须允许自动复跑 |
| 本沙箱 Excel 集成测试 **≈43–70 秒/用例** | COM 启动开销主导；全量 148 项约 3 小时 |
| 日期分组字段名**随 Office 界面语言本地化** | 中文版生成 `天(Date)`/`月(Date)`/`年(Date)`，英文版为 `Days` 等。断言英文子串会在中文 Excel 上**必然假红** |
| `PivotTable.Index` 之类标识**不能**当稳定 ID | 见上：判定同一性要靠内容比对 |

---

## 7. 构建与发布链路

### 为什么必须走 `build.sh`

沙箱 shell 缺少 Windows 核心环境变量（`APPDATA` / `PROGRAMDATA` / `ALLUSERSPROFILE` /
`SystemRoot` / `windir` / `COMSPEC` / `ProgramFiles(x86)`），裸调 `dotnet` 会因
`Value cannot be null (Parameter 'path1')` 失败。`build.sh` 补齐环境块后再调固定 SDK。

```bash
bash build.sh build -c Release src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj
```

> 给 `build.sh` 的项目路径**必须传相对路径**——脚本自己会 `cd` 到仓库根，
> 传 POSIX 绝对路径在经 `env` 转发时不被 MSYS 转换，MSBuild 会把 `/d` 读成开关报 `MSB1001`。

### 一键发布

```bash
bash release.sh              # 版本递增 → 构建 → 部署 → 打包 → 自校验
bash release.sh --no-bump    # 保持当前版本
bash release.sh --no-build   # 跳过编译，只重新部署与打包
```

`release.sh` 的每一步都不是可选的：

| 步骤 | 为什么不能省 |
|---|---|
| **杀进程** | 更新二进制前必须杀掉运行中的 `Sbroenne.ExcelMcp.McpServer.exe`，否则 DLL 被锁，`cp` 报 `Device or resource busy`。Git Bash 下须 `MSYS_NO_PATHCONV=1 taskkill /F /T /PID <pid>` |
| **覆盖式复制** | **不要**递归删除部署目录（沙箱对递归删除 fail-closed）→ 覆盖式复制 + 陈旧文件报告 + `--version` 自校验兜底 |
| **Python `zipfile` 打包** | **不要**用 PowerShell `Compress-Archive`：本沙箱会被 safe-delete 钩子打断进程树，产出 **0 字节 zip** |
| **`--version` 自校验** | 部署完必须能报出版本号，否则等于没部署 |

### 版本号

单一来源：`Directory.Build.props` 的 `<Version>`，四段数字（如 `3.0.0.0`）。
`tools/pack.py` 的 `next_version` 做四段递增；`selftest` 子命令内置版本算术断言：

```bash
python tools/pack.py selftest    # 断言 5 个版本算术用例
python tools/pack.py show
python tools/pack.py bump
```

> 坑：`Directory.Build.props` 的 XML 注释里**不能出现 `--`**（如 `--version`），
> 会触发 `MSB4024: XML comment cannot contain '--'`。

产物一律**带版本号后缀**：`jyyj-mcp-3.0.0.0-win-x64.zip` + `jyyj-mcp-3.0.0.0-manifest.json`。

> **产物「带版本号」不等于「可用」**：发布只发**完整 zip**，不发裸 exe。
> .NET apphost（`.exe`）必须与 `*.dll` + `*.deps.json` + `*.runtimeconfig.json` 同目录，
> 单独拷出的 exe 必报 `The application to execute does not exist`（退出码 2147516570）。
> `pack.py` 用 `assert_bundle_runnable()` + CLI 实跑冒烟守卫这一点。

### 开源发布

```bash
python tools/publish.py check     # 索引 vs 枚举集一致性守卫
python tools/publish.py push      # 暂存 → 提交 → 推送
```

**绝对不要在工作区根 `git init`**：根目录里嵌着上游 checkout（`mcp-server-excel/.git`），
`git add -A` 只会把它记成一个 **gitlink（mode 160000）**——1100+ 源码文件静默变成一条引用，
推上去编译不了且**零报错**。`publish.py` 用**枚举式暂存**到 `_publish/` 独立树（该树内无嵌套 `.git`）。
自检：

```bash
cd _publish && git ls-files -s | awk '$1=="160000"'   # 必须为空
```

另一个坑：上游存在「**已跟踪但被自己 `.gitignore` 匹配**」的文件
（实测 `llm-tests/aitest-reports/results.json`），发布树里普通 `git add -A` 会丢掉它，
仓库静默少文件。处置：用 `git add -A -f`，并对「git 索引 vs 枚举集」做**集合差断言**。

---

## 8. 验证体系与反假绿守卫

**永远绿的守卫比没有守卫更危险。** 本项目所有守卫都做了结构断言，且失败必须进失败集。

### 四层验证

| 层 | 工具 | 断言什么 |
|---|---|---|
| 单元测试 | `dotnet test`（299 个 .cs） | 纯逻辑：引用转义、版本算术、参数变换 |
| **集成回归** | `tools/run_regression_v5.sh` + 分片清单 | 真机 Excel COM：**148 用例**，A=50 / B=15 / C=36 / D=14 / E=33 |
| **schema 审计** | `tools/audit_schema.py` | 31 工具 / 328 动作 / 534 参数全 snake_case、无缺描述、错误可操作 |
| **边界探针** | `tools/probe_edge_cases.py` | 15 个边界用例，**每个带期望值**，匹配不上就是缺陷 |
| 文档计数 | `tools/check_doc_counts.py` | 15 处文档里的「31 tools / 328 operations」口径一致 |

### 反假绿的三道硬约束

这几条都是**用真实的假绿事故换来的**：

**① 分片运行器必须校验汇总行**

`dotnet test` 对「filter 零匹配」返回 **0**（成功）。所以
`--filter` 写错 → 静默空跑 → 报告"0 失败"。运行器**必须**在日志里找到
`已通过!` / `失败!` 汇总行，缺失即判失败。

**② 预期用例数必须按 filter 的「包含语义」回算**

参数化用例（如 `CreateScenarioSummary_AddsReportWorksheet(reportType: Summary|PivotTable)`）
剥掉括号后是**一条**清单项，却会**执行两个**用例——实测 A 片 49 行 → 50 用例。
且候选全集必须是**全部**清单行：参数变体行可能不含你的关键词
（`Summary` 变体不含 `PivotTable`），只按关键词过滤会漏算。
`tools/gen_regression_shards.py` 已按此实现（产物 `.txt` 清单 + `.count` 预期数）。

**③ filter 只能用真实类名/方法名**

不能用 `.cs` 文件名后缀。`FullyQualifiedName~X.Slicers` 这类会**匹配 0 个用例**。
真实用例名以 `dotnet test --list-tests` 的权威清单（`_tlist_v5.log`，1088 项）为准。

> 另注：`dotnet vstest <dll> --list-tests` 在 VSTest 18.7.0 下被拒
> （「参数 --list-tests 无效」）→ 列清单只能走 `dotnet test <csproj> --list-tests`。

### 探针必须是「可证伪」的

`probe_edge_cases.py` 的每个 `Case` 都带 `expect_ok`（该成功还是该失败）与
`expect_regex`（错误文本里必须出现的短语）。教训：

> 早期版本里 `run()` 对 `server.call` 抛异常一律记 failure，
> 于是**所有"期望失败"的用例全变成假 FAIL**（9/13）。修复方式是把异常文本**照期望判定**，
> 而不是假设最坏。

同类的另一次事故：探针曾在 400 字符处截断回复，
而服务端的补救提示当时排在第 437 字符——**截断**读成了**缺失**，
差点把已修好的功能判成没修。两处一起改了（提示前移 + 上限放宽到 1200）。

---

## 9. 扩展指南

### 加一个动作者

1. **在 Core 接口加方法**。动作名默认由方法名推导（`SetLayout` → `set-layout`），
   不符预期时加 `[ServiceAction("custom-action")]`。
   参数类型决定 MCP 面形状（见第 3 节表）。
2. **不要手写 schema**。生成器会推导参数绑定与描述；手写的会被覆盖或冲突。
3. **更新计数**：
   - `tests/ExcelMcp.McpServer.Tests` 里 `McpToolSurfaceTests` 的操作数常量；
   - 15 处文档计数（跑 `python tools/check_doc_counts.py` 看哪些不一致）。
4. **跑守卫**：
   ```bash
   python tools/check_doc_counts.py          # 文档口径
   python tools/audit_schema.py              # 参数命名/描述/错误可操作性
   python tools/probe_edge_cases.py          # 边界行为
   ```
5. **真机验证**：把新动作写进对应分片清单，或临时用 `--filter` 跑真实用例名。

### 加一个工具（新域）

除上述之外：

1. 在 `src/ExcelMcp.McpServer/Tools/` 新建 `ExcelXxxTool.cs`，带 `[McpServerToolType]`；
   **工具类必须能被 MCP 框架发现**（参考现有 `ExcelWorksheetTool.cs` 的写法）。
2. 每个参数补 `[Description]`——`audit_schema.py` 会把缺描述的报成 finding。
3. 更新 `ServerInstructions`（`Program.cs`）里的工具概览，否则模型不知道它存在。

### 改提示词 / 技能文档

`skills/excel-mcp/SKILL.md`、`skills/excel-cli/SKILL.md` 由
`*.sbn` 模板在构建期渲染。**改模板，不要改产物**。
`skills/shared/*.md`（如 `slicer.md`）是手写内容，会被拼进生成的提示词。

### 改文档计数后

```bash
python tools/check_doc_counts.py   # canonical = manifest.totalOperations − diag(3) + file(5)
```

---

## 10. 诊断工具速查

| 工具 | 用途 | 典型调用 |
|---|---|---|
| `probe_mcp.py` | stdio JSON-RPC 探针，校验已部署 server 的工具 schema | `python probe_mcp.py` |
| `tools/audit_schema.py` | 全量 schema 审计（命名/描述/体积/错误可操作性） | `--json out.json`、`--full` |
| `tools/probe_edge_cases.py` | 15 个边界用例，带期望值 | `--list` 看用例 id |
| `tools/probe_chart_types.py` | 图表类型全量探针（84 枚举） | 真机跑一遍确认可用集 |
| `tools/probe_modern_stock_shot.py` | 现代图 / 股价图 / 截图专项 | 改图表路径后必跑 |
| `tools/probe_slicer_skin.py` | 切片器内置样式扫描 | `--style <名>` 单测 / `--all` 扫 12 种 |
| `tools/demo_slicer_link.py` | 切片器多表联动端到端演示 | `--show` 前台停留 180s 供手动点击 |
| `tools/gen_regression_shards.py` | 从权威清单生成分片 + 预期用例数 | 改测试面后重跑 |
| `tools/run_regression_v5.sh` | 分片回归运行器（汇总行断言 + 执行数断言 + 瞬态复跑） | 报告落 `_regression_v5_report.txt` |
| `tools/check_doc_counts.py` | 文档计数守卫 | 改文档后必跑 |
| `tools/pack.py` | 版本读写与产物收集 | `show` / `bump` / `collect` / `selftest` |
| `tools/publish.py` | 开源发布（枚举式暂存 + 索引守卫） | `check` / `push` |

### 定 COM 层问题的最快路径：pywin32 对照法

遇到"只有某个参数组合失败"的怪问题时，**用 pywin32 复刻失败，再逐参数消除变量**。
本项目定死股价图 `PlotBy` 那个坑就是这么做的（`tools/_diag_stock.py`）：
先复刻出 `0xB0D7018E`，再一个参数一个参数地加回去，直到定位到 `PlotBy`。

比读源码猜快得多，因为失败集合会随夹具列数漂移，源码里看不出因果。

---

## 相关文档

| 文档 | 内容 |
|---|---|
| [USER-GUIDE.md](USER-GUIDE.md) | 用户手册：安装、接入、31 工具用法、列表格式、配方、排查 |
| [README.md](README.md) | 项目概览、相对上游做了什么 |
| [DEPLOYMENT.md](DEPLOYMENT.md) | 部署手册：版本、补丁表、MCP 配置、验证证据、重建步骤 |
| [ADVERSARIAL-REVIEW.md](ADVERSARIAL-REVIEW.md) | 对抗性审查报告：三视角问题清单 + 真机证据 + 处置 |
| `mcp-server-excel/` | 上游源码 + 本地补丁（含上游自身文档） |

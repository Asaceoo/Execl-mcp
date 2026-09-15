# jyyj-mcp 助手

**通过 MCP（Model Context Protocol）协议驱动真实 Microsoft Excel 的本地增强版 ExcelMcp。**

上游项目：[`sbroenne/mcp-server-excel`](https://github.com/sbroenne/mcp-server-excel)（MIT，v2.0.8）。
本仓库是它的**下游构建（downstream build）**：在完整保留上游能力的基础上，修掉了本机实测中发现的 5 类缺陷，
并补上一套可复现的构建 / 发布 / 回归验证工具链。

> 与上游的关系：本仓库不是上游的官方分支，也不代表上游立场。
> 上游版权声明已按 MIT 要求完整保留，详见 [`LICENSE`](LICENSE)。

---

## 它是什么

ExcelMcp 与「解析 xlsx 文件的库」不是一类东西：它**通过 Excel 官方 COM 接口驱动真正的 Excel 进程**。
因此透视表、宏、图表、数据模型、条件格式、Power Query 这些"只有 Excel 自己才懂"的东西都能被创建和保留下来，
而不是在读写过程中被降级成静态值。

对外暴露 **31 个工具 / 328 个操作**，覆盖工作表、区域与格式、Excel 表格、透视表与切片器、Power Query (M)、
DAX 与数据模型、VBA、图表与看板、条件格式、数据验证、查询表、命名区域、XML 映射、PowerShell 与 Python in Excel 等。

---

## 本仓库相对上游做了什么

### 1. 切片器多表联动（上游原生做不到的能力）

Excel 的一个切片器只能驱动**共享同一 PivotCache** 的多个透视表，上游没有暴露这个能力。
本仓库补上了三个动作，全链路用原生 COM 实现：

| 工具 | 新增面 | 作用 |
|---|---|---|
| `pivottable` | 参数 `share_cache_from` | 让新建透视表复用已有透视表的 PivotCache |
| `slicer` | 动作 `connect-pivots` | 把一个切片器接到更多透视表上 |
| `slicer` | 动作 `disconnect-pivots` | 安全断开（至少保留一个连接，重复条目自动去重） |

### 2. 图表创建路径修复（可用图表 25 → 37 种）

上游的图表创建走单一的 `AddChart` + `SetSourceData` 路径，现代图（Treemap / 直方图 / 帕累托 / 箱形图 /
瀑布图 / 漏斗图 / 区域地图 / 柱线组合图）与股价图（HLC / OHLC / VHLC / VOHLC）在这条路径上必然失败。
新增 `ChartCreationPath.cs` 按图表类型分流：

- **现代图**：`Shapes.AddChart2(-1, code)` + `SeriesCollection.NewSeries()` 手动挂系列。
- **股价图**：先建柱形图 → `SetSourceData(range, PlotBy=2)` → 改 `ChartType`。
  `PlotBy` 必须显式传，否则 `AddChart(51)` 会先用整区 `CurrentRegion` 自动填充，切换类型时被 Excel 拒绝。
- 失败自动清理残留 Shape；创建后回读 `ChartType` 校验，**不接受静默回退**。

修复后真机可用 **37 种**图表（基础 25 + 现代 8 + 股价 4）。

### 3. `chart_config add-series` 传 Range 对象

原先把地址字符串直接赋给 `Series.Values`，恒定抛 `0x800A03EC`。现改为按图表所在工作表解析成 Range 对象再赋值。

### 4. 参数绑定错误透传

MCP SDK 在**参数绑定阶段**抛出异常时不会进入工具内部的 catch，导致 6 种不同的调用错误被吞成同一句话。
新增 `ToolErrorSurfaceFilter`（注册在 `SessionIdentityFilter` 之后）把它们转成可读文本，
调用方现在能看到"哪个参数、合法值有哪些"。

### 5. 现代图读取容错

现代图的 `series.Values` / `XValues` / `Formula` 与 `chart.PivotLayout` 在**读取**时抛 `E_NOTIMPL`
（迟到绑定器表现为 `NotImplementedException`），会让 `chart list` / `chart read` 整体崩溃。
现已在 `CanHandle` 同时捕获 `COMException` 与 `NotImplementedException`，读取路径降级为 `SafeText`——
现代图可以正常列出与读取，写入能力不受影响。

### 附带的可复现工具链

| 脚本 | 作用 |
|---|---|
| `build.sh` | 补全 Windows 核心环境变量后调用固定的 .NET SDK 构建（沙箱裸调 dotnet 会因 `path1` 为 null 失败） |
| `release.sh` | 一键发布：版本递增 → 构建 → 部署 → 打包 → `--version` 自校验 |
| `tools/pack.py` | 版本读写（`show` / `bump`）与产物收集（`collect`），产物一律带版本号后缀 |
| `tools/check_doc_counts.py` | 文档计数守卫（`scripts/check-doc-counts.ps1` 的跨平台复刻） |
| `tools/gen_regression_shards.py` | 从权威测试清单生成分片，并**按 filter 包含语义**回算预期用例数 |
| `tools/run_regression_v5.sh` | 分片回归运行器：汇总行断言 + 执行数断言 + 瞬态 COM 自动复跑 |
| `tools/probe_*.py` | 真机探针：图表类型 / 现代图+股价图+截图 / 切片器皮肤 |

---

## 环境要求

- **Windows x64**（COM 自动化只存在于 Windows）
- **Microsoft Excel 2016 或更高版本（桌面版）** —— 必须真装 Excel，WPS 不适用
- 目标工作簿在操作时必须**在 Excel 里关闭**（COM 需要独占访问）
- 仅从源码构建时需要 **.NET 10 SDK**

---

## 快速开始（用预编译包）

1. 从 Releases 下载 `jyyj-mcp-<version>-win-x64.zip` 并解压到任意目录（例如 `D:\jyyj-mcp`）。
2. 在 MCP 客户端里注册：

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

3. 重启客户端，在连接器/工具列表里确认看到 **31 个工具**。

> **无头运行是设计如此**：`file` 工具创建/打开会话时 `show` 参数默认 `false`，Excel 在后台运行、窗口不可见。
> 需要在屏幕上看着它操作时，显式传 `show: true`（或对已开会话用 `window` 工具的 `show` / `bring-to-front`）。
> 注意 `screenshot` 工具的 `capture-sheet` **必须**在可见会话下使用。

### 真实例子：一句 `slicer connect-pivots` 让一个切片器驱动多张透视表

```python
# 1. 第二个透视表复用第一个的缓存 —— 这是"一个切片器驱动多表"的前提
pivottable(action="create", session_id=sid, sheet_name="Pivot2",
           source_range_address="Data!A1:D100",
           share_cache_from="Pivot1")

# 2. 把已有切片器接到第二个透视表上
slicer(action="connect-pivots", session_id=sid, slicer_name="区域",
       pivot_table_names=["Pivot1", "Pivot2"])
```

---

## 从源码构建

```bash
git clone https://github.com/Asaceoo/Execl-mcp.git
cd Execl-mcp
```

默认布局（任何一项都可以用环境变量覆盖）：

```
<workspace>/build.sh              # 构建入口
<workspace>/mcp-server-excel/     # 待构建的源码树
<workspace>/.dotnet10/dotnet.exe  # 固定的 .NET SDK
<workspace>/.nuget-packages/      # 重定向的 NuGet 缓存（不写用户目录）
```

```bash
# 环境变量覆盖（默认取脚本所在目录）
export EXCEL_MCP_ROOT=/d/my/jyyj-mcp        # 工作区根
export EXCEL_MCP_REPO=$EXCEL_MCP_ROOT/mcp-server-excel
export EXCEL_MCP_DOTNET=$EXCEL_MCP_ROOT/.dotnet10/dotnet.exe
export EXCEL_MCP_NUGET=$EXCEL_MCP_ROOT/.nuget-packages

# 构建（项目路径必须传相对路径，且相对的是源码树而非工作区）
bash build.sh build -c Release src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj
bash build.sh build -c Release src/ExcelMcp.CLI/ExcelMcp.CLI.csproj

# 一键发布：版本递增 → 构建 → 部署到 excel-mcp-bin/ → 产物落 dist/ → 自校验
bash release.sh              # 自动递增版本后缀
bash release.sh --no-bump    # 保持当前版本
bash release.sh --no-build   # 跳过编译，只重新部署与打包
```

`Directory.Build.props` 里的 `<Version>` 是版本号的**唯一来源**，所有产物都带该版本号后缀。

---

## 仓库结构

```
.
├── README.md                  # 本文件
├── LICENSE                    # MIT（上游原文 + 本仓库修改声明）
├── DEPLOYMENT.md              # 部署手册：补丁表 / MCP 配置 / 验证证据 / 重建步骤
├── ADVERSARIAL-REVIEW.md      # 对抗性审查报告：问题清单 + 真机证据 + 处置
├── build.sh / release.sh      # 构建与一键发布
├── probe_mcp.py               # stdio JSON-RPC 探针：校验已部署 server 的工具 schema
├── tools/                     # 探针、分片回归、打包脚本
├── dist/                      # （构建产物，不入库）
├── excel-mcp-bin/             # （部署目录，不入库）
└── mcp-server-excel/          # 上游 v2.0.8 源码 + 本地补丁（含其自身 .gitignore）
```

**关于 `mcp-server-excel/.github/workflows/`**：上游的 11 个 CI 工作流依赖其自托管 Windows runner、
sealed release secrets 与 gh-pages / 插件市场发布目标，本仓库不具备这些条件。原样纳入会导致**每次推送都失败**
并向仓库所有者发送失败通知，因此**有意排除**（已在 `.gitignore` 中显式注明），而不是悄悄删掉。
工作流原文见上游仓库。

---

## 已知限制

| 限制 | 说明 |
|---|---|
| 旭日图（Sunburst） | Excel 自身会回退成其它类型，工具**明确报错**（`Excel fell back to a different chart type`）而不是假报成功 |
| `screenshot capture-sheet` | 必须 `show: true` 的可见会话；无头会话下捕获必然失败 |
| 连续自动化 | 约 1 小时后 Excel COM 可能出现瞬态不可用（`0x800706BA`），停下重试即可，不是本项目的缺陷 |
| 一个工作表一个透视表 | 集成测试里每个透视表必须独占一张工作表，这是 Excel 的限制而非工具限制 |
| VBA | 需要在 Excel 里手工开启"信任对 VBA 工程对象模型的访问" |

---

## 许可与致谢

本项目基于 **[ExcelMcp](https://github.com/sbroenne/mcp-server-excel)**（作者 Stefan Broenner）构建，
上游以 **MIT License** 发布。上游版权声明已**完整保留**在原文件中，本仓库的修改另行声明，
详见 [`LICENSE`](LICENSE)。

感谢上游作者把 Excel COM 自动化的复杂性收敛成一套稳定的工具面 —— 没有它就没有这个下游构建。

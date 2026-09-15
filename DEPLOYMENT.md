# jyyj-mcp 助手 本地部署说明

**版本**：jyyj-mcp 助手 **v3.0.0.0**（四段版本号）
**基线**：上游 `sbroenne/mcp-server-excel` 2.0.8 + 本地补丁（切片器多透视表联动 + 图表创建路径修复 + 错误透传）
**状态**：已构建、已通过三视角对抗性审查、已真机验证、已部署
**部署日期**：2026-09-15
**审查报告**：见 `ADVERSARIAL-REVIEW.md`（F1–F7 共 37 项问题、逐项取证与处置）

> **版本号来源**：`mcp-server-excel/Directory.Build.props` 的 `<Version>` —— **单一来源**。
> 打包时由 `release.sh` 自动递增（支持四段递增，如 `3.0.0.0` → `3.0.0.1`），
> 产物一律带版本号后缀，不产出无版本号文件。
> `--version` 输出 `jyyj-mcp 助手 v3.0.0.0`；MCP `initialize` 的 `serverInfo.version`
> 同样报 `3.0.0.0`（`Directory.Build.props` 里 `<Version>` / `<AssemblyVersion>` / `<FileVersion>`
> 三者同步，避免"横幅说 A、协议说 B"）。
> 版本算术由 `python tools/pack.py selftest` 断言（含 `3.0.0.0` → `3.0.0.1` 等 5 个用例）。
> **编码提示**：`--version` 横幅含中文，经 Windows 控制台代码页输出为 **GBK/cp936** 而非 UTF-8；
> 消费该输出的脚本必须按 cp936 回退解码（`tools/pack.py` 的 `run_version_check` 已如此处理），
> 否则会抛 `UnicodeDecodeError` 并让自校验静默拿到空串。

---

## 一、补丁内容（相对上游 2.0.8）

上游的切片器**无法驱动多个透视表** —— 切片器只能联动共享同一个 `PivotCache` 的透视表，而产品原本每次创建都新建独立缓存，且没有任何 API 能把多个透视表接到一个切片器上。本补丁补齐该能力。

| # | 变更 | 文件 |
|---|---|---|
| A | 新增 `share_cache_from` 参数：`create-from-range` / `create-from-table` / `create-from-datamodel` 可复用既有透视表的 `PivotCache` | `Commands/PivotTable/IPivotTableCommands.cs`、`PivotTableCommands.Create.cs`、`Models/PivotTableTypes.cs` |
| B | 新增切片器动作 `connect-pivots` / `disconnect-pivots` | `Commands/Slicer/ISlicerCommands.cs`、`SlicerCommands.cs`、`Commands/PivotTable/PivotTableCommands.Slicers.cs` |
| C | 修正 `create-slicer` 的误导性 `WorkflowHint`（原先指向一个不存在的「连接更多透视表」能力） | `PivotTableCommands.Slicers.cs` |
| D | **修复上游源生成器缺陷**：`ServiceRegistryGenerator.cs` 转义描述时只替换 `\n` 未清 `\r`；本仓库为 CRLF，任何多行 `<param>` 文档都会截断生成的字符串字面量并注入裸 C# 标识符（48 个 CS0246）。现按同文件已有写法补 `.Replace("\r", "")` | `src/ExcelMcp.Generators/ServiceRegistryGenerator.cs` |
| E | 同步文档计数（新增 2 个 action → 326→**328**）：15 处文档标题 + `docs/features` 章节求和 + `McpToolSurfaceTests.ExpectedOperationCount` | 见 `git status` |

**审查轮补丁（slicerlink.1 → .2）**，均为真机取证后修正：

| ID | 变更 |
|---|---|
| 输入契约 | 新增 `TryCleanPivotTableNames`：请求按 `OrdinalIgnoreCase` 去重 + 去空白后再计数。修复「重复名污染『至少保留一个连接』算术」与「对同一表调用两次 `RemovePivotTable` 抛 `0x8002000B`」 |
| 部分失败语义 | `connect-pivots` 失败信息改为披露**被拒清单 + 已生效清单 + 当前实时连接列表**（原先只列被拒项，调用方会以为整批未生效） |
| 空白请求 | 无可用名字的请求直接失败，不再走「无待新增项」分支返回 `Success=true` |
| 源身份判定 | `SourceDataConflict` 删除前缀/后缀容差，改为归一化后严格相等；`NormalizeSourceRef` **保留空白**（仅剔 `$` 与引号）。修复「表名互为后缀 → 跨表共享被静默接受」 |
| 提示一致性 | 抽出 `BuildSharedCacheHint`，三条创建路径统一回填 `WorkflowHint` |
| 错误路径 | `disconnect-pivots` 的 `RemovePivotTable` 包 `COMException`，与 `connect-pivots` 策略一致，不再把裸错误码抛给调用方 |

**图表修复轮补丁（slicerlink.4 → .5）**，真机取证后修正（详见 `ADVERSARIAL-REVIEW.md` F6-1~F6-5、F5-9）：

| ID | 变更 |
|---|---|
| F6-1 | 新增 `Commands/Chart/ChartCreationPath.cs` 统一图表创建路径：现代图（树图/瀑布/漏斗/直方图/帕累托/箱线/地图/组合 8 种）走 `Shapes.AddChart2(-1, code)` + `SeriesCollection.NewSeries()` 手动挂系列；股价图 4 种走「先建柱形图 → `SetSourceData(range, PlotByColumns)` → 改 `ChartType`」。**可用图表从 25 种提升到 37 种**。失败自动清理残留 Shape |
| F6-2 | `RegularChartStrategy.AddSeries` 改为解析 Range **对象**赋给 `Series.Values`/`XValues`（原先传地址字符串，恒抛 `0x800A03EC`） |
| F6-5 | 现代图 `read`/`list` 崩溃：`CanHandle` 补捕 `NotImplementedException`；`GetDetailedInfo` 的 `PivotLayout`、series 读取（E_NOTIMPL）降级为 `SafeText` |
| F5-9 | 新增 `McpServer/ToolErrorSurfaceFilter.cs` 请求过滤器，把被 SDK 参数绑定阶段吞掉的 `range_edit`/`range_format` 真实错误透传给调用方 |
| 测试 | 新增 `tests/.../ChartCommandsTests.CreationPath.cs` 12 用例（含股价图按 HLC=3/OHLC·VHLC=4/VOHLC=5 精确匹配列数，防假绿）；夹具数据扩到 5 列数值 |

**兼容性**：本轮全部为缺陷修复与内部路径重构，未新增/删除任何 tool 或 action，操作数仍为 328。

---

### 看板配色兜底路（VBA 换切片器皮肤，实测可用）

MCP 原生 `slicer` 工具没有 `set-style` 动作；需要按主题定制看板配色时，走 **`vba` 工具兜底**：

1. 工作簿须为 `.xlsm`，且信任中心已勾选「信任对 VBA 工程对象模型的访问」（见「六、使用前提」）；
2. `vba import` 导入模块（VBA 内**按名引用** `Worksheets("Sheet1")`，不要用 `Worksheets(1)`——`worksheet create` 会把新表插到活动表前面导致索引位移）；
3. `vba run "模块.过程"` 执行 `sl.Style = "SlicerStyleLightN"` / `"SlicerStyleDarkN"`；
4. 保存用 `file close` + `save: true`（`file` 工具没有 save action）。

**真机实测结论**（`tools/probe_slicer_skin.py --all`，一次会话扫 12 个内置样式）：`SlicerStyleLight1-6` 与 `SlicerStyleDark1-6` **12/12 全部应用成功**（含样式回读校验）。单一样式验证：`python tools/probe_slicer_skin.py --style SlicerStyleDark3`。探针产物在 `_demo/slicer-skin-*.xlsm`。

---

## 二、目录清单

| 路径 | 用途 |
|---|---|
| `mcp-server-excel\` | 源码（浅克隆 + 本地补丁） |
| `.dotnet10\` | 隔离 .NET SDK **10.0.302 + 10.0.401**；运行时 NETCore.App / WindowsDesktop.App 10.0.10、10.0.12 |
| `.nuget-packages\` | NuGet 包缓存（重定向，避免写系统 `~/.nuget` 被拒） |
| `excel-mcp-bin\` | **部署目录**（MCP 实际加载此目录；含 `VERSION.txt` 自描述） |
| `dist\` | 发布产物：便携 zip（唯一可分发单元）+ manifest |
| `build.sh` | 构建包装脚本：补齐沙箱缺失的 Windows 环境变量后转调 `.dotnet10/dotnet.exe` |
| `release.sh` | **一键发布**：版本递增 → 构建 → 部署 → 打包 → 自校验 |
| `tools\pack.py` | 版本读写（`show` / `bump`）与产物收集（`collect`） |
| `tools\check_doc_counts.py` | 文档计数守卫（复刻 `scripts/check-doc-counts.ps1`，本沙箱无法跑 PowerShell） |
| `probe_mcp.py` | MCP stdio 协议探针：校验 `tools/list` 是否暴露新动作与参数 |
| `ADVERSARIAL-REVIEW.md` | 三视角对抗性审查报告（问题清单 + 真机证据 + 处置） |
| `_research\` | 对标调研资料存档 |

> `.dotnet10` 与 `.nuget-packages` 仅在**重新构建**时需要；纯运行只需 `excel-mcp-bin` + 运行时。

---

## 三、一键发布

```bash
bash /d/execl-mcp/release.sh              # 递增版本 → 构建 → 部署 → 打包 → 自校验
bash release.sh --no-bump    # 用当前版本重新打包
bash release.sh --test       # 额外跑核心测试套件
```

链路设计要点：

1. `pack.py bump` 递增 `<Version>` 的尾部构建号（`…slicerlink.2` → `…slicerlink.3`）并把新版本打到 stdout；
2. 构建与打包共用同一个版本值；`collect` **断言** `Directory.Build.props` 的版本等于传入版本，不等直接失败（挡住「构建完又 bump」的错位）；
3. 部署前先 `taskkill` 掉占用 `excel-mcp-bin` 的进程（否则覆盖时报 `Device or resource busy`）；
4. 覆盖式复制（**不递归删除** —— 沙箱对递归删除 fail-closed），列出「只在目标存在」的陈旧文件；
5. 压缩用 Python `zipfile` 进程内完成（PowerShell `Compress-Archive` 在本沙箱会被 safe-delete 钩子打断进程树，产出 0 字节 zip）；
6. **部署后实跑 `--version` 与目标版本比对，不一致即非零退出** —— 这一条同时充当陈旧文件守卫。

---

## 四、MCP 配置

已写入 MCP 客户端的配置（本机路径为 `%USERPROFILE%\.workbuddy\mcp.json`；现有 `mempalace` / `git` / `playwright` / `sequentialthinking` / `wigolo` 均未改动）：

```json
"jyyj-mcp": {
  "command": "<workspace>\\excel-mcp-bin\\Sbroenne.ExcelMcp.McpServer.exe",
  "args": [],
  "env": {
    "DOTNET_ROOT": "<workspace>\\.dotnet10",
    "DOTNET_CLI_TELEMETRY_OPTOUT": "1"
  },
  "timeout": 600,
  "disabled": false
}
```

> **连接器键名与部署目录名**：键名已随项目改名从 `excel-mcp` 改为 `jyyj-mcp`（客户端里显示的名字取自
> server 的 `serverInfo.name`，同样为 `jyyj-mcp`）。**部署目录 `excel-mcp-bin` 刻意未改名** ——
> 因为 `excel-mcp-bin` 里的可执行文件是 `Sbroenne.ExcelMcp.McpServer.exe`，程序集名与 C# 命名空间
> 按设计保持不变（改动它们会波及 328 个操作名与 1088 个测试），只改父目录名得不到命名上的一致性，
> 却要动 `build.sh` / `release.sh` / `pack.py` / 探针 / 本文档五处路径。等哪天真要改程序集名时一并处理。

**激活方式**：该 MCP **不会自动生效**。需在连接器管理页右上角打开「自定义连接器」入口，在 `jyyj-mcp` 上点「信任」，然后重启客户端。

**更新已部署二进制时**：运行中的 `Sbroenne.ExcelMcp.McpServer.exe` 会锁住 `excel-mcp-bin` 下的 DLL，导致复制报 `Device or resource busy`。`release.sh` 已自动处理；手动操作时需先结束进程（Git Bash 下必须加 `MSYS_NO_PATHCONV=1` 前缀），再覆盖部署，最后重启客户端。

---

## 五、验证结果（实测证据）

### 1. 进程级 / 产物级

```
> excel-mcp-bin\Sbroenne.ExcelMcp.McpServer.exe --version
jyyj-mcp 助手 v3.0.0.0
```

MCP `initialize` 回读（`tools/audit_schema.py`）：`serverInfo: name=jyyj-mcp version=3.0.0.0`，
工具面 31 个工具 / 328 个动作 / 534 个参数。

`dist/` 产物（由 `release.sh` 自动生成，zip 与 manifest 带版本号后缀）：
`jyyj-mcp-3.0.0.0-win-x64.zip`（7,382,504 B）、`jyyj-mcp-3.0.0.0-manifest.json`
（历史产物 `excel-mcp-2.0.8-slicerlink.*` 与 `jyyj-mcp-2.0.8-jyyj.1` 保留在同目录，便于回滚比对）。
`excel-mcp-bin/VERSION.txt` 与 `<Version>` 同源于 `Directory.Build.props`。

**zip 是唯一可分发单元**，内含 `excel-mcp-bin/` 全量（server + CLI 两个入口及其 apphost 四件套）。
**不再单独产出裸 `.exe`**：.NET apphost 只是加载器，脱离同目录的 `*.dll` / `*.deps.json` / `*.runtimeconfig.json`
必然报 `The application to execute does not exist`（退出码 2147516570）——详见审查报告 **F5-7**。
`pack.py` 用 `assert_bundle_runnable()` 守卫该不变式，并对 CLI 实跑 `--version` 冒烟。

### 2. MCP 协议级（`probe_mcp.py` / `tools/audit_schema.py`，stdio JSON-RPC）

| 方法 | 结果 |
|---|---|
| `initialize` | `name=jyyj-mcp`，`version=3.0.0.0` |
| `tools/list` | **31 个工具**（数量未变，靠 action 扩展） |
| 参数命名 | **534 / 534 全 snake_case**，无 camelCase 混用 |
| 参数描述 | 无缺失（`screenshot` 的 3 个参数已补齐，见 F7-7）；`audit_schema.py` 报 **FINDINGS: none** |
| 错误可操作性 | 4 / 4 探针 OK（缺 action、非法 enum、错参数名、未知工具均给出可执行信息） |
| `slicer` 工具 | action 枚举 **10** 项，含 `connect-pivots` / `disconnect-pivots`；参数含 `slicer_name`、`pivot_table_names` |
| `pivottable` 工具 | 参数含 `share_cache_from`；action 枚举 10 项 |

### 3. Excel COM 真机集成测试

`PivotTableCommandsTests.SlicerLink` 共 **11 项**，覆盖：

| 测试 | 验证内容 |
|---|---|
| `SharedCache_LetsOneSlicerFilterEveryPivotTable` | 3 个共享缓存透视表被一个切片器同时过滤（合计 650 → 325） |
| `IndependentCaches_ConnectingIsRejectedWithoutSideEffects` | 独立缓存被拒绝 + 可操作报错 + 无副作用 |
| `DisconnectPivots_RemovesConnectionButKeepsAtLeastOne` | 断开生效、禁止断到 0 连接 |
| `ShareCacheFromDifferentSource_IsRejected` | 跨源共享被拒（校验按 R1C1 规范化） |
| `RawAddPivotTableAcrossCaches_IsDeterministic` | 原始 COM 跨缓存行为确定性 |
| `DisconnectDuplicateNames_KeepTheLastConnectionGuardAccurate` | **审查新增**：重复名不污染「至少保留一个」算术 |
| `ConnectBlankOnlyRequest_IsRejectedInsteadOfReportedAsDone` | **审查新增**：无可用名字的请求不报成功 |
| `PartialConnectFailure_DisclosesTheStateItLeftBehind` | **审查新增**：部分失败披露已生效连接 |
| `ShareCacheFromSuffixSheetName_IsRejectedAsADifferentSource` | **审查新增**：表名互为后缀视为不同源 |
| `ShareCacheFromSpaceInSheetName_MatchesOnlyTheSameSheet` | **审查新增**：引号/锚点归一化、空白不归一化 |
| `WholeColumnSource_RecordsTheNotationTheGuardMustCompare` | **审查新增**：整列源两侧记法一致（`C1:C4`） |

**回归结果（实测）**：本沙箱实测 **43–70 秒/用例**（Excel COM 启动开销占主导）。已执行并全绿的部分：

| 分片 | 用例数 | 结果 | 日志 |
|---|---|---|---|
| 切片器面（`SlicerLink_`+`ListSlicers_`+`CreateSlicer_`+`SetSlicerSelection_`+`DeleteSlicer_`） | 25 | **25 通过 / 0 失败** | `_r_a_slicer.log` |
| `SlicerLink`（补丁面） | 11 | **11 通过 / 0 失败** | `_r_slicer.log` |
| `GroupByDate`（区域设置修复后） | 4 | **4 通过 / 0 失败** | `_r_gbdate2.log` |
| 字段面（`Add*Field_`/`SetField*`/`SortField_`/`RemoveField_`） | 27 | **27 通过 / 0 失败** | `_r_d_field.log` |
| 计算字段 / 计算成员 / 总计 / 布局 / 分组 | 36 | **36 通过 / 0 失败** | `_r_e_calc.log` |
| 缓存选项 / 数据准备 / 钻取 / 刷新 / 列表 / 删除 等杂项 | 18 | **18 通过 / 0 失败**（欠跑 16 项经核对 0 遗漏） | `_r_f_rest.log` |

> 至此 Slicer / PivotTable 面 **146 项全部覆盖，0 失败**。

> 末批 `f_misc` 原用 `List_`/`GetInfo_`/`Delete_` 等**过泛前缀**，实际匹配 126 个、35 分钟只跑完 30 个即被墙钟截断，
> 输出虽是「0 失败」但覆盖率仅 24%，不能作为通过依据。故改用差集 + 完整方法名精确重跑（见审查报告 **F5-8**）。
| 创建路径 `CreateFrom*`（含 `share_cache_from` 补丁面） | 34 | **34 通过 / 0 失败**（首跑受 Excel COM 瞬态不可用影响，复跑确认） | `_r_c_create2.log` |
| 表格切片器面 | 16 | **16 通过 / 0 失败**（首跑 15 通过 + 1 项 `800706BA` 环境抖动，单独复跑该项通过后归零） | `_r_b_tblslc.log`、`_r_tblslc_one.log` |
| PivotTable 面（首批） | 20 | **20 通过 / 0 失败** | `_r_pivot.log` |

> **未完成部分不作通过声明**：142 项全集串行约需 1.7–2.8 小时，首批分片被 15 分钟墙钟上限截断（`Slicer` 面 43 项完成 11 项、`PivotTable` 面完成 20 项）。补充分片（`_r_s_tbl` / `_r_s_pvt` / `_r_p1`…`_r_p5`）以 40 分钟/片运行，见各日志末尾的 `已通过!` / `失败!` 汇总行。
>
> ⚠️ 上表为 **slicerlink.1–.3 时期**的分批取证，部分旧日志已无汇总行（当时的零匹配/截断现场）。
> **当前版本以本文档下方的「v.5 全量回归」为准**：147 项一次性按权威清单分片重跑，逐片断言执行数，0 确定性失败。

**v.5 全量回归（2026-09-15，`tools/run_regression_v5.sh`，147 项 / 148 用例）**

清单来自当前构建产物的 `dotnet test --list-tests`（权威来源 `_tlist_v5.log`），分片过滤器由
`tools/gen_regression_shards.py` 生成（`FullyQualifiedName~<完整类名>.<方法名>`，参数化用例剥离括号参数——
VSTest 过滤器语法里括号是运算符）。每片都断言 **汇总行存在** 且 **执行数 == 预期用例数**：

| 分片 | 覆盖 | 预期 | 实测 | 结果 | 日志 |
|---|---|---|---|---|---|
| A_lifecycle | 创建/删除/列表/取数/缓存/数据准备/钻取/刷新 | 50 | 50 | **50 通过 / 0 失败** | `_rr_v5_A_lifecycle_a1.log` |
| B_fields | `Add*Field`/`SortField`/`SetField*`/`RemoveField` | 15 | 15 | **15 通过 / 0 失败** | `_rr_v5_B_fields_a1.log` |
| C_layout_grouping | 布局/小计/总计/分组/表格切片器 | 36 | 36 | **36 通过 / 0 失败** | `_rr_v5_C_layout_grouping_a1.log` |
| D_olap | OLAP 消歧/计算成员/数据模型/多维字段 | 14 | 14 | **14 通过 / 0 失败** | `_rr_v5_D_olap_a1.log` |
| E_slicers_charts | 切片器 + `SlicerLink_` + PivotChart 创建 | 33 | 33 | 首跑 31/2（瞬态）→ **复跑 33 通过 / 0 失败** | `_rr_v5_E_slicers_charts_a2.log` |

> **E 片首跑 2 项失败经判定为环境瞬态**：同片在清残留 `EXCEL.EXE` + 歇 60 秒后**整片复跑 33/33 全绿**，
> 与既往「连续自动化约 1 小时后 COM 瞬态不可用」的已知模式一致。**确定性回归为 0**。
> 唯一一次计数告警（A 片 50 ≠ 49）已定位为参数化用例口径：`AnalysisCommandsTests.CreateScenarioSummary_AddsReportWorksheet`
> 有 `Summary` / `PivotTable` 两个变体，两个都通过；生成器已改为按 **filter 包含语义**回算预期，
> 现在 A 片预期即为 50，不再误报。

**回归工具的两个反假绿设计**（此前踩过坑，现固化为守卫）：
1. 期待值取自 `.count`（由清单与 filter 语义共同推导），**执行数不等即判 MISMATCH**；
2. 日志缺失「已通过!/失败!」汇总行一律判**失败**（零匹配时 `dotnet test` 退出码是 0，只看退出码会静默假绿）。

**工具面契约 5/5 通过**：`McpToolSurfaceTests`（31 工具 / 328 操作）。

**文档计数校验通过**：`tools/check_doc_counts.py` → canonical = 31 工具 / 328 操作，15 条文档标题 + 工具面交叉校验 + `--help` 派生校验全部一致。

### 4. 图表能力真机复验（v.5，`tools/probe_modern_stock_shot.py`）

针对「现代图 / 股价图 / 截图」三项做聚焦复验（全量探针 64 项太慢，此项约 10 分钟）：

| 组 | 用例 | 结果 |
|---|---|---|
| MODERN | 树状图 / 直方图 / 帕累托 / 箱形图 / 瀑布图 / 漏斗图 / 区域地图 / 柱线组合图 | **8/8 通过** |
| MODERN（限制） | 旭日图 | **Excel 侧限制**：Excel 自身回退到别的类型，工具**明确报错**（`Excel fell back to a different chart type ... requested Sunburst`），不再假报成功 |
| STOCK | HLC(3 列) / OHLC(4 列) / VHLC(4 列) / VOHLC(5 列) | **4/4 通过** |
| READBACK | 12 张图逐张回读 `chartType` + `chart read` 现代图/股价图 | **14/14 通过**——类型与请求完全一致（无静默回退）；现代图 `read` 不再崩（F6-5 确认修复） |
| SCREENSHOT | 可见会话（`show: true`）下 `capture-sheet` | **通过**：1352×196 px `image/png`，23,722 B |

**截图的调用要点**（此前误判为「不可用」）：

- 参数只有 `sheet_name` + `quality`；`--output` 是 **CLI 专用**旗标，MCP 面没有该参数；
- 服务端返回 **两个内容块**：`type: image`（base64 PNG）+ `type: text`（摘要，如 `Screenshot: $A$1:$M$7 on 'Charts' (1352x196px)`）。
  **只拼接 text 的客户端看不到图**——`tools/demo_slicer_link.py` 的 `Server.call` 已改为把 image 块保留在 `_images` 字段；
- 会话必须可见：无头会话下捕获必失败，这是设计如此（见 §六 使用前提）。

### 5. 真机上得到的关键事实（已写入代码注释，避免后人重踩）

- **`PivotCache.Index` 不可用于判断缓存身份**：实测两个独立缓存都返回 `0`，同一共享缓存返回 `0/1`。故 `connect-pivots` 改为「应用后回读切片器实际连接列表」的行为校验，不依赖任何身份探测。
- **Excel 用 R1C1 记法存储 range 型 `PivotCache.SourceData`**：`'SalesData'!A1:D6` 存回为 `SalesData!R1C1:R6C4`。按 A1 字符串比较会对**合法共享产生误拒**，故源冲突校验改为经 Excel 转成 R1C1 后比较。
- **整列源的记法两侧一致**：`Range("A:D").Address(R1C1)` 与缓存 `SourceData` 都是 `C1:C4` —— 曾担心整列源会误拒，实测推翻。
- **Excel 对跨缓存连接抛 `COMException 0x800A03EC`**（且连接不生效）。`connect-pivots` 捕获该异常并回读校验，把裸错误码转成可操作信息。
- **对已移除条目重复调用 `RemovePivotTable` 抛 `COMException 0x8002000B`（无效索引）**，不静默忽略。
- **同一工作表不能放置第二个透视表**（Excel 报「已存在数据透视表 X」）。集成测试因此把每个透视表放到独立工作表。
- **日期分组字段名随 Office 界面语言本地化**：英文版生成 `Days` / `Months` / `Years` / `Quarters`，中文版生成 `天(Date)` / `月(Date)` / `年(Date)` / `季度(Date)`。断言英文子串会让测试在中文 Excel 上**必然假红**，从而掩盖真实回归 —— 已抽出 `HasGroupedDateField` 同时接受两种命名。

---

## 六、使用前提（硬约束）

1. **Windows + Excel 2016+ + 交互式桌面**（不可用于无头/服务端批处理）
2. **独占访问**：操作前需关闭所有 Excel 窗口（Excel COM 限制）
3. **VBA 功能**需在信任中心勾选「信任对 VBA 工程对象模型的访问」，且工作簿须为 `.xlsm`
4. `range` 的 `set-values` / `get-values` **必须传 `sheet_name`**，否则报 `InvalidInput`
5. **切片器联动三步走**：先建第一个透视表 → 其余用 `pivottable` `share_cache_from` 指向它 → `slicer` `create-slicer` → `slicer` `connect-pivots` 挂其余。

---

## 七、重新构建（手动路径）

```bash
cd <workspace>                                  # 例如 /d/execl-mcp
bash build.sh build -c Release src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj
cp -r mcp-server-excel/src/ExcelMcp.McpServer/bin/Release/net10.0-windows/. excel-mcp-bin/
```

> `build.sh` 的项目路径是**相对源码树**的（它自己会 `cd` 进去），而 `cp` 的路径是**相对工作区**的。
> 换机器时用 `EXCEL_MCP_ROOT` / `EXCEL_MCP_REPO` / `EXCEL_MCP_DOTNET` / `EXCEL_MCP_NUGET`
> 覆盖默认布局即可，不必改脚本。

（推荐直接用 `release.sh`：它会顺带 bump 版本、杀占用进程、打版本化产物并做 `--version` 自校验。）

> **注意**：不要用 `dotnet publish`。本环境 NuGet 源缺少 SDK 10.0.401 要求的 Runtime Pack `10.0.12`，会报 `NU1102`。`build` 的输出目录已是完整可运行的应用。

---

## 八、构建期已知问题与绕过

| 问题 | 根因 | 处置 |
|---|---|---|
| `CS9057` 分析器版本更新 | 项目固定 `Microsoft.CodeAnalysis.CSharp 5.9.0`，SDK 10.0.302 内置编译器仅 5.6.0 | 升级到 SDK **10.0.401** |
| `NU1102` 找不到 Runtime Pack | `publish` 需要的 `Microsoft.*.App.Runtime.win-x64 = 10.0.12` 未同步到源 | 放弃 `publish`，用 `build` 输出目录作部署目录 |
| restore 报 `Access to the path is denied` | 默认包缓存 `%USERPROFILE%\.nuget\packages` 在工作区外 | 设 `NUGET_PACKAGES` 重定向；偶发时报错则重试收敛 |
| SDK 版本不满足 `global.json` | 本机原有 `dotnet10` 目录为 10.0.300（< 10.0.302） | 下载并存 10.0.401 到 `.dotnet10` |
| 多行 `<param>` 文档导致生成代码 48 个 `CS0246` | 源生成器只替换 `\n` 未清 `\r`（CRLF 仓库截断字符串字面量） | 已修 `ServiceRegistryGenerator.cs`（见补丁 D） |
| `cp` 报 `Device or resource busy` | 运行中的 MCP 进程锁住部署目录 | `release.sh` 自动 `taskkill`；手动时加 `MSYS_NO_PATHCONV=1` |
| `MSBUILD : error MSB1001: 未知开关。开关:/d/execl-mcp/...csproj` | `release.sh` 向 `build.sh` 传 POSIX 绝对路径，经 `env` 转发未被 MSYS 转换 | 改传**相对路径**（`build.sh` 自身会 `cd` 到仓库根）—— 见 F5-1 |
| `release.sh: line N: unexpected EOF while looking for matching '"'` | **脚本运行途中被编辑**：bash 按字节偏移增量读取脚本，中途改动导致解析错位 | 不要在长跑脚本运行期间编辑该脚本；改完先 `bash -n` 再跑 |
| `GroupByDate_*` 4 项稳定失败（`Expected to find Days field`） | 中文版 Excel 生成的字段名为 `天(Date)` 等，测试断言英文子串 | 已改为双语命名容错 —— 见 F5-4 |
| 分批回归中整批报 `COMException 0x800706BA RPC 服务器不可用` | **连续自动化约 1 小时后 Excel COM 瞬态不可用**；失败全在夹具建会话阶段、未进入断言，且同一瞬间批量发生 | 停止后续空跑 → `tasklist` 确认无残留 `EXCEL.EXE` / `testhost` → 单独复跑该分片（实测 34/34 恢复全绿）—— 见 F5-6 |
| 分批回归秒结束、退出码 0 却无用例统计 | `VSTest --filter` 误用 `.cs` 文件名后缀 → 匹配 0 用例，而 `dotnet test` 对零匹配返回 0 | 筛选改用 `--list-tests` 导出的**真实方法名前缀**；运行器校验日志须出现 `已通过!`/`失败!` 汇总行 —— 见 F5-5 |
| 回归运行器报 `UnicodeDecodeError: 'utf-8' codec can't decode byte 0xd6`，随后 `deployed binary reports ''` | `--version` 横幅含中文，经 Windows 控制台代码页输出为 **GBK**；`subprocess.run(text=True)` 按 UTF-8 解码，异常发生在读取线程里，主流程只拿到空串 | `run_version_check` 改为捕获**字节**并按 `utf-8 → cp936 → mbcs → latin-1` 逐级回退解码（`tools/pack.py`）。这是"守卫静默拿到空值"的典型面貌——自校验没报错，只是永远比不过 |

---

## 九、改名与开源发布（2.0.8-slicerlink.5 → 2.0.8-jyyj.1）

> 本节记录**当时**的改名轮次，版本号的历史值（`2.0.8-jyyj.1`）保持原样不改，
> 以便与当时的产物对得上。版本号后来演进为四段制 `3.0.0.0`，见 **第十节**。

**项目名**：ExcelMcp 本地构建 → **jyyj-mcp 助手**
**对外仓库**：https://github.com/Asaceoo/Execl-mcp

### 1. 改名落点（只改用户可见身份，不动程序集名）

| 类别 | 落点 | 改成 |
|---|---|---|
| MCP 运行时 | `Program.cs` 的 `ServerInfo.Name` | `excel-mcp` → `jyyj-mcp` |
| MCP 运行时 | `ServerInfo` 的 `ServerInstructions` 首行 | 标注为 ExcelMcp 的下游构建 |
| 控制台 | `BuildHelpText()` / `ShowVersionAsync()` 横幅 | `jyyj-mcp 助手 v{version}`（顺带把 `Usage` 里的 `mcp-excel.exe` 改成真实文件名） |
| 程序集元数据 | `Directory.Build.props` | `Product` = `jyyj-mcp`；`Authors` / `Copyright` 保留上游并追加本仓库；`RepositoryUrl` / `PackageProjectUrl` 指向本仓库 |
| 更新检查 | 两个 `NuGetVersionChecker.cs` | GitHub Releases 端点与 User-Agent 指向本仓库（否则会误报上游 `2.0.9` 可升级） |
| 分发元数据 | `mcpb/manifest.json`、`.mcp/server.json`、`vscode-extension/package.json` | 名称 / 展示名 / homepage / repository 指向本仓库 |
| 守卫同步 | `scripts/check-doc-counts.ps1` 第 6 节 | canonical URL 由上游文档站改为本仓库，4 处元数据同步 |
| 产物命名 | `tools/pack.py` | `jyyj-mcp-<version>-win-x64.zip` / `jyyj-mcp-<version>-manifest.json` |

**刻意不改**：C# 命名空间、`AssemblyName`（`Sbroenne.ExcelMcp.McpServer` / `excelcli`）、
`PackageId`、31 个 MCP 工具名与 328 个操作名。改这些会级联波及全部测试与部署链路，收益为零。

### 2. 开源化的配套改动

- `build.sh` / `release.sh` / `tools/run_regression_v5.sh` 去掉了硬编码的工作区绝对路径与开发者 home 路径，
  改为从脚本自身位置推导，并支持 `EXCEL_MCP_ROOT` / `EXCEL_MCP_REPO` / `EXCEL_MCP_DOTNET`
  / `EXCEL_MCP_NUGET` / `EXCEL_MCP_PYTHON` 覆盖；
  已有的 `build-dotnet-in-sandbox` 经验（四个 Windows 核心环境变量缺失）仍内建在 `build.sh` 里。
- 新增仓库根 `README.md`（面向使用者）与 `LICENSE`（MIT 原文 + 修改声明）。
- 新增根 `.gitignore`，显式排除 SDK / NuGet 缓存 / `bin` / `obj` / 部署目录 / 产物 / 探针产物 / 日志，
  并**显式注明**为何排除上游的 11 个 CI 工作流。

### 3. 改名后的验证（2.0.8-jyyj.1）

| 验证项 | 结果 |
|---|---|
| 构建 | 0 警告 0 错误 |
| `release.sh` 自校验 | `version check: jyyj-mcp 助手 v2.0.8-jyyj.1`；部署 139 文件 |
| MCP `initialize` 回读 | `serverInfo: name=jyyj-mcp version=2.0.8.0`；`tools advertised: 31` |
| `McpToolSurfaceTests` | 5/5 通过（含 `--help` 横幅派生计数断言） |
| 真机端到端探针 | 27 passed / 1 failed（唯一失败为旭日图这一已知 Excel 侧限制，工具如实报错） |
| 分项 | MODERN 8/8、STOCK 4/4、READBACK 14/14、SCREENSHOT 1/1（1352×196px PNG） |

---

## 十、版本演进到四段制 3.0.0.0

### 1. 为什么换版本号形态

`2.0.8-jyyj.1` 这种"上游版本 + 预发布后缀"在语义上把本仓库绑死在上游的版本号上，
而本仓库的改动面（新增多表联动动作、重写图表创建路径、新增错误透传层）已经超出"补丁"的范畴，
继续用后缀会让 ① 排序混乱（`2.0.8-jyyj.10` < `2.0.8-jyyj.9` 的字符串序问题），
② 与上游将来真实的 `2.0.9` 难以区分。改为**四段数字版本**：

| 项 | 值 |
|---|---|
| `Directory.Build.props` → `<Version>` | `3.0.0.0` |
| `<AssemblyVersion>` | `3.0.0.0`（与 `<Version>` 同步，避免协议回读与横幅不一致） |
| `<FileVersion>` | `3.0.0.0` |
| 产物 | `jyyj-mcp-3.0.0.0-win-x64.zip`、`jyyj-mcp-3.0.0.0-manifest.json` |

`tools/pack.py` 的 `next_version` 支持四段递增（`3.0.0.0` → `3.0.0.1`），
并新增 `selftest` 子命令断言 5 个版本算术用例（含 `.9 → .10` 这类进位边界）：

```bash
python tools/pack.py selftest     # 断言版本算术
python tools/pack.py show         # 读当前版本
python tools/pack.py bump         # 递增
```

> 坑：`Directory.Build.props` 是 **XML**，注释里不能出现 `--`
> （写 `--version` 会触发 `MSB4024: XML comment cannot contain '--'`）。
> 注释里引用命令请写成不带双横线的形式（如 `python tools/pack.py selftest`）。

### 2. 本轮补丁增量（相对 2.0.8-jyyj.1）

| 类别 | 内容 |
|---|---|
| 修正 | 股价图形状拒绝的**裸 COM 码** → 带图表类型与列数规则的可操作错误（F7-2） |
| 修正 | 参数形状补救提示**位置前移**（长参数列表之后 → 失败句之后），避免被客户端的输出上限吃掉（F7-3） |
| 修正 | 补齐 `ChartCreationPath` 缺失的 `IsStockShapeRefusal` / `ColumnCountPhrase`（F7-1） |
| 修正 | 工作表名单引号转义：抽出 `Core/Utilities/SheetReference.cs` + `SheetReferenceTests` 单测（F7-10） |
| 清理 | 死代码 `NeedsSeriesAttach` / `IsStockChart`、冗余包装 `IsTextValue2`（F7-6） |
| 文档 | `screenshot` 三参数补 `[Description]`；`chart` 描述由 "70+ types" 改为 84 枚举 / 37 可用（F7-7、F7-9） |
| 文档 | 新增 **`USER-GUIDE.md`**（用户手册）与 **`TECHNICAL-MANUAL.md`**（技术手册） |
| 工具 | `tools/demo_slicer_link.py` / `tools/probe_edge_cases.py` 的输出上限 400 → 1200（F7-4） |

### 3. 验证结果（3.0.0.0，全部为工具实测输出）

| 验证项 | 命令 | 结果 |
|---|---|---|
| 构建 | `bash release.sh --no-bump` | **0 警告 0 错误** |
| 部署 + 打包 + 自校验 | 同上 | 139 文件；`jyyj-mcp-3.0.0.0-win-x64.zip`（7,382,504 B）；`version check: jyyj-mcp 助手 v3.0.0.0` |
| MCP schema 审计 | `python tools/audit_schema.py --json _demo/schema-audit.json` | 31 工具 / 328 动作 / 534 参数全 snake_case；**FINDINGS: none** |
| 边界探针 | `python tools/probe_edge_cases.py` | **15/15 matched expectation** |
| 文档计数守卫 | `python tools/check_doc_counts.py` | **PASS**（15 条标题 + 工具面交叉校验） |
| 分片集成回归 | `bash tools/run_regression_v5.sh` | **148 用例全过**（A=50 / B=15 / C=36 / D=14 / E=33） |

### 4. 文档结构（本版定型）

| 文件 | 读者 | 职责 |
|---|---|---|
| `README.md` | 所有人 | 项目概览 + 文档索引 |
| `USER-GUIDE.md` | 使用者 | 装、接、用、查 |
| `TECHNICAL-MANUAL.md` | 维护者 | 架构、原理、实测事实、扩展、验证 |
| `DEPLOYMENT.md`（本文件） | 部署者 | 版本、补丁表、配置、证据、重建 |
| `ADVERSARIAL-REVIEW.md` | 维护者 | 问题清单与处置（F1–F7） |

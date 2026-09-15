# 切片器联动补丁 — 三视角对抗性审查报告

> 审查对象：`share_cache_from`（Create 路径）+ `connect-pivots` / `disconnect-pivots`（Slicer 路径）
> 版本：`2.0.8-slicerlink.1` → `2.0.8-slicerlink.3`
> 方法论：三视角独立提问 → **真机取证**（能复现的一律写成测试）→ 修正 → 复验 → 清单归零

## 视角定义

| 视角 | 关注点 | 提问方式 |
|---|---|---|
| **① 工具实现者** | COM 正确性、句柄释放、调用顺序、异常路径、幂等性 | 「这段代码在 Excel 真实对象模型下会不会炸？」 |
| **② 审查者** | 一致性、死代码、注释与实现是否相符、文档锚点是否漂移 | 「两处代码说的是同一件事吗？」 |
| **③ AI 使用方** | 参数可发现性、错误信息可执行性、返回结构自洽、能否自愈 | 「只看 schema 和报错，我能不能一次做对？」 |
| **④ 发布与验证闭环**（本轮新增） | 打包链路可执行性、**验证器自身是否可信**、测试是否因环境而假红 | 「发布脚本和守卫脚本，自己有没有被验过？」 |

## 结论

发现问题 **26 项**：确认并修复 **18 项**，真机取证后**推翻 1 项**（F1-5），审查者/发布链路自身引入的缺陷 **7 项**（F4-1、F5-1…F5-7，均已修），不计入代码的问题 **2 项**（F2-6 文档锚点、F3-4 与 F1-4 合并）。**v2.0.8-slicerlink.5 已清零待修**：F5-9（错误透传，`ToolErrorSurfaceFilter`）、F6-1（现代/股价图表创建路径，`ChartCreationPath`）、F6-2（`add-series` 传 Range 对象）、F6-3（正确命名固化为探针）全部修复并真机复验；修复过程中又发现并修复 **2 项**（F6-4 股价图 `SetSourceData` 缺 `PlotBy`、F6-5 现代图 `read/list` 读取路径抛 `NotImplementedException`）。

关键收获：**10 项只有真机跑才暴露的问题**（F1-1 / F1-4 / F2-2 / F3-3 / F5-1 / F5-2 / F5-3 / F5-4 / F5-5 / F5-7）全部由真实报错确认，无一项来自静态推测。

第二轮（视角 ④）的核心教训：**验证器本身必须被对抗性审查**，且同一缺陷出现了**三次**（F5-2 / F5-3 / F5-5）。`probe_mcp.py` 曾对缺失参数打印 `MISS` 却整体报 `PASS`；分片回归用文件名后缀当筛选条件导致**零匹配却退出码 0**，6 个分片变成空跑假绿。一个永远绿、或根本没跑的守卫，比没有守卫更危险。

---

## 一、问题清单与处置

### 视角 ① 工具实现者

| ID | 严重度 | 问题 | 真机证据 | 处置 |
|---|---|---|---|---|
| F1-1 | **高** | `DisconnectPivotTables` 未对 `toRemove` 去重：同一名字的大小写变体被计入两次 → ① `connected.Count - toRemove.Count < 1` 被污染，**误报「必须保留至少一个」**；② `RemovePivotTable` 对同一表调用两次，第二次抛裸 COM 异常 | `System.Runtime.InteropServices.COMException : 无效索引。 (0x8002000B (DISP_E_BADINDEX))` 抛在 `PivotTableCommands.<>c__DisplayClass58_0.<DisconnectPivotTables>b__0` | **已修**：新增 `TryCleanPivotTableNames`，按 `OrdinalIgnoreCase` 去重并去空白后再计数 |
| F1-2 | 低 | `ConnectPivotTables` 的 `toAdd` 未去重 → 白一次 COM 调用，`WorkflowHint` 出现重复名字 | 同 F1-1 的复现路径 | **已修**：同一 helper |
| F1-3 | 中 | `DisconnectPivotTables` 的 `RemovePivotTable` 未包异常处理，与 `Connect` 的「吞 COM 异常 + 回读校验」策略不一致 → 裸 `0x800A03EC/0x8002000B` 穿透给调用方，违背该 action 的立项目的 | F1-1 的同一处堆栈 | **已修**：包 `COMException`，由回读校验给出可执行错误 |
| F1-4 | 中 | `ConnectPivotTables` 部分成功时错误信息误导：已生效的连接不回滚、错误信息也不说明，调用方会认为「什么都没变」 | 实测 `Success=False` 且 `Live connections: [MixP1, MixP2]`，而原错误信息只提 `MixSolo` | **已修**：失败信息区分 `applied` / `refused`，并回显切片器当前实时连接列表 |
| F1-5 | 高（假设） | 猜测 `BuildExpectedRangeSourceRef` 对**整列**源范围会给出与 Excel 存储值不同的 R1C1，从而误拒合法共享 | **真机取证推翻**：`Range('A:D').Address(R1C1) = 'C1:C4'`，`PivotCache.SourceData = 'SalesData!C1:C4'` —— 两侧记法一致，无误拒风险 | **推翻**：保留现状，新增特征化测试固化该事实 |

### 视角 ② 审查者

| ID | 严重度 | 问题 | 真机证据 | 处置 |
|---|---|---|---|---|
| F2-1 | 中 | `WorkflowHint` 只在 `CreateFromRange` 设置，`CreateFromTable` / `CreateFromDataModel` 复用缓存时不回填 → 三路径行为不一致，走 table/datamodel 的调用方拿不到 `connect-pivots` 引导 | `grep WorkflowHint Create.cs` 仅 1 处命中 | **已修**：抽出 `BuildSharedCacheHint`，三路径统一调用 |
| F2-2 | 中 | `SourceDataConflict` 的 `EndsWith` 后缀启发式**静默放过不同源**；而它本要服务的 ListObject 场景已整体传 `null` → 该分支无服务对象、只贡献风险 | 新增测试 `Assert.Throws() Failure: No exception was thrown` —— 跨表共享被**静默接受** | **已修**：删除后缀容差，改为归一化后严格相等 |
| F2-3 | 中 | `NormalizeSourceRef` 剔除空白字符会把不同表名抹平 | 与 F2-2 同源（`SalesData` / `Sales Data`） | **已修**：保留空白、仅剔除 `$` 与引号；新增双向测试固化该决策 |
| F2-4 | 低 | 注释称 `RemovePivotTable` 按 MS 文档接受 name/index，而文档签名是 `RemovePivotTable(PivotTable)` | 权威文档核对 | **已修**：注释改为实测口径并指向固化该行为的测试 |
| F2-5 | 低 | `CreateFromTable` 连续两段重复解释「为什么传 `null`」 | 代码比对 | **已修**：合并措辞，并点明「记法未被任何测试固定下来」 |
| F2-6 | 低 | 技能/记忆引用 `Slicers.cs:49`、`Create.cs:49` 等**行号锚点**，补丁后必然漂移 | 补丁后行号已偏移 | **已修**：技能改为函数名锚点 |

### 视角 ③ AI 使用方

| ID | 严重度 | 问题 | 真机证据 | 处置 |
|---|---|---|---|---|
| F3-1 | 中 | 两步流程（① `share_cache_from` 建表 → ② `create-slicer` → ③ `connect-pivots` 挂其余）在 schema 里只以半句形式散落，未成体系 | 读生成后的工具描述 | **已修**：`pivottable` 描述与类型注释补「3 步」流程；`slicer` 描述补 `connect-pivots` / `disconnect-pivots` 双向用法与「至少保留一个连接」约束 |
| F3-2 | 低 | `connect-pivots` 失败分支未回显实时连接列表（`disconnect-pivots` 有）→ 无法自愈 | 见 F1-4 | **已修**：与 F1-4 合并 |
| F3-3 | 中 | 纯空白名字被静默忽略：`["", "  "]` 会走到 `added.Count == 0` 分支返回 **`Success=true` + 「已连接」** → 请求一堆垃圾名字却报成功 | `Assert.False(result.Success, "A request containing no usable PivotTable name must not report success.")` 失败 | **已修**：无法用名字直接拒绝，`connect` / `disconnect` 对称生效 |
| F3-4 | 低 | 失败信息未说明可操作前置动作的完成度 | 见 F1-4 | **已修**：与 F1-4 合并 |

### 视角 ② 附带：审查过程自身引入的缺陷

| ID | 严重度 | 问题 | 取证 | 处置 |
|---|---|---|---|---|
| F4-1 | **高** | 新写的 `tools/pack.py` 从 **MCP 服务器输出目录**取 `excelcli.exe`，而该文件只存在于 CLI 项目输出目录 → `collect` 必然 `FileNotFoundError` | `ls src/ExcelMcp.McpServer/bin/Release/net10.0-windows/excelcli.exe` → `No such file or directory` | **已修**：每个可执行文件从各自项目输出目录取，缺失即显式报错 |
| F4-2 | 低 | `pack.py` 在写入 `VERSION.txt` **之前**统计文件数，报告值与压缩包内容不一致 | 代码审查 | **已修**：先写 `VERSION.txt` 再统计 |
| F4-3 | 低 | `release.sh` 用 `$PACK` 字符串展开，python 路径含空格时会被词分割 | 代码审查 | **已修**：改为带引号的 `pack()` 函数封装 |
| F4-4 | 低 | 失败信息 `"{X} were connected"` 在单个名字时语法错误 | 真机输出 `MixP2 were connected before the request failed` | **已修**：单复数分开措辞 |

### 视角 ④ 发布与验证闭环（本轮新增）

| ID | 严重度 | 问题 | 取证 | 处置 |
|---|---|---|---|---|
| F5-1 | **高** | `release.sh` 向 `build.sh` 传 POSIX 绝对路径项目文件，经 `env` 转发未被 MSYS 转换，MSBuild 把开头的 `/d` 读成开关 → 一键发布第一步即失败 | `MSBUILD : error MSB1001: 未知开关。开关:/d/execl-mcp/mcp-server-excel/src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj` | **已修**：改传相对路径（`build.sh` 自身会 `cd` 到仓库根），并补注释固化该约束 |
| F5-2 | **高** | `probe_mcp.py` 用「序列化 JSON 含子串」代替结构断言，且唯一失败项**只打印 `MISS`、不进失败集** → 探针在参数缺失时仍报 `PASS` | 运行输出 `MISS slicer param: pivotTableNames` 后紧跟 `PASS` | **已修**：改为断言 `inputSchema.properties` 的真实键（`pivot_table_names` / `slicer_name` / `share_cache_from`），缺失即入失败集 |
| F5-3 | 高 | 探针期望的参数名是驼峰 `pivotTableNames`，而生成 schema 实际广告的是下划线 `pivot_table_names` → 期望值本身错误，属「验证器与被验证物脱节」 | `tools/list` 实测 `SLICER props: [action, clear_first, column_name, ..., pivot_table_names, position, ...]` | **已修**：期望值改为实测名；并把「以序列化子串代结构校验」这一反模式标注在代码注释中 |
| F5-4 | 中 | `GroupByDate_*` 4 项长期红：断言英文子串 `Days/Months/Years/Quarters`，而中文版 Excel 生成的字段名为 `天(Date)` / `月(Date)` / `年(Date)` / `季度(Date)` → 分组功能正常但测试必红，**掩盖真实回归** | `Expected to find Days field after grouping. Actual fields: Region, Product, Sales, Date, 天(Date)` | **已修**：抽出 `HasGroupedDateField(fields, english, localized)`，两种命名均接受；并借 `git status` 独立确认分组实现文件（`RegularPivotTableFieldStrategy.cs`）**未被本补丁改动**，故排除回归嫌疑 |
| F5-5 | **高** | 回归分片用**文件名后缀**当筛选条件（`FullyQualifiedName~PivotTableCommandsTests.Slicers`、`...Creation` 等），而这些只是 `.cs` 文件后缀、不是 FQN 的一部分 → **匹配 0 个用例**；`dotnet test` 对「零匹配」返回 **0**，分片于是静默报成功 | 7 个分片全程 1m49s 结束、`EXIT=0`、日志中**没有任何** `已通过!` 汇总行；对比用真实类名 `PivotTableOlapDisambiguationTests` 的分片正常跑出 `通过:4` | **已修**：改用**真实方法名前缀**（`_tlist.log` 的 146 个用例名）重建分片；运行器加硬校验 —— 日志中缺 `已通过!`/`失败!` 汇总行一律判为失败并显式打印 |
| **F5-7** | **高** | 为满足「产物带版本号」，`pack.py` 把两个 `.exe` **单独**拷进 `dist/`。但 .NET apphost 只是加载器，必须与 `*.dll` / `*.deps.json` / `*.runtimeconfig.json` 同目录 → **产出两个永远无法启动的死文件**；同时 zip 只打包 server 输出、**不含 CLI**，与 manifest 声明的三产物不一致 | `> dist\Sbroenne.ExcelMcp.McpServer-2.0.8-slicerlink.3.exe --version` → `The application to execute does not exist: 'D:\execl-mcp\dist\Sbroenne.ExcelMcp.McpServer.dll'`，**退出码 2147516570**；`excelcli` 同样报 `...dist\excelcli.dll`；zip 内 `含 excelcli.exe: False` | **已修**：①`copy_entry_point()` 把 CLI 的 apphost 四件套并入部署目录（与 server 共享依赖集），zip 成为**单一自洽分发包**；②删除裸 exe 产出逻辑；③新增 `assert_bundle_runnable()`，缺任一入口四件套即发布失败；④新增 CLI 冒烟（无输出即非零退出）。复跑实测：`deployed 139 files`、`zip 7,374,611 B`、`version check OK`、`cli check: ✓ You're running the latest version: 2.0.8-slicerlink.3` |
| **F5-8** | **中** | 末批分片 `f_misc` 用 `List_` / `GetInfo_` / `Delete_` / `Refresh_` / `GetData_` 等**过泛前缀**建筛选 → 实际匹配 **126** 个用例（远超 Slicer/PivotTable 面），35 分钟墙钟内只跑完 **30** 个即被截断。输出仍是 `失败:0，通过:30`，**看似全绿，实则覆盖率 24%、且无法证明目标用例已跑** | `--list-tests` 同条件权威枚举 = **126**；`_tlist.log`（Slicer/PivotTable 面）差集后真正欠覆盖 = **16**；批次 `EXIT=124` 截断 | **处置**：改用「已覆盖前缀**差集**」+ **完整方法名**精确建批，并断言执行总数等于清单长度（16），杜绝「匹配过多被截断 / 匹配过少假绿」两种失真。结果见 `_r_f_rest.log` |
| F5-6 | 中（非代码） | 连续自动化约 1 小时后，Excel COM 服务瞬态不可用：22 项 `CreateFrom*` 用例全部报 `COMException 0x800706BA RPC 服务器不可用`，外加 1 项表格切片器用例同样报错 | ①失败全在**夹具建会话阶段**（`DataModelPivotTableFixture.InitializeCore` → `AllocateComObject`），Polly 重试耗尽，**未进入任何断言**；②22 项失败**集中在 `00:00:18.80–18.84` 这 40 毫秒内同时发生**，属夹具一次性初始化失败，不是逐条退化；③日志中表征产品路径的 `0x800A03EC` 出现 **0** 次；④16 分钟前的 `a_slicer` 仍是 25/25 全绿 | **判为环境抖动，已复跑确认**：停止后续空跑批次 → 确认无残留 `EXCEL.EXE`/`testhost` → 单独复跑同一筛选条件 → **34/34 全部通过、0 失败**（`_r_c_create2.log`，22m56s）。与代码无关 |

| **F6-1** | **高** | 图表面「现代图 + 股价图」整类不可用：`CreateFromRange` 一律走 `Shapes.AddChart(XlChartType:...)`，创建后再 `Chart.SetSourceData()`。对 Excel 2016+ 的现代图（树图/瀑布/漏斗/直方图/帕累托/箱线/地图/组合）与 4 种股价图，这两个调用分别抛 `NotImplementedException` 与 `0x800A03EC` → **枚举里声明的 84 种类型，实际只有 25 种建得出来**，且失败文案是泛化异常，调用方无从判断是类型不支持还是数据错了 | MCP 探针 `tools/probe_chart_types.py`：`BASIC 18/18`、`3D 4/4`、`SHAPES 3/3` 全过；`MODERN 1/9`（旭日图还是**假成功**，回读 `ChartType=-4111`）、`STOCK 0/4`。pywin32 直连对照组 `tools/verify_addchart2.py` + `tools/verify_chart_paths.py`：①`AddChart(117)` 与 `AddChart2(-1,117)` 都能建出来，但 `SetSourceData` 必抛 `0x800A03EC`；②改用 `SeriesCollection.NewSeries()` + `Values=Range` → **8 种现代图全部成功且类型不回落**；③股价图用「先建柱形图 → `SetSourceData` → 改 `ChartType=88/89`」→ **OK**。即 Excel 本身支持，是**调用路径**不支持 | **未修**（已定位根因）：建议 `RegularChartStrategy` 增加两条路径——现代图用 `AddChart2` + 手动 `NewSeries` 挂系列；股价图用「先建柱形图再改类型」。另需给 `list`/`read` 的类型回读加校验，避免旭日图这类「报成功但类型回落」 |
| **F6-2** | 中 | `RegularChartStrategy.AddSeries` 把**地址字符串**赋给 `Series.Values`（`newSeries.Values = valuesRange;`），而 `Series.Values` 只接受 Range 对象或数组 → `chart_config add-series` 在任何工作簿上都抛 `0x800A03EC` | 干净环境复现（`tools/_verify4.py`）：新建工作簿 → 建簇状柱形图 → `add-series`，三种参数组合（`series_name+values_range` / 仅 `values_range` / 加 `category_range`）**全部** `COMException 0x800A03EC`；同一脚本用 pywin32 传 **Range 对象**则成功 | **未修**：改为从图表所在表取 `Range` 后再赋值（对照 pywin32 成功路径），并补一条真机用例 |
| F6-3 | 低（文档） | 工具名与参数名在文档/示例里是驼峰（`chartconfig`、`visible`），实际 MCP 面是下划线（`chart_config`、`source_range`、`show_major`、`destination_sheet`）。按驼峰写调用会得到 `Unknown tool` 或「参数不合法」 | `tools/list` 实测 31 工具名；`chart_config` 的 schema 参数键均为 snake_case | **未修**：把正确命名写进 `tools/probe_chart_types.py` 作为可执行示例（`_verify4.py` 亦已固化三种参数组合） |

> **本轮反复出现的同一缺陷类别**（F5-2 / F5-3 / F5-5）：**验证环节本身没跑、或跑了却报绿**。
> 三次都是「工具输出被当作通过证据，而工具并未真正校验任何东西」。结论：任何守卫/探针/分片，
> 都必须 (a) 对真实结构断言，(b) 失败（含「零匹配」「无结果行」）进入失败集，(c) 被截断时不计为通过。

---

## 二、兜底校验清单（结果）

- [x] 真机测试通过：新增 6 项（重复名 / 空白名 / 部分成功 / 后缀表名 / 含空格表名 / 整列记法）
- [x] 真机测试通过：`SlicerLink` 全组 11 项（`_r_slicer.log`：`通过:11，失败:0`）
- [x] 真机测试通过：PivotTable 面已执行 20 项（`_r_pivot.log`：`通过:20，失败:0`）
- [x] `GroupByDate` 4 项由「区域设置脆弱」修复为通过（`_r_gbdate2.log`）
- [x] **切片器面全覆盖 25/25 通过**（`_r_a_slicer.log`，15m42s）：`SlicerLink_` 11 + `ListSlicers_` 5 + `CreateSlicer_` 4 + `SetSlicerSelection_` 3 + `DeleteSlicer_` 2（**含本补丁全部改动面**）
- [x] **字段面全覆盖 27/27 通过**：`Add*Field_` / `SetField*`（格式/汇总/名称/筛选）/ `SortField_` / `RemoveField_` / `ListFields_`（`_r_d_field.log`，18m21s）
- [x] **创建路径全覆盖 34/34 通过（含 `share_cache_from` 补丁面）**：首跑因 Excel COM 瞬态不可用报 22 项失败（F5-6），**复跑确认 34/34 全绿**（`_r_c_create2.log`，22m56s）
- [x] **表格切片器面 16/16 通过**：首跑 15 通过 + 1 项 `CreateTableSlicer_InvalidColumn_ReturnsError` 报 `COMException 800706BA`（`AllocateComObject` 阶段即失败，未进入任何断言）；单独复跑该项 **1/1 通过**（`_r_tblslc_one.log`，1m15s）→ 同属 F5-6 环境抖动
- [x] **末批杂项 18/18 通过**（`_r_f_rest.log`，15m43s，`EXIT=0` 无截断）；以「已覆盖前缀差集」+ 完整方法名建批，并核对**欠覆盖 16 项全部在内、0 遗漏**（多跑的 2 项为同名方法跨类的 `ChartCommandsTests.List_*`）。Slicer / PivotTable 面至此**全部 146 项覆盖完毕**（F5-8 已闭环）
- [x] **计算字段 / 计算成员 / 总计 / 布局 / 分组面 36/36 通过**（`_r_e_calc.log`，22m33s）
- [ ] **全量回归仍在推进**：本沙箱实测 **43–70 秒/用例**（Excel COM 启动开销），146 项串行约需 2–3 小时，且连续自动化 ~1 小时后会触发 Excel COM 瞬态不可用（F5-6）。**已跑完的分片除环境抖动外 0 失败**；未跑完部分不声称通过
- [x] 分片回归的筛选条件已修正为**真实方法名前缀**（F5-5），运行器对「零匹配 / 无结果行」显式判失败
- [x] 修正后的分片日志：`_r_a_slicer.log` / `_r_b_tblslc.log` / `_r_c_create.log` / `_r_d_field.log` / `_r_e_calc.log` / `_r_f_misc.log`（后 4 片仍在跑）
- [x] `McpToolSurfaceTests` 通过（操作数守卫 328）
- [x] 文档计数守卫通过（31 工具 / 328 操作，15 条标题 + 工具面 + `--help` 派生）
- [x] MCP `tools/list` 探针通过（5 项结构断言，含参数键与 action 枚举）
- [x] `--version` 与 `Directory.Build.props` 一致（由 `pack.py collect` 断言，`v2.0.8-slicerlink.3`）
- [x] 无 `COMException` 裸抛（`connect` / `disconnect` 均已转为可执行错误信息）
- [x] 注释与实现一致（无「文档说 A、代码做 B」）
- [x] 打包产物全部带版本号后缀，且版本号自动递增（`.2` → `.3` 实测）
- [x] **产物不只要有版本号，还必须能启动**：zip 需含每个入口的 apphost 四件套（`assert_bundle_runnable` 守卫），并对 CLI 实跑 `--version` 冒烟（F5-7）
- [x] 验证器自身经对抗性检查（F5-2 / F5-3：探针不得「打印 MISS 却报 PASS」）

## 三、新增测试（可复现的真实报错来源）

| 测试 | 固化的事实 |
|---|---|
| `SlicerLink_DisconnectDuplicateNames_KeepTheLastConnectionGuardAccurate` | 重复名不得污染「至少保留一个连接」的算术，也不得对同一表调用两次删除 |
| `SlicerLink_ConnectBlankOnlyRequest_IsRejectedInsteadOfReportedAsDone` | 无可用名字的请求不得报成功（connect / disconnect 对称） |
| `SlicerLink_PartialConnectFailure_DisclosesTheStateItLeftBehind` | 部分失败必须披露已生效的连接 |
| `SlicerLink_ShareCacheFromSuffixSheetName_IsRejectedAsADifferentSource` | 表名互为后缀（`SalesData` ⊂ `RawSalesData`）视为不同源 |
| `SlicerLink_ShareCacheFromSpaceInSheetName_MatchesOnlyTheSameSheet` | 引号/锚点归一化，**空白不归一化**（`Sales Data` ≠ `SalesData`） |
| `SlicerLink_WholeColumnSource_RecordsTheNotationTheGuardMustCompare` | 整列源两侧记法一致（`C1:C4`），无误拒风险 |

## F6 追加（v2.0.8-slicerlink.5 修复轮发现并修复）

| 编号 | 严重度 | 问题 | 真机证据 | 处置 |
|---|---|---|---|---|
| **F6-4** | **高** | 股价图创建的隐性坑：`Shapes.AddChart(51)` 会先用**数据区整个 CurrentRegion** 自动填充图表；随后 `SetSourceData(子区域)` 若**不传 `PlotBy`**，系列数仍是整区的（如 A1:E6=5 列），切 `ChartType=88/89/90` 时 Excel 拒绝并抛 scode `0xB0D7018E/0x190/0x191`——失败集合随夹具数据列数漂移，极具迷惑性 | pywin32 对照（`tools/_diag_stock.py`）：同一布局下不传 `PlotBy` → HLC/OHLC/VHLC 必失败（scode 与集成测试报错逐位一致）、VOHLC 恰好列数匹配而「假绿」；显式传 `PlotBy=2`（xlColumns）→ **4/4 全过且类型不回落** | **已修**：`ChartCreationPath.CreateShape` 股价路径显式传 `PlotByColumns`；集成测试用例改为按变体列数建数据（HLC=3 列、OHLC/VHLC=4 列、VOHLC=5 列，股价图不接受类别列），分片复跑 **12/12 通过** |
| **F6-5** | 中 | 现代图（树图/瀑布/漏斗等）建出来后，`chart list` / `read` 会在两处崩：①策略判定 `CanHandle` 只捕获 `COMException`，而迟到绑定器把「无 `PivotLayout` 成员」报成 `NotImplementedException`；②`GetDetailedInfo` 回读 `series.Values/XValues/Formula` 时 Excel 对现代图抛 `E_NOTIMPL`（**写入可用、读取被拒**） | 属性逐个探测（`tools/verify_modern_chart_properties.py`）：现代图 `PivotLayout`、`Series.Values/XValues/Formula` 的**读取**全部抛 `NotImplementedException`；同一系列**写入** `Values` 却成功 | **已修**：两个策略的 `CanHandle` 增补 `NotImplementedException` 分支；`GetDetailedInfo` 三处读取降级为 `SafeText`（失败返回空串，不拖垮整个 inspection）。部署后探针 `MODERN 8/9`（唯一失败为旭日图的本机 COM 限制，与代码无关） |
| F6-2 复验 | — | `add-series` 修复后 | 部署版探针 `CONFIG 17/17`（含 `add-series`） | **闭环** |
| F5-9 复验 | — | 错误透传修复后 | 探针对非法参数已能看到「Tool 'file' rejected the arguments before execution…」级别的具体原因，不再是统一泛化文案 | **闭环** |

**图表能力终态（部署版 v2.0.8-slicerlink.5 探针 62/64）**：BASIC 18/18、3D 4/4、SHAPES 3/3、STOCK 4/4、MODERN 8/9（旭日图为本机 COM 限制）、CONFIG 17/17、DASHBOARD 8/9（截图需 `show:true` 前台会话）。**可用图表 25 → 37 种**。

---

## F7 追加（v3.0.0.0 轮：三视角复审 + 工具闭环取证）

本轮仍按三视角展开（工具实现者 / 审查者 / AI 使用方），**每个结论都由真机或工具输出取证**，
不采信"读源码觉得没问题"。版本推进到 **3.0.0.0**（四段版本号）。

### 复审发现与处置

| 编号 | 视角 | 严重度 | 问题 | 证据 | 处置 |
|---|---|---|---|---|---|
| **F7-1** | 审查者 | **高** | **本仓库当前源码树编译不过**：`ChartCreationPath.cs` 调用了 `IsStockShapeRefusal(...)` 与 `ColumnCountPhrase(...)`，但这两个方法**在文件中不存在**（上一轮编辑只写入调用点与辅助函数体，未定义方法本身）。改动晚于上一次构建，因此"上一个产物是好的"掩盖了这个状态 | `grep` 全文只有第 79/88 行两处**调用**、无定义；`release.sh` 重跑构建是唯一能暴露它的动作 | **已修**：补齐 `IsStockShapeRefusal`（匹配 `0xB0D7018E` / `0xB0D70190` / `0xB0D70191`）与 `ColumnCountPhrase`。构建复验 **0 警告 0 错误** |
| **F7-2** | 工具实现者 | 中 | 股价图列数不匹配时只抛**裸 COM 码**：`COMException: 0xB0D7018E`——既没说是哪个图表类型，也没说列数规则。这正是当初定这个坑要花一次 bisect 的原因 | `tools/probe_edge_cases.py` 的 `CH-STOCK-WRONGCOLS`：`observed: 'COMException: 0xB0D7018E'`，期望短语 `stock\|series` 缺席 → 判 FAIL | **已修**：分级错误文案。命中形状拒绝码 → 说明"该变体需要几列、你给了几列"；其他 COM 失败 → 只报操作与码，**不把列数规则扣到无证据的失败上**。对 `0x800A03EC` 刻意不匹配，因为该分支上它也可能是"源区域不可用" |
| **F7-3** | AI 使用方 | 中 | **可操作的补救提示排在长列表之后**：`ToolErrorSurfaceFilter` 把"参数形状写错了、应该怎么改"追加在 `Accepted argument(s): …`（11 个参数名）**之后**，实测落在第 437 字符；而客户端输出上限（本仓库探针为 400）恰好把它切掉 → **最有用的一句话最先被丢弃** | 直接 stdio JSON-RPC 取证：完整错误文本 674 字符，`remedy position: 270`（修复前 >437）、`accepted-args position: 437` | **已修**：补救提示**前移到失败句之后**、参数列表之前。这条同时解释了一个假象——探针此前报"提示缺失"，实际是**截断**被读成了**缺失** |
| **F7-4** | 审查者 | 低 | 诊断工具自身的 400 字符上限会制造假阴性 | `tools/demo_slicer_link.py` 两处 `[:400]`、`tools/probe_edge_cases.py` 一处 `[:400]` | **已修**：统一放宽到 **1200**，并在注释里写明"截断曾被读成缺失" |
| **F7-5** | AI 使用方 | 低 | **README 旗舰示例有三处错误**：`pivottable(action="create", …)` 的动作不存在（真实动作是 `create-from-range` / `create-from-table` / `create-from-datamodel`）；`source_range_address` 不是该工具的参数（真实是 `source_sheet` + `source_range`）；缺 `destination_sheet` / `destination_cell` | schema 导出：`pivottable` 的 action enum 里没有 `create`，参数表里没有 `source_range_address` | **已修**：示例改为可直接运行的真实调用，并补一条"动作名与参数名"提示，指向用户手册速查表 |
| **F7-6** | 审查者 | 低 | 死代码：`ChartCreationPath.NeedsSeriesAttach` / `IsStockChart` 声明后**零引用**。因为成员是 `internal`，编译器不告警，只能靠人工或审查发现 | 全文 `grep` 只有定义、无调用；无 `CS0169`/`CS0414` 告警 | **已修**：删除；同时删掉冗余包装 `IsTextValue2`（直接调 `IsText`） |
| **F7-7** | AI 使用方 | 低 | `screenshot` 工具的三个参数（`sheet_name` / `range_address` / `quality`）**没有 `[Description]`**——模型不知道该填什么 | `tools/audit_schema.py` 报 3 个参数缺描述 | **已修**：补齐 `[Description]`。复审后 `audit_schema.py` 报 **FINDINGS: none** |
| **F7-8** | 审查者 | 低 | 改名残留：`ToolErrorSurfaceFilter.cs` 与 `ExcelToolsBase.cs` 的日志前缀仍是 `[ExcelMcp]`，与项目名 `jyyj-mcp` 不一致 | 全文 `grep "\[ExcelMcp\]"` | **已修**：统一为 `[jyyj-mcp]` |
| **F7-9** | AI 使用方 | 低 | `chart` 工具描述写 "70+ types"，与实际不符（枚举 84、真机可用 37） | 工具描述文本 | **已修**：改为 "84 types are enumerable and 37 are verified to work (…) after the local creation-path patch" |
| **F7-10** | 工具实现者 | 中 | 图表数据源地址**未转义工作表名单引号**：`$"'{sheetName}'!{range}"`。表名含 `'`（如 `O'Brien`）时抛 `0x800A03EC`。Excel 规则是**单引号加倍**（`'O''Brien'`） | `CH-QUOTE-APOSTROPHE` 用例：修复前 `COMException 0x800A03EC`；修复后 `success`，且 `With Space` / `Plain` 两个控制组仍通过 | **已修**：抽出共享工具 `Core/Utilities/SheetReference.cs`（`QuoteSheetName` + `BuildRangeReference`），配 `SheetReferenceTests` 单测（Plain / With Space / `O'Brien` / `'Quoted'` / `a'b'c`） |
| **F7-11** | — | 低 | 项目记忆 `MEMORY.md` 超过注入上限（18,059 B）被系统截断，导致关键约束进入不了上下文 | 注入时的 `ACTION REQUIRED` 提示 | **已修**：合并去重重写为约 6 KB，约束条目按"违反必踩坑"排序保留 |

### 本轮验证结果（全部为工具实测输出）

| 验证项 | 命令 | 结果 |
|---|---|---|
| 构建 | `bash release.sh --no-bump` | **0 警告 / 0 错误**（两个项目） |
| 部署 | 同上 | **139 文件**（覆盖式复制，无递归删除） |
| 打包 + 自校验 | 同上 | `jyyj-mcp-3.0.0.0-win-x64.zip`（7,382,504 B）；`version check: jyyj-mcp 助手 v3.0.0.0` |
| MCP schema 审计 | `python tools/audit_schema.py --json _demo/schema-audit.json` | 31 工具 / 328 动作 / 534 参数，**命名 534/534 snake_case**；错误可操作性 4/4 OK；**FINDINGS: none** |
| **边界探针** | `python tools/probe_edge_cases.py` | **15/15 matched expectation**（修复前 13/15；两个 FAIL 分别为 F7-2 真缺陷与 F7-3 截断假象） |
| 文档计数守卫 | `python tools/check_doc_counts.py` | **PASS**——15 条标题 + 工具面交叉校验一致（31 工具 / 328 操作） |
| 分片集成回归（v.5 基线） | `bash tools/run_regression_v5.sh` | **148 用例全过**：A=50 / B=15 / C=36 / D=14 / E=33；E 片一次瞬态 COM 失败后自动复跑通过 |

### 本轮固化的两条工程原则

**① 错误文本里，可操作的补救必须排在最前。**
F7-3 的本质不是"少了一句话"，而是**排序错了**。任何"先铺一堆上下文、最后才说要怎么改"的错误消息，
在真实链路上（客户端截断、终端滚动、日志行、模型上下文预算）都会退化成不可用。
诊断细节应该排在补救之后，因为细节只在人要看的时候才有价值，补救是每次都要用的。

**② 中途编辑必须立刻过一遍编译。**
F7-1 的成因是"写入调用点"与"写入方法定义"分成了两次编辑，中间产物不可编译，
而两次构建之间没有人碰过它，于是坏状态静默存活。
规则：**任何跨文件/跨方法的改动，落盘后第一件事是构建**，不要等"下一轮统一验证"。

### 迭代轨迹（三视角 → 收敛）

```
视角① 工具实现者 ─┐
视角② 审查者     ─┼─► 12 项发现（含 2 项审查工具自身缺陷）─► 逐项修复 ─► 复验
视角③ AI 使用方  ─┘                                              │
                                                                 ▼
                                          audit FINDINGS none · probe 15/15 · 计数 PASS · 回归 148/148
```

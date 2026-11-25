# 构建记录（决策与修复留痕）

## 2026-11-24 · 首版

### 目标
把 PowerPoint 的切换/动画 OOXML 逆向后，做成 AI 可直接驱动的生成能力（MCP）。

### 决策
1. **用真机采样代替 schema 推断**。先枚举 COM 类型库拿名字（192 切换 / 155 动画），
   再让 PowerPoint 逐个写文件，抽 XML 对照。534 页样本全部来自 PowerPoint 本身。
2. **目录 id 取 COM 枚举名**（`ppEffectCoverLeftUp` → `cover-left-up`），不取元素名。
   原因：`glitter` 的 diamond/hexagon 两族共享 `<p14:glitter>`，靠 `pattern` 区分；
   `strips` 的两个枚举名映射同一 XML。用元素名会撞车，用枚举名天然唯一。
3. **中文别名用"族 + 方向"规则生成**，覆盖 PowerPoint 中文 UI 叫法（淡出/覆盖（左上）/涡流（左）…），
   少数特例（卷页单页左、轮子 1 根辐条）单独登记。
4. **两套写入路径**：内容用 python-pptx（保证可编辑），切换/动画用 lxml 直接写 `p:transition`/`p:timing`。
   不在 python-pptx 的对象模型里硬加东西，避免升级冲突。
5. **动效风格套餐**（subtle/bold/cinematic/tech）：AI 不必逐页想效果，选风格即可；
   仍然可以逐页覆盖。

### 采样中得到的关键结论（已写进 schema/ooxml_structure.json）
- 切换效果是 `p:transition` 的子元素，不是属性。
- `p14/p15/p159` 元素必须 `mc:AlternateContent` 包裹；`p14:dur` 一旦使用也必须包裹。
- 动画触发写在 `nodeType`（clickEffect/withEffect/afterEffect），`with`/`after` 追加进当前 click 组。
- 效果时长落在行为节点 `cTn@dur`，而 `p:set` 内的 `dur="1"` 是瞬时标记不能改。
- `p:cTn@id` 全页唯一，注入时整树重编号。
- 文本分级 = 每段一个 click 组 + `p:txEl/p:pRg` + `bldP build="p"`。

### 修复
1. `MSPPT.OLB` 取常量：`PyITypeInfo.GetConstant` 在 pywin32 里不存在，
   `PyVARDESC.value` 直接就是常量值 —— 改用它。
2. `glitter-left` 别名指向了不存在的 id（真实 id 是 `glitter-diamond-left`），修正别名表。
3. 动画目标不存在时原本抛 `KeyError` 导致整份构建失败 → 降级为 warning 并跳过该条，
   同时在 warning 里列出本页可用目标名。
4. 分节版式的 `index` 大号数字框高不足（`BoundHeight` 131pt > 框 115pt），
   调高框体并下移标题，`layout_check` 归零。
5. MCP server 启动失败：本机 anyio 3.7.1 与 mcp 1.16 不兼容
   （`create_memory_object_stream[...]` 在 3.x 是函数）。升级 anyio 到 4.16.0 后 stdio 正常。
6. `Slide.Export` 需要绝对路径，否则报"找不到文件"；`export_png` 改为 `abspath`。

### 验证证据
- `scripts/matrix_test.py`：24/24（真实 PowerPoint 回读 `EntryEffect` 与预设）
- `validate`：0 problem；`layout_check`：0 溢出
- `scripts/test_stdio.py`：MCP 协议端到端，15 工具
- `render_png`：逐页 `ink_ratio` 0.04–0.41，页面非空白

---

## 2026-11-24 · 第二轮：形状层（定位 / 属性 / 组合 / 移动 / 缩放）

### 新增采样（第 4 层）
`rev/acquire_shapes.py`：7 页真机采样，覆盖组合、旋转、翻转、透明度、线宽、渐变、
组整体移动/缩放、组动画与组内子动画的 target。产物 `rev/out/shapes_map.json`。

结论（已写进 `schema/ooxml_structure.json` 的 `shape_layer`）：
1. `p:grpSp` 的子形状**直接挂在 grpSp 下**，没有 `p:spTree` 包裹。
2. 组合时 `off = chOff = 包围盒左上`、`ext = chExt = 包围盒尺寸`，子形状 off/ext 不变。
3. 组移动只改 `off`，组缩放只改 `ext`；视觉 = `off + (child_off − chOff) × (ext/chExt)`。
4. 组动画 target 仍是 `<p:spTgt spid="组id"/>`，没有 `p:grpSpTgt`。
5. `rot` 单位 60000/度；`flipH/flipV="1"`；透明度 `a:alpha@val` 百分比×1000。

### 新增能力
- `mcp_ppt/shapes.py`：单位解析（英寸/`%`/`in`/`pt`/`cm`/`emu` + 9 种锚点）、
  `place` / `move` / `resize` / `set_props` / `group_shapes` / `ungroup_shape` /
  `align_shapes`（含分布）/ `duplicate_shape` / `delete_shape` / `apply_ops` / `shape_info`。
- `add_element`：12 种几何（含 `line` 用 `a:cxnSp` 真直线）。
- deck 模型新增 `slides[].elements` / `slides[].ops` / 元素的 `group` 字段；
  `patch_deck` 现在同时支持元素、shape ops、切换、动画。
- 动画 target 支持 `group:<组名>`。
- MCP 工具 +3：`add_elements` / `edit_elements` / `element_info`（共 18 个）。

### 修复
1. `shape_tree` 解析组内子形状时错误地去找 `p:spTree`，改为直接遍历 `grpSp` 子节点。
2. 动画 target 解析新增 `group:` / `name:` 前缀；`shape:<id>` 分支被误删后补回。
3. `ungroup` 的临时注释节点会残留在 spTree 里，改为插入后立即清理。
4. `delete_shape` 在编辑中被误删一行，已修复。

### 验证
- `scripts/test_shapes.py`：组合 → 平移 → 缩放 → 解组 → 对齐，COM 回读几何一致。
- `scripts/test_layout.py`：9 条 ops 全 ok；组动画 `(23, 3, 'row')`、子形状 `(61, 3, 'nodeC')`。
- `scripts/test_shape_tools.py`：`add_elements` → `edit_elements` → `element_info` 全通。
- `scripts/run_all.py`：8 步全绿。

---

## 2026-11-24 · 第三轮：绘图属性（填充 / 线条 / 效果 / 文本 / 图片）

### 新增采样（第 5 层）
`rev/acquire_styles.py`：6 页真机采样，逐项钉死填充（solid+alpha / 双色渐变 / 单色渐变 / 图案 / 无）、
线条（虚线 / 箭头 / 透明度 / 粗细）、效果（外阴影 / 发光 / 柔化边缘 / 3D）、
文本（run 级粗斜体下划线、中西文字体、段前段后、行距、项目符号、缩进）、图片裁剪与亮度、表格样式。
产物 `rev/out/styles_map.json` + `styles_raw/`。

关键结论（写进 `schema/ooxml_structure.json` 的 `shape_layer.style`）：
1. 渐变 `a:gs@pos` = 1/1000 %，`a:lin@ang` = 1/60000 度，径向用 `a:path/a:fillToRect`。
2. `a:effectLst` 子元素顺序固定：blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge。
3. `blurRad`/`rad` = 磅×12700；`Shadow.Transparency` 要换算成 `alpha = 1-x`。
4. 文本 `sz` = 1/100 pt；段落 `lnSpc@spcPct` = %×1000、`spcBef/spcAft@spcPts` = 1/100 pt；缩进是 EMU。
5. 图片 `p:blipFill`（**p 命名空间**）里的 `a:srcRect` = 比例×100000；COM 的 `CropXxx` 单位是磅。
6. **COM 的 RGB 是 BGR**（`0x3366CC` → `srgbClr val="CC6633"`）。

### 新增能力
- `mcp_ppt/dml.py`：`set_fill`（solid/gradient/radial/none）、`set_line`（虚线/箭头/透明/连接）、
  `set_effects`（外阴影/内阴影/发光/柔化）、`set_text_body`（段落 + run 重建）、
  `set_crop`、`set_lum`；全部按 schema 顺序插入（`insert_ordered`）。
- `set_props` 扩展：`fill` 接受 dict、`effects`、`line_*` 全套、`paragraphs`/`runs`、`crop`、`brightness`/`contrast`。
- `add_element` 改为"除结构键外全部透传 set_props"，新属性无需再改白名单。
- MCP 工具 +1：`set_text`（19 个）。

### 修复
1. **`_color_el` 生成嵌套颜色**：`_color_el("a:srgbClr", …)` 产出 `<a:srgbClr><a:srgbClr/></a:srgbClr>`，
   导致渐变、阴影、发光三处文件被 PowerPoint 判为损坏（能保存、python-pptx 能读，但 PowerPoint 打不开）。
   拆出 `_clr()` 专门生成单层颜色节点，并用 `scripts/diag_styles.py` 逐项二分验证 11/11 可打开。
2. `set_crop` 找的是 `a:blipFill`，实际是 `p:blipFill`，裁剪一直静默失败。
3. `element_info` / `element_info` 的 docstring 在编辑中被误删后补回。

### 验证
- `scripts/diag_styles.py`：11 个单变量文件全部能被 PowerPoint 打开。
- `scripts/test_styles.py`：COM 回读 `fill.Type=3`（渐变）、`Shadow.Visible=-1`、`crop=(48,24,24)`；
  `validate` 0 problem、`layout_check` 0 溢出。
- `scripts/test_deck_styles.py`：同一套属性通过 deck 模型写出，组动画 `(23, 3, 'badge')`。
- `scripts/run_all.py`：11 步全绿。

---

## 2026-11-24 · 第四轮：数据层（表格单元格 + 原生图表）

### 新增采样（第 6 层）
`rev/acquire_table_chart.py` + `rev/probe_chart_anim.py`：4 页真机样本，覆盖
表格合并/列宽行高/单元格填充/四边与对角边框/表样式，以及图表部件与图表动画的组合探查。

关键结论（写进 `schema/ooxml_structure.json` 的 `shape_layer.style.table/chart`）：
1. 表格合并：主单元格 `gridSpan="2"`，被吞的单元格保留 `a:tc` 并写 `hMerge="1"`。
2. 边框在 `a:tcPr` 内，顺序 lnL → lnR → lnT → lnB → lnTlToBr → lnBlToTr → 填充。
3. 列宽 `a:gridCol@w`、行高 `a:tr@h`；表样式是 `a:tblPr/a:tableStyleId{GUID}`。
4. 图表 = `p:graphicFrame` + `ppt/charts/chartN.xml` + `ppt/embeddings/*.xlsx`。
5. **图表分级动画不可用**：`msoAnimateChartBySeries/Category/AllAtOnce` 在 `AddEffect` 上全部 `Invalid request`。

### 新增能力
- `mcp_ppt/table.py`：`build_table` / `set_cell` / `set_borders` / `merge` / `style_table` /
  `column_widths` / `row_heights` / `edit_table` / `table_info`。
- `mcp_ppt/chart.py`：`build_chart`（13 种图表 + 配色 + 数据标签 + 轴格式 + 柱间距）/ `chart_info`。
- 版式新增 `chart`，`table` 版式重写为走 table 构建器（表头/斑马纹/列宽）。
- `add_element` 支持 `kind: table|chart`；`apply_ops` 支持 `op: table`（cells/merge/style/widths/heights）。
- MCP 工具 +3：`add_table` / `add_chart` / `edit_table`（共 22）；`element_info` 对表格/图表附带内容摘要。

### 修复
1. **python-pptx 的 `Emu` 在算术后退化成 `int`**（`Emu - int*Emu`），导致 `isinstance(x, Length)` 判断失效，
   版式传入的 EMU 被当英寸再乘 914400，表格宽度写成 9.6e12 → PowerPoint 打不开。
   新增 `to_emu_len()`（Length / 大整数 / 英寸 / 带单位字符串），table 与 chart 统一走它。
2. `apply_ops` 不认 `op: table`，表格编辑 ops 被当作未知操作跳过。

### 验证
- `scripts/diag_data2.py`：6 个用例（纯 python-pptx / 构建器 / 版式）全部能被 PowerPoint 打开。
- `scripts/test_data.py`：3 页（2 图表 + 1 表格），ops 全 ok，`validate` 0 problem、`layout_check` 0 溢出，
  图表动画回读 `(10,1,'title')`/`(22,3,'chart')`。
- `scripts/test_data_tools.py`：`add_table → edit_table → add_chart → element_info` 全通，工具 22 个。
- `scripts/run_all.py`：14 步全绿。

---

## 2026-11-25 · 第五轮：主题与放映设置

### 新增采样（第 7 层）
`rev/acquire_theme.py`：theme1.xml 的 clrScheme/fontScheme、PowerPoint 改主题色后的差异、
改放映设置后的 presProps.xml。

关键结论（写进 `schema/ooxml_structure.json` 的 `theme_and_show`）：
1. 12 个配色键里只有 dk1/lt1 是 `sysClr+lastClr`，其余 10 个是 `srgbClr@val`。
2. 中文要同时改 `<a:ea>` 与 `<a:font script="Hans">`。
3. **放映设置不在 presentation.xml，而在 ppt/presProps.xml 的 `<p:showPr>`**；
   循环是 `loop="1"`，展台是 `<p:kiosk/><p:sldAll/>`。
4. COM 的 `ThemeColorScheme.Colors(n)`：1=dk1 2=lt1 3=dk2 4=lt2 5=accent1…，RGB 仍是 BGR。

### 新增能力
- `mcp_ppt/theme.py`：`apply_theme`（改 clrScheme/fontScheme）、`extract_theme`（从模板提取）、
  `set_show` / `read_show`（新建或改写 presProps.xml）。
- DSL 新增 `theme.colors` / `theme.fonts` / `theme_from` / `show`；引擎自绘色也跟随 `accent1/accent2` 与字体。
- MCP 工具 +3：`set_theme` / `set_show` / `read_theme`（共 25）。

### 修复
1. python-pptx 里 theme 部件是普通 `Part`（blob）而非 `XmlPart`，`_element` 为 `None` 导致改主题静默无效；
   改为"解析 blob → 改 → 写回 `_blob`"，同时兼容 XmlPart。
2. `presProps.xml` 在默认模板里不存在，需新建 Part 并 `relate_to(RT.PRES_PROPS)`。

### 验证
- `scripts/test_theme.py`：5 个配色键 COM 回读与写入一致（BGR 还原后）、`LoopUntilStopped=-1`、展台 `ShowType=3`、
  `validate` 0 problem、`layout_check` 0 溢出。
- 工具层：`set_theme` / `set_show` / `read_theme` 串联，readback 含 `useTimings="0"`。
- `scripts/run_all.py`：15 步全绿。

---

## 2026-11-25 · 第六轮：Morph 关键帧流程（用多页元素变换做复杂动画）

### 新增采样（第 8 层）
`rev/acquire_morph.py`：3 页 `!!hero`/`!!side` 逐帧位移+放大+变色，页间 `p159:morph`，`advTm=1500`。
验证结论：连续 morph 被接受（`EntryEffect=3954`）、`!!` 前缀名称原样保留、`advTm` 与 `p14:dur` 共存。

### 新增能力
- `mcp_ppt/morph.py`：`plan`（关键帧 → 时间轴）、`build_morph_deck`（N 页序列 + morph + advTm）、
  `decompose`（元素拆解 + 跨页编码识别）、`assign_codes`（批量 `!!` 编码）。
- MCP 工具 +4：`decompose_deck` / `assign_codes` / `plan_morph` / `build_morph`（共 29）。
- 文档 `docs/MORPH_WORKFLOW.md`：四步流程、关键帧模型、时间轴公式、与对象动画的分工。

### 验证
- `scripts/test_morph.py`：4 帧 / 6.4s，4/4 页 `EntryEffect=3954`，元素位置逐帧变化，`validate` 0 problem。
- `scripts/morph_codecraft_demo.py`：拆出 5 个章节页 → 生成 4 帧章节转场（01→02 缩放+变色+文字变形）。
- `scripts/run_all.py`：16 步全绿。

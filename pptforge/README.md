# pptforge —— 让 AI 生成的 PPT 自带切换与动画

把 PowerPoint 的演示结构逆向拆开，编译成一张"效果目录"，再用 MCP 暴露给任意 LLM 客户端。
生成结果是**原生 .pptx**：文字可编辑，切换与对象动画在放映时真的会动。

```
逆向层  rev/            真机采样 → schema/transitions.json + schema/animations.json
目录层  schema/         150 切换效果 · 150 动画预设 · 结构化 OOXML 地图
引擎层  mcp_ppt/        JSON 模型 → pptx（python-pptx 画内容 + lxml 写 transition/timing）
协议层  mcp_ppt/server.py  15 个 MCP 工具（stdio）
```

## 为什么需要它

`python-pptx` 的对象模型里没有幻灯片切换，也没有对象动画——它保存出来的 pptx 能打开，
但放映时没有过渡、段落不会逐条出现。切换与动画必须直接写 OOXML：

- 切换是 `<p:sld>` 下的 `<p:transition>`（或 `mc:AlternateContent` 包裹的扩展效果）
- 动画是同一页的 `<p:timing>` 时间轴（`tmRoot → mainSeq → click 组 → 效果 par → 行为节点`）

两套 XML 都有隐式规则（元素顺序、命名空间前缀、`p14:dur` 的容器要求、`cTn@id` 的唯一性、
文本分级要用 `p:txEl/p:pRg`）。这些规则不写在任何 API 里，我们把它们从 PowerPoint 自己
写出的文件里采样出来，编译成目录，再由引擎按目录拼装。

## 30 秒上手

```bash
cd E:\pptx\pptforge
python scripts/test_tools.py            # 生成 samples/pptforge-demo.pptx 并自检
```

MCP 客户端接线（Claude Desktop / Cursor / 自研 Agent 均适用）：

```json
{
  "mcpServers": {
    "pptforge": {
      "command": "python",
      "args": ["-m", "mcp_ppt.server"],
      "cwd": "E:\\pptx\\pptforge",
      "env": { "PYTHONIOENCODING": "utf-8" }
    }
  }
}
```

最省事的调用顺序（LLM 侧）：

1. `deck_model_template()` → 拿到 JSON 模型的说明与样板
2. `catalog_transitions("涡流")` / `catalog_animations("飞入")` → 挑效果名（中英文都认）
3. `build_deck_from_outline(...)` 或 `build_deck(model, out_path, motion_style="cinematic")`
4. `validate()` / `layout_check()` → 结构 + 版式自检
5. `render_png()` / `export_pdf()` → 出图

## MCP 工具

| 工具 | 作用 |
| --- | --- |
| `capability_report` | 效果目录规模、本机 PowerPoint COM 是否可用 |
| `catalog_transitions` / `catalog_animations` | 列出可用切换 / 动画（id、中文名、OOXML 元素） |
| `motion_styles` | 4 套动效风格的具体配置 |
| `deck_model_template` | deck JSON 模型说明 + 可改写样板 |
| `build_deck` | 按模型生成整套 pptx |
| `build_deck_from_outline` | 最简大纲直接出片（标题 + 每节要点） |
| `set_transition` | 给已有 pptx 的指定页设置切换 |
| `add_animations` | 重写某一页的对象动画（不动内容），目标可以是组 |
| `add_elements` | 在某页新增形状/文本框/图片，可顺带组合 |
| `edit_elements` | 批量形状操作：定位、移动、缩放、属性、组合、解组、对齐、复制、删除 |
| `set_text` | 重写文本框内容与格式（段内多 run 混排、项目符号、行距） |
| `add_table` | 新增表格（表头样式、列宽行高、单元格样式） |
| `edit_table` | 编辑表格：单元格 / 合并 / 表样式 / 列宽 / 行高 |
| `add_chart` | 新增原生图表（柱/条/线/面积/饼/环/雷达/散点，数据内嵌工作簿） |
| `set_theme` | 改 theme1.xml 的配色方案与字体方案 |
| `set_show` | 写放映设置（循环/展台/旁白/计时） |
| `read_theme` | 读出主题配色/字体与放映设置（照着模板做） |
| `decompose_deck` | 拆解每页元素 + 列出跨页同编码元素（morph 可插值的对象） |
| `assign_codes` | 把元素改名为 `!!code`，让 morph 强制按名称匹配 |
| `plan_morph` | 只算关键帧时间轴：每帧起止时间与元素状态 |
| `build_morph` | 关键帧 → N 页 morph 序列（自动换片） |
| `element_info` | 读回形状几何/文本/旋转翻转/组内结构（英寸 + EMU） |
| `inspect` | 读回每页形状 / 切换元素 / 动画 preset 与触发 |
| `validate` | 结构校验（zip、动画形状引用、XML 顺序） |
| `layout_check` | 用 PowerPoint 的 `BoundHeight` 判断文本溢出与越界 |
| `slide_xml` | 导出单页 XML（逆向调试） |
| `render_png` / `export_pdf` | 导图 / 导 PDF（需本机 PowerPoint） |

## deck 模型

```json
{
  "theme": "dark",
  "footer": "公司 · 2026",
  "defaults": { "transition": { "effect": "fade", "duration": 0.7 } },
  "slides": [
    { "layout": "title", "title": "季度复盘", "subtitle": "Q3 结果与下一步",
      "transition": { "effect": "morph", "duration": 1.2 },
      "animations": [
        { "target": "title", "effect": "fade", "trigger": "click", "duration": 0.8 },
        { "target": "subtitle", "effect": "fade", "trigger": "after_previous" } ] },
    { "layout": "bullets", "kicker": "进展", "title": "三件事做成了",
      "bullets": ["留存率提升 12%", { "text": "来自新引导流程", "level": 1 }, "成本下降 8%"],
      "animations": [
        { "target": "title", "effect": "fade", "trigger": "click" },
        { "target": "bullets", "effect": "fly", "trigger": "after_previous",
          "build": "paragraph", "duration": 0.5 } ] },
    { "layout": "metrics", "title": "关键指标",
      "cards": [{ "value": "12%", "label": "留存提升" }, { "value": "8%", "label": "成本下降" }],
      "animations": [{ "target": ["card1", "card2"], "effect": "zoom", "trigger": "click" }] }
  ]
}
```

**版式**：`title` 封面 / `section` 分节 / `bullets` 列表 / `metrics` 指标卡 / `quote` 引用 /
`image` 图文 / `table` 表格 / `chart` 图表 / `blank` 空白。

**动画目标**：形状语义名——`title`、`subtitle`、`kicker`、`bullets`、`card1..cardN`、
`cardValue1..`、`cardLabel1..`、`accent-bar`、`title-rule`、`image`、`table`、`footer`；
也支持 `*`（全部）与 `shape:<id>`。目标不存在时只记 warning，不中断构建。

**触发**：`click`（单击时，默认）/ `with_previous`（与上一动画同时）/ `after_previous`（上一动画之后）。

**分段**：`build: "paragraph"` 让文本框按段落逐条出现（写入 `p:txEl/p:pRg` 与 `build="p"`）。

**其它参数**：`duration`（秒）、`delay`（秒）、`repeat`（次数）、`smooth_start`/`smooth_end`（0–1，
映射到 `accel`/`decel` 百分比）、`exit: true`（退出动画）。

**切换参数**：`effect`、`duration`（秒，写 `p14:dur`）、`speed`（`slow`/`med`/`fast`）、
`advance_after`（秒，自动换片）、`advance_on_click: false`（关闭点击换片）。

**风格**：`motion_style` 取 `subtle` / `bold` / `cinematic` / `tech`，给没写动效的页自动配切换与动画。

## 形状层：定位 · 属性 · 组合 · 移动 · 缩放

模板版式之外的自由排版走 `elements` + `ops`（`build_deck` 里写，或用 `add_elements` / `edit_elements` 改已有文件）。

```json
{ "layout": "blank", "background": "#071426",
  "elements": [
    { "name": "n1", "kind": "ellipse", "x": "18%", "y": "45%", "w": 1.3, "h": 1.3,
      "text": "A", "fill": "#36C2FF", "align": "center", "group": "nodes" },
    { "name": "n2", "kind": "ellipse", "x": "40%", "y": "45%", "w": 1.3, "h": 1.3,
      "text": "B", "fill": "#36C2FF", "align": "center", "group": "nodes" },
    { "name": "wire", "kind": "line", "x": 3.4, "y": 4.3, "w": 1.2, "h": 0,
      "line_color": "#36C2FF", "line_width": 2 } ],
  "ops": [
    { "op": "place", "target": "wire", "x": "18%", "y": "50%", "anchor": "center" },
    { "op": "resize", "target": "group:nodes", "scale": 1.15, "anchor": "center" },
    { "op": "props", "target": "n1", "rot": 8, "alpha": 0.85, "fill": "#FFD166" },
    { "op": "duplicate", "target": "n2", "name": "n3", "dx": 2.4 },
    { "op": "align", "targets": ["n1", "n2", "n3"], "mode": "distribute-h", "to": "slide" },
    { "op": "group", "targets": ["n1", "n2"], "name": "row" } ],
  "animations": [ { "target": "group:row", "effect": "fade", "trigger": "click" } ] }
```

**单位**：数字 = 英寸；字符串支持 `"50%"`（相对幻灯片或 `to` 指定的参照）、`"1.5in"`、`"120pt"`、`"3cm"`、`"914400emu"`。
**锚点** `anchor`：`left-top`（默认）/ `center` / `right-top` / `left-bottom` / `top` / `bottom` …（resize 时表示"哪个点不动"，place 时表示"用哪个点对齐坐标"）。

**元素种类** `kind`：`rect` / `round-rect` / `ellipse` / `triangle` / `diamond` / `star` / `pentagon` / `chevron` / `arrow` / `line` / `text` / `image` / `table` / `chart`。

**外观属性**（`add_elements` 的元素里写，或 `edit_elements` 的 `{"op":"props", ...}` 里写）：

| 属性 | 取值 |
| --- | --- |
| `fill` | `"#RRGGBB"` / `"none"` / `{type:"solid",color,alpha}` / `{type:"gradient",angle:45,stops:[{pos:0,color:"#4C8DFF"},{pos:100,color:"#6FE3C4"}],style:"linear"\|"radial"}` |
| `alpha` | 0–1（只改透明度时保留原色） |
| `effects` | `{shadow:{blur:14,dist:10,dir:45,color:"#000",alpha:0.5,inner:false}, glow:{radius:14,color,alpha}, soft_edge:12}`；`"clear"` 清空 |
| `line_*` | `line_color` / `line_width`(pt) / `line_dash`(`dash`/`dot`/`lgDash`…) / `line_alpha` / `line_head`/`line_tail`(`triangle`/`arrow`/`stealth`/`diamond`/`oval`) / `join` / `cap` / `line_none` |
| `paragraphs` | `[{align, level, line_spacing, space_before, space_after, bullet:{char,color,size_pct}\|"none", runs:[{text,bold,italic,underline,strike,size,color,alpha,font,font_ea,baseline,spc}]}]` |
| `text` / `lines` / `runs` | 单段快捷写法（`runs` = 单段多 run） |
| `crop` | 图片裁剪比例 `{l,t,r,b}`（0–1） |
| `brightness` / `contrast` | 图片亮度/对比度（-1–1） |
| `rot` / `flip_h` / `flip_v` / `z_order` | 旋转（度）/ 翻转 / 层级（`front`/`back`/`forward`/`backward`） |

**ops 清单**：`place`（重新定位）、`move`（dx/dy 或 to_x/to_y）、`resize`（scale 或 w/h，可指定固定锚点）、
`props`（属性编辑，含上表全部键）、`group` / `ungroup`、`align`（`left`/`right`/`center-h`/`top`/`bottom`/`center-v`/`center`/`distribute-h`/`distribute-v`，`to` 可为形状名或 `slide`）、
`duplicate`（复制，组会连同子形状重编号）、`delete`。每条独立执行，失败只记 warning 并继续。

**组合的 OOXML 语义**（从真机采样得到，见 [REVERSE_REPORT](docs/REVERSE_REPORT.md#4-形状层pgrpsp)）：

| 操作 | XML 变化 | 子形状 |
| --- | --- | --- |
| 组合 | 生成 `<p:grpSp>`：`off = chOff = 包围盒左上`，`ext = chExt = 包围盒尺寸` | 坐标不动 |
| 组移动 | 只改 `a:off` | 不动（视觉随组平移） |
| 组缩放 | 只改 `a:ext`（`chExt` 不变） | 不动（视觉按 `ext/chExt` 缩放） |
| 解组 | 按 `off + (child_off − chOff) × (ext/chExt)` 换算回父坐标，组的 `rot`/`flip` 折算到每个子形状 | 重新落位 |
| 组动画 | 仍是 `<p:spTgt spid="组id"/>` | 组内单形状动画用子形状 id |

## 数据层：表格与图表

```json
{ "layout": "chart", "title": "季度收支",
  "chart": { "type": "column", "categories": ["Q1","Q2","Q3","Q4"],
             "series": [{ "name": "收入", "values": [12,18,15,22], "color": "#4C8DFF" },
                        { "name": "成本", "values": [8,9.5,11,12.5], "color": "#6FE3C4" }],
             "legend": "bottom", "data_labels": true, "number_format": "0.0" },
  "animations": [{ "target": "title", "effect": "fade", "trigger": "click" },
                 { "target": "chart", "effect": "wipe", "trigger": "after_previous" }] }

{ "layout": "table", "title": "分渠道表现",
  "rows": [["渠道","Q3","Q4","变化"], ["直营","18.0","22.0","+22%"], ["分销","9.5","12.5","+32%"]],
  "header": true, "col_widths": [2.6,2.2,2.2,2.2], "font_size": 15,
  "ops": [{ "op": "table", "target": "table",
            "cells": [{ "row": 1, "col": 3, "text": "+22%", "color": "#6FE3C4" }] },
          { "op": "table", "target": "table", "merge": { "row1": 3, "col1": 0, "row2": 3, "col2": 3 } },
          { "op": "table", "target": "table", "style": { "first_row": true, "band_row": true } }] }
```

**图表** `type`：`column` / `column-stacked` / `column-stacked-100` / `bar` / `bar-stacked` / `line` /
`line-markers` / `area` / `area-stacked` / `pie` / `doughnut` / `radar` / `scatter`。
参数：`title`、`legend`(`bottom`/`right`/`top`/`left`/`none`)、`data_labels`、`number_format`、
`gridlines`、`gap_width`、`overlap`、`colors`（逐点配色，饼图用）；每个系列可给 `color` 或 `colors`。
数据写进内嵌工作簿（`ppt/embeddings/Microsoft_Excel_SheetN.xlsx`），PowerPoint 里能"编辑数据"。

**表格**：`rows`（二维数组）、`header`、`col_widths` / `row_heights`（英寸）、`font_size`、
`cell` / `header_cell` / `body_cell`（`fill`/`color`/`size`/`bold`/`align`/`valign`/`borders`/`margins`）、
`style{first_row,band_row,style_id}`。
`borders` 支持 `left/right/top/bottom/diag_down/diag_up`，值为 `{"color","width","dash"}` 或 `"none"`。
单元格级编辑走 `ops: [{"op":"table", "cells":[…], "merge":{…}, "style":{…}, "widths":[…], "heights":[…]}]`。

**动画**：图表和表格都是 `graphicFrame`，动画目标直接写名字或 `shape:<id>`。
图表的"按系列/按类别"分级动画在 COM 上不可用（详见边界），只做整体动画。

## 主题与放映设置

```json
{ "theme": { "preset": "dark", "name": "Corporate",
             "colors": { "accent1": "#FF7A45", "accent2": "#6FE3C4", "dk2": "#1B2A4A", "hlink": "#36C2FF" },
             "fonts":  { "major": "Microsoft YaHei", "minor": "Microsoft YaHei", "hans": "Microsoft YaHei" } },
  "show":  { "loop": true, "show_type": "kiosk", "narration": false, "timings": true } }
```

`theme` 会同时写两处：`ppt/theme/theme1.xml`（`clrScheme` + `fontScheme`，影响所有引用 `schemeClr` 的对象）
与引擎自绘的颜色（`accent1 → accent`、`accent2 → accent_2`、`major/minor → font_title/font_body`），保证画面与主题一致。
`theme_from: "template.pptx"` 可直接把已有模板的配色与字体搬过来（`read_theme` 先看再搬）。

`show` 写进 `ppt/presProps.xml` 的 `showPr`（**不是** `presentation.xml`）：
`loop` 循环、`show_type`（`speaker`/`browse`/`kiosk` 展台）、`narration` 旁白、`animation`、`timings`、`scrollbar`。

## Morph 关键帧流程（复杂动画）

对象动画只能表达"一条时间轴上的进出强调"。要做"元素从 A 漂到 B、途中放大变色、再归位"这种连续叙事，
用 **多页元素变换 + `morph` 平滑切换 + `advTm` 自动换片**：每帧画成一页，页间自动插值。

四步（详见 [docs/MORPH_WORKFLOW.md](docs/MORPH_WORKFLOW.md)）：

| 步骤 | 工具 | 产物 |
| --- | --- | --- |
| ① 读取与拆解 | `decompose_deck(path)` | 每页元素清单（id/名称/位置/文本）+ 跨页同编码元素 |
| ② 元素编码 | `assign_codes(path, {"标题 5": "hero"})` | 同一元素跨页统一叫 `!!hero`（morph 强制匹配） |
| ③ 每页 plan | `plan_morph(model)` | 时间轴：每帧起止时间 + 各元素状态（只算不写） |
| ④ 具体设计 | `build_morph(model, out)` | N 页 morph 序列，每页 `advTm` 自动换片 |

```json
{ "frame_duration": [2.0, 1.2, 1.2],
  "transition": { "effect": "morph", "duration": 1.2 },
  "elements": [
    { "name": "num", "kind": "text", "text": "01", "size": 220, "color": "#4C8DFF",
      "keyframes": [
        { "x": 4.4, "y": 1.6, "w": 4.5, "h": 3.6 },
        { "x": 0.9, "y": 0.7, "w": 1.5, "h": 1.2, "size": 72, "color": "#6FE3C4" },
        { "x": 0.9, "y": 0.7, "w": 1.5, "h": 1.2, "text": "02", "size": 72, "color": "#FFD166" } ] } ] }
```

`keyframes` 就是**起点 / 中间点（0…n 个）/ 终点**；省略的字段沿用上一帧。

## 逆向结论（要点）

| 层 | 事实 |
| --- | --- |
| 包 | 页序来自 `presentation.xml` 的 `p:sldIdLst`，经 rels 映射到 `slides/slideN.xml`；文件编号 ≠ 页序 |
| 页 | 子元素顺序固定：`cSld → clrMapOvr → transition|AlternateContent → timing → extLst` |
| 切换 | 效果是 `p:transition` 的**子元素**（`<p:fade thruBlk="1"/>`、`<p14:vortex dir="l"/>`），不是属性 |
| 切换 | `p14`/`p15`/`p159` 效果必须包 `mc:AlternateContent/mc:Choice(Requires=前缀)` + `mc:Fallback`（`p:fade`） |
| 切换 | 精确时长 `p14:dur`（毫秒）是扩展属性，一旦使用**也**必须走 `mc:AlternateContent` |
| 切换 | `spd="slow|med|fast"` ↔ COM `Speed=1/2/3`；`fast` 是默认值，PowerPoint 直接省略 |
| 动画 | 骨架 `tmRoot(1) → seq → mainSeq(2) → click 组 → 效果容器 → 效果 par(presetID/Class/Subtype) → 行为节点` |
| 动画 | 触发写在 `nodeType`：`clickEffect` / `withEffect` / `afterEffect`；`with`/`after` 追加进当前 click 组 |
| 动画 | 时长落在行为节点 `cTn@dur`（毫秒）；`<p:set>` 里的 `dur="1"` 是瞬时标记，不可改 |
| 动画 | `p:cTn@id` 全页唯一，注入时必须整树重编号 |
| 动画 | 文本分级：每段一个 click 组，`p:spTgt/p:txEl/p:pRg st,end`，并在 `p:bldLst` 写 `build="p"` |
| 动画 | 动作路径 `p:animMotion`，`path` 用相对坐标（0..1 相对形状），`pathEditMode="relative"` |
| 形状 | 几何在 `a:xfrm`：`a:off`/`a:ext`（EMU），`rot` 度×60000，`flipH`/`flipV="1"`，透明度 `a:alpha` 百分比×1000 |
| 组合 | `p:grpSp` 的 `a:xfrm` 同时有 `off/ext`（父坐标）与 `chOff/chExt`（子坐标），组合时 1:1；移动只改 `off`、缩放只改 `ext` |
| 填充 | 纯色 `a:solidFill`，渐变 `a:gradFill`（`a:gs@pos` 为 1/1000 %、`a:lin@ang` 为 1/60000 度），图案 `a:pattFill`，无 `a:noFill` |
| 效果 | `a:effectLst` 子元素顺序固定（blur → glow → innerShdw → outerShdw → reflection → softEdge）；`blurRad/rad = 磅×12700` |
| 文本 | run 用 `a:rPr@sz`（1/100 pt）、`a:latin`+`a:ea` 分字体；段落 `a:pPr` 的 `lnSpc@spcPct`（%×1000）、`spcBef/spcAft@spcPts`（1/100 pt） |
| 图片 | `p:blipFill`（p 命名空间）里的 `a:srcRect` 用比例×100000；COM 的 `CropXxx` 单位是磅，两者不能混 |
| 单位 | `dur`/`delay`/`advTm` 毫秒；`repeatCount` 次数×1000；`accel`/`decel` 百分比×1000；COM 的 `RGB` 是 BGR 顺序 |

完整骨架与全部 150+150 条效果见 [`schema/ooxml_structure.json`](schema/ooxml_structure.json)、
[`schema/transitions.json`](schema/transitions.json)、[`schema/animations.json`](schema/animations.json)，
过程记录见 [`docs/REVERSE_REPORT.md`](docs/REVERSE_REPORT.md)。

## 验证

| 检查 | 结果 |
| --- | --- |
| 验收矩阵（15 种切换 + 9 种动画场景，真机回读 `EntryEffect`/`EffectType`/`TriggerType`） | 24/24 一致 |
| 形状层（组合/移动/缩放/属性/对齐，真机回读几何与 `GroupItems`） | 通过，组 `groupItems=3`，组动画被识别 |
| 绘图属性（11 个单变量文件逐个用 PowerPoint 打开 + 完整 deck 回读） | 11/11 可打开；渐变 `fill.Type=3`、阴影 `Shadow.Visible=-1`、裁剪读回一致 |
| 数据层（表格/图表版式 + 单元格编辑 + 图表动画） | 表格 Ops 全 ok、`chart1.xml` + 内嵌 xlsx 齐全、图表动画回读 `(22,3,'chart')` |
| 主题与放映（theme1.xml / presProps.xml + 真机回读） | 5 个配色键回读一致（BGR 已还原）、`loop=-1`、展台 `type=3` |
| Morph 关键帧（4 帧序列 + 真机回读） | 4/4 页 `EntryEffect=3954`、元素位置逐帧变化、`validate` 0 problem |
| 布局层（elements + ops + 组动画挂载） | 9 条 ops 全 ok，组动画 `(23, 3, 'row')` |
| 生成件结构校验（`validate`） | 0 problem（动画形状引用、XML 顺序） |
| 版式体检（`layout_check`，PowerPoint `BoundHeight`） | 0 溢出 |
| MCP stdio 端到端（initialize → tools/list → tools/call） | 19 工具可用，生成件 0 problem |
| 渲染 | `render_png` 导出成功，逐页 `ink_ratio` 0.04–0.41（页面非空白） |

复跑：

```bash
python scripts/run_all.py         # 全量回归（16 步）
python scripts/matrix_test.py     # 验收矩阵（需本机 PowerPoint）
python scripts/test_shapes.py     # 形状层：组合/移动/缩放/属性
python scripts/test_layout.py     # 布局层：elements + ops + 组动画
python scripts/test_styles.py     # 绘图属性：渐变/效果/富文本/裁剪
python scripts/test_data.py       # 数据层：表格 / 图表 / 图表动画
python scripts/test_theme.py      # 主题配色/字体 + 放映设置
python scripts/test_morph.py      # morph 关键帧序列
python scripts/diag_styles.py     # 属性单变量二分（判断哪个写法会让 PowerPoint 打不开）
python scripts/test_stdio.py      # MCP 协议端到端
python rev/build_catalog.py       # 从采样重编译效果目录
```

## 目录

```
pptforge/
├─ mcp_ppt/            引擎与 MCP 服务
│  ├─ catalog.py       效果目录索引与别名解析（中英文）
│  ├─ ooxml.py         transition / timing 写入与读回
│  ├─ shapes.py        形状层：定位/移动/缩放/组合/属性/对齐/复制
│  ├─ dml.py           DrawingML 属性：填充/线条/效果/富文本/图片裁剪
│  ├─ table.py         表格：单元格/边框/合并/表样式/行高列宽
│  ├─ chart.py         图表：原生 chart + 数据内嵌工作簿
│  ├─ theme.py         主题（theme1.xml）+ 放映设置（presProps.xml）
│  ├─ morph.py         morph 关键帧：拆解/编码/时间轴/多页序列
│  ├─ deck.py          JSON 模型 → pptx；patch_deck 只改动效与形状
│  ├─ design.py        版式与设计系统（4 套主题）
│  ├─ motion.py        动效风格套餐
│  ├─ inspect.py       结构读取与校验
│  ├─ layout.py        PowerPoint 版式体检 / 像素统计
│  ├─ comtools.py      COM：回读、导图、导 PDF（可选）
│  └─ server.py        MCP server（stdio）
├─ schema/             逆向产物：transitions / animations / ooxml_structure
├─ rev/                采样与编译脚本、原始证据（out/）
├─ scripts/            冒烟、矩阵、stdio、工具联调
├─ samples/            demo 输出
├─ config/mcp.json     MCP 接线示例
```

## 环境

- Python 3.10+（本机 3.13.7）
- `python-pptx>=1.0.2`、`lxml>=5`、`mcp>=1.16`、`anyio>=4.5`（mcp 1.16 在 anyio 3.x 下启动会失败）
- 可选：`pywin32` + 本机 PowerPoint——用于真机回读、导图/导 PDF、版式体检；没有也能生成

## 边界

- 没采样的动画类型：媒体播放类（`MediaPlay/Pause/Stop`，需媒体形状）、图示分级（`msoAnimateDiagram*`）、
  "单击对象时"的交互序列（`InteractiveSequences`）。这三种在 `catalog` 里没有条目，调用会明确报"未知效果"。
- 图表**分级动画**（按系列/按类别）在 COM 的 `AddEffect` 上一律 `Invalid request`（7/8/10 三档实测全失败），
  引擎只做整体动画；表格边框支持四边 + 两条对角线，没有单元格 3D 与斜线阴影。
- `add_animations` 是**重写整页动画**语义，不做增量追加。
- 生成件依赖 PowerPoint 2010+ 的扩展命名空间；WPS/Keynote 对 `p14/p15/p159` 效果可能回退成淡出（同 `mc:Fallback` 的行为）。
- `layout_check` / `render_png` / `export_pdf` 需要本机 PowerPoint；纯结构生成不需要。
- 形状层的已知取舍：组合时包围盒按未旋转的 `off/ext` 计算（旋转形状的视觉外框会略大）；
  `ungroup` 只展开一层（嵌套组需要多次调用），组的旋转/翻转会折算到子形状的 `rot`/`flip`；
  `line` 用 `a:cxnSp` 直线，`w/h` 表示终点相对起点的位移（负值会自动交换端点）；
  `duplicate` 对组会递归重编号子形状 id，但动画不会自动复制。


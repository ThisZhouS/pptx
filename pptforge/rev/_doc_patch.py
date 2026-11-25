"""一次性文档补丁：把数据层内容写进 README / REVERSE_REPORT / CHANGELOG。"""
import io
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def patch(path: str, pairs: list[tuple[str, str]]) -> None:
    full = os.path.join(ROOT, path)
    with io.open(full, encoding="utf-8") as fh:
        text = fh.read()
    for old, new in pairs:
        if old not in text:
            print(f"  [miss] {path}: {old[:50]!r}")
            continue
        text = text.replace(old, new)
    with io.open(full, "w", encoding="utf-8") as fh:
        fh.write(text)
    print(f"  [ok] {path}")


DATA_SECTION = """## 数据层：表格与图表

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

## 逆向结论（要点）"""

patch("README.md", [
    ("`table` 表格 / `blank` 空白。", "`table` 表格 / `chart` 图表 / `blank` 空白。"),
    ("`arrow` / `line` / `text` / `image`。", "`arrow` / `line` / `text` / `image` / `table` / `chart`。"),
    ("| `set_text` | 重写文本框内容与格式（段内多 run 混排、项目符号、行距） |",
     "| `set_text` | 重写文本框内容与格式（段内多 run 混排、项目符号、行距） |\n"
     "| `add_table` | 新增表格（表头样式、列宽行高、单元格样式） |\n"
     "| `edit_table` | 编辑表格：单元格 / 合并 / 表样式 / 列宽 / 行高 |\n"
     "| `add_chart` | 新增原生图表（柱/条/线/面积/饼/环/雷达/散点，数据内嵌工作簿） |"),
    ("| 绘图属性（11 个单变量文件逐个用 PowerPoint 打开 + 完整 deck 回读） | 11/11 可打开；渐变 `fill.Type=3`、阴影 `Shadow.Visible=-1`、裁剪读回一致 |",
     "| 绘图属性（11 个单变量文件逐个用 PowerPoint 打开 + 完整 deck 回读） | 11/11 可打开；渐变 `fill.Type=3`、阴影 `Shadow.Visible=-1`、裁剪读回一致 |\n"
     "| 数据层（表格/图表版式 + 单元格编辑 + 图表动画） | 表格 Ops 全 ok、`chart1.xml` + 内嵌 xlsx 齐全、图表动画回读 `(22,3,'chart')` |"),
    ("│  ├─ dml.py           DrawingML 属性：填充/线条/效果/富文本/图片裁剪",
     "│  ├─ dml.py           DrawingML 属性：填充/线条/效果/富文本/图片裁剪\n"
     "│  ├─ table.py         表格：单元格/边框/合并/表样式/行高列宽\n"
     "│  ├─ chart.py         图表：原生 chart + 数据内嵌工作簿"),
    ("python scripts/test_styles.py     # 绘图属性：渐变/效果/富文本/裁剪",
     "python scripts/test_styles.py     # 绘图属性：渐变/效果/富文本/裁剪\n"
     "python scripts/test_data.py       # 数据层：表格 / 图表 / 图表动画"),
    ("## 逆向结论（要点）", DATA_SECTION),
    ("- `duplicate` 对组会递归重编号子形状 id，但动画不会自动复制。",
     "- `duplicate` 对组会递归重编号子形状 id，但动画不会自动复制。\n"
     "- 图表**分级动画**（按系列/按类别）在 COM 的 `AddEffect` 上一律 `Invalid request`（7/8/10 三档实测全失败），\n"
     "  引擎只做整体动画；表格边框支持四边 + 两条对角线，没有单元格 3D 与斜线阴影。"),
])

patch("docs/CHANGELOG.md", [
    ("- `scripts/run_all.py`：11 步全绿。",
     "- `scripts/run_all.py`：11 步全绿。\n\n"
     "---\n\n"
     "## 2026-11-24 · 第四轮：数据层（表格单元格 + 原生图表）\n\n"
     "### 新增采样（第 6 层）\n"
     "`rev/acquire_table_chart.py` + `rev/probe_chart_anim.py`：4 页真机样本，覆盖\n"
     "表格合并/列宽行高/单元格填充/四边与对角边框/表样式，以及图表部件与图表动画的组合探查。\n\n"
     "关键结论（写进 `schema/ooxml_structure.json` 的 `shape_layer.style.table/chart`）：\n"
     "1. 表格合并：主单元格 `gridSpan=\"2\"`，被吞的单元格保留 `a:tc` 并写 `hMerge=\"1\"`。\n"
     "2. 边框在 `a:tcPr` 内，顺序 lnL → lnR → lnT → lnB → lnTlToBr → lnBlToTr → 填充。\n"
     "3. 列宽 `a:gridCol@w`、行高 `a:tr@h`；表样式是 `a:tblPr/a:tableStyleId{GUID}`。\n"
     "4. 图表 = `p:graphicFrame` + `ppt/charts/chartN.xml` + `ppt/embeddings/*.xlsx`。\n"
     "5. **图表分级动画不可用**：`msoAnimateChartBySeries/Category/AllAtOnce` 在 `AddEffect` 上全部 `Invalid request`。\n\n"
     "### 新增能力\n"
     "- `mcp_ppt/table.py`：`build_table` / `set_cell` / `set_borders` / `merge` / `style_table` /\n"
     "  `column_widths` / `row_heights` / `edit_table` / `table_info`。\n"
     "- `mcp_ppt/chart.py`：`build_chart`（13 种图表 + 配色 + 数据标签 + 轴格式 + 柱间距）/ `chart_info`。\n"
     "- 版式新增 `chart`，`table` 版式重写为走 table 构建器（表头/斑马纹/列宽）。\n"
     "- `add_element` 支持 `kind: table|chart`；`apply_ops` 支持 `op: table`（cells/merge/style/widths/heights）。\n"
     "- MCP 工具 +3：`add_table` / `add_chart` / `edit_table`（共 22）；`element_info` 对表格/图表附带内容摘要。\n\n"
     "### 修复\n"
     "1. **python-pptx 的 `Emu` 在算术后退化成 `int`**（`Emu - int*Emu`），导致 `isinstance(x, Length)` 判断失效，\n"
     "   版式传入的 EMU 被当英寸再乘 914400，表格宽度写成 9.6e12 → PowerPoint 打不开。\n"
     "   新增 `to_emu_len()`（Length / 大整数 / 英寸 / 带单位字符串），table 与 chart 统一走它。\n"
     "2. `apply_ops` 不认 `op: table`，表格编辑 ops 被当作未知操作跳过。\n\n"
     "### 验证\n"
     "- `scripts/diag_data2.py`：6 个用例（纯 python-pptx / 构建器 / 版式）全部能被 PowerPoint 打开。\n"
     "- `scripts/test_data.py`：3 页（2 图表 + 1 表格），ops 全 ok，`validate` 0 problem、`layout_check` 0 溢出，\n"
     "  图表动画回读 `(10,1,'title')`/`(22,3,'chart')`。\n"
     "- `scripts/test_data_tools.py`：`add_table → edit_table → add_chart → element_info` 全通，工具 22 个。"),
])
print("done")
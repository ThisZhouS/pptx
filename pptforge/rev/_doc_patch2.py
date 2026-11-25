"""第二轮文档补丁：README 边界 + REVERSE_REPORT 第 6 章与编号。"""
import io
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load(p):
    with io.open(os.path.join(ROOT, p), encoding="utf-8") as fh:
        return fh.read()


def save(p, text):
    with io.open(os.path.join(ROOT, p), "w", encoding="utf-8") as fh:
        fh.write(text)
    print(f"[ok] {p}")


# ---------------- README 边界
t = load("README.md")
t = t.replace(
    "- 没采样的动画类型：媒体播放类（`MediaPlay/Pause/Stop`，需媒体形状）、图表/图示分级（`msoAnimateChart*`）、\n"
    '  "单击对象时"的交互序列（`InteractiveSequences`）。这三种在 `catalog` 里没有条目，调用会明确报"未知效果"。',
    "- 没采样的动画类型：媒体播放类（`MediaPlay/Pause/Stop`，需媒体形状）、图示分级（`msoAnimateDiagram*`）、\n"
    '  "单击对象时"的交互序列（`InteractiveSequences`）。这三种在 `catalog` 里没有条目，调用会明确报"未知效果"。\n'
    "- 图表**分级动画**（按系列/按类别）在 COM 的 `AddEffect` 上一律 `Invalid request`（7/8/10 三档实测全失败），\n"
    "  引擎只做整体动画；表格边框支持四边 + 两条对角线，没有单元格 3D 与斜线阴影。")
save("README.md", t)

# ---------------- REVERSE_REPORT
p = "docs/REVERSE_REPORT.md"
t = load(p)

chapter6 = """## 6. 表格与图表

### 6.1 表格

```xml
<a:tbl>
  <a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5940675A-…}</a:tableStyleId></a:tblPr>
  <a:tblGrid><a:gridCol w="1778000"/>…</a:tblGrid>
  <a:tr h="846667">
    <a:tc><a:txBody>…</a:txBody><a:tcPr/></a:tc>
    <a:tc hMerge="1"><a:txBody>…</a:txBody><a:tcPr/></a:tc>
    <a:tc><a:txBody>…</a:txBody>
      <a:tcPr>
        <a:lnL w="19050" cmpd="sng"><a:solidFill><a:srgbClr val="457AFF"/></a:solidFill><a:prstDash val="solid"/></a:lnL>
        <a:lnR …/><a:lnT …/><a:lnB …/><a:lnTlToBr …/><a:lnBlToTr …/>
      </a:tcPr>
    </a:tc>
  </a:tr>
</a:tbl>
```

- 合并：主单元格写 `gridSpan="2"`，被吞掉的单元格保留 `a:tc` 并写 `hMerge="1"`（纵向是 `rowSpan` / `vMerge`）。
- `a:tcPr` 的子元素顺序：`lnL → lnR → lnT → lnB → lnTlToBr → lnBlToTr → 填充 → extLst`。
- 列宽 `a:gridCol@w`、行高 `a:tr@h` 都是 EMU；表样式是 `a:tblPr/a:tableStyleId{GUID}`。
- 边框实测：`Cell(2,2).Borders.Item(ppBorderBottom).Weight=1.5` → `<a:lnB w="19050" cmpd="sng">`。

### 6.2 图表

图表不是形状，而是 `p:graphicFrame` 指向独立部件：

```
ppt/charts/chart1.xml                       c:chartSpace / c:chart / c:plotArea / c:barChart…
ppt/charts/_rels/chart1.xml.rels
ppt/embeddings/Microsoft_Excel_Sheet1.xlsx  数据源（"编辑数据"打开的就是它）
```

所以图表必须带内嵌工作簿，否则 PowerPoint 打开会报错或降级成图片。
生成端用 python-pptx 的 `add_chart`（自动写 xlsx），只需补配色、图例、数据标签、轴格式。

### 6.3 图表动画：一个 COM 死路

| 尝试 | 结果 |
| --- | --- |
| `AddEffect(chart, msoAnimEffectFade, msoAnimateLevelNone, click)` | OK（整体动画） |
| `AddEffect(chart, msoAnimEffectFade, msoAnimateChartBySeries=10, click)` | `Invalid request` |
| `AddEffect(chart, 21/2/1, 10)` 各效果组合 | 全部 `Invalid request` |
| `AddEffect(chart, 10, msoAnimateChartByCategory=8 / AllAtOnce=7)` | 全部 `Invalid request` |

结论：这一版 PowerPoint 的 COM 不接受图表分级参数，`Sequence.AddEffect` 只给整体效果。
"按系列出现"只能在 UI 里做，或手写 `p:bldGraphic/bldSub/a:bldChart` —— 后者无真机样本可比对，
不放进目录，避免产出 PowerPoint 会拒绝的文件。

## 7. 目录编译"""

t = t.replace("## 6. 目录编译", chapter6)
t = t.replace("## 7. 闭环验证", "## 8. 闭环验证")
t = t.replace("## 8. 未覆盖", "## 9. 未覆盖")
t = t.replace("## 9. 复现", "## 10. 复现")
t = t.replace("| ⑨ 闭环验证 |", "| ⑩ 闭环验证 |")
t = t.replace("| ⑧ 编译目录 |",
              "| ⑧ 数据层采样 | 表格合并/边框/列宽行高；图表部件与图表动画组合探查 | `rev/acquire_table_chart.py`、`rev/probe_chart_anim.py` | `gt_table_chart.pptx`(3 页)、`gt_table_border.pptx`、`table_chart_map.json` |\n| ⑨ 编译目录 |")
t = t.replace("采样总量：**547 页**（197 + 12 + 154 + 150 + 21 + 7 + 6）",
              "采样总量：**551 页**（197 + 12 + 154 + 150 + 21 + 7 + 6 + 4）")
t = t.replace("python rev/build_structure_map.py        # ⑧b 结构地图",
              "python rev/acquire_table_chart.py       # ⑧ 3 页表格/图表 + 1 页边框\npython rev/build_structure_map.py        # ⑨b 结构地图")
t = t.replace("python scripts/matrix_test.py           # ⑨ 闭环验证",
              "python scripts/matrix_test.py           # ⑩ 闭环验证\npython scripts/test_data.py             # 表格/图表闭环")
t = t.replace("| 组内时序（`p:grpSpTgt` 子目标） | PowerPoint 2024 对组动画直接写组的 `spTgt`，不用 `grpSpTgt` | 与 PowerPoint 行为一致 |",
              "| 组内时序（`p:grpSpTgt` 子目标） | PowerPoint 2024 对组动画直接写组的 `spTgt`，不用 `grpSpTgt` | 与 PowerPoint 行为一致 |\n"
              "| 图表分级动画（`p:bldGraphic/bldSub`） | COM 拒绝该参数，没有真机样本可对照 | 只做整体动画 |")
save(p, t)
print("done")
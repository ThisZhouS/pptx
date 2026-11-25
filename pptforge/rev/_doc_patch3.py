"""文档与结构地图补丁：主题层（theme/presProps）。"""
import io
import os

ROOT = r"E:\pptx\pptforge"


def patch(rel, pairs):
    p = os.path.join(ROOT, rel)
    t = io.open(p, encoding="utf-8").read()
    for a, b in pairs:
        if a not in t:
            print(f"  [miss] {rel}: {a[:60]!r}")
            continue
        t = t.replace(a, b)
    io.open(p, "w", encoding="utf-8").write(t)
    print(f"  [ok] {rel}")


# ---------------------------------------------------------------- 结构地图源码
patch("rev/build_structure_map.py", [
    ('"drawing_styles": 6,', '"drawing_styles": 6,\n                    "theme_and_presprops": 3,'),
    ('"图表是 p:graphicFrame',
     '"放映设置不在 presentation.xml，而在 ppt/presProps.xml 的 p:presentationPr/p:showPr。",\n'
     '                "theme1.xml 的 dk1/lt1 用 sysClr+lastClr，其余 10 个键用 srgbClr@val；改主题只需改这些值。",\n'
     '                "图表是 p:graphicFrame'),
    ('''        "evidence": {
            "sampled_pptx"''', '''        "theme_and_show": {
            "theme_part": "ppt/theme/theme1.xml（每个 slideMaster 一份）",
            "clrScheme": {
                "dk1": "<a:dk1><a:sysClr val=\\"windowText\\" lastClr=\\"000000\\"/></a:dk1>",
                "lt1": "<a:lt1><a:sysClr val=\\"window\\" lastClr=\\"FFFFFF\\"/></a:lt1>",
                "others": "dk2 / lt2 / accent1..accent6 / hlink / folHlink 都是 <a:srgbClr val=\\"RRGGBB\\"/>",
                "measured": "Theme.ThemeColorScheme.Colors(5).RGB=0xFF7A45 → accent1 val=\\"FF7A45\\"（COM 索引 5=accent1）",
            },
            "fontScheme": {
                "structure": "<a:majorFont><a:latin typeface=\\"…\\"/><a:ea typeface=\\"…\\"/><a:cs typeface=\\"\\"/><a:font script=\\"Hans\\" typeface=\\"宋体\\"/>…</a:majorFont>",
                "note": "中文走 <a:font script=\\"Hans\\"> 与 <a:ea>，只改 latin 不动中文",
            },
            "presProps": {
                "part": "ppt/presProps.xml（CT_PresentationProperties）",
                "skeleton": "<p:presentationPr><p:showPr loop=\\"1\\"><p:present/><p:sldAll/><p:penClr><a:prstClr val=\\"red\\"/></p:penClr></p:showPr></p:presentationPr>",
                "attrs": {"loop": "循环放映", "showNarration": "0 关旁白", "showAnimation": "0 关动画",
                          "useTimings": "0 不用排练计时", "showScrollbar": "1 显示滚动条"},
                "modes": {"present": "演讲者放映", "browse": "观众自行浏览", "kiosk": "展台浏览", "sldAll": "全部幻灯片"},
                "measured": "LoopUntilStopped=-1 / ShowType=kiosk → loop=\\"1\\" + <p:kiosk/><p:sldAll/>；COM 回读 type=3",
            },
            "com_index": "ppColorSchemeIndex: 1=dk1 2=lt1 3=dk2 4=lt2 5=accent1 6=accent2 …；RGB 仍是 BGR 顺序",
            "evidence": ["rev/out/gt_theme.pptx", "rev/out/theme_map.json", "rev/out/theme_raw/theme1.xml"],
        },
        "evidence": {
            "sampled_pptx"'''),
    ('"rev/out/anim_variants.json", "rev/out/shapes_map.json"]',
     '"rev/out/anim_variants.json", "rev/out/shapes_map.json", "rev/out/theme_map.json"]'),
])

# ---------------------------------------------------------------- run_all
patch("scripts/run_all.py", [
    ('    ("数据层 MCP 工具联调", ["scripts/test_data_tools.py"]),',
     '    ("数据层 MCP 工具联调", ["scripts/test_data_tools.py"]),\n'
     '    ("主题与放映设置", ["scripts/test_theme.py"]),'),
])

# ---------------------------------------------------------------- README
THEME_SECTION = """## 主题与放映设置

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

## 逆向结论（要点）"""

patch("README.md", [
    ("## 逆向结论（要点）", THEME_SECTION),
    ("| `add_chart` | 新增原生图表（柱/条/线/面积/饼/环/雷达/散点，数据内嵌工作簿） |",
     "| `add_chart` | 新增原生图表（柱/条/线/面积/饼/环/雷达/散点，数据内嵌工作簿） |\n"
     "| `set_theme` | 改 theme1.xml 的配色方案与字体方案 |\n"
     "| `set_show` | 写放映设置（循环/展台/旁白/计时） |\n"
     "| `read_theme` | 读出主题配色/字体与放映设置（照着模板做） |"),
    ("| 数据层（表格/图表版式 + 单元格编辑 + 图表动画） | 表格 Ops 全 ok、`chart1.xml` + 内嵌 xlsx 齐全、图表动画回读 `(22,3,'chart')` |",
     "| 数据层（表格/图表版式 + 单元格编辑 + 图表动画） | 表格 Ops 全 ok、`chart1.xml` + 内嵌 xlsx 齐全、图表动画回读 `(22,3,'chart')` |\n"
     "| 主题与放映（theme1.xml / presProps.xml + 真机回读） | 5 个配色键回读一致（BGR 已还原）、`loop=-1`、展台 `type=3` |"),
    ("│  ├─ chart.py         图表：原生 chart + 数据内嵌工作簿",
     "│  ├─ chart.py         图表：原生 chart + 数据内嵌工作簿\n"
     "│  ├─ theme.py         主题（theme1.xml）+ 放映设置（presProps.xml）"),
    ("python scripts/test_data.py       # 数据层：表格 / 图表 / 图表动画",
     "python scripts/test_data.py       # 数据层：表格 / 图表 / 图表动画\n"
     "python scripts/test_theme.py      # 主题配色/字体 + 放映设置"),
    ("python scripts/run_all.py         # 全量回归（14 步）",
     "python scripts/run_all.py         # 全量回归（15 步）"),
])

# ---------------------------------------------------------------- REVERSE_REPORT
CHAPTER7 = """## 7. 主题与放映设置

### 7.1 theme1.xml

```xml
<a:theme><a:themeElements>
  <a:clrScheme name="Office">
    <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
    <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="1F497D"/></a:dk2> … <a:accent1>…<a:accent6>
    <a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="Office">
    <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>
      <a:font script="Hans" typeface="宋体"/>…</a:majorFont>
    <a:minorFont>…</a:minorFont>
  </a:fontScheme>
  <a:fmtScheme>…</a:fmtScheme>
</a:themeElements></a:theme>
```

- 12 个配色键里，只有 `dk1`/`lt1` 是 `sysClr`（`val="windowText"/"window"` + `lastClr`），其余 10 个是 `srgbClr@val`。
- 改主题只需改这些值，`fmtScheme` 与 `fontScheme` 的 script 列表不用动（结构合法）。
- 中文走两层：`<a:ea typeface>` 与 `<a:font script="Hans">` —— 只改 `latin` 不动中文会"看起来没生效"。
- COM 侧 `Theme.ThemeColorScheme.Colors(n)` 的索引是：1=dk1 2=lt1 3=dk2 4=lt2 **5=accent1** 6=accent2 …，
  且 `RGB` 仍是 BGR（设 `0xFF7A45` → XML `val="457AFF"`）。

### 7.2 presProps.xml（放映设置不在 presentation.xml）

```xml
<p:presentationPr xmlns:a="…" xmlns:r="…" xmlns:p="…">
  <p:showPr loop="1" showNarration="0">
    <p:kiosk/><p:sldAll/>
    <p:penClr><a:prstClr val="red"/></p:penClr>
  </p:showPr>
</p:presentationPr>
```

| COM 设置 | XML |
| --- | --- |
| `SlideShowSettings.LoopUntilStopped = msoTrue` | `showPr@loop="1"` |
| `ShowWithNarration = 0` | `showPr@showNarration="0"`（默认就是关，省不省看版本） |
| `AdvanceMode = ppSlideShowUseSlideTimings` | 默认，`useTimings` 省略 |
| `ShowType = ppShowTypeKiosk(3)` | `<p:kiosk/>` + `<p:sldAll/>`（演讲者是 `<p:present/>`，浏览是 `<p:browse/>`） |

实测：写入 `loop + kiosk` 后 PowerPoint 回读 `LoopUntilStopped=-1`、`ShowType=3`，说明手写 `presProps.xml` 被完全接受。
注意这个部件在 python-pptx 默认模板里**不存在**，需要新建 Part 并挂 `presentation → presProps` 关系。

## 8. 目录编译"""

patch("docs/REVERSE_REPORT.md", [
    ("## 7. 目录编译", CHAPTER7),
    ("## 8. 闭环验证", "## 9. 闭环验证"),
    ("## 9. 未覆盖", "## 10. 未覆盖"),
    ("## 10. 复现", "## 11. 复现"),
    ("| ⑩ 闭环验证 |", "| ⑪ 闭环验证 |"),
    ("| ⑨ 编译目录 |",
     "| ⑨ 主题与放映采样 | theme1.xml 配色/字体；presProps.xml 的 showPr 与展示方式 | `rev/acquire_theme.py` | `gt_theme.pptx`、`theme_map.json` |\n| ⑩ 编译目录 |"),
    ("采样总量：**551 页**（197 + 12 + 154 + 150 + 21 + 7 + 6 + 4）",
     "采样总量：**554 页**（197 + 12 + 154 + 150 + 21 + 7 + 6 + 4 + 3）"),
    ("python rev/acquire_table_chart.py       # ⑧ 3 页表格/图表 + 1 页边框",
     "python rev/acquire_table_chart.py       # ⑧ 3 页表格/图表 + 1 页边框\npython rev/acquire_theme.py             # ⑨ 主题部件 + 放映设置"),
    ("python scripts/test_data.py             # 表格/图表闭环",
     "python scripts/test_data.py             # 表格/图表闭环\npython scripts/test_theme.py            # 主题/放映闭环"),
])

# ---------------------------------------------------------------- CHANGELOG
patch("docs/CHANGELOG.md", [
    ("- `scripts/run_all.py`：14 步全绿。",
     "- `scripts/run_all.py`：14 步全绿。\n\n"
     "---\n\n"
     "## 2026-11-25 · 第五轮：主题与放映设置\n\n"
     "### 新增采样（第 7 层）\n"
     "`rev/acquire_theme.py`：theme1.xml 的 clrScheme/fontScheme、PowerPoint 改主题色后的差异、\n"
     "改放映设置后的 presProps.xml。\n\n"
     "关键结论（写进 `schema/ooxml_structure.json` 的 `theme_and_show`）：\n"
     "1. 12 个配色键里只有 dk1/lt1 是 `sysClr+lastClr`，其余 10 个是 `srgbClr@val`。\n"
     "2. 中文要同时改 `<a:ea>` 与 `<a:font script=\"Hans\">`。\n"
     "3. **放映设置不在 presentation.xml，而在 ppt/presProps.xml 的 `<p:showPr>`**；\n"
     "   循环是 `loop=\"1\"`，展台是 `<p:kiosk/><p:sldAll/>`。\n"
     "4. COM 的 `ThemeColorScheme.Colors(n)`：1=dk1 2=lt1 3=dk2 4=lt2 5=accent1…，RGB 仍是 BGR。\n\n"
     "### 新增能力\n"
     "- `mcp_ppt/theme.py`：`apply_theme`（改 clrScheme/fontScheme）、`extract_theme`（从模板提取）、\n"
     "  `set_show` / `read_show`（新建或改写 presProps.xml）。\n"
     "- DSL 新增 `theme.colors` / `theme.fonts` / `theme_from` / `show`；引擎自绘色也跟随 `accent1/accent2` 与字体。\n"
     "- MCP 工具 +3：`set_theme` / `set_show` / `read_theme`（共 25）。\n\n"
     "### 修复\n"
     "1. python-pptx 里 theme 部件是普通 `Part`（blob）而非 `XmlPart`，`_element` 为 `None` 导致改主题静默无效；\n"
     "   改为\"解析 blob → 改 → 写回 `_blob`\"，同时兼容 XmlPart。\n"
     "2. `presProps.xml` 在默认模板里不存在，需新建 Part 并 `relate_to(RT.PRES_PROPS)`。\n\n"
     "### 验证\n"
     "- `scripts/test_theme.py`：5 个配色键 COM 回读与写入一致（BGR 还原后）、`LoopUntilStopped=-1`、展台 `ShowType=3`、\n"
     "  `validate` 0 problem、`layout_check` 0 溢出。\n"
     "- 工具层：`set_theme` / `set_show` / `read_theme` 串联，readback 含 `useTimings=\"0\"`。\n"
     "- `scripts/run_all.py`：15 步全绿。"),
])
print("done")
"""文档补丁：morph 关键帧流程。"""
import io
import os

ROOT = r"E:\pptx\pptforge"


def patch(rel, pairs):
    p = os.path.join(ROOT, rel)
    t = io.open(p, encoding="utf-8").read()
    for a, b in pairs:
        if a not in t:
            print(f"  [miss] {rel}: {a[:50]!r}")
            continue
        t = t.replace(a, b)
    io.open(p, "w", encoding="utf-8").write(t)
    print(f"  [ok] {rel}")


MORPH_SECTION = """## Morph 关键帧流程（复杂动画）

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

## 逆向结论（要点）"""

patch("README.md", [
    ("## 逆向结论（要点）", MORPH_SECTION),
    ("| `read_theme` | 读出主题配色/字体与放映设置（照着模板做） |",
     "| `read_theme` | 读出主题配色/字体与放映设置（照着模板做） |\n"
     "| `decompose_deck` | 拆解每页元素 + 列出跨页同编码元素（morph 可插值的对象） |\n"
     "| `assign_codes` | 把元素改名为 `!!code`，让 morph 强制按名称匹配 |\n"
     "| `plan_morph` | 只算关键帧时间轴：每帧起止时间与元素状态 |\n"
     "| `build_morph` | 关键帧 → N 页 morph 序列（自动换片） |"),
    ("| 主题与放映（theme1.xml / presProps.xml + 真机回读） | 5 个配色键回读一致（BGR 已还原）、`loop=-1`、展台 `type=3` |",
     "| 主题与放映（theme1.xml / presProps.xml + 真机回读） | 5 个配色键回读一致（BGR 已还原）、`loop=-1`、展台 `type=3` |\n"
     "| Morph 关键帧（4 帧序列 + 真机回读） | 4/4 页 `EntryEffect=3954`、元素位置逐帧变化、`validate` 0 problem |"),
    ("│  ├─ theme.py         主题（theme1.xml）+ 放映设置（presProps.xml）",
     "│  ├─ theme.py         主题（theme1.xml）+ 放映设置（presProps.xml）\n"
     "│  ├─ morph.py         morph 关键帧：拆解/编码/时间轴/多页序列"),
    ("python scripts/test_theme.py      # 主题配色/字体 + 放映设置",
     "python scripts/test_theme.py      # 主题配色/字体 + 放映设置\n"
     "python scripts/test_morph.py      # morph 关键帧序列"),
    ("python scripts/run_all.py         # 全量回归（15 步）",
     "python scripts/run_all.py         # 全量回归（16 步）"),
])

patch("scripts/run_all.py", [
    ('    ("主题与放映设置", ["scripts/test_theme.py"]),',
     '    ("主题与放映设置", ["scripts/test_theme.py"]),\n'
     '    ("Morph 关键帧序列", ["scripts/test_morph.py"]),'),
])

patch("docs/CHANGELOG.md", [
    ("- `scripts/run_all.py`：15 步全绿。",
     "- `scripts/run_all.py`：15 步全绿。\n\n"
     "---\n\n"
     "## 2026-11-25 · 第六轮：Morph 关键帧流程（用多页元素变换做复杂动画）\n\n"
     "### 新增采样（第 8 层）\n"
     "`rev/acquire_morph.py`：3 页 `!!hero`/`!!side` 逐帧位移+放大+变色，页间 `p159:morph`，`advTm=1500`。\n"
     "验证结论：连续 morph 被接受（`EntryEffect=3954`）、`!!` 前缀名称原样保留、`advTm` 与 `p14:dur` 共存。\n\n"
     "### 新增能力\n"
     "- `mcp_ppt/morph.py`：`plan`（关键帧 → 时间轴）、`build_morph_deck`（N 页序列 + morph + advTm）、\n"
     "  `decompose`（元素拆解 + 跨页编码识别）、`assign_codes`（批量 `!!` 编码）。\n"
     "- MCP 工具 +4：`decompose_deck` / `assign_codes` / `plan_morph` / `build_morph`（共 29）。\n"
     "- 文档 `docs/MORPH_WORKFLOW.md`：四步流程、关键帧模型、时间轴公式、与对象动画的分工。\n\n"
     "### 验证\n"
     "- `scripts/test_morph.py`：4 帧 / 6.4s，4/4 页 `EntryEffect=3954`，元素位置逐帧变化，`validate` 0 problem。\n"
     "- `scripts/morph_codecraft_demo.py`：拆出 5 个章节页 → 生成 4 帧章节转场（01→02 缩放+变色+文字变形）。\n"
     "- `scripts/run_all.py`：16 步全绿。"),
])
print("done")
"""追加"高动画模板时序配方"一节，并修正"默认不加动画"的结论。"""
import io

P = r"E:\pptx\pptforge\docs\TEMPLATE_LEARNING.md"
EXTRA = """

## 八、高动画模板的时序配方（按需求聚焦"要动画"）

上一节说"高端模板 0 动画"，但既然目标是**要动画**，就该看有动画的那批。
筛出动画密度 ≥4 条/页的 **246 个模板**，取前 14 个拆时序（`scripts/analyze_timing.py`，数据 `samples/timing_recipes.json`）。

| 模板 | 动画/页 | click 组 | with | after | 组内条数 | 最大 delay | 预设组合 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 黑白灰动态 03 | **34.4** | 0 | 30.2 | 0 | 30.2 | **3000ms** | fade+wipe+peek+fly+circle |
| 黑白灰动态 25 | 29.0 | 0 | 31.4 | 2.9 | 34.3 | 1500ms | fade+faded-zoom+fly+descend |
| 高级极简 48 | 21.0 | 0 | 0 | 13.4 | 13.4 | 0 | fly+… |
| 高级极简 45 | 20.8 | 0 | 0 | 16.5 | 16.5 | 0 | wipe+faded-zoom+fly+fade+ascend |
| 欧美简约 35 | 19.6 | 1.3 | 2.3 | 16.7 | 17.1 | 100ms | fade+wipe+faded-zoom+dissolve |
| 欧美简约 5 | 18.5 | 0.8 | 16.8 | 0 | 16.9 | 0 | **custom(路径) ×202** |
| 高级极简 43 | 16.6 | 0 | 13.6 | 0.4 | 14.0 | **7000ms** | wipe+fly+draw+descend |
| 中国风 13 | 16.4 | 1.0 | 14.0 | 0 | 15.0 | 0 | ascend+wipe+dissolve+… |
| 日系和风 19 | 16.3 | 0 | 5.8 | 10.8 | 16.6 | 500ms | wipe+faded-zoom+peek+… |
| 马卡龙 3 | 15.9 | 1.7 | 3.2 | 13.0 | 15.8 | 0 | faded-zoom+fly+fade+draw |
| 黑白灰动态 15 | 15.2 | 0 | 13.8 | 0.5 | 14.2 | 2750ms | fly+circle+faded-zoom+wipe |
| 日系和风 18 | 15.1 | 0 | 3.1 | 12.6 | 15.7 | 0 | faded-zoom+wipe+ascend+draw |

### 四条可复用的规则

1. **几乎不用 click 分组**：click 组平均 **0–1.7 个**，其余全是 `withEffect`（同一条时间轴自动连播）。
   观众不需要点 20 次才能看完一页。
2. **靠 delay 拉瀑布**：最大 delay 从 0 到 **7000ms**（黑白灰动态 03 是 3000ms、高级极简 43 是 7000ms）。
   一页的动画可以持续 3–7 秒，节奏完全由 delay 控制，而不是靠 "after previous" 串行。
3. **一页混用 4–6 种预设**：fade / wipe / fly / ascend / faded-zoom / peek / custom-path 交替，
   而不是全篇一种效果。欧美简约 5 号甚至 202 次全用**自定义路径**。
4. **动画条数 ≈ 形状数**：34.4 条/页对应 20–30 个形状，即"每个元素都动"。

### 已落到引擎

`motion.py` 的 `subtle` 套餐按这套配方重写（此前是"每页 3–5 条、1 个 click、全部 stagger 0.1"的保守版）：

| 页型 | 编排 |
| --- | --- |
| 封面 | 装饰条 wipe → 标题 **fly** → 副标题 **ascend** → meta fade → 细线 wipe（1 click + 4 with，delay 递增到 0.25） |
| 内容 | kicker fade → 标题 **fly** → 细线 **wipe** → 正文 **ascend** + `build: paragraph`（逐段浮出）→ 页脚 fade |
| 指标 | kicker → 标题 → 细线 → 卡片 **faded-zoom** → 卡内数值 fade → 卡内标签 fade（**16–20 条/页**） |

`codecraft-v3.pptx` 重建后：**162 个动画 / 19 页**（原先 118），
预设混用 fade 58 + wipe 42 + ascend 39 + faded-zoom 13 + fly 10，
触发 withEffect 143 / click 19 —— 与模板配方同一数量级。

### 仍差的

模板每页 20–30 个形状，我的版式只有 5–20 个，所以动画条数上限低（内容页 8 条 vs 模板 16 条）。
要追平，得先补形状密度（装饰层），再谈动画条数。
"""
with io.open(P, encoding="utf-8") as fh:
    text = fh.read()
if "八、高动画模板的时序配方" not in text:
    # 修正原第六节的措辞
    text = text.replace(
        "1. **默认不加动画**：新增 `motion_style: \"none\"` 档，对应\"零动画\"那一批高端模板——版式做扎实，动效留白。",
        "1. ~~默认不加动画~~ → 按需求改为**默认走高动画配方**（见第八节）：每页全元素动画 + 1 个 click + delay 瀑布。")
    with io.open(P, "w", encoding="utf-8") as fh:
        fh.write(text + EXTRA)
    print("appended")
else:
    print("already present")
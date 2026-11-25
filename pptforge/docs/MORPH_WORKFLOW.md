# Morph 关键帧流程（用多页元素变换做复杂动画）

对象动画（`p:timing`）只适合"一条时间轴上的进出强调"，做不了"元素从 A 漂到 B、途中放大变色、
再缩小归位"这种连续叙事。这种效果用 **morph 平滑切换 + 多页元素变换 + 自动换片** 更自然：

```
页1（起点） --morph 1.2s--> 页2（中间点1） --morph--> 页3（中间点2） --> 页4（终点）
        └── 每页 advTm 1.6s 自动换片，整段不点一下就播完 ──┘
```

PowerPoint 的 morph 按"对象名称"匹配：相邻两页同名对象会被插值（位置、尺寸、旋转、填充、文字），
名称以 `!!` 开头时强制按名称匹配。所以同一逻辑元素在每页都叫 `!!hero`，动画就发生在它自己身上。

## 四步流程

| 步骤 | 做什么 | 工具 |
| --- | --- | --- |
| ① 读取与拆解 | 逐页读元素：id、名称、标签、位置尺寸、文本 | `decompose_deck(path)` |
| ② 编码 | 把"同一元素"在跨页统一命名成 `!!code`；放大/缩小的同一元素也共用这个码 | `assign_codes(path, {"标题 5": "hero"})` |
| ③ 每页 plan | 列出这一页要动哪些元素、用哪种过渡、每帧停留多久；泛用 `morph`，用多页 + 自动换片串成时间轴 | `plan_morph(model)`（只算不写） |
| ④ 具体设计 | 每个元素写 起点 / 中间点(0…n 个) / 终点，再落盘 | `build_morph(model, out_path)` |

## 模型

```json
{
  "theme": "dark",
  "frame_duration": 1.6,
  "transition": { "effect": "morph", "duration": 1.2 },
  "base": {
    "elements": [
      { "name": "title", "kind": "text", "x": 0.9, "y": 0.6, "w": 8, "h": 0.7,
        "text": "关键帧叙事", "size": 26, "color": "#EAF0FF" }
    ]
  },
  "elements": [
    { "name": "hero", "kind": "round-rect", "fill": "#4C8DFF", "text": "HERO",
      "keyframes": [
        { "x": 1.0, "y": 1.6, "w": 2.4, "h": 1.2 },
        { "x": 4.6, "y": 3.0, "w": 4.6, "h": 2.2, "fill": "#6FE3C4", "rot": 8 },
        { "x": 8.6, "y": 1.4, "w": 3.4, "h": 1.6, "fill": "#FFD166", "rot": -6 },
        { "x": 5.2, "y": 4.6, "w": 2.2, "h": 2.2, "alpha": 0.75 }
      ] }
  ]
}
```

- **`keyframes` 就是"起点 / 中间点… / 终点"**：第 1 个是起点、最后一个是终点，中间写几个都行，0 个也行。
- 每帧只写变化的部分，**省略的字段沿用上一帧**（上面第 4 帧没写 `fill`，就保持第 3 帧的黄色）。
- 可动的属性：`x` `y` `w` `h`（英寸或 `"50%"`）、`rot`、`fill`、`alpha`、`text`、`size`、`color`、`lines`。
- `base.elements` 是每帧都在的静态元素（标题、页码），它们也会带 `!!` 前缀，保证 morph 时"不闪"。
- `frame_duration` 可以是数字（每帧一样）或数组（逐帧不同，比如首帧停 2s、其他 1s）。
- 页数 = 最长关键帧链的长度，短的元素在后续帧保持最后一帧状态。

## 时间轴

```
T(第 i 帧) = Σ frame_duration[0..i-1]
总时长     = Σ frame_duration
每页 advTm = frame_duration[i] × 1000（毫秒）
```

`plan_morph` 会直接算给你：每帧的 `start` / `duration` / 各元素状态。放映时把
`SlideShowSettings.AdvanceMode` 设为"使用排练计时"（`set_show` 的 `timings: true`），
整段就会自动播放。

## 实际效果示例

`samples/morph-sequence.pptx`（4 帧 / 6.4s）：

| 帧 | 起始 | hero（方块） | dot（圆） |
| --- | --- | --- | --- |
| 1 | 0.0s | (1.0, 1.6) 2.4×1.2 蓝 | (10.6, 1.6) 1.1 青 |
| 2 | 1.6s | (4.6, 3.0) 4.6×2.2 绿 rot 8° | (9.2, 4.2) 1.8 橙 |
| 3 | 3.2s | (8.6, 1.4) 3.4×1.6 黄 rot -6° | (2.2, 4.2) 0.9 半透明 |
| 4 | 4.8s | (5.2, 4.6) 2.2×2.2 紫 75% | (11.2, 5.2) 1.4 绿 |

四页的切换都是 `p159:morph`（PowerPoint 回读 `EntryEffect=3954`），元素名 `!!hero` / `!!dot`
四页一致 —— 这是插值能发生的前提。

## 用在已有 PPT 上

对 CodeCraft AI 那份 19 页：

```python
decompose_deck("CodeCraft AI.pptx")          # ① 看每页有什么
assign_codes("CodeCraft AI.pptx", {"标题 5": "hero"})   # ② 给要变形的元素编码
plan_morph({...})                             # ③ 先审时间轴
build_morph({...}, "out.pptx")                # ④ 落盘
```

判断标准：

- **适合**：章节转场（大号编号 01→02 缩放移位）、单品讲解（产品图从满屏缩到角落、正文跟进）、
  数据强调（柱状图里某一根抽出来放大）。
- **不适合**：内容页本身信息密度高、每页版式不同——这些页没有共同元素，morph 退化成普通淡入。

所以更稳的做法是：静态内容页保留普通淡入（`fade`），**只在"叙事段落"处插入一个 2–4 页的
morph 小序列**，讲完再接回普通页。这样既不破坏原版式，又能拿到"高级感"。

## 与对象动画的分工

| 需求 | 用什么 |
| --- | --- |
| 逐条出现的要点 | 对象动画 + `build: "paragraph"`（`add_animations`） |
| 元素沿路径移动 / 连续变形 / 多段叙事 | morph 关键帧序列（`build_morph`） |
| 页面之间的转场 | 切换（`set_transition`，`morph` / `fade` / `cover-left`…） |
| 自动播放整段 | `show: {loop, timings}` + 每页 `advTm` |

morph 序列里**不要再放对象动画**：两者叠加会让时间轴对不上（morph 是页级过渡，对象动画是页内时间轴）。
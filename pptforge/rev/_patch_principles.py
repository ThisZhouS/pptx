"""把"复制页做 morph"的落地说明补进设计原则文档。"""
import io

P = r"E:\pptx\pptforge\docs\DESIGN_PRINCIPLES.md"
EXTRA = """

## 六、代码里怎么落实（`scripts/morph_copy_mode.py`）

正确做法已经实现为脚本，三步：

1. **编码**：先把参与 morph 的元素改名为 `!!num` / `!!title`（对应官方 `!!` 强制匹配约定）。
2. **复制**：用 PowerPoint 的 `Slide.Duplicate()` 复制原页（不是 `add_slide(layout)`），
   然后只改元素几何 —— 字体、bodyPr、占位符样式、层级全部继承，morph 只插值一个维度。
3. **排序与切换**：复制出来的帧默认排在原页之后，用 `Slide.MoveTo()` 移到原页之前，
   形成 `起点 → 中间点 → 终点（原页）`；首帧用 `fade 800`，之后两帧用 `morph 2000`；
   同时清空这几页的页内对象动画（morph 是页级过渡，叠加会打架）。

实测（`samples/morph-copy-demo.pptx`，基于 CodeCraft 第 3 页）：

| 帧 | 页码 | `!!num` 位置 | 切换 |
| --- | --- | --- | --- |
| 起点 | 3 | (0.40, **3.98**) 2.68×2.32 | fade 800 |
| 中间点 | 4 | (0.40, **3.28**) 2.68×2.32 | morph 2000 |
| 终点 | 5 | 原位（占位符继承 layout，2.57） | morph 2000 |

三帧**尺寸完全一致、只改 y**，所以插值是一条竖直上浮线 —— 不会再出现"斜穿页面"的错位感。

## 七、两条硬规则（来自这次排查）

1. **morph 帧必须由复制产生**，不要新建同版式页 —— 后者会因 `bodyPr` 继承差异造成文字错位。
2. **morph 帧里不要放对象动画** —— 页级过渡与页内时间轴会互相打断，观感是"动一下卡住"。
"""
with io.open(P, encoding="utf-8") as fh:
    text = fh.read()
if "六、代码里怎么落实" not in text:
    with io.open(P, "w", encoding="utf-8") as fh:
        fh.write(text + EXTRA)
    print("appended")
else:
    print("already present")
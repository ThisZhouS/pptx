"""把 morph 章节改造记录追加进 CodeCraft 报告。"""
import io

P = r"E:\pptx\pptforge\docs\CODECRAFT_OPTIMIZATION.md"
EXTRA = """

---

## 六、追加：章节页改造为 morph 转场序列

产物：`samples/codecraft-morph-chapters.pptx`（27 页）

按四步流程（见 `docs/MORPH_WORKFLOW.md`）把 4 个章节页各扩成 3 帧：

| 步骤 | 动作 |
| --- | --- |
| ① 拆解 | `decompose_deck` 找出 4 个章节页（第 3/6/12/16 页，含 `01`–`04` 编号），再用 COM 读出编号与标题的实际几何（编号 100pt、标题 48pt） |
| ② 编码 | 章节页的编号/标题改名 `!!num` / `!!title`；新帧用同一 layout 的占位符、同名 → morph 可插值 |
| ③ plan | 每章 3 帧：起点（编号居中放大 `(3.9,1.5) 5.6×4.2`，标题在下方 `y=5.9`）→ 中间点（编号缩到左上 `(0.9,0.55) 1.6×1.3`，标题上移 `y=0.7`）→ 终点（原章节页） |
| ④ 编写 | 同 layout 新建两帧插到章节页之前；三页都设 `morph` 1.2s + `advTm`（帧 1.4s / 章节页 2.0s）；章节页清空页内对象动画 |

最终页序（27 页）：

```
1 封面 · 2 目录 · 3-5 章节01 三帧 · 6-7 内容 · 8-10 章节02 三帧
· 11-15 内容 · 16-18 章节03 三帧 · 19-21 内容 · 22-24 章节04 三帧 · 25-27 内容/结尾
```

验证：

| 检查 | 结果 |
| --- | --- |
| 12 页 morph（4 章 × 3 帧） | COM 回读 `EntryEffect=3954`，其余页仍是 `fade` / `fade-smoothly` |
| 元素编码 | `!!num` / `!!title` 出现在 3/4/5、8/9/10、16/17/18、22/23/24 各三页 |
| 内容页对象动画 | 101 个动画原样保留（只清空了章节页的） |
| `validate` | 27 页 / 27 页有切换 / 0 problem |
| 渲染 | 27 页 PNG 正常；章节帧背景与章节页一致（同 layout 继承） |

踩到的坑：帧的最终页码不能用"插入那一刻的索引"记录 —— 从后往前插时，后续插入会让已处理章节的页码继续右移。
正确算法是 `最终页码 = 原页码 + 2×(排在前面的章节数) + 2`（自己在前面插的两帧）。
第一版用中间态页码，结果把 morph 打到了内容页上（虽然内容页动画没丢，但有两页章节帧没被设上）。
"""

with io.open(P, encoding="utf-8") as fh:
    text = fh.read()
if "六、追加：章节页改造" not in text:
    text += EXTRA
    with io.open(P, "w", encoding="utf-8") as fh:
        fh.write(text)
    print("appended")
else:
    print("already present")
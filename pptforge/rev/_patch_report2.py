"""把"依据真实模板重新调参"的记录写进 CodeCraft 报告。"""
import io

P = r"E:\pptx\pptforge\docs\CODECRAFT_OPTIMIZATION.md"
EXTRA = """

---

## 七、重调：拿真实设计师模板做参照

上一版效果差的根因不是代码，是**参数靠感觉**。这次下了两个人类设计师做的带动画模板做基准
（`samples/templates/`，来自 [Creative Venus](https://creativevenus.com/)，免费可商用、免注册直链）：

| 模板 | 页数 | 切换 | 对象动画 |
| --- | --- | --- | --- |
| `blue-glow-ai-platform` | 5 | `p:fade` | 146 个 |
| `movement-elements` | 10 | `p159:morph` × 10，**`p14:dur=2000`** | **0 个** |

### 模板给出的三条硬参数

1. **进入效果几乎所有都用 `presetID 42`「上升」(Ascend)** —— 146 个动画里 124 个是它。
   不是原地淡入，是"从下方轻微上浮 + 淡入"。
2. **时长 800ms 是主流**（372 个 800ms，600–700ms 只有 19 个）。我原来用 400–600ms，偏快。
3. **触发用 `withEffect` 占 141/146** —— 靠 **delay 错峰**做瀑布，而不是 `after_previous` 一个等一个。
   我原来是 61 个 afterEffect 串行，节奏发"顿"。

模板的作者原话（页面正文）：每页动效预算 ≈3 秒、单次 0.5 秒够、只给值得的页加、
"十页连续运动会让观众等着而不是看着"。

### 改了什么

| 项 | 旧 | 新 |
| --- | --- | --- |
| 进入效果 | `fade` 64 / `zoom` 24 / `wipe` 13 | **`ascend` 90（100%）** |
| 时长 | 500ms×112 / 600ms×60 / 400ms×54 | **800ms × 90** |
| 触发 | `afterEffect` 61（串行） | **`withEffect` 78 + stagger 0.12s** |
| 章节切换 | `morph` 1.2s | **`morph` 2.0s**（对齐模板的 2000ms） |
| 章节帧位移 | 从屏幕中央飞向左上（跨度大） | 原地**上浮 1.4 英寸**，尺寸不变 |

引擎侧改动：`motion.py` 四个套餐重写（`ascend` + 0.8s + `stagger`），
`deck._anim_items` 支持 `stagger` 累积 delay，`morph.py` 默认 `duration 2.0` / `frame_duration 2.5`。

### 顺带修掉的两个真 bug

1. `patch_deck` 里 `if patch.get("animations")` —— 空列表是 falsy，**传 `[]` 根本清不掉动画**。
   改成 `if "animations" in patch`。
2. 即使进了分支，也写成 `set_slide_timing(...) if items else {"items": 0}` —— 空列表时**直接短路返回**，
   旧的 `p:timing` 没被移除。改成无条件调用（函数内部本来就会先移除旧 timing）。

如果没有第 2 条，章节页会同时带着 morph 和页内动画 —— 时间轴打架，正是"看着别扭"的来源之一。

### 现在的参数对照

| 指标 | 设计师模板 | CodeCraft 新版 |
| --- | --- | --- |
| 进入预设 | ascend 85% | ascend 100% |
| 时长 | 800ms | 800ms |
| withEffect 占比 | 97% | 87% |
| morph 时长 | 2000ms | 2000ms |

产物：

| 文件 | 页数 | 动效 |
| --- | --- | --- |
| `codecraft-optimized-v2.pptx` | 19 | 全部 `fade 800` + `ascend 800` 错峰（每页 1.3s 打完） |
| `codecraft-morph-chapters-v2.pptx` → 覆盖为 `codecraft-morph-chapters.pptx` | 27 | 12 页 `morph 2000` 章节序列 + 15 页 `fade 800` |

`validate`：19 页 / 98 个动画 / 0 problem；27 页 / 90 个动画（章节页清空）/ 0 problem。
"""
with io.open(P, encoding="utf-8") as fh:
    text = fh.read()
if "七、重调" not in text:
    with io.open(P, "w", encoding="utf-8") as fh:
        fh.write(text + EXTRA)
    print("appended")
else:
    print("already present")
"""把「全量 2120」与动画分档写进 SKILL.md 和 README。"""
import io
import os

SK = r"E:\pptx\pptx-skill\SKILL.md"
RD = r"E:\pptx\README.md"

# ---------- SKILL.md
t = io.open(SK, encoding="utf-8").read()
t = t.replace("从 `Resource/` 里 **2125 个真实模板**拆出来的素材库 + 生成引擎。",
              "从 `Resource/` 里 **2120 个真实模板（全量，不再抽样）** 拆出来的素材库 + 生成引擎。")
t = t.replace("| `styles.json` | 6 种风格 × 配色/动效/版式偏好 + 任务映射 | 62 板块画像 |",
              "| `styles.json` | 6 种风格 × 配色/动效/版式偏好 + 任务映射 | 62 板块 × 2120 文件 |")
t = t.replace("| `motion.json` | 时长分布、预设配比、click/with 规则、delay 瀑布、14 个模板的实测配方 | 594 个模板的动画统计 |",
              "| `motion.json` | 时长分布、预设配比、click/with 规则、delay 瀑布、动画分档、14 个模板的实测配方 | **2120 个模板** |")
t = t.replace("| `decor.json` | 装饰元素规格（chip/dot/rule/page-num）+ 62 板块画像 | 444 个内容页 |",
              "| `decor.json` | 装饰元素规格（chip/dot/rule/page-num）+ 62 板块画像 | 444 个内容页 + 2120 文件 |")
t = t.replace("| `layouts.json` | 边距、覆盖率、对齐线、形状密度、列数、图文关系、字号层级、栅格 | 444 个内容页 + 15407 个形状 |",
              "| `layouts.json` | 边距、覆盖率、对齐线、形状密度分布（P10/50/90）、列数、图文关系、字号层级、栅格 | 444 内容页 + 15407 形状 + 2120 文件 |")

# 插入动画分档小节
anchor = "## 生成参数（内置在引擎里）"
if "## 动画分档（全量数据）" not in t:
    block = '''## 动画分档（全量 2120 个模板实测）

| 档位 | 动画/页 | 文件数 | 占比 | 对应 recipe |
| --- | --- | --- | --- | --- |
| zero | 0 | 792 | **37%** | `none` |
| low | 0–3 | 367 | 17% | `subtle` |
| mid | 3–8 | 618 | 29% | `rich` |
| high | 8–15 | 285 | 13% | `rich` |
| extreme | 15+ | 58 | 3% | `rich` |

**全量的动画中位数只有 1.8 条/页**（P90 是 9.8）。也就是说「大部分模板几乎不动」，
真正高频动画的只占 16%。风格表里的 `极简/文艺` = zero 档，`商务/国潮` = low 档，
`科技/沉稳深色` = mid/high 档 —— 覆盖从 0 到 18 条/页的全区间。

'''
    t = t.replace(anchor, block + anchor, 1)

io.open(SK, "w", encoding="utf-8").write(t)
print("[ok] SKILL.md")

# ---------- README
r = io.open(RD, encoding="utf-8").read()
r = r.replace("从 2125 个真实 PPT 模板里拆出**可复用的设计参数**",
              "从 **2120 个真实 PPT 模板（全量）** 里拆出**可复用的设计参数**")
r = r.replace("| 动效时长 | 500ms 占 54238 次（主流），1000ms 30178 |",
              "| 动效时长 | 500ms 占 165962 次（主流），1000ms 114903 |")
r = r.replace("| 进入预设 | fly 10436 / wipe 7918 / ascend 6715 / fade 5880 / faded-zoom 4778 |",
              "| 进入预设 | fly 29522 / wipe 26761 / ascend 25371 / fade 20220 / faded-zoom 14508 |")
r = r.replace("| 时长与密度 | 高动画模板 15–34 条/页，动画条数 ≈ 形状数 |",
              "| 动画分档 | 零动画 37% / 低 17% / 中 29% / 高 13% / 极高 3%；全量中位 1.8 条/页 |")
r = r.replace("| 切换 | none 5380 页最多；morph 只用于图片推镜/轮播 |",
              "| 切换 | none 16644 页最多（35%）；morph 1584 页，集中在图片轮播 |")
io.open(RD, "w", encoding="utf-8").write(r)
print("[ok] README.md")
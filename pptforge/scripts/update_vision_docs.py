"""把视觉画像写进 SKILL.md / README，并最终提交。"""
import io
import os

SK = r"E:\pptx\pptx-skill\SKILL.md"
RD = r"E:\pptx\README.md"

t = io.open(SK, encoding="utf-8").read()
anchor = "## 配色（全量 2120 个主题实测）"
if "## 视觉画像（296 个首页渲染实测）" not in t:
    block = '''## 视觉画像（296 个首页渲染实测）

从 2120 个模板均匀抽 300 个，渲染首页后按像素统计分类：

| 风格标签 | 占比 |
| --- | --- |
| **低饱和（莫兰迪/灰调）** | **46.6%** |
| 信息密集 | 22.3% |
| 图片/色块主导 | 10.5% |
| 高饱和（马卡龙/潮流） | 6.4% |
| 深色极简 | 5.7% |
| 均衡型 | 4.4% |
| 浅色留白 | **4.1%** |

视觉指纹中位：`ink_ratio` 0.575（信息密度中高）、`brightness` 0.796（偏亮）、
`saturation` **0.167（很低）**、`edge_density` 0.141。

两个反直觉的结论：

1. **主流不是"大留白"** —— 真正的"浅色留白"只占 4.1%，近一半是"低饱和 + 中等密度"。
   `极简` 风格是这个 skill 提供的选项，不是模板的主流。
2. **整体饱和度极低（0.167）但 accent 饱和度很高（0.58）** —— 彩色只占很小面积。
   这解释了为什么"大面积白 + 小面积强调色"看起来最像专业模板。

'''
    t = t.replace(anchor, block + anchor, 1)
io.open(SK, "w", encoding="utf-8").write(t)
print("[ok] SKILL.md 视觉画像")

r = io.open(RD, encoding="utf-8").read()
if "视觉画像" not in r:
    r = r.replace("| 栅格 | 按 0.25/0.5/1.0 英寸对齐，**不用 12 列栅格** |",
                  "| 栅格 | 按 0.25/0.5/1.0 英寸对齐，**不用 12 列栅格** |\n"
                  "| 视觉画像 | 低饱和 46.6%、信息密集 22.3%、浅色留白仅 4.1%；整体饱和度中位 0.167 |")
io.open(RD, "w", encoding="utf-8").write(r)
print("[ok] README 视觉画像行")
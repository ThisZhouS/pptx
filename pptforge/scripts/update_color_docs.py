"""把全量配色结论写进 SKILL.md 与 README。"""
import io

SK = r"E:\pptx\pptx-skill\SKILL.md"
RD = r"E:\pptx\README.md"

t = io.open(SK, encoding="utf-8").read()
anchor = "## 生成参数（内置在引擎里）"
if "## 配色（全量 2120 个主题实测）" not in t:
    block = '''## 配色（全量 2120 个主题实测）

| 项 | 数据 |
| --- | --- |
| 深/浅底 | **浅底 2116 / 深底 4** —— 几乎全是白底 |
| accent 色相 | **蓝 81.7%**（1733）、红橙 10%、青 5.7%、绿 1.9%、紫 0.6% |
| accent 饱和度/明度 | 蓝 0.58/0.77、青 0.64/0.62、红橙 0.50/0.80 |
| accent 与背景对比 | 中位 **4.03**（蓝 4.03 / 红橙 3.99 / 青 3.39） |

注意：模板自身的 accent 对比多数**低于 WCAG AA 的 4.5**（只适合大标题）。
生成端用 `accent_from_hsv(min_contrast=4.5)` 往下压明度，得到比模板略深、但正文也可用的强调色。

'''
    t = t.replace(anchor, block + anchor, 1)
io.open(SK, "w", encoding="utf-8").write(t)
print("[ok] SKILL.md 配色段")

r = io.open(RD, encoding="utf-8").read()
r = r.replace("| 配色 | 91/91 浅底黑字，accent 与背景对比 4.7（WCAG AA），蓝色占 69% |",
              "| 配色 | 浅底 2116/2120，蓝色占 **81.7%**，accent 与背景对比中位 4.03 |")
io.open(RD, "w", encoding="utf-8").write(r)
print("[ok] README.md 配色行")
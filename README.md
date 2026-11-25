# pptx

从 **2120 个真实 PPT 模板（全量）** 里拆出**可复用的设计参数**，再用引擎从零生成对标水准的 .pptx。

```
pptx-skill/     ← agent skill：按任务类型与风格生成 pptx（本仓库主产出）
pptforge/       ← 逆向与生成引擎：OOXML 结构、动效、版式、校验工具链
Resource/       ← 原始模板库（2125 个 / 47GB，不入库）
```

## pptx-skill（先看这个）

按任务类型与风格选素材，从大纲生成 pptx：

```bash
cd pptx-skill
python scripts/build.py --demo --out demo.pptx --style 科技
python scripts/build.py my-outline.json --out deck.pptx --task 汇报
```

- 6 种风格：极简 / 商务 / 科技 / 国潮 / 文艺 / 沉稳深色（`assets/styles.json`）
- 素材库：配色（91 主题）、版式参数（444 页）、动效配方（594 模板 + 14 个时序样本）、装饰规格
- 详细用法见 [`pptx-skill/SKILL.md`](pptx-skill/SKILL.md)

## pptforge（引擎与逆向）

把 PowerPoint 的 OOXML 拆成可编程能力，全部结论来自真机采样（PowerPoint 16 自己写出的 XML）：

| 层 | 能力 |
| --- | --- |
| 切换 | 150 种效果（含 p14/p15/p159 扩展与 mc:AlternateContent 规则） |
| 动画 | 150 个预设、三种触发、按段落分级、动作路径、退出动画 |
| 形状 | 定位/移动/缩放/组合/解组/对齐/属性编辑 |
| 绘图 | 纯色/渐变/无填充、阴影/发光/柔化、虚线箭头、富文本、图片裁剪 |
| 数据 | 表格（单元格/边框/合并/表样式）、图表（13 种 + 内嵌工作簿） |
| 主题 | theme1.xml 配色与字体、presProps.xml 放映设置 |
| morph | 关键帧序列（复制页生成帧，避免元素错位） |

文档：[逆向报告](pptforge/docs/REVERSE_REPORT.md) ·
[模板拆解](pptforge/docs/TEMPLATE_LEARNING.md) ·
[布局与风格](pptforge/docs/LAYOUT_STYLE.md) ·
[UI 布局](pptforge/docs/UI_LAYOUT.md) ·
[从零对标](pptforge/docs/FROM_ZERO.md) ·
[设计原则](pptforge/docs/DESIGN_PRINCIPLES.md)

## 关键数据（来自 2125 个模板的抽样拆解）

| 维度 | 结论 |
| --- | --- |
| 动效时长 | 500ms 占 165962 次（主流），1000ms 114903 |
| 进入预设 | fly 29522 / wipe 26761 / ascend 25371 / fade 20220 / faded-zoom 14508 |
| 触发方式 | 每页 0–1.7 个 click 组，其余 `with_previous` + delay 瀑布 |
| 动画分档 | 零动画 37% / 低 17% / 中 29% / 高 13% / 极高 3%；全量中位 1.8 条/页 |
| 切换 | none 16644 页最多（35%）；morph 1584 页，集中在图片轮播 |
| 边距 | 上 0.16in / 左右 0.48in / 下 0.60in（上紧下松） |
| 字号 | 11 档，最大字 = 正文 5 倍 |
| 配色 | 91/91 浅底黑字，accent 与背景对比 4.7（WCAG AA），蓝色占 69% |
| 栅格 | 按 0.25/0.5/1.0 英寸对齐，**不用 12 列栅格** |

## 环境

Python 3.10+、`python-pptx>=1.0.2`、`lxml`。可选 `pywin32` + 本机 PowerPoint：
用于真机回读校验、导图/导 PDF、`layout_check` 版式体检。

## 复现全部拆解

见 [`pptforge/docs/FROM_ZERO.md`](pptforge/docs/FROM_ZERO.md) 第六节的八条命令。
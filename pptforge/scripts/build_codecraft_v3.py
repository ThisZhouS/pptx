"""重头设计 CodeCraft 演示：用引擎重建版式与动效（不再对原件打补丁）。

设计约束（见 docs/DESIGN_PRINCIPLES.md）：
  · 对齐：全篇同一左基线（0.92in），标题与正文共享
  · 间距：同级等距，层级之间用 2 倍间距
  · 层级：全篇三档字号（标题 34 / 要点 20 / 说明 15）
  · 留白：四边等宽，内容不贴边
动效：全篇 fade 800 切换；每页一组 ascend 800 错峰入场（不加强调、不用 morph 空转页）
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_ppt import comtools
from mcp_ppt.deck import build_deck
from mcp_ppt.inspect import validate_deck
from mcp_ppt.layout import png_stats

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples", "codecraft-v3.pptx")

FOOT = "CodeCraft AI · 2025"

MODEL = {
    "theme": {
        "preset": "midnight",
        "colors": {"accent1": "#36C2FF", "accent2": "#FFD166"},
        "fonts": {"major": "Microsoft YaHei", "minor": "Microsoft YaHei"},
    },
    "footer": FOOT,
    "defaults": {"transition": {"effect": "fade", "duration": 0.8}},
    "slides": [
        # 1 封面
        {"layout": "title",
         "title": "CodeCraft AI",
         "subtitle": "基于大语言模型的自主软件架构优化与重构引擎",
         "meta": ["2025 / 10 / 30", "汇报人 · Redamancy"],
         "title_size": 66},

        # 2 目录
        {"layout": "metrics", "kicker": "CONTENTS", "title": "四个部分",
         "cards": [{"value": "01", "label": "核心理念与痛点"},
                   {"value": "02", "label": "核心技术栈与产品"},
                   {"value": "03", "label": "创新性与壁垒"},
                   {"value": "04", "label": "商业模式"}]},

        # 3 章节 01
        {"layout": "section", "index": 1, "title": "核心理念与痛点",
         "subtitle": "从「代码补全」到「系统级医生」"},

        # 4 理念
        {"layout": "bullets", "kicker": "01 · 理念", "title": "与现有工具的定位差异",
         "bullets": [
             {"text": "系统级医生，而非代码补全", "level": 0},
             {"text": "Copilot 类工具补全单行代码；CodeCraft AI 把整个仓库当作一个病人，先诊断再开方。", "level": 1},
             {"text": "深度分析代码仓库", "level": 0},
             {"text": "不只读单个文件，而是理解架构分层、数据流走向与依赖关系，为重构提供全局依据。", "level": 1},
             {"text": "输出可直接合并的重构方案", "level": 0},
             {"text": "自主识别代码坏味、性能瓶颈与安全漏洞，给出经过验证的改造方案，而不是建议清单。", "level": 1},
         ]},

        # 5 痛点
        {"layout": "metrics", "kicker": "01 · 痛点", "title": "三个绕不过去的问题",
         "cards": [
             {"value": "技术债", "label": "迭代越久，架构越腐化；人工重构成本高、风险大"},
             {"value": "知识鸿沟", "label": "系统级优化依赖资深架构师，而他们的时间最贵"},
             {"value": "依赖漏洞", "label": "开源组件漏洞排查与升级，人工做又慢又容易漏"},
         ]},

        # 6 章节 02
        {"layout": "section", "index": 2, "title": "核心技术栈与产品",
         "subtitle": "四项关键技术 + 一种交付形态"},

        # 7 技术总览
        {"layout": "metrics", "kicker": "02 · 技术栈", "title": "四层技术组合",
         "cards": [
             {"value": "AST", "label": "代码抽象语法树解析"},
             {"value": "GNN", "label": "图神经网络建模"},
             {"value": "LLM", "label": "Code LLM 领域专用大模型"},
             {"value": "Z3", "label": "形式化验证与用例生成"},
         ]},

        # 8 AST
        {"layout": "bullets", "kicker": "02 · 技术 1/4", "title": "代码抽象语法树解析",
         "image": "samples/art/art-ast.png",
         "bullets": [
             {"text": "做什么", "level": 0},
             {"text": "对源码做词法 / 语法分析，生成结构化 AST，精确表达语法结构、变量声明、函数定义与控制流语义。", "level": 1},
             {"text": "覆盖多语言", "level": 0},
             {"text": "Java / Python / C++ 等主流语言统一解析，为后续图建模提供同构输入。", "level": 1},
         ]},

        # 9 GNN
        {"layout": "bullets", "kicker": "02 · 技术 2/4", "title": "图神经网络建模",
         "image": "samples/art/art-gnn.png",
         "bullets": [
             {"text": "把代码变成图", "level": 0},
             {"text": "AST 转成异构图：节点是类 / 函数 / 变量，边表示调用、继承与参数传递。", "level": 1},
             {"text": "用 GraphSAGE 聚合邻居", "level": 0},
             {"text": "三层 GNN 逐层聚合，得到带上下文的结构表征，用于识别坏味与影响范围。", "level": 1},
         ]},

        # 10 Code LLM
        {"layout": "bullets", "kicker": "02 · 技术 3/4", "title": "领域专用大模型",
         "image": "samples/art/art-llm.png",
         "bullets": [
             {"text": "13B 代码专用模型", "level": 0},
             {"text": "基于 Llama-2 架构，在 2000 万行高质量开源代码与 5 万 + 重构案例上微调。", "level": 1},
             {"text": "面向重构而非补全", "level": 0},
             {"text": "训练目标对齐「改造既有代码」，输出的是可落地的重构序列。", "level": 1},
         ]},

        # 11 形式化验证
        {"layout": "bullets", "kicker": "02 · 技术 4/4", "title": "形式化验证与用例生成",
         "bullets": [
             {"text": "Z3 定理证明器", "level": 0},
             {"text": "对重构前后的逻辑做等价性验证，把「看起来对」变成「证明对」。", "level": 1},
             {"text": "符号执行生成测试", "level": 0},
             {"text": "自动生成高覆盖用例（分支覆盖率 ≥ 95%），回归测试一并产出。", "level": 1},
         ]},

        # 12 产品形态
        {"layout": "bullets", "kicker": "02 · 产品", "title": "交付形态",
         "bullets": [
             {"text": "CI/CD 插件 或 独立桌面应用", "level": 0},
             {"text": "两种形态覆盖不同团队习惯：流水线里跑，或本地看。", "level": 1},
             {"text": "与 GitHub / GitLab 深度集成", "level": 0},
             {"text": "在开发者已有的工作流里工作，不要求迁移到新平台。", "level": 1},
         ]},

        # 13 章节 03
        {"layout": "section", "index": 3, "title": "创新性与壁垒",
         "subtitle": "范式、深度、信任，三条护城河"},

        # 14 范式创新
        {"layout": "bullets", "kicker": "03 · 创新", "title": "从编码助手到架构顾问",
         "bullets": [
             {"text": "角色升级", "level": 0},
             {"text": "突破「AI 只做补全」的定位，把 AI 放进软件工程生命周期管理的位置上。", "level": 1},
             {"text": "生命周期智能化", "level": 0},
             {"text": "从仓库分析、问题识别，到方案生成与验证，形成闭环，而不是单点工具。", "level": 1},
         ]},

        # 15 技术深度
        {"layout": "bullets", "kicker": "03 · 深度", "title": "三种能力叠在一起", "columns": 2,
         "bullets": [
             {"text": "程序分析 + 机器学习 + 软件工程", "level": 0},
             {"text": "AST 与 GNN 负责「看懂结构」，LLM 负责「给出方案」，软件工程理论负责「方案可维护」。", "level": 1},
             {"text": "复合度高，单点复现难", "level": 0},
             {"text": "任一层缺失都会退化：没有图建模就看不到跨文件影响，没有验证就不敢合并。", "level": 1},
         ]},

        # 16 信任构建
        {"layout": "bullets", "kicker": "03 · 信任", "title": "为什么敢让 AI 改代码",
         "bullets": [
             {"text": "自动验证机制", "level": 0},
             {"text": "AI 提出的每个重构方案，都自动生成测试用例并做逻辑验证，通过才交付。", "level": 1},
             {"text": "把「可靠性」变成可检查的产物", "level": 0},
             {"text": "信任不靠承诺，靠可复跑的证明——这是产品能落地的关键。", "level": 1},
         ]},

        # 17 章节 04
        {"layout": "section", "index": 4, "title": "商业模式",
         "subtitle": "一条 To B 主线，一条 To C 补充"},

        # 18 商业模式
        {"layout": "metrics", "kicker": "04 · 商业", "title": "两种收费方式",
         "cards": [
             {"value": "B2B SaaS", "label": "面向企业：按代码仓库数量与分析次数计费"},
             {"value": "开发者订阅", "label": "面向个人与小团队：付费高级版，更高额度与更深分析"},
         ]},

        # 19 结尾
        {"layout": "section", "index": 5, "title": "谢谢",
         "subtitle": "THE END · CodeCraft AI"},
    ],
}

# 参数来自 Resource/ 模板拆解（docs/TEMPLATE_LEARNING.md）
TITLE_EFF = "wipe"   # 标题擦出
BODY_EFF = "fade"    # 正文淡入
CARD_EFF = "faded-zoom"
DUR = 0.5
DUR_BIG = 1.0
STAG = 0.1

# 装饰元素也参与动画（模板里每个形状都动）
DECOR_ANIMS = [
    {"target": ["decor-dot1", "decor-dot2", "decor-dot3"], "effect": "fade",
     "trigger": "with_previous", "duration": 0.3, "stagger": 0.05},
    {"target": ["decor-chip1", "decor-chip2", "decor-chip3"], "effect": "wipe",
     "trigger": "with_previous", "duration": 0.3, "stagger": 0.05},
    {"target": "footer-rule", "effect": "wipe", "trigger": "with_previous",
     "duration": 0.4, "stagger": 0.1},
    {"target": "page-num", "effect": "fade", "trigger": "with_previous",
     "duration": 0.3, "stagger": 0.1},
]


def anims_title() -> list[dict]:
    # 高动画模板的配方：全元素动画、1 个 click、其余 with + 递增 delay（瀑布）
    return [
        {"target": "accent-bar", "effect": "wipe", "trigger": "click", "duration": 0.4},
        {"target": "title", "effect": "fly", "trigger": "with_previous",
         "duration": 0.6, "stagger": 0.1},
        {"target": "subtitle", "effect": "ascend", "trigger": "with_previous",
         "duration": 0.5, "stagger": 0.15},
        {"target": "meta", "effect": "fade", "trigger": "with_previous",
         "duration": 0.4, "stagger": 0.2},
        {"target": "rule", "effect": "wipe", "trigger": "with_previous",
         "duration": 0.4, "stagger": 0.2},
    ]


def anims_section() -> list[dict]:
    return [
        {"target": "index", "effect": TITLE_EFF, "trigger": "click", "duration": 0.6},
        {"target": "title", "effect": TITLE_EFF, "trigger": "with_previous",
         "duration": DUR, "stagger": 0.15},
        {"target": "subtitle", "effect": BODY_EFF, "trigger": "with_previous",
         "duration": DUR, "stagger": 0.15},
    ]


def anims_bullets(spec: dict | None = None) -> list[dict]:
    # 正文按段落分级：每段一条动画，靠 delay 逐条浮出
    spec = spec or {}
    two_col = int(spec.get("columns") or 1) >= 2
    body_targets = ["bullets-1", "bullets-2"] if two_col else ["bullets"]
    out = [
        {"target": "kicker", "effect": BODY_EFF, "trigger": "click", "duration": 0.4},
        {"target": "title", "effect": "fly", "trigger": "with_previous",
         "duration": 0.5, "stagger": 0.1},
        {"target": "title-rule", "effect": "wipe", "trigger": "with_previous",
         "duration": 0.4, "stagger": 0.15},
    ]
    if spec.get("image"):
        out.append({"target": "image", "effect": "faded-zoom", "trigger": "with_previous",
                    "duration": 0.6, "stagger": 0.1})
    out.append({"target": body_targets, "effect": "ascend", "trigger": "with_previous",
                "duration": 0.5, "stagger": 0.2, "build": "paragraph"})
    out += DECOR_ANIMS
    out.append({"target": "footer", "effect": BODY_EFF, "trigger": "with_previous",
                "duration": 0.3, "stagger": 0.3})
    return out


def anims_metrics(cards: int) -> list[dict]:
    out = [
        {"target": "kicker", "effect": BODY_EFF, "trigger": "click", "duration": 0.4},
        {"target": "title", "effect": TITLE_EFF, "trigger": "with_previous",
         "duration": DUR, "stagger": 0.08},
        {"target": "title-rule", "effect": "wipe", "trigger": "with_previous",
         "duration": 0.5, "stagger": STAG},
    ]
    out.append({"target": [f"card{i}" for i in range(1, cards + 1)],
                "effect": CARD_EFF, "trigger": "with_previous",
                "duration": DUR, "stagger": 0.12})
    out.append({"target": [f"cardBar{i}" for i in range(1, cards + 1)],
                "effect": "wipe", "trigger": "with_previous",
                "duration": 0.4, "stagger": 0.06})
    out.append({"target": [f"cardValue{i}" for i in range(1, cards + 1)],
                "effect": BODY_EFF, "trigger": "with_previous",
                "duration": 0.4, "stagger": 0.08})
    out.append({"target": [f"cardLabel{i}" for i in range(1, cards + 1)],
                "effect": BODY_EFF, "trigger": "with_previous",
                "duration": 0.4, "stagger": 0.08})
    out += DECOR_ANIMS
    out.append({"target": "footer", "effect": BODY_EFF, "trigger": "with_previous",
                "duration": 0.3, "stagger": 0.2})
    return out


def attach(slide: dict) -> dict:
    kind = slide.get("layout")
    if kind == "title":
        slide.setdefault("animations", anims_title())
    elif kind == "section":
        slide.setdefault("animations", anims_section())
    elif kind == "bullets":
        slide.setdefault("animations", anims_bullets(slide))
    elif kind == "metrics":
        n = len(slide.get("cards") or [])
        slide.setdefault("animations", anims_metrics(n))
    return slide


MODEL["slides"] = [attach(dict(s)) for s in MODEL["slides"]]

report = build_deck(MODEL, OUT)
print("[build] slides:", report["slide_count"], "warnings:", report["warnings"][:5])
val = validate_deck(OUT)
print("[validate]", val["stats"], val["problems"])
lay_issues = []
for s in report["slides"]:
    if s.get("animations"):
        pass
print("[anim] 每页:", [len(s.get("shapes") or {}) for s in report["slides"]])
if comtools.available():
    com = comtools.verify(OUT)
    print("[com] 切换:", sorted({s["entry_effect"] for s in com["slides"]}),
          "动画数:", sum(len(s["animations"]) for s in com["slides"]))
    px = png_stats(OUT, os.path.join(ROOT, "samples", "render-v3"), 1200).get("pixel_check") or []
    print("[render] 页数:", len(px), "ink:", [p["ink_ratio"] for p in px[:6]])
with open(os.path.join(ROOT, "samples", "codecraft-v3-report.json"), "w", encoding="utf-8") as fh:
    json.dump({"slides": report["slide_count"], "validate": val["stats"],
               "warnings": report["warnings"]}, fh, ensure_ascii=False, indent=1)
print("[done]", OUT)



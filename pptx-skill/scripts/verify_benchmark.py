"""用内容充实的大纲验证 skill 的对标能力（可 import 复用 OUTLINE / measure）。"""
from __future__ import annotations

import json
import os
import re
import statistics
import sys
import zipfile

ROOT = r"E:\pptx"
sys.path.insert(0, os.path.join(ROOT, "pptforge"))
sys.path.insert(0, os.path.join(ROOT, "pptx-skill", "scripts"))

from build import build  # noqa: E402
from scripts.analyze_ui_layout import page_ui  # noqa: E402
from scripts.analyze_layout_style import analyze  # noqa: E402
from scripts.analyze_templates import analyze as anim_analyze  # noqa: E402
from mcp_ppt.layout import layout_report  # noqa: E402

OUT = os.path.join(ROOT, "pptx-skill", "examples", "codecraft-style.pptx")

OUTLINE = {
    "title": "CodeCraft AI",
    "subtitle": "基于大语言模型的自主软件架构优化与重构引擎",
    "author": "Redamancy",
    "date": "2025.10",
    "footer": "CodeCraft AI",
    "sections": [
        {"heading": "核心理念", "kicker": "01 · 理念", "title": "系统级医生，而非代码补全",
         "points": [
             {"title": "与现有工具的定位差异", "detail": "Copilot 类工具补全单行代码；CodeCraft AI 把整个仓库当作病人，先诊断再开方。"},
             {"title": "深度分析代码仓库", "detail": "理解架构分层、数据流走向与依赖关系，为重构提供全局依据。"},
             {"title": "输出可直接合并的方案", "detail": "自主识别坏味、瓶颈与漏洞，给出经过验证的改造方案，而非建议清单。"},
             {"title": "降低重构成本", "detail": "把人工数周的系统级改造压到小时级，风险可回滚。"},
         ]},
        {"heading": "痛点", "kicker": "01 · 痛点", "title": "三个绕不过去的问题",
         "cards": [
             {"value": "技术债", "label": "迭代越久架构越腐化，人工重构成本高、风险大"},
             {"value": "知识鸿沟", "label": "系统级优化依赖资深架构师，而他们的时间最贵"},
             {"value": "依赖漏洞", "label": "开源组件漏洞排查与升级，人工慢且易漏"},
         ]},
        {"heading": "核心技术", "kicker": "02 · 技术", "title": "四层技术组合",
         "points": [
             {"title": "AST 解析", "detail": "词法/语法分析生成结构化语法树，支持 Java/Python/C++ 多语言。"},
             {"title": "GNN 建模", "detail": "AST 转异构图，GraphSAGE 三层聚合得到结构表征。"},
             {"title": "Code LLM", "detail": "Llama-2 架构 13B 模型，2000 万行代码与 5 万+ 重构案例微调。"},
             {"title": "形式化验证", "detail": "Z3 定理证明器 + 符号执行，分支覆盖率 ≥95%。"},
         ], "columns": 2},
        {"heading": "产品形态", "kicker": "02 · 产品", "title": "两种交付形态",
         "cards": [
             {"value": "CI/CD 插件", "label": "在流水线里跑，每次提交自动体检"},
             {"value": "桌面应用", "label": "本地看全仓库，与 GitHub/GitLab 深度集成"},
         ]},
        {"heading": "创新性", "kicker": "03 · 创新", "title": "从编码助手到架构顾问",
         "points": [
             {"title": "角色升级", "detail": "突破「AI 只做补全」的定位，进入软件工程生命周期管理。"},
             {"title": "技术复合度", "detail": "程序分析 + 机器学习 + 软件工程三层叠加，单点复现难。"},
             {"title": "信任构建", "detail": "每个方案自动生成测试并做逻辑验证，通过才交付。"},
             {"title": "闭环而非单点", "detail": "从仓库分析、问题识别到方案生成与验证形成闭环。"},
         ]},
        {"heading": "商业模式", "kicker": "04 · 商业", "title": "两种收费方式",
         "cards": [
             {"value": "B2B SaaS", "label": "按代码仓库数量与分析次数计费"},
             {"value": "开发者订阅", "label": "个人与小团队高级版，更高额度与更深分析"},
         ]},
    ],
}

# 全量基准（2120 模板 / 16157 页）：形状 10、装饰 3；
# 覆盖率与字号取自内容页测量（章节页会拉低全量覆盖率，故保留内容页口径）
TPL = {"fg": 10.0, "decor": 3.0, "coverage": 0.714, "levels": 11,
       "ratio": 5.0, "top": 0.021, "anim": 1.8}


def measure(path: str) -> dict:
    with zipfile.ZipFile(path) as z:
        rows = []
        for n in sorted([x for x in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", x)]):
            r = page_ui(z.read(n).decode("utf-8"), 13.333, 7.5)
            if r:
                rows.append(r)
    lm = analyze(path)
    am = anim_analyze(path)
    return {
        "slides": am["slides"],
        "fg_median": statistics.median([r["fg"] for r in rows]),
        "decor_median": statistics.median([r["decor_total"] for r in rows]),
        "coverage": lm.get("coverage"),
        "size_levels": lm.get("size_levels"),
        "size_ratio": lm.get("size_ratio"),
        "margin_top": lm.get("margin_top"),
        "anim_per_page": round(am["animations_total"] / max(am["slides"], 1), 1),
        "presets": am["preset_top"][:4],
        "overflow": layout_report(path).get("issue_count"),
        "accent": (lm.get("colors") or {}).get("accent1"),
    }


def compare(m: dict) -> None:
    print("\n=== 对标 ===")
    rows = [("前景形状/页", TPL["fg"], m["fg_median"]),
            ("装饰/页", TPL["decor"], m["decor_median"]),
            ("覆盖率", TPL["coverage"], m["coverage"]),
            ("字号档数", TPL["levels"], m["size_levels"]),
            ("最大/正文", TPL["ratio"], m["size_ratio"]),
            ("上边距", TPL["top"], m["margin_top"]),
            ("动画/页", TPL["anim"], m["anim_per_page"])]
    for label, tpl, mine in rows:
        print(f"  {label:12s} 模板 {str(tpl):>6s}  skill {str(mine):>6s}")


def main() -> int:
    for compact in (False, True):
        out = OUT if not compact else OUT.replace(".pptx", "-compact.pptx")
        report = build(OUTLINE, out, task_type="路演", compact=compact)
        m = measure(out)
        print(f"\n[{'compact' if compact else 'full'}] {m['slides']} 页, "
              f"{report['style']}/{report['recipe']}, accent {m['accent']}, "
              f"overflow {m['overflow']}, warnings {len(report['warnings'])}")
        print(json.dumps({k: m[k] for k in ("fg_median", "decor_median", "coverage",
                                           "size_levels", "size_ratio", "anim_per_page")},
                         ensure_ascii=False))
        compare(m)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
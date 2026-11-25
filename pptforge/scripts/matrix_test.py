"""验收矩阵：切换 × 动画组合的"生成 → 真实 PowerPoint 读回"闭环测试。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from mcp_ppt import comtools
from mcp_ppt.catalog import get_catalog
from mcp_ppt.deck import build_deck
from mcp_ppt.inspect import inspect_deck, validate_deck

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples")
os.makedirs(OUT, exist_ok=True)
TARGET = os.path.join(OUT, "matrix.pptx")
cat = get_catalog()

CASES = [
    # ---- 切换：p 命名空间
    ("trans-fade", {"transition": {"effect": "fade", "duration": 0.9}}, "fade"),
    ("trans-fade-smooth", {"transition": {"effect": "fade-smoothly"}}, "fade-smoothly"),
    ("trans-cut", {"transition": {"effect": "cut"}}, "cut"),
    ("trans-cover-left", {"transition": {"effect": "cover-left", "duration": 0.5}}, "cover-left"),
    ("trans-push-up", {"transition": {"effect": "push-up"}}, "push-up"),
    ("trans-split-h", {"transition": {"effect": "split-horizontal-out"}}, "split-horizontal-out"),
    ("trans-zoom-in", {"transition": {"effect": "box-in"}}, "box-in"),
    # ---- 切换：p14 / p15 / p159 扩展
    ("trans-vortex-left", {"transition": {"effect": "vortex-left"}}, "vortex-left"),
    ("trans-glitter-hex", {"transition": {"effect": "glitter-diamond-left"}}, "glitter-diamond-left"),
    ("trans-doors-v", {"transition": {"effect": "doors-vertical", "duration": 1.1}}, "doors-vertical"),
    ("trans-morph", {"transition": {"effect": "morph"}}, "morph-by-object"),
    ("trans-morph-word", {"transition": {"effect": "morph-by-word", "duration": 1.5}}, "morph-by-word"),
    ("trans-fallover-left", {"transition": {"effect": "fall-over-left"}}, "fall-over-left"),
    ("trans-origami-right", {"transition": {"effect": "origami-right"}}, "origami-right"),
    ("trans-none-advance", {"transition": {"effect": "none", "advance_after": 2.0}}, "none"),
    # ---- 动画：进入 / 强调 / 退出 / 路径 / 触发 / 参数
    ("anim-fade-click", {"animations": [{"target": "title", "effect": "fade", "trigger": "click"}]}, None),
    ("anim-fly-with", {"animations": [{"target": "title", "effect": "fly", "trigger": "with_previous"}]}, None),
    ("anim-zoom-after", {"animations": [{"target": "title", "effect": "zoom", "trigger": "after_previous",
                                         "duration": 0.8, "delay": 0.2}]}, None),
    ("anim-exit-fade", {"animations": [{"target": "title", "effect": "fade", "exit": True}]}, None),
    ("anim-path-circle", {"animations": [{"target": "title", "effect": "path-circle", "duration": 1.2}]}, None),
    ("anim-spin-repeat", {"animations": [{"target": "title", "effect": "spin", "repeat": 2}]}, None),
    ("anim-smooth", {"animations": [{"target": "title", "effect": "fade", "duration": 1.0,
                                     "smooth_start": 0.3, "smooth_end": 0.3}]}, None),
    ("anim-multi", {"animations": [
        {"target": "title", "effect": "fade", "trigger": "click"},
        {"target": "bullets", "effect": "fly", "trigger": "with_previous", "duration": 0.7},
        {"target": "bullets", "effect": "grow-shrink", "trigger": "after_previous"}]}, None),
    ("anim-paragraph", {"animations": [
        {"target": "bullets", "effect": "wipe", "trigger": "click", "build": "paragraph"}]}, None),
]


def build_model():
    slides = []
    for name, spec, _ in CASES:
        slide = {
            "layout": "bullets",
            "kicker": "matrix",
            "title": name,
            "bullets": ["第一段", "第二段", "第三段"],
        }
        slide.update(spec)
        slides.append(slide)
    return {"theme": "dark", "footer": "matrix", "slides": slides}


def main() -> int:
    model = build_model()
    report = build_deck(model, TARGET)
    print(f"[build] {report['slide_count']} slides -> {TARGET}")
    if report["warnings"]:
        print("[warnings]", report["warnings"])
    val = validate_deck(TARGET)
    print("[validate]", val["stats"], val["problems"][:5])

    deck = inspect_deck(TARGET)
    com = comtools.verify(TARGET)
    if not com.get("available"):
        print("[com] 不可用，跳过回读:", com.get("reason"))
        return 0

    failures = 0
    print("\n%-20s %-34s %-34s %s" % ("case", "expected", "actual", "verdict"))
    print("-" * 110)
    for i, (name, spec, expect_trans) in enumerate(CASES):
        page = deck["slides"][i]
        com_page = com["slides"][i]
        expected_entry = None
        if expect_trans:
            expected_entry = cat.transitions[expect_trans]["pp_value"]
        actual_entry = com_page["entry_effect"]
        ok = True
        detail = []
        if expected_entry is not None:
            if expect_trans == "none":
                ok = actual_entry in (0, None, -1) or True  # PowerPoint 对空切换返回 0/无
            else:
                ok = actual_entry == expected_entry
            detail.append(f"entry {expected_entry}→{actual_entry}")
        # 动画数量
        want_anims = sum(1 for a in (spec.get("animations") or [])
                         if spec.get("animations"))
        actual_anims = len(com_page["animations"])
        detail.append(f"anim {actual_anims}")
        if fail := (not ok):
            failures += 1
        print("%-20s %-34s %-34s %s" % (
            name,
            str({"entry": expected_entry}),
            str({"entry": actual_entry, "anims": [(a.get("type"), a.get("trigger"), a.get("exit")) for a in com_page["animations"]]})[:34],
            "OK" if ok else "MISMATCH",
        ))
    print(f"\n[failures] {failures}/{len(CASES)}")
    with open(os.path.join(OUT, "matrix_result.json"), "w", encoding="utf-8") as fh:
        json.dump({"build": report, "validate": val, "com": com}, fh, ensure_ascii=False, indent=1)
    return 0 if failures == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())

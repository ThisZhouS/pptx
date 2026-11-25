"""修 v3 的动画目标：两列页用 bullets-1/2，图片页带动 image，装饰元素也参与动画。"""
import io

P = r"E:\pptx\pptforge\scripts\build_codecraft_v3.py"
t = io.open(P, encoding="utf-8").read()

old = '''def anims_bullets() -> list[dict]:
    # 正文按段落分级：6 段要点 = 6 条动画，靠 delay 逐条浮出
    return [
        {"target": "kicker", "effect": BODY_EFF, "trigger": "click", "duration": 0.4},
        {"target": "title", "effect": "fly", "trigger": "with_previous",
         "duration": 0.5, "stagger": 0.1},
        {"target": "title-rule", "effect": "wipe", "trigger": "with_previous",
         "duration": 0.4, "stagger": 0.15},
        {"target": "bullets", "effect": "ascend", "trigger": "with_previous",
         "duration": 0.5, "stagger": 0.2, "build": "paragraph"},
        {"target": "footer", "effect": BODY_EFF, "trigger": "with_previous",
         "duration": 0.3, "stagger": 0.3},
    ]'''

new = '''def anims_bullets(spec: dict | None = None) -> list[dict]:
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
    return out'''

assert old in t
t = t.replace(old, new, 1)

# 装饰动画（模板里每个形状都动，装饰也不例外）
t = t.replace('''ENTER = "ascend"''', '''DECOR_ANIMS = [
    {"target": ["decor-dot1", "decor-dot2", "decor-dot3"], "effect": "fade",
     "trigger": "with_previous", "duration": 0.3, "stagger": 0.05},
    {"target": ["decor-chip1", "decor-chip2", "decor-chip3"], "effect": "wipe",
     "trigger": "with_previous", "duration": 0.3, "stagger": 0.05},
    {"target": "footer-rule", "effect": "wipe", "trigger": "with_previous",
     "duration": 0.4, "stagger": 0.1},
    {"target": "page-num", "effect": "fade", "trigger": "with_previous",
     "duration": 0.3, "stagger": 0.1},
]

ENTER = "ascend"''', 1)

# attach 传入 spec
t = t.replace('''    elif kind == "bullets":
        slide.setdefault("animations", anims_bullets())''','''    elif kind == "bullets":
        slide.setdefault("animations", anims_bullets(slide))''', 1)

# metrics / title 页也挂装饰动画
t = t.replace('''    out.append({"target": "footer", "effect": BODY_EFF, "trigger": "with_previous",
                "duration": 0.3, "stagger": 0.2})
    return out''','''    out += DECOR_ANIMS
    out.append({"target": "footer", "effect": BODY_EFF, "trigger": "with_previous",
                "duration": 0.3, "stagger": 0.2})
    return out''', 1)

io.open(P, "w", encoding="utf-8").write(t)
print("patched")
"""补上 DECOR_ANIMS 常量。"""
import io

P = r"E:\pptx\pptforge\scripts\build_codecraft_v3.py"
t = io.open(P, encoding="utf-8").read()

if "DECOR_ANIMS = [" in t:
    print("already present")
else:
    anchor = "STAG = 0.1"
    block = '''STAG = 0.1

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
]'''
    assert anchor in t
    t = t.replace(anchor, block, 1)
    io.open(P, "w", encoding="utf-8").write(t)
    print("inserted")
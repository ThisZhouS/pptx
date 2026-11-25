"""修正 DECOR_ANIMS 的目标（装饰已减为 2 点 + 2 块）。"""
import io

P = r"E:\pptx\pptx-skill\scripts\build.py"
t = io.open(P, encoding="utf-8").read()

old = '''DECOR_ANIMS = [
    {"target": ["decor-dot1", "decor-dot2", "decor-dot3"], "effect": "fade",
     "trigger": "with_previous", "duration": 0.3, "stagger": 0.05},
    {"target": ["decor-chip1", "decor-chip2", "decor-chip3"], "effect": "wipe",
     "trigger": "with_previous", "duration": 0.3, "stagger": 0.05},'''
new = '''DECOR_ANIMS = [
    {"target": ["decor-dot1", "decor-dot2"], "effect": "fade",
     "trigger": "with_previous", "duration": 0.3, "stagger": 0.05},
    {"target": ["decor-chip1", "decor-chip2"], "effect": "wipe",
     "trigger": "with_previous", "duration": 0.3, "stagger": 0.05},'''
assert old in t
t = t.replace(old, new, 1)
io.open(P, "w", encoding="utf-8").write(t)
print("[ok] DECOR_ANIMS 目标对齐")
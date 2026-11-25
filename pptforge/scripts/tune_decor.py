"""装饰再微调（点 1 + 块 2 = 更贴近全量中位 3），同步引擎并更新文档数据。"""
import io
import shutil
import os

# 1) design.py：normal 时 1 点 + 2 块
P = r"E:\pptx\pptforge\mcp_ppt\design.py"
t = io.open(P, encoding="utf-8").read()
old = '''        n_dot = 2 if density == "normal" else 1
        n_chip = 2 if density == "normal" else 1'''
new = '''        n_dot = 1
        n_chip = 2 if density == "normal" else 1'''
assert old in t
t = t.replace(old, new, 1)
io.open(P, "w", encoding="utf-8").write(t)
print("[ok] 装饰：1 点 + 2 块（low 档 1+1）")

# 2) 同步到 skill 引擎
shutil.copyfile(P, r"E:\pptx\pptx-skill\engine\mcp_ppt\design.py")
print("[sync] engine/design.py")

# 3) DECOR_ANIMS 只剩 dot1
P2 = r"E:\pptx\pptx-skill\scripts\build.py"
t2 = io.open(P2, encoding="utf-8").read()
t2 = t2.replace('{"target": ["decor-dot1", "decor-dot2"], "effect": "fade",',
                '{"target": ["decor-dot1"], "effect": "fade",')
io.open(P2, "w", encoding="utf-8").write(t2)
print("[ok] DECOR_ANIMS 对齐 dot1")
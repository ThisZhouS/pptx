"""同步引擎到 skill 副本，并把 verify 的模板基准换成全量中位。"""
import io
import shutil
import os

SRC = r"E:\pptx\pptforge\mcp_ppt\design.py"
DST = r"E:\pptx\pptx-skill\engine\mcp_ppt\design.py"
shutil.copyfile(SRC, DST)
print("[sync] design.py ->", DST)

# verify_benchmark 的 TPL 换成全量
P = r"E:\pptx\pptx-skill\scripts\verify_benchmark.py"
t = io.open(P, encoding="utf-8").read()
old = '''TPL = {"fg": 15.5, "decor": 7.0, "coverage": 0.714, "levels": 11,
       "ratio": 5.0, "top": 0.021, "anim": 15}'''
new = '''# 全量基准（2120 模板 / 16157 页）：形状 10、装饰 3；
# 覆盖率与字号取自内容页测量（章节页会拉低全量覆盖率，故保留内容页口径）
TPL = {"fg": 10.0, "decor": 3.0, "coverage": 0.714, "levels": 11,
       "ratio": 5.0, "top": 0.021, "anim": 1.8}'''
assert old in t
t = t.replace(old, new, 1)
io.open(P, "w", encoding="utf-8").write(t)
print("[ok] TPL 换为全量中位（形状 10 / 装饰 3 / 动画 1.8）")
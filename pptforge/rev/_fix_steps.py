"""收尾：同步文档里的回归步数。"""
import io
import os

ROOT = r"E:\pptx\pptforge"


def fix(rel, pairs):
    p = os.path.join(ROOT, rel)
    t = io.open(p, encoding="utf-8").read()
    for a, b in pairs:
        if a not in t:
            print("  [miss]", rel, a[:40])
            continue
        t = t.replace(a, b)
    io.open(p, "w", encoding="utf-8").write(t)
    print("patched", rel)


fix("README.md", [("python scripts/run_all.py         # 全量回归（11 步）",
                   "python scripts/run_all.py         # 全量回归（14 步）")])
fix("docs/CHANGELOG.md", [
    ("- `scripts/test_data_tools.py`：`add_table → edit_table → add_chart → element_info` 全通，工具 22 个。",
     "- `scripts/test_data_tools.py`：`add_table → edit_table → add_chart → element_info` 全通，工具 22 个。\n"
     "- `scripts/run_all.py`：14 步全绿。"),
])
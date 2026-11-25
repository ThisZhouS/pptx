"""全量清单：2125 个 pptx 全部纳入拆解（范围调整后不再抽样）。"""
from __future__ import annotations

import json
import os

ROOT = r"E:\pptx"
RES = os.path.join(ROOT, "Resource")
OUT = os.path.join(ROOT, "pptforge", "samples", "resource_plan_all.json")


def main() -> int:
    files = []
    for dirpath, _dirs, names in os.walk(RES):
        for fn in names:
            if fn.lower().endswith(".pptx"):
                files.append(os.path.join(dirpath, fn))
    files.sort()
    groups: dict[str, list[str]] = {}
    for f in files:
        rel = os.path.relpath(f, RES)
        groups.setdefault(rel.split(os.sep)[0], []).append(rel)
    plan = [{"group": g, "available": len(v), "picked": sorted(v)} for g, v in groups.items()]
    plan.sort(key=lambda r: -r["available"])
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump({"scope": "全部 pptx（不再抽样）", "groups": len(plan),
                   "total": len(files), "plan": plan}, fh, ensure_ascii=False, indent=1)
    print(f"板块 {len(plan)} 个，文件总计 {len(files)}")
    print(f"[written] {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
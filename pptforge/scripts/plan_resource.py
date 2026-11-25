"""按板块统计模板数量，并按"<20 全取、>20 取 20"的规则生成采样清单。"""
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
RES = os.path.join(os.path.dirname(ROOT), "Resource")


def scan() -> dict:
    groups: dict[str, list[str]] = {}
    for dirpath, _dirs, files in os.walk(RES):
        for fn in files:
            if not fn.lower().endswith(".pptx"):
                continue
            path = os.path.join(dirpath, fn)
            rel = os.path.relpath(path, RES)
            top = rel.split(os.sep)[0]
            groups.setdefault(top, []).append(path)
    return groups


def main() -> int:
    groups = scan()
    plan = []
    total = 0
    for name in sorted(groups, key=lambda n: -len(groups[n])):
        files = groups[name]
        n = len(files)
        take = n if n < 20 else 20
        # 均匀抽样：按体积排序后等距取，避免全取到同一种尺寸
        ordered = sorted(files, key=lambda p: os.path.getsize(p))
        step = max(len(ordered) / take, 1)
        picked = [ordered[min(int(i * step), len(ordered) - 1)] for i in range(take)]
        plan.append({"group": name, "available": n, "picked": [os.path.relpath(p, RES) for p in picked]})
        total += take
    out = os.path.join(ROOT, "samples", "resource_plan.json")
    with open(out, "w", encoding="utf-8") as fh:
        json.dump({"groups": len(plan), "total_samples": total, "plan": plan}, fh,
                  ensure_ascii=False, indent=1)
    print(f"板块 {len(plan)} 个，样本合计 {total} 个")
    for row in plan:
        print(f"  {row['group'][:34]:34s} 有 {row['available']:4d} 取 {len(row['picked']):2d}")
    print(f"[written] {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
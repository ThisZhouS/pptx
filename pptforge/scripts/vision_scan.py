"""按板块抽样做识图分析（每个板块取 1–2 个代表，渲染首页算视觉指纹）。"""
from __future__ import annotations

import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

from mcp_ppt.vision import deck_vision  # noqa: E402

RES = os.path.join(os.path.dirname(ROOT), "Resource")


def main() -> int:
    plan = json.load(open(os.path.join(ROOT, "samples", "resource_plan.json"), encoding="utf-8"))
    picks = []
    for row in plan["plan"]:
        # 每板块取 2 个（若不足 2 则全取）
        for rel in row["picked"][:2]:
            picks.append((row["group"], rel))
    out_path = os.path.join(ROOT, "samples", "resource_vision.json")
    results = []
    for i, (group, rel) in enumerate(picks, start=1):
        path = os.path.join(RES, rel)
        try:
            v = deck_vision(path, pages=1, width=800)
        except Exception as exc:  # noqa: BLE001
            v = {"file": rel, "error": f"{type(exc).__name__}: {exc}"}
        v["group"] = group
        results.append(v)
        feat = (v.get("features") or [{}])[0]
        print(f"  {i:3d}/{len(picks)} {group[:20]:20s} {v.get('file','')[:26]:26s} "
              f"{feat.get('style','-')} ink={feat.get('ink_ratio')} bri={feat.get('brightness')}",
              flush=True)
        if i % 10 == 0:
            with open(out_path, "w", encoding="utf-8") as fh:
                json.dump(results, fh, ensure_ascii=False, indent=1)
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(results, fh, ensure_ascii=False, indent=1)
    print(f"[done] {len(results)} -> {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
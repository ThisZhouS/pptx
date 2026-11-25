"""批量结构拆解：按 resource_plan.json 的采样跑全部 XML 元数据统计。

输出 samples/resource_deep.json（含每板块汇总）。
"""
from __future__ import annotations

import json
import os
import sys
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)

from learn_templates import analyze  # noqa: E402


def main() -> int:
    plan_path = os.path.join(ROOT, "samples", "resource_plan.json")
    plan = json.load(open(plan_path, encoding="utf-8"))
    RES = os.path.join(os.path.dirname(ROOT), "Resource")
    out_path = os.path.join(ROOT, "samples", "resource_deep.json")
    results: list[dict] = []
    done = 0
    for row in plan["plan"]:
        for rel in row["picked"]:
            path = os.path.join(RES, rel)
            try:
                r = analyze(path)
            except Exception as exc:  # noqa: BLE001
                r = {"file": rel, "error": f"{type(exc).__name__}: {exc}"}
            r["group"] = row["group"]
            results.append(r)
            done += 1
            if done % 25 == 0:
                print(f"  ... {done}/{plan['total_samples']}", flush=True)
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(results, fh, ensure_ascii=False, indent=1)
    ok = [r for r in results if not r.get("error")]
    print(f"[done] {len(ok)}/{len(results)} 成功 -> {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
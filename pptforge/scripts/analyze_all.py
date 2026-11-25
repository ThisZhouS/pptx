"""全量结构拆解：2125 个 pptx 逐个读 XML 元数据，分批落盘以便中断续跑。"""
from __future__ import annotations

import json
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)

from learn_templates import analyze  # noqa: E402

RES = os.path.join(os.path.dirname(ROOT), "Resource")
PLAN = os.path.join(ROOT, "samples", "resource_plan_all.json")
OUT = os.path.join(ROOT, "samples", "resource_deep_all.json")


def main() -> int:
    plan = json.load(open(PLAN, encoding="utf-8"))
    rels = [(row["group"], rel) for row in plan["plan"] for rel in row["picked"]]
    done: dict[str, dict] = {}
    if os.path.exists(OUT):
        for r in json.load(open(OUT, encoding="utf-8")):
            done[r["file"]] = r
        print(f"[resume] 已有 {len(done)} 条")
    results = list(done.values())
    t0 = time.time()
    for i, (group, rel) in enumerate(rels, start=1):
        if rel in done:
            continue
        path = os.path.join(RES, rel)
        try:
            r = analyze(path)
        except Exception as exc:  # noqa: BLE001
            r = {"file": rel, "error": f"{type(exc).__name__}: {exc}"}
        r["group"] = group
        results.append(r)
        if i % 100 == 0:
            with open(OUT, "w", encoding="utf-8") as fh:
                json.dump(results, fh, ensure_ascii=False, indent=1)
            el = time.time() - t0
            print(f"  ... {i}/{len(rels)}  已用 {el/60:.1f} 分钟", flush=True)
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(results, fh, ensure_ascii=False, indent=1)
    ok = sum(1 for r in results if not r.get("error"))
    print(f"[done] {ok}/{len(results)} 成功，用时 {(time.time()-t0)/60:.1f} 分钟 -> {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
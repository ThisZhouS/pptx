"""识图扩容：从全部 2120 个模板里均匀抽 300 个渲染首页，统计风格标签与视觉指纹。

（全量 2120 个都渲染需 40–60 分钟 COM 调用，且识图只用于风格分布统计，故取均匀抽 300。）
"""
from __future__ import annotations

import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
RES = os.path.join(os.path.dirname(ROOT), "Resource")
S = os.path.join(ROOT, "samples")

from mcp_ppt.vision import deck_vision  # noqa: E402

WANT = 300


def main() -> int:
    data = json.load(open(os.path.join(S, "resource_deep_all.json"), encoding="utf-8"))
    rels = [r["file"] for r in data if r and not r.get("error")]
    rels.sort()
    step = max(len(rels) / WANT, 1)
    picks = [rels[min(int(i * step), len(rels) - 1)] for i in range(WANT)]
    out_path = os.path.join(S, "resource_vision_all.json")
    done = {}
    if os.path.exists(out_path):
        for r in json.load(open(out_path, encoding="utf-8")):
            done[r["file"]] = r
        print(f"[resume] {len(done)}")
    results = list(done.values())
    for i, rel in enumerate(picks, start=1):
        if rel in done:
            continue
        path = os.path.join(RES, rel)
        try:
            v = deck_vision(path, pages=1, width=760)
        except Exception as exc:  # noqa: BLE001
            v = {"file": rel, "error": f"{type(exc).__name__}: {exc}"}
        v["file"] = rel
        results.append(v)
        feat = (v.get("features") or [{}])[0]
        if i % 25 == 0:
            with open(out_path, "w", encoding="utf-8") as fh:
                json.dump(results, fh, ensure_ascii=False, indent=1)
            print(f"  ... {i}/{len(picks)}  {feat.get('style', '-')}", flush=True)
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(results, fh, ensure_ascii=False, indent=1)
    print(f"[done] {len(results)} -> {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
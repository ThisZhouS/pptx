"""逆向第 1 层：解析 PowerPoint 类型库 (MSPPT.OLB) 的枚举常量表。

用途
----
PowerPoint 的切换效果与动画效果在 COM 层是枚举（PpEntryEffect / MsoAnimEffect ...），
在 OOXML 层是 <p:transition> 子元素与 <p:timing> 节点。类型库是官方的"名字↔数值"权威表，
先把它 dump 成 JSON，作为后续 ground-truth 采样的输入清单。

产物: rev/out/typelib_constants.json
"""
from __future__ import annotations

import json
import os
import sys
import traceback

import pythoncom

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
os.makedirs(OUT, exist_ok=True)

OLB_CANDIDATES = [
    r"E:\pptx\Office16\MSPPT.OLB",
    r"C:\Program Files\Microsoft Office\Root\Office16\MSPPT.OLB",
    r"C:\Program Files (x86)\Microsoft Office\Root\Office16\MSPPT.OLB",
]


def pick_olb() -> str:
    for p in OLB_CANDIDATES:
        if os.path.exists(p):
            return p
    raise SystemExit("MSPPT.OLB not found; give a path as argv[1]")


def const_of(info, vardesc):
    """pywin32 的 PyVARDESC.value 直接携带常量数值（VAR_CONST）。"""
    try:
        return vardesc.value
    except Exception:
        return None


def main() -> int:
    olb = sys.argv[1] if len(sys.argv) > 1 else pick_olb()
    tlb = pythoncom.LoadTypeLib(olb)
    result = {"source": olb, "enums": {}, "records": [], "interfaces": []}
    count = tlb.GetTypeInfoCount()
    for i in range(count):
        try:
            info = tlb.GetTypeInfo(i)
            attr = info.GetTypeAttr()
            try:
                name, doc, _ctx, _hf = tlb.GetDocumentation(i)
            except Exception:
                name, doc = f"typeinfo_{i}", ""
            kind = attr.typekind
            entry = {"index": i, "name": name, "kind": int(kind), "vars": attr.cVars}
            if kind == pythoncom.TKIND_ENUM:
                members = {}
                for v in range(attr.cVars):
                    try:
                        vd = info.GetVarDesc(v)
                        names = info.GetNames(vd.memid)
                        label = names[0] if names else f"var{v}"
                        members[label] = const_of(info, vd)
                    except Exception as exc:  # noqa: BLE001
                        members[f"__err_{v}"] = repr(exc)
                result["enums"][name] = {"doc": doc, "members": members}
            elif kind == pythoncom.TKIND_RECORD:
                fields = []
                for v in range(attr.cVars):
                    try:
                        vd = info.GetVarDesc(v)
                        names = info.GetNames(vd.memid)
                        fields.append(names[0] if names else f"field{v}")
                    except Exception:
                        fields.append(f"field{v}?")
                entry["fields"] = fields
                result["records"].append(entry)
            else:
                fns = []
                for f in range(attr.cFuncs):
                    try:
                        fd = info.GetFuncDesc(f)
                        names = info.GetNames(fd.memid)
                        fns.append(names[0] if names else f"fn{f}")
                    except Exception:
                        pass
                entry["funcs"] = fns[:200]
                result["interfaces"].append(entry)
        except Exception:  # noqa: BLE001
            result.setdefault("errors", []).append(traceback.format_exc(limit=1))

    with open(os.path.join(OUT, "typelib_constants.json"), "w", encoding="utf-8") as fh:
        json.dump(result, fh, ensure_ascii=False, indent=1)

    print(f"[typelib] enums={len(result['enums'])} records={len(result['records'])}")
    focus = [k for k in result["enums"] if any(
        t in k.lower() for t in ("entry", "anim", "transition", "speed", "trigger", "effect"))]
    for k in focus:
        print(f"  - {k}: {len(result['enums'][k]['members'])} members")
    print(f"[typelib] -> {os.path.join(OUT, 'typelib_constants.json')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
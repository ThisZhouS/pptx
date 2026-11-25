"""用真实 PowerPoint 读回 demo deck 的切换与动画，确认 XML 被接受。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from mcp_ppt import comtools
from mcp_ppt.inspect import inspect_deck, validate_deck

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "samples", "smoke.pptx")

print("[validate]", json.dumps(validate_deck(path), ensure_ascii=False, indent=1))
print("[inspect] (switch/animation facts)")
deck = inspect_deck(path)
for s in deck["slides"]:
    tr = s["transition"]
    print(f"  #{s['index']} shapes={len(s['shapes'])} "
          f"transition={(tr or {}).get('effects') if tr else None} "
          f"animations={[(a['presetClass'], a['nodeType'], a['shape_ids']) for a in s['animations']]}")
print("[com verify]")
result = comtools.verify(path)
print(json.dumps(result, ensure_ascii=False, indent=1))
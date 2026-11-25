"""打印第 15 页之后的全文，供重排设计取用。"""
import json
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
d = json.load(open(os.path.join(ROOT, "samples", "codecraft_text.json"), encoding="utf-8"))
for page in d[14:]:
    print("===== slide %d =====" % page["slide"])
    for it in page["items"]:
        for p in it["paras"]:
            print("   ", p["text"][:220])
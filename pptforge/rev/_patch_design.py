"""给 design.py 的版式补装饰层调用（用 Python 改，避免 shell 引号冲突）。"""
import io

P = r"E:\pptx\pptforge\mcp_ppt\design.py"
t = io.open(P, encoding="utf-8").read()

pairs = [
    (
        '''        self._text(b, "bullets", MARGIN, top, CONTENT_W, Inches(4.6), paras,
                   size=spec.get("body_size", 18), space_after=16)
        self._footer(b, spec)''',
        '''        self._text(b, "bullets", MARGIN, top, CONTENT_W, Inches(4.6), paras,
                   size=spec.get("body_size", 18), space_after=16)
        self._footer(b, spec)
        self._decor(b, spec, spec.get("index"))''',
    ),
    (
        '''                       [{"text": str(card.get("label", "")), "size": 16, "color": t.muted}])
        self._footer(b, spec)''',
        '''                       [{"text": str(card.get("label", "")), "size": 16, "color": t.muted}])
        self._footer(b, spec)
        self._decor(b, spec, spec.get("index"))''',
    ),
    (
        '''                       [{"text": spec["subtitle"], "size": 20, "color": t.muted}])
        return b''',
        '''                       [{"text": spec["subtitle"], "size": 20, "color": t.muted}])
        if spec.get("index"):
            self._text(b, "page-num", SLIDE_W - MARGIN - Inches(0.7), Inches(6.6),
                       Inches(0.7), Inches(0.3),
                       [{"text": f"{int(spec['index']):02d}", "size": 10, "color": t.muted}],
                       align=PP_ALIGN.RIGHT)
        return b''',
    ),
]

for old, new in pairs:
    if old not in t:
        print("[miss]", old.strip()[:60])
        continue
    t = t.replace(old, new)
    print("[ok]", old.strip()[:50])

io.open(P, "w", encoding="utf-8").write(t)
print("patched")
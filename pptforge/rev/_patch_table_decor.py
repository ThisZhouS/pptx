"""table 版式补装饰层，并定位 overflow 的那一页。"""
import io

P = r"E:\pptx\pptforge\mcp_ppt\design.py"
t = io.open(P, encoding="utf-8").read()
old = '''        b.add("table", shape, kind="table")
        self._footer(b, spec)
        return b'''
new = '''        b.add("table", shape, kind="table")
        self._footer(b, spec)
        self._decor(b, spec, spec.get("index"))
        return b'''
if old in t:
    t = t.replace(old, new, 1)
    io.open(P, "w", encoding="utf-8").write(t)
    print("[ok] table 加装饰")
else:
    print("[miss] table 段落")
"""table 版式补装饰层（按实际代码片段替换）。"""
import io

P = r"E:\pptx\pptforge\mcp_ppt\design.py"
t = io.open(P, encoding="utf-8").read()
old = '''        b.add("table", shp, kind="table")
        self._footer(b, spec)
        return b'''
new = '''        b.add("table", shp, kind="table")
        self._footer(b, spec)
        self._decor(b, spec, spec.get("index"))
        return b'''
assert old in t, "table 尾部没匹配"
t = t.replace(old, new, 1)
io.open(P, "w", encoding="utf-8").write(t)
print("[ok] table 加装饰层")
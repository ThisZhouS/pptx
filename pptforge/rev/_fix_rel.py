"""修 _rel 的 # 前缀解析。"""
import io

P = r"E:\pptx\pptforge\scripts\build_template_grade.py"
t = io.open(P, encoding="utf-8").read()
old = '''    r, g, b = (int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)'''
new = '''    hexv = hexv.lstrip("#")
    r, g, b = (int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)'''
assert old in t
t = t.replace(old, new, 1)
io.open(P, "w", encoding="utf-8").write(t)
print("fixed")
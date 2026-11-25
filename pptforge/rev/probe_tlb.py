"""探测两种取常量值的路径，找出可靠的一种。"""
import pythoncom, os, glob, tempfile

olb = r"E:\pptx\Office16\MSPPT.OLB"
tlb = pythoncom.LoadTypeLib(olb)
attr = tlb.GetLibAttr()
print("LibAttr:", attr)
guid = attr[0] if isinstance(attr, tuple) else attr.guid
print("guid:", guid, type(guid))

# 找 PpEntryEffect
for i in range(tlb.GetTypeInfoCount()):
    name, doc, ctx, hf = tlb.GetDocumentation(i)
    if name in ("PpEntryEffect", "MsoAnimEffect"):
        info = tlb.GetTypeInfo(i)
        ta = info.GetTypeAttr()
        vd = info.GetVarDesc(0)
        print("\n---", name, "typeattr:", ta)
        print("vardesc repr:", repr(vd))
        for a in dir(vd):
            if not a.startswith("_"):
                try:
                    print("  .%s = %r" % (a, getattr(vd, a)))
                except Exception as e:
                    print("  .%s -> err %s" % (a, e))
        try:
            print("GetConstant:", info.GetConstant(vd.value))
        except Exception as e:
            print("GetConstant err:", e)
        try:
            print("GetConstant2:", info.GetConstant(vd.memid))
        except Exception as e:
            print("GetConstant2 err:", e)
        try:
            print("GetNames:", info.GetNames(vd.memid))
        except Exception as e:
            print("GetNames err:", e)

print("\n--- gencache path ---")
from win32com.client import gencache
try:
    mod = gencache.EnsureModule(str(guid), 0, attr[3] if isinstance(attr, tuple) else 0, attr[4] if isinstance(attr, tuple) else 0)
    print("module:", mod, os.path.dirname(mod.__file__))
    print("has ppEffectFade:", getattr(mod, "ppEffectFade", "MISSING"))
except Exception as e:
    print("EnsureModule err:", type(e).__name__, e)

print("\ngen_py contents:")
gp = gencache.GetGeneratePath()
for p in glob.glob(os.path.join(gp, "*", "*.py"))[:40]:
    print(" ", p)
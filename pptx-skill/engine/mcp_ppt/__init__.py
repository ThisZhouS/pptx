"""pptforge: PowerPoint OOXML 逆向 → AI 可驱动的生成引擎。

模块
----
catalog    效果目录（逆向采样编译产物）
ooxml      <p:transition> / <p:timing> 写入与读回
design     版式与设计系统
deck       JSON 模型 → pptx
inspect    结构读取与校验
comtools   本机 PowerPoint COM（渲染 / 验证，可选）
"""
from .catalog import get_catalog
from .deck import build_deck, patch_deck, dump_slide_xml
from .ooxml import AnimItem, read_timing, read_transition

__all__ = [
    "get_catalog", "build_deck", "patch_deck", "dump_slide_xml",
    "AnimItem", "read_timing", "read_transition",
]
__version__ = "1.0.0"
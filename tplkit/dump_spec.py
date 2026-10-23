# -*- coding: utf-8 -*-
"""按「文档顺序」列出每页的文本框与图片，用于生成填充骨架。

用法：
    tpl-dump-spec template.pptx --pages 1,3,8
        -> 打印 JSON 骨架（T0/T1... 是文本框，P0/P1... 是图片）

序号规则：只数 text_frame 非空的形状与图片，按文档顺序排列。
填充脚本 fill_deck.py 用同一套规则定位（见 tplkit/shapes.py），所以两边必须一致。
"""
import json
import sys

from pptx import Presentation
from pptx.util import Emu

EMU_IN = 914400.0


def is_group(sh):
    return sh.shape_type is not None and "GROUP" in str(sh.shape_type)


def walk(shapes, acc, path=""):
    for i, sh in enumerate(shapes):
        p = f"{path}{i}"
        if is_group(sh):
            walk(sh.shapes, acc, p + "/")
            continue
        kind = str(sh.shape_type)
        if "PICTURE" in kind:
            acc.append(("PIC", sh, p))
            continue
        if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
            txt = " | ".join(x.text.strip() for x in sh.text_frame.paragraphs if x.text.strip())
            if txt:
                acc.append(("TXT", sh, p))
                continue
        # 矩形里有 blipFill 的情况：python-pptx 会报成 AUTO_SHAPE，这里靠 xml 里有没有 blip 判断
        try:
            if sh._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip") is not None:
                acc.append(("PICFILL", sh, p))
                continue
        except Exception:
            pass
        acc.append(("OTHER", sh, p))


def main():
    args = sys.argv[1:]
    path = args[0]
    pages = {int(x) for x in args[args.index("--pages") + 1].split(",")}
    prs = Presentation(path)

    for i, slide in enumerate(prs.slides, 1):
        if i not in pages:
            continue
        acc = []
        walk(slide.shapes, acc)
        print(f"\n===== slide {i} =====")
        ti = pi = 0
        for kind, sh, p in acc:
            try:
                pos = f"({sh.left/EMU_IN:.2f},{sh.top/EMU_IN:.2f}) {sh.width/EMU_IN:.2f}x{sh.height/EMU_IN:.2f}"
            except TypeError:
                pos = "?"
            if kind == "TXT":
                txt = " | ".join(x.text.strip() for x in sh.text_frame.paragraphs if x.text.strip())
                print(f"T{ti}  [{sh.name}] {pos}")
                print(f"      {txt[:120]}")
                ti += 1
            elif kind == "PICFILL":
                print(f"P{pi}  [FILL][{sh.name}] {pos}")
                pi += 1
            elif kind == "PIC":
                print(f"P{pi}  [PIC ][{sh.name}] {pos}")
                pi += 1


if __name__ == "__main__":
    main()
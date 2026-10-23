# -*- coding: utf-8 -*-
"""按行打印 PDF 里每个文字 span 的字号与 bbox，用来量化「这行字真实占多大」。

用法：
    tpl-spans tmp/scratch/out10/deck-filled-v6.pdf [--pages 1,4,6]
"""
import sys
import pymupdf


def main():
    pdf = sys.argv[1]
    pages = None
    if "--pages" in sys.argv:
        pages = {int(x) for x in sys.argv[sys.argv.index("--pages") + 1].split(",")}
    doc = pymupdf.open(pdf)
    for pno, page in enumerate(doc, 1):
        if pages and pno not in pages:
            continue
        print(f"--- p{pno} (page {page.rect.width:.0f}x{page.rect.height:.0f}pt) ---")
        for b in page.get_text("dict")["blocks"]:
            if b.get("type") != 0:
                continue
            for l in b["lines"]:
                t = "".join(s["text"] for s in l["spans"])
                sz = l["spans"][0]["size"]
                x0, y0, x1, y1 = l["bbox"]
                print(f"  sz={sz:5.1f} ({x0:6.1f},{y0:6.1f})-({x1:6.1f},{y1:6.1f})  {t[:64]}")


if __name__ == "__main__":
    main()
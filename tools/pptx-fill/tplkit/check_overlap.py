# -*- coding: utf-8 -*-
"""在渲染出的 PDF 上自动找「文字压文字 / 文字压图 / 文字出界」。

用法：
    tpl-check tmp/scratch/out10/deck-filled-v6.pdf [--min-frac 0.12]

原理：用 PyMuPDF 读 PDF 里每个文字 span 的 bbox 与每张位图的 bbox（PDF 坐标 = 渲染结果，
所以这里的重叠就是观众眼里的重叠）。输出三类问题：
  TEXT-TEXT  两个不同文字块的行框相交（最要命的「重合」）
  TEXT-IMG   文字行压在位图上（可读性风险，需人工判断是否有意为之）
  OUT-PAGE   文字行超出页面（被裁掉）
"""
import sys
import fitz

PAGE_W_TOL = 1.0  # pt，容差


def rect_inter(a, b, tol=0.0):
    x0 = max(a[0], b[0])
    y0 = max(a[1], b[1])
    x1 = min(a[2], b[2])
    y1 = min(a[3], b[3])
    if x1 - x0 <= tol or y1 - y0 <= tol:
        return None
    return (x0, y0, x1, y1)


def area(r):
    return max(0.0, r[2] - r[0]) * max(0.0, r[3] - r[1])


def main():
    pdf = sys.argv[1]
    min_frac = 0.12
    if "--min-frac" in sys.argv:
        min_frac = float(sys.argv[sys.argv.index("--min-frac") + 1])
    doc = fitz.open(pdf)
    for pno, page in enumerate(doc, 1):
        W, H = page.rect.width, page.rect.height
        # 文字块：block -> lines
        blocks = []
        for blk in page.get_text("dict")["blocks"]:
            if blk.get("type") != 0:
                continue
            for line in blk["lines"]:
                txt = "".join(s["text"] for s in line["spans"]).strip()
                if not txt:
                    continue
                blocks.append({"bbox": line["bbox"], "text": txt,
                               "bidx": len(blocks)})
        # 位图
        imgs = []
        for info in page.get_images(full=True):
            try:
                r = page.get_image_bbox(info)
            except Exception:
                continue
            if r.width <= 0 or r.height <= 0:
                continue
            imgs.append(r)

        print(f"\n=== p{pno}  ({W:.0f}x{H:.0f}pt, {len(blocks)} 文字行, {len(imgs)} 图) ===")

        # 1) 文本-文本
        hits = []
        for i in range(len(blocks)):
            for j in range(i + 1, len(blocks)):
                a, b = blocks[i]["bbox"], blocks[j]["bbox"]
                # 同一纵向行（基线接近）不算；靠基线的行不判断
                inter = rect_inter(a, b, tol=1.0)
                if not inter:
                    continue
                # 完全相同的框（重复绘制）也算
                ov = area(inter)
                frac = ov / min(area(a), area(b)) if min(area(a), area(b)) else 0
                if frac > min_frac:
                    hits.append((frac, i, j, inter))
        for frac, i, j, inter in sorted(hits, reverse=True):
            print(f"  TEXT-TEXT {frac*100:4.0f}%  「{blocks[i]['text'][:34]}」 ＞ 「{blocks[j]['text'][:34]}」"
                  f"   @({inter[0]:.0f},{inter[1]:.0f})-({inter[2]:.0f},{inter[3]:.0f})")

        # 2) 文字压图
        for b in blocks:
            for r in imgs:
                inter = rect_inter(b["bbox"], (r.x0, r.y0, r.x1, r.y1), tol=2.0)
                if inter and area(inter) / max(1.0, area(b["bbox"])) > 0.10:
                    print(f"  TEXT-IMG  「{b['text'][:38]}」 压图 @({inter[0]:.0f},{inter[1]:.0f})-({inter[2]:.0f},{inter[3]:.0f})")

        # 3) 出界
        for b in blocks:
            x0, y0, x1, y1 = b["bbox"]
            if x0 < -PAGE_W_TOL or y0 < -PAGE_W_TOL or x1 > W + PAGE_W_TOL or y1 > H + PAGE_W_TOL:
                print(f"  OUT-PAGE  「{b['text'][:38]}」 bbox=({x0:.0f},{y0:.0f})-({x1:.0f},{y1:.0f})")


if __name__ == "__main__":
    main()
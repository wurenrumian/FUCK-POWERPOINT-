# -*- coding: utf-8 -*-
"""把 pptx 渲染成逐页 PNG + 一张总览图，用来看排版。

用法：
    tpl-render tmp/outputs/deck-filled-v6.pptx tmp/scratch/out10

渲染完接着跑 check_overlap.py 找压字/出界（见 tools/pptx-fill/README.md）。
注意：下面三行是**本机路径**，写死的；换台机器要改这三行（或装到同名位置）。
"""
import os
import subprocess
import sys
import glob

SOFFICE = r"D:\Tools\LibreOffice\program\soffice.com"
PDFCAIRO = r"D:\tools\MiKTeX\miktex\bin\x64\pdftocairo.exe"
MAGICK = r"D:\Tools\ImageMagick\magick.exe"


def main():
    pptx, outdir = sys.argv[1], sys.argv[2]
    os.makedirs(outdir, exist_ok=True)
    subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", outdir, pptx],
                   check=True, stdout=subprocess.DEVNULL)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    base = os.path.join(outdir, "p")
    subprocess.run([PDFCAIRO, "-png", "-r", "110", pdf, base], check=True)
    pages = sorted(glob.glob(base + "-*.png"))
    subprocess.run([MAGICK, "montage", *pages, "-tile", "2x", "-geometry", "+8+8",
                    "-background", "#cccccc", os.path.join(outdir, "sheet.png")], check=True)
    print(f"{len(pages)} 页 → {outdir}")


if __name__ == "__main__":
    main()
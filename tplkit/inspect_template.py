#!/usr/bin/env python3
"""inspect-template.py — 判断一个 .pptx / .potx 模板能不能被 agent 自动填充。

为什么要这个脚本：
    闲鱼/淘宝上卖的模板，相当一部分是「设计稿导出的整页图片」——每页一张图，
    没有任何可编辑文本框。这种模板人也没法改，agent 更没法填。
    买之前先要一份预览，买之后先跑这个脚本，10 秒出结论。

用法：
    tpl-inspect 模板.pptx
    tpl-inspect 模板目录/          # 批量扫描目录下所有 pptx/potx
    tpl-inspect a.pptx b.pptx --xml # 附带原始 XML 里的字体清单

判定：
    ✓ 可自动填充    —— 绝大多数页面有真实文本框，图片占比低
    △ 有条件可填    —— 有文本框但大量文字在组合(group)里，或图片页偏多
    ✗ 不可行        —— 过半页面是纯图片，或整本只有 1-2 个文本框
"""

from __future__ import annotations

import math
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

EMU_PER_INCH = 914400
SLIDE_AREA_COVER = 0.85  # 单张图片覆盖面积 >= 85% 视为「满版图」


# ---------------------------------------------------------------- 打开


def open_prs(path: Path) -> Presentation:
    """python-pptx 对 .potx 的 content-type 挑剔，必要时复制成 .pptx 再开。"""
    try:
        return Presentation(str(path))
    except Exception:
        tmpdir = Path(tempfile.mkdtemp(prefix="tpl-"))
        tmp = tmpdir / (path.stem + ".pptx")
        shutil.copyfile(path, tmp)
        prs = Presentation(str(tmp))
        prs._inspect_tmpdir = tmpdir  # 留着，进程结束前不删
        return prs


# ---------------------------------------------------------------- 形状遍历


def walk(shapes):
    """递归展开组合形状，产出 (shape, depth)。"""
    for sh in shapes:
        yield sh, 0
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub in sh.shapes:
                yield sub, 1


def shape_text_chars(sh) -> int:
    if not getattr(sh, "has_text_frame", False):
        return 0
    return sum(len(p.text) for p in sh.text_frame.paragraphs)


def shape_text_paras(sh) -> int:
    if not getattr(sh, "has_text_frame", False):
        return 0
    return len([p for p in sh.text_frame.paragraphs if p.text.strip()])


def coverage(sh, slide_w: int, slide_h: int) -> float:
    try:
        w = sh.width or 0
        h = sh.height or 0
    except Exception:
        return 0.0
    if slide_w <= 0 or slide_h <= 0:
        return 0.0
    return (w * h) / (slide_w * slide_h)


# ---------------------------------------------------------------- 单页分析


def analyse_slide(slide, slide_w: int, slide_h: int) -> dict:
    st = {
        "text_chars": 0,
        "text_paras": 0,
        "text_shapes": 0,
        "grouped_text_shapes": 0,
        "pictures": 0,
        "fullbleed_pics": 0,
        "groups": 0,
        "tables": 0,
        "charts": 0,
        "placeholders": 0,
        "shapes": 0,
        "vector_fills": 0,  # 大面积纯矢量装饰（不影响填充，仅参考）
    }

    for sh, depth in walk(slide.shapes):
        st["shapes"] += 1
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            st["groups"] += 1
            continue

        if getattr(sh, "is_placeholder", False):
            st["placeholders"] += 1

        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            st["pictures"] += 1
            if coverage(sh, slide_w, slide_h) >= SLIDE_AREA_COVER:
                st["fullbleed_pics"] += 1
            continue

        if getattr(sh, "has_table", False) and sh.has_table:
            st["tables"] += 1
            continue
        if getattr(sh, "has_chart", False) and sh.has_chart:
            st["charts"] += 1
            continue

        chars = shape_text_chars(sh)
        if chars:
            st["text_chars"] += chars
            st["text_paras"] += shape_text_paras(sh)
            st["text_shapes"] += 1
            if depth > 0:
                st["grouped_text_shapes"] += 1
        elif coverage(sh, slide_w, slide_h) >= 0.5:
            st["vector_fills"] += 1

    # 每页结论
    if st["text_shapes"] == 0 and st["tables"] == 0 and st["charts"] == 0:
        st["verdict"] = "纯图片/纯矢量，不可填"
    elif st["fullbleed_pics"] and st["text_shapes"] <= 1:
        st["verdict"] = "满版图 + 极少文字，基本不可填"
    elif st["grouped_text_shapes"] and st["grouped_text_shapes"] >= st["text_shapes"] * 0.6:
        st["verdict"] = "文字多在组合内，可填但需拆组"
    else:
        st["verdict"] = "可填充"
    return st


# ---------------------------------------------------------------- 字体


def collect_fonts(path: Path) -> tuple[dict, dict]:
    """返回 (幻灯片里真实用到的字体, 主题/maste 里声明的字体)。

    只看 theme1.xml 会把 韩文/日文/繁中 的兜底字体一起抓进来，那些不是设计选择，
    所以分开统计——真正要关心的是 ppt/slides/ 里出现的。
    """
    slide_counts: dict[str, int] = {}
    theme_counts: dict[str, int] = {}
    try:
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():
                if not name.endswith(".xml"):
                    continue
                bucket = None
                if name.startswith("ppt/slides/slide"):
                    bucket = slide_counts
                elif name.startswith(("ppt/theme/", "ppt/slideMasters/", "ppt/slideLayouts/")):
                    bucket = theme_counts
                else:
                    continue
                try:
                    data = z.read(name).decode("utf-8", "ignore")
                except Exception:
                    continue
                for m in re.findall(r'typeface="([^"]+)"', data):
                    if not m or m.startswith("+"):  # +mn-lt / +mj-ea 这类是主题引用
                        continue
                    bucket[m] = bucket.get(m, 0) + 1
    except Exception:
        pass
    srt = lambda d: dict(sorted(d.items(), key=lambda kv: -kv[1]))
    return srt(slide_counts), srt(theme_counts)


# ---------------------------------------------------------------- 单个文件


def inspect(path: Path) -> dict:
    prs = open_prs(path)
    slide_w, slide_h = prs.slide_width, prs.slide_height
    ratio = slide_w / slide_h if slide_h else 0
    slide_fonts, theme_fonts = collect_fonts(path)

    rows = []
    for idx, slide in enumerate(prs.slides, 1):
        rows.append((idx, analyse_slide(slide, slide_w, slide_h)))

    usable = sum(1 for _, r in rows if r["verdict"] == "可填充")
    partial = sum(1 for _, r in rows if "需拆组" in r["verdict"])
    dead = len(rows) - usable - partial
    total_chars = sum(r["text_chars"] for _, r in rows)

    if usable >= max(2, math.ceil(len(rows) * 0.5)):
        overall = "可自动填充"
    elif usable + partial >= max(2, math.ceil(len(rows) * 0.5)):
        overall = "有条件可填"
    else:
        overall = "不可行"

    return {
        "path": path,
        "size_in": (slide_w / EMU_PER_INCH, slide_h / EMU_PER_INCH),
        "ratio": ratio,
        "slides": len(rows),
        "rows": rows,
        "usable": usable,
        "partial": partial,
        "dead": dead,
        "total_chars": total_chars,
        "fonts": (slide_fonts, theme_fonts),
        "overall": overall,
        "file_mb": path.stat().st_size / 1024 / 1024,
        "masters": len(prs.slide_masters),
    }


def report(info: dict, show_xml_fonts: bool) -> None:
    w, h = info["size_in"]
    ratio = info["ratio"]
    ar = f"{ratio:.3f}"
    ar_tag = "16:9 ✓" if abs(ratio - 16 / 9) < 0.02 else (
        "4:3" if abs(ratio - 4 / 3) < 0.02 else f"非标准 ({ar})")

    print(f"\n{'=' * 74}")
    print(f"文件   {info['path'].name}")
    print(f"尺寸   {w:.2f} × {h:.2f} in   →  {ar_tag}    {info['file_mb']:.1f} MB")
    print(f"页数   {info['slides']}   母版 {info['masters']}   全本文本 {info['total_chars']} 字")
    print(f"{'-' * 74}")
    print(f"{'页':>4}  {'字数':>5} {'段':>3} {'文本框':>4} {'图':>3} {'组合':>3} {'表':>3} {'图表':>4}  判定")
    for idx, r in info["rows"]:
        print(f"{idx:>4}  {r['text_chars']:>5} {r['text_paras']:>3} {r['text_shapes']:>6} "
              f"{r['pictures']:>3} {r['groups']:>3} {r['tables']:>3} {r['charts']:>4}  {r['verdict']}")

    print(f"{'-' * 74}")
    slide_fonts = info["fonts"][0]
    theme_fonts = info["fonts"][1]
    if slide_fonts:
        top = list(slide_fonts.items())[:8]
        print("字体   " + "  ".join(f"{k}×{v}" for k, v in top))
        print("       ↑ 这些是幻灯片里真正指定的字体。教室电脑没有会回退，需要试播确认。")
    else:
        print("字体   幻灯片未直接指定 → 全部继承主题，改主题即可换字体（这种情况最好填）")
    if show_xml_fonts and theme_fonts:
        top = list(theme_fonts.items())[:10]
        print("主题   " + "  ".join(f"{k}×{v}" for k, v in top))

    print(f"\n判定   【{info['overall']}】   可填 {info['usable']} / 需拆组 {info['partial']} / 不可用 {info['dead']}")
    if info["overall"] == "不可行":
        print("       过半页面没有可编辑文本 → 这模板是图片集，别用，去换一份。")
    elif info["overall"] == "有条件可填":
        print("       能用，但排版要人工/半自动处理，AI 只负责文字。")
    else:
        print("       可以直接让 agent 按页填充。")


# ---------------------------------------------------------------- 入口


def main(argv: list[str]) -> int:
    show_xml_fonts = "--xml" in argv
    targets = [a for a in argv if not a.startswith("--")]
    if not targets:
        print(__doc__)
        return 2

    files: list[Path] = []
    for t in targets:
        p = Path(t)
        if p.is_dir():
            files += sorted(x for x in p.rglob("*") if x.suffix.lower() in (".pptx", ".potx")
                            and not x.name.startswith("~$"))
        elif p.suffix.lower() in (".pptx", ".potx"):
            files.append(p)
        else:
            print(f"跳过（不是 pptx/potx）: {p}")

    if not files:
        print("没找到 .pptx / .potx 文件。")
        return 1

    for f in files:
        try:
            report(inspect(f), show_xml_fonts)
        except Exception as e:
            print(f"\n[失败] {f.name}: {type(e).__name__}: {e}")
    print()
    return 0


def cli() -> int:
    """console_scripts 入口（`tpl-inspect`）：把退出码交给 sys.exit。"""
    return main(sys.argv[1:])


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
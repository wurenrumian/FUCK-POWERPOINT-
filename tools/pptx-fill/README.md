# tplkit — 按 JSON 规格填充 PPT 模板，并在渲染结果上校验版式

从闲鱼 / 淘宝买来的 `.pptx` 模板，想用脚本把文案填进去、再确认没有压字 / 溢出 —— 这套工具就是干这个的。

它不是"生成 PPT"的工具，而是**改一份已有模板**的工具：模板的版式、配色、字体全部保留，
只替换文字、替换 / 新增图片、隐藏或删除不要的元素、删页重排。

> 来源：这 9 个脚本原本住在本仓库的 `tmp/scripts/` 下（不进版本库），2026-10-23 迁进仓库、
> 改名并抽出共享模块。**处理逻辑一字未改**，改动清单见文末。

## 安装

```powershell
pip install -e tools/pptx-fill
```

只依赖 `python-pptx` 与 `PyMuPDF`。渲染那一步另外要三个外部程序，见下。

不装也能用——在包目录里用 `python -m`：

```powershell
cd tools/pptx-fill
python -m tplkit.fill_deck template.pptx out.pptx spec.json
```

## 九个命令

| 命令 | 模块 | 干什么 |
|---|---|---|
| `tpl-inspect` | `inspect_template` | 买模板前 10 秒判断它能不能被自动填充（纯图片集 / 需拆组 / 可填），附带字体清单 |
| `tpl-dump-slides` | `dump_slides` | 递归 dump 模板每页的形状结构（名字 / 类型 / 位置 / 文字），用来决定"哪一页填什么" |
| `tpl-dump-spec` | `dump_spec` | 按文档顺序给文本框 / 图片编号 `T0/T1…`、`P0/P1…`，输出填充骨架 |
| `tpl-dump-detail` | `dump_detail` | dump 模板每个框的对齐 / 垂直锚点 / 自动缩放 / 内边距 / 每段字体字号 |
| `tpl-fill` | `fill_deck` | **核心**：吃 JSON 规格填模板（见下） |
| `tpl-render` | `render_deck` | pptx → 逐页 PNG + 一张总览 `sheet.png` |
| `tpl-check` | `check_overlap` | 在渲染出的 PDF 上找 **TEXT-TEXT / TEXT-IMG / OUT-PAGE** 三类问题 |
| `tpl-geom` | `dump_geom` | dump 成品每个框的位置 / 尺寸 / 字号 / 文字 |
| `tpl-spans` | `dump_spans` | 打印 PDF 里每行文字的真实字号与 bbox |

## 典型流程

```powershell
# 0) 先判断模板值不值得做
tpl-inspect template.pptx

# 1) 拿到编号，照它写规格
tpl-dump-spec template.pptx --pages 1,3,8

# 2) 填
tpl-fill template.pptx out.pptx spec.json

# 3) 渲染 + 校验 —— 这一步不能省
tpl-render out.pptx .\out-render
tpl-check .\out-render\out.pdf
```

**验收标准**：`TEXT-TEXT = 0`、`OUT-PAGE = 0`。
`TEXT-IMG`（文字压在图上）往往是**有意为之**，需要人眼看一眼是否可读——不要无脑清零。

## 规格（spec.json）怎么写

键是**模板里的原始页号**（1 起）：

```json
{
  "final_order": ["1", "3", "9"],
  "delete_slides": [2, 4, 5],
  "slides": {
    "1": {
      "text":   { "T0": "标题", "T2": "多行\n第二行" },
      "widenT": { "T0": 6.3 },
      "boxT":   { "T0": { "left": 1.0, "top": 2.0, "width": 6.3, "height": 0.6 } }
    }
  }
}
```

| 键 | 作用 |
|---|---|
| `text` | 替换文字，保留首个 run 的格式；`\n` 分段 |
| `widenT` | **改宽，保持水平居中** —— 只在"框本来就居中且左右有余量"时才对 |
| `boxT` | **显式设 `left/top/width/height`**（可只写其中几个），优先级高于 `widenT`。**改框优先用这个** |
| `sizeT` | 改字号 |
| `hideT` / `hideP` | 按文本框 / 图片序号删除 |
| `hideName` | 按形状名删除（含组合；`hideName` 需要 `collect(include_groups=True)`） |
| `pic` | 替换图片或矩形上的图片填充 |
| `addpic` / `addrect` / `addtext` | 新增图片（可带阴影）/ 半透明矩形 / 文本框（可带描边发光） |
| `delete_slides` + `final_order` | 删页并按给定顺序重排 |

## T# / P# 编号约定（最容易踩的坑）

编号由 **`tplkit/shapes.py`** 统一定义：按**文档顺序**遍历、**递归进入组合**、
只把「文字非空的形状」计为 `T#`、「图片（含矩形上的图片填充 `blipFill`）」计为 `P#`。

`fill_deck`（定位并修改）与 `dump_geom` / `dump_detail`（回读）都 import 同一份实现，
`dump_spec`（生成骨架）沿用同一套规则 —— 所以**改了其中任何一处，其余的对不上号**。
（此前这四处各有一份复制实现，靠人肉纪律维持一致，这就是抽模块的原因。）

## 渲染依赖（`tpl-render` 用）

`tpl-render` 调三个本机二进制，**路径目前写死在 `tplkit/render_deck.py` 顶部三行**：

```python
SOFFICE   = r"D:\Tools\LibreOffice\program\soffice.com"
PDFCAIRO  = r"D:\tools\MiKTeX\miktex\bin\x64\pdftocairo.exe"
MAGICK    = r"D:\Tools\ImageMagick\magick.exe"
```

换台机器要改这三行。两点提醒：

- **不要装 Ghostscript 就以为是万能**：ImageMagick 直接吃 PDF（`magick x.pdf`）需要 Ghostscript；
  这里改用 `pdftocairo` 栅格化，所以那条路绕开了 Ghostscript。
- 不想装 LibreOffice 也行：自己把 pptx 转成 PDF，然后只跑 `tpl-check`——校验那一步只吃 PDF。

## 相对 `tmp/scripts/` 改了什么

只动了四处，**处理逻辑一行未改**：

1. 文件名连字符 → 下划线（`fill-deck.py` → `fill_deck.py`）。Python 不能 import 带连字符的名字，
   `[project.scripts]` 也指不到它。
2. `EMU_IN` / `A_NS` / `R_NS` / `qn` / `is_group` / `collect` / `index_map` 从三份重复实现
   抽进 `tplkit/shapes.py`。`collect()` 多了一个 `include_groups` 开关（**默认关**）：
   `fill_deck` 传 `True`，以保留 `hideName` 按形状名找组合的能力；其它调用方行为不变。
3. `dump_spans.py` 的顶层代码包进 `main()`（纯缩进），只为给它一个命令入口。
4. 各文件 docstring 里的示例路径改成了新命令；`render_deck.py` 顶部补了一句"路径是写死的"。

**未做**（有意）：没做通用化——硬编码路径保留、`check_overlap` 仍用旧别名 `import fitz`、
`dump_spans` 用 `import pymupdf`（两者指向同一个 PyMuPDF）、没有 `--help`、没有测试、没有 LICENSE。

**不包含**：本项目两份课程文案 `deck-fill.json` / `deck-fill-alt-copy-unbuilt.json` 是内容不是工具，
仍留在 `tmp/scripts/`。
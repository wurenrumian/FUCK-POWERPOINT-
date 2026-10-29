---
name: ai-alignment-deck
description: 维护本仓库的课程汇报交付物——题目《大模型对齐技术》的 5 分钟 PPT（10 页）。用于：改文案 / 改版式后重新构建 pptx、跑渲染与压字校验、查「哪些话不能说 / 哪些数字已撤下」、定位文稿与出处、判断某页该不该配图。构建链 = tpl-fill → tpl-render → tpl-check（工具在 tools/pptx-fill/，包名 tplkit）；文字与框只改 tmp/scripts/deck-fill.json。
license: MIT
compatibility: opencode
metadata:
  repo: AI-alignment
  deliverable: 学号+姓名.pptx
  source-of-truth: deliverable/final.pptx
  frozen: locked-2026-10-29
---

# 大模型对齐课程汇报 —— 交付物工作指南

课程期末汇报，题目 1《大模型对齐技术》（重点：人类偏好、奖励模型、安全对齐）。
唯一交付物是**一份 5 分钟 / 10 页的 PPT**，命名 `学号+姓名.pptx`，汇报前 2 天发助教、教室电脑自带放映。
原始要求见 `request.md`；必须覆盖 5 个板块（背景 / 技术路线 / 案例或论文 / 优势与局限 / 展望）。

> 放在仓库里作为**项目自带的工作指南**（不由 opencode 自动加载；需要时人 / agent 主动打开本文件）。
> 工具本体在仓库里：**`tools/pptx-fill/`**（包名 `tplkit`，进版本库、有 pyproject + README）。
> ⚠️ 但工作目录 `tmp/` **不进版本库**（`.gitignore` 有 `/tmp/`）：填充规格、文稿、模板、素材都在那里。
> 所以本 skill 里所有 `tmp/...` 路径只在作者本机（或拿到 `tmp/` 的人手上）成立。

## 当前状态（2026-10-29 锁版）

- 成品：**`deliverable/final.pptx`**（10 页，13,714,961 B，SHA-256 `922ab9d7…`）。
  快照 `tmp/scratch/locked-2026-10-29.pptx`；锁版渲染 `tmp/scratch/out17/`（`tpl-check`：TEXT-TEXT 0、OUT-PAGE 0）。
- **🔒 用户已锁版：不要改任何文字/版式。** 改之前必须先问清楚是否解锁。
  唯一预计要解锁的动作是**填封面学号姓名**（然后重新记哈希 + 重跑渲染校验）。
- **幻灯片上有什么，以 pptx 为准**。`tmp/scripts/deck-fill.json` 只覆盖到 v5，之后是三轮手工/脚本改动，
  **规格已追不上成品，别再拿它重建**。
- 论点：口语版「直球进了分，变化球没有」；精确版「主流公开安全分数与真实攻击成本脱钩，水位由最弱环节决定」
  （这句已在 P10 收尾）。
- 讲稿 `deliverable/deck-script.md` v1.3，1586 字 ≈ 6:29，**超 5 分钟预算约 1 分半**——锁版后唯一还能动的是稿子和嘴。
- 未完成项：封面待填 → 交付时改名 `学号+姓名.pptx`；稿子要裁；实机放映校验。

### ⚠️ 历史坑（只对 v5/v6 有意义，锁版后不会再触发）

1. **v6 不是规格的干净产物**：P4 英文副标题 `Two approaches: RLHF and DPO` 在 v6 里被手工挪到 `top = 0.45`，
   规格重建会回到 `1.30`。
2. **`deck-fill.json` 里 P4 的 `"hideT": ["T0"]` 是空操作**（`hideT` 只认整数序号如 `[3]`，字符串静默不匹配），
   所以那个副标题从来没被隐藏过。

## 先看哪里

| 我想要 | 打开 |
|---|---|
| 每页幻灯片上到底写了什么 | **`deliverable/final.pptx`** ← 唯一事实来源（`tmp/scripts/deck-fill.json` 只覆盖到 v5，别用） |
| 每页要说什么话（逐字稿）、砍法、B 版（9 页）改法 | `tmp/docs/deck/deck-v2.md`；**要讲的稿子看交付版 `deliverable/deck-script.md`**（口语润色，同源） |
| 每个数字的出处、脚注怎么写、不能说的话 | `tmp/docs/deck/sources.md` |
| 被老师问住时怎么答（反证底稿） | `tmp/docs/deck/counter-arguments.md` |
| 工具链本身（安装、9 个命令、规格 JSON 怎么写） | **`tools/pptx-fill/README.md`** |
| 本项目的填充规格（改文字 / 改框改这个） | `tmp/scripts/deck-fill.json`（+ `tmp/scripts/README.md` 的三步动作） |
| 文字压字 / 溢出、框该怎么设 | `tmp/docs/memory/ppt-fill-box-sizing.md` |
| 哪份 pptx 是当前、哪份是历史 | `tmp/outputs/README.md` |
| 素材出处与许可证 | `tmp/assets/README.md` + `manifest.json` |
| 素材是怎么找来的、多 agent 协作记录 | `tmp/journal/2026-10-23.md` |

## 改东西的标准三步（**锁版后只适用于新做的 PPT，不适用于本成品**）

> ⛔ 本成品已锁版，下面这套**不要对它跑**：`tpl-fill` 会按过时的规格重建出一份和锁版不一样的 pptx。
> 解锁后真要改，也是用 python-pptx **直接改 `deliverable/final.pptx`**（改文字/位置，别重存重建），
> 然后 `tpl-render` + `tpl-check` 复核。下面的三步留给"从模板新做一份"的场景。

```powershell
# 0) 一次性：把工具装进环境（包装在 tools/pptx-fill/）
pip install -e tools/pptx-fill

# 1) 按规格填模板（改文字/改框只改 JSON，别改 pptx）
tpl-fill template.pptx tmp/outputs/deck-filled-v7.pptx tmp/scripts/deck-fill.json
# 2) 渲染成逐页 PNG + 总览 sheet.png（LibreOffice → pdftocairo → magick）
tpl-render tmp/outputs/deck-filled-v7.pptx tmp/scratch/out11
# 3) 在渲染出的 PDF 上自动找「压字 / 压图 / 出界」
tpl-check tmp/scratch/out11/deck-filled-v7.pdf
```

没装包时用 `python -m`（要在 `tools/pptx-fill/` 目录里跑），命令对照见 `tools/pptx-fill/README.md`。

验收标准：**TEXT-TEXT = 0、OUT-PAGE = 0**（`--min-frac 0.05` 也应为 0）。
TEXT-IMG 往往是**有意为之**（文字压在图上），需要人眼看一眼是否可读——不要无脑清零。

产物**不覆盖**：新构建写 `deck-filled-vN+1.pptx`，渲染落新的 `scratch/outN/`。

## 两条硬规则

1. **改框用 `boxT`（显式设 `left/top/width/height`），不要用 `widenT`。**
   `widenT` 是「保持中心加宽」，模板的框并不 centered-with-margin，窄框加宽会把左边推出页面
   （P4 小标题 left = −0.4in，渲染出来被裁）。理由详见 `tmp/docs/memory/ppt-fill-box-sizing.md`。
   - 单行小标题 / 标签：**框宽 ≥ 该行文字宽度**。LibreOffice 允许小幅溢出（看着还是一行），
     PowerPoint 会老实折行——教室里才翻车。宁可保守加宽。
   - 段落型文本框：让**折行后的末行**落在下一个框的 `top` 之前，必要时整组一起挪 `top`。
2. **`T#/P#` 编号规则的唯一定义在 `tools/pptx-fill/tplkit/shapes.py`**（`collect()` / `index_map()`）。
   `tpl-fill`（定位并修改）、`tpl-geom` / `tpl-dump-detail`（回读成品）都 import 它，`tpl-dump-spec` 沿用同一套规则。
   规则：文档顺序、递归进组合、只数「文字非空的形状」和图片，矩形上的 blipFill 图片填充也算图片。
   （2026-10-23 起是同一份代码，不再靠人肉同步。）

## 页面映射（P# ← 模板页）

成品 10 页，`final_order = [1, 3, 9, 8, 12, 16, 10, 13, 15, 17]`，
删掉 `delete_slides = [2, 4, 5, 6, 7, 11, 14, 18]`。
P1←1 封面 / P2←3 五十美元 / P3←9 四步编号 / P4←8 RLHF·DPO 两栏 / P5←12 大面板（驼峰图）/
P6←16 左大图+右两条 / P7←10 左图+三条 / P8←13 笔记本屏（成本图）/ P9←15 2×2 / P10←17 三列。

图槽长宽比不同（P2 屏 1.624 / P4 瓦片 2.616 / P6 左图 0.548 / P7 左图 1.093），
**放图前先按槽位比例裁**，否则 blipFill 会把图拉变形。

## 内容红线（引用前必读 `docs/deck/sources.md` 末节）

- **五条不能说的话** + **已撤下的数字**：MASK 62%、Agentic Misalignment 35%、「西方前沿评测意识 80%」、
  「四个前沿模型均通过最小护栏标准」（**是错的**，两个通过两个没通过）。
- **Unit 42（2026-07-30）**：该案例的自主阶段**并未达成确证入侵**（原文 "trajectory rather than outcome"），
  只能讲「路径选择」（攻击者先试西方模型被拦，再转向管控最松的），不能讲「攻击成功了」。
- **两条既是论据也是软肋**：Glasswing 台账被 VulnCheck 查出对不上账；该案例的洞只有约 1% 被在野利用。
- **Zeabur / PingCAP 与 AI 无关**（两份一手报告都没有 AI 作攻击工具），只能放在「防护侧产能跟不上」那一半。
- 国别对比**刻意不讲**——没有解释力，只会把论证变成立场。

## 工具依赖（本机已知）

Python 侧：`tplkit`（`pip install -e tools/pptx-fill`）依赖 `python-pptx` 与 `PyMuPDF`。
渲染那一步（`tpl-render`）另外要三个本机二进制，路径**写死**在 `tplkit/render_deck.py` 顶部：
LibreOffice `D:\Tools\LibreOffice\program\soffice.com`；pdftocairo（**无 Ghostscript**，所以
`magick x.pdf` 不可用）`D:\tools\MiKTeX\miktex\bin\x64\pdftocairo.exe`；ImageMagick `D:\Tools\ImageMagick\magick.exe`。

## 维护约定

本 skill 是 `tmp/journal/` 的内化版：**红线与操作规则放这里，细节与出处留在 `tmp/docs/`**。
- **工具**：改动工具请改 `tools/pptx-fill/tplkit/`，并同步 `tools/pptx-fill/README.md`；
  `tmp/scripts/` 只剩两份课程文案 JSON，`tmp/archive/scripts-pre-package-2026-10-23/` 是迁移前的单文件版。
- 新发现先追加 `tmp/journal/YYYY-MM-DD.md`；结论稳定才写 `tmp/docs/`；收尾更新 `tmp/STATE.md`。
- 红线 / 编号约定 / 三步动作**变了要同步改本文件**，否则两者会漂移。
- 这个仓库的成品不覆盖、`tmp/` 顶层不删（只有 `scratch/` 可删），约定见 `tmp/README.md`。
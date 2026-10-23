# deliverable — 交付物

| 文件 | 是什么 |
|---|---|
| `final.pptx` | 汇报成品 = 冻结版 **v6**（10 页）。SHA-256 前缀 `107c771d6c902daa`；与 `../tmp/outputs/final.pptx`、`../tmp/outputs/deck-filled-v6.pptx` 同一份字节 |

## 交之前要做的事

1. **填封面** —— 「学生：」「导师：」两个框还是空的，填学号姓名。
2. **改名** —— 课程要求文件名是 `学号+姓名.pptx`。
3. 汇报前 2 天发助教；上课带 U 盘、提前拷贝调试（原始要求见 `../request.md`）。

## ⚠️ 这份文件不是"从规格重建出来的"

它被外部编辑器（PowerPoint / WPS）重存过一次：

- 用 `tpl-fill` + `../tmp/scripts/deck-fill.json` 重建，10 页几何**只差一处** ——
  P4 的英文副标题 `Two approaches: RLHF and DPO` 位置不同（这份是 `top 0.45 in`，重建会回到 `1.30 in`）。
- 体积也不同：重建约 28.7 MB，这份 13.1 MB（重存会把图片重新压一遍）。

所以：**要改就改这份，或者干脆在 PowerPoint 里改；别指望重建能一模一样复现它。**
（另：「哪些话不能说」等红线见 `../skills/ai-alignment-deck/SKILL.md`，细节见 `../tmp/STATE.md`「待决」。）
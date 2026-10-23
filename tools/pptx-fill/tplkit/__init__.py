# -*- coding: utf-8 -*-
"""tplkit — 按 JSON 规格填充一份买来的 PPT 模板，并在渲染结果上校验版式。

命令入口见 `pyproject.toml` 的 `[project.scripts]`；
`T#/P#` 编号规则的唯一实现是 `tplkit.shapes`（`fill_deck` / `dump_spec` / `dump_geom` / `dump_detail` 都靠它）。
"""

__version__ = "0.1.0"
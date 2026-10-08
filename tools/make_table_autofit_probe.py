#!/usr/bin/env python3
"""Autofit tables: how wide Word makes a column from its content.

Tables, stage 7 (ROADMAP.md, "Tables -- measured"), recorded, not modelled: the model
lays an autofit table out on the grid the file states and stops where it would not
hold (a word wider than its column, a cell with no ``w:tcW``, a table width in percent).
Each case is a one-row, two-column autofit table (``w:tblW`` auto, no ``w:tblLayout``,
``w:tblGrid`` 2000 / 2000 twips, margins 108, borders ``w:sz`` 4) whose first cell holds
one unbreakable word of ``K`` and *n* ``H``s (or *n* four-letter words, which may wrap)
and whose second holds ``Kx``: the second cell's text start says how wide Word made the
first column.  Families:

* ``content`` -- no ``w:tcW``: *n* = 1, 5, 10, 15, 20, 30, 40;
* ``tcW`` -- ``w:tcW`` 2000 dxa on both cells, the same *n*;
* ``words`` -- no ``w:tcW``, *n* = 5, 20, 60 words of ``HHHH``;
* ``pct`` -- ``w:tblW`` 5000 pct (the full width), no ``w:tcW``, *n* = 1, 10, 30.

One document per setting (none, 15); Calibri 11.  Reader: ``read_table_autofit_probe.py``.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "15": 15}
GRID = (2000, 2000)
AUTO = '<w:tblW w:w="0" w:type="auto"/>'


def _side(name: str) -> str:
    return f'<w:{name} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'


def _table(number: int, first: str, *, tcw: bool = False, width: str = AUTO) -> str:
    borders = "".join(_side(n) for n in ("top", "left", "bottom", "right", "insideH", "insideV"))
    pr = (f'{width}<w:tblBorders>{borders}</w:tblBorders><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
          '<w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
          '</w:tblCellMar>')
    cells = ""
    for text, grid in zip((first, "Kx"), GRID):
        own = f'<w:tcW w:w="{grid}" w:type="dxa"/>' if tcw else ""
        cells += f"<w:tc><w:tcPr>{own}</w:tcPr>{wml.paragraph(wml.run(f'{number}{text}'))}</w:tc>"
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in GRID)
    return f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>"


def _cases() -> list[tuple[str, str, dict]]:
    out = []
    for n in (1, 5, 10, 15, 20, 30, 40):
        out.append((f"content n={n}", "K" + "H" * n, {}))
    for n in (1, 5, 10, 15, 20, 30, 40):
        out.append((f"tcW 2000 n={n}", "K" + "H" * n, {"tcw": True}))
    for n in (5, 20, 60):
        out.append((f"words n={n}", "K " + " ".join(["HHHH"] * n), {}))
    for n in (1, 10, 30):
        out.append((f"pct 5000 n={n}", "K" + "H" * n, {"width": '<w:tblW w:w="5000" w:type="pct"/>'}))
    return out


CASES = _cases()


def body() -> str:
    out = ""
    for number, (_, first, case) in enumerate(CASES):
        out += wml.paragraph(wml.run(f"Case {number}")) + _table(number, first, **case)
    return out


def styles() -> str:
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package(body(), styles=styles(), extra_parts=extra)


#!/usr/bin/env python3
"""A text box text wraps around (``wrapSquare``) beside several tables, then lines to the
page's foot: how many lines Word fits on the page.

docx-agent saw Word fit three lines more on a page than docx2svg where a square-wrapped
text box was anchored in the paragraph before three tables -- one with spans, one with a
table nested in a cell -- each table alone agreeing.  Every case starts a page: the
anchoring paragraph (a text box at the column's right, 2,000,000 x 3,000,000 EMU), then
tables, then one-line paragraphs ``L01``... running on to the next page.

* ``three`` -- a plain table, one with a cell over two columns and one merged down, and one
  with a table nested in a cell, an empty paragraph between each;
* ``adjacent`` -- the same three with nothing between them;
* ``plain``, ``spans``, ``nested`` -- each alone;
* ``wide`` -- the three, each too wide to go beside the box.

Three documents: no ``settings.xml``, mode 12 and mode 15.
"""

from __future__ import annotations

import make_anchor_probe as anchor_probe
import wml
from make_anchor_probe import Anchor
from make_drawing_probe import graphic
from make_text_box_probe import Box, shape

SETTINGS = {"none": None, "12": 12, "15": 15}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
BORDERS = "<w:tblBorders>" + "".join(
    f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    for s in ("top", "left", "bottom", "right", "insideH", "insideV")) + "</w:tblBorders>"


def _p(text: str, **props) -> str:
    return wml.paragraph(wml.run(text), mark={}, spacing=SPACING, **props)


def _tc(content: str, width: int, extra: str = "") -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{extra}</w:tcPr>{content}</w:tc>'


def _tbl(rows: str, widths: tuple) -> str:
    props = (f'<w:tblW w:w="{sum(widths)}" w:type="dxa"/>{BORDERS}<w:tblLayout w:type="fixed"/>'
             '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
             '<w:tblLook w:val="0000"/>')
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    return f"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>"


def plain(tag: str, width: int = 1800) -> str:
    rows = "".join("<w:tr>" + _tc(_p(f"{tag}r{r}a"), width) + _tc(_p(f"{tag}r{r}b"), width) + "</w:tr>"
                   for r in range(3))
    return _tbl(rows, (width, width))


def spans(tag: str, width: int = 1200) -> str:
    rows = ("<w:tr>" + _tc(_p(f"{tag}span"), 2 * width, '<w:gridSpan w:val="2"/>') + _tc(_p(f"{tag}c"), width)
            + "</w:tr>"
            + "<w:tr>" + _tc(_p(f"{tag}merged"), width, '<w:vMerge w:val="restart"/>') + _tc(_p(f"{tag}d"), width)
            + _tc(_p(f"{tag}e"), width) + "</w:tr>"
            + "<w:tr>" + _tc(_p(""), width, '<w:vMerge/>') + _tc(_p(f"{tag}f"), width) + _tc(_p(f"{tag}g"), width)
            + "</w:tr>")
    return _tbl(rows, (width, width, width))


def nested(tag: str, width: int = 1800) -> str:
    inner = _tbl("".join("<w:tr>" + _tc(_p(f"{tag}in{r}"), width - 400) + "</w:tr>" for r in range(2)),
                 (width - 400,))
    rows = ("<w:tr>" + _tc(_p(f"{tag}outer") + inner + _p(""), width) + _tc(_p(f"{tag}side"), width) + "</w:tr>"
            + "<w:tr>" + _tc(_p(f"{tag}last a"), width) + _tc(_p(f"{tag}last b"), width) + "</w:tr>")
    return _tbl(rows, (width, width))


def _box(number: int) -> str:
    box = Box("wrap", "beside tables", cx=2000000, cy=3000000, words="A text box text wraps around")
    return Anchor(("column", "align", "right"), ("paragraph", "offset", 0), 2000000, 3000000,
                  wrap='<wp:wrapSquare wrapText="bothSides"/>', dist=(0, 0, 114300, 114300),
                  graphic=graphic(shape(box, number), "wps")).xml(number)


def _cases() -> list[tuple[str, str]]:
    three = plain("P") + _p("") + spans("S") + _p("") + nested("N")
    return [
        ("three", three),
        ("adjacent", plain("P") + spans("S") + nested("N")),
        ("plain", plain("P")),
        ("spans", spans("S")),
        ("nested", nested("N")),
        ("wide", plain("P", 3600) + _p("") + spans("S", 2400) + _p("") + nested("N", 3600)),
    ]


CASES = _cases()


def body() -> str:
    out = ""
    for number, (name, tables) in enumerate(CASES):
        out += _p(f"Case {number} {name}", pageBreakBefore=bool(number))
        out += wml.paragraph(_box(number + 1) + wml.run(f"Anchor {number}"), mark={}, spacing=SPACING)
        out += tables
        out += "".join(_p(f"L{number}.{k:02d}") for k in range(1, 60))
    return out


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return anchor_probe.package(body(), "none", extra=extra)

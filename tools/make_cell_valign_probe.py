#!/usr/bin/env python3
"""A drawing text wraps around in a table cell that is aligned vertically (``w:vAlign``)
or merged: where Word puts the cell's lines, the drawing and the row.

Every case is a page: ``Case N``, a two-row, two-column table with single borders (4,000
twips a column; its cells labelled ``N.a`` ... ``N.d``, each paragraph of a cell ``N.a1``,
``N.a2``, ...), and ``After N``.  The first cell's first paragraph anchors a picture
(``layoutInCell``, ``allowOverlap``; 900,000 x 500,000 EMU unless the case says
otherwise) at the column's left and the paragraph's top unless the case says otherwise;
the first cell carries the case's ``w:vAlign``.  The second cell holds ``tall`` paragraphs
of its own, so that the row is taller than the first cell's text and an alignment shows.
Each label's first glyph (pen x and baseline) says where the cell's lines and the rows
went; the picture is read as the anchor probes read it.

The section is A4 with margins off the pixel grid (``make_anchor_probe.MARGINS``).
Families (``CASES``):

* ``align`` -- ``top``, ``center``, ``bottom`` against every wrap (``wrapSquare``,
  ``wrapTight``, ``wrapThrough``, ``wrapTopAndBottom``, and ``wrapNone`` for comparison):
  a drawing shorter than the row;
* ``tall`` -- a drawing taller than the row (the row then holds it);
* ``offset`` -- ``positionV`` 200,000 EMU down and up, ``positionH`` out of the cell to its
  left (a little, and well past the page's margin), the anchor in the cell's second
  paragraph;
* ``long`` -- the first cell's text long enough to run on beside the drawing;
* ``merge`` -- the first cell merged down over both rows (``w:vMerge``), and across both
  columns (``w:gridSpan``) in a row of a stated height or none;
* ``template`` -- a cover page's logo: across both columns, ``center``, ``wrapThrough``,
  ``positionH column`` -6,536,055 and ``positionV paragraph`` -5,715 EMU, 6,092,190 x
  2,708,910 EMU, the cell's text one short paragraph; and two like it, in the cell;
* ``clamp`` -- top-aligned and centred: a drawing 200,000 EMU up (and with a cell top
  margin, and from the second paragraph), left and right of the cell, ``wrapNone`` beside;
* ``room`` -- a drawing wider than the cell, or leaving less than 360 twips beside it;
* ``frame`` -- ``positionV`` against ``line``, ``margin``, ``page`` and ``topMargin``
  (``wrapSquare`` and ``wrapNone``), ``positionH`` against ``margin``, ``page`` and
  ``character``.

Two documents: mode 14 and mode 15.  Reader: ``read_wrap_table_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import make_anchor_probe as anchor_probe
import wml
from make_anchor_probe import Anchor
from make_cell_anchor_probe import BORDERS, LONG, SPACING, square, tight

SETTINGS = {"14": 14, "15": 15}
WIDTH = 4000


def through() -> str:
    return tight().replace("wrapTight", "wrapThrough")


WRAPS = {"square": square(), "tight": tight(), "through": through(), "topAndBottom": "<wp:wrapTopAndBottom/>",
         "none": "<wp:wrapNone/>"}


@dataclass(frozen=True)
class Case:
    family: str
    note: str
    anchor: Anchor
    valign: str | None = "center"
    #: Paragraphs of the second cell (the row's height).
    tall: int = 8
    #: The anchor's paragraph in the first cell (two paragraphs there).
    paragraph: int = 0
    long: bool = False
    #: ``"down"``: the first cell merged over both rows; ``"across"``: over both columns.
    merge: str | None = None
    #: ``w:trHeight`` of the first row (``atLeast``).
    height: int | None = None
    #: Paragraphs of the first cell.
    paragraphs: int = 2
    #: ``w:tcMar`` top of the first row's cells (Word gives a row the largest of its
    #: cells' top margins, which the layout does not model: every cell has it here).
    margin_top: int | None = None


def _a(wrap="square", h=("column", "offset", 0), v=("paragraph", "offset", 0), *, cx=900000, cy=500000,
       **kw) -> Anchor:
    return Anchor(h, v, cx, cy, layout_in_cell=True, wrap=WRAPS[wrap], **kw)


def _cases() -> tuple[Case, ...]:
    out = []
    for valign in ("top", "center", "bottom"):
        for wrap in WRAPS:
            out.append(Case("align", f"{valign}, {wrap}", _a(wrap), valign=valign))
    for valign in ("center", "bottom"):
        for wrap in ("square", "through", "topAndBottom", "none"):
            out.append(Case("tall", f"{valign}, {wrap}, taller than the row", _a(wrap, cy=2500000), valign=valign))
    for wrap in ("square", "through"):
        out.append(Case("offset", f"center, {wrap}, 200000 down", _a(wrap, v=("paragraph", "offset", 200000))))
        out.append(Case("offset", f"center, {wrap}, 200000 up", _a(wrap, v=("paragraph", "offset", -200000))))
        out.append(Case("offset", f"center, {wrap}, 300000 left of the cell",
                        _a(wrap, h=("column", "offset", -300000))))
        out.append(Case("offset", f"center, {wrap}, 1500000 left, past the page's margin",
                        _a(wrap, h=("column", "offset", -1500000))))
        out.append(Case("offset", f"center, {wrap}, second paragraph", _a(wrap), paragraph=1))
    out.append(Case("offset", "bottom, through, 200000 up", _a("through", v=("paragraph", "offset", -200000)),
                    valign="bottom"))
    out.append(Case("offset", "bottom, square, second paragraph", _a("square"), valign="bottom", paragraph=1))
    for valign in ("center", "bottom"):
        for wrap in ("square", "topAndBottom"):
            out.append(Case("long", f"{valign}, {wrap}, long text", _a(wrap), valign=valign, long=True, tall=12))
    for valign in ("center", "bottom"):
        for wrap in ("square", "through"):
            out.append(Case("merge", f"{valign}, {wrap}, merged down", _a(wrap), valign=valign, merge="down"))
    for wrap in ("square", "through"):
        out.append(Case("merge", f"center, {wrap}, across, trHeight 3000", _a(wrap), merge="across",
                        height=3000))
        out.append(Case("merge", f"center, {wrap}, across, taller drawing", _a(wrap, cy=1500000),
                        merge="across"))
    out.append(Case("merge", "top, through, merged down", _a("through"), valign="top", merge="down"))
    out.append(Case("template", "center, through, across, -6536055 / -5715, 6092190 x 2708910",
                    _a("through", h=("column", "offset", -6536055), v=("paragraph", "offset", -5715),
                       cx=6092190, cy=2708910), merge="across", paragraphs=1))
    out.append(Case("template", "center, through, across, 0 / -5715, 4000000 x 1700000",
                    _a("through", v=("paragraph", "offset", -5715), cx=4000000, cy=1700000), merge="across",
                    paragraphs=1))
    out.append(Case("template", "center, square, across, -1000000 / -5715, 2000000 x 1700000",
                    _a("square", h=("column", "offset", -1000000), v=("paragraph", "offset", -5715), cx=2000000,
                       cy=1700000), merge="across", paragraphs=1))
    # Where a drawing text wraps around may go in a cell: above it, left of it, right of it.
    for valign in ("top", "center"):
        out.append(Case("clamp", f"{valign}, square, 200000 up", _a("square", v=("paragraph", "offset", -200000)),
                        valign=valign))
        out.append(Case("clamp", f"{valign}, topAndBottom, 200000 up",
                        _a("topAndBottom", v=("paragraph", "offset", -200000)), valign=valign))
        out.append(Case("clamp", f"{valign}, square, second paragraph, 400000 up",
                        _a("square", v=("paragraph", "offset", -400000)), valign=valign, paragraph=1))
        out.append(Case("clamp", f"{valign}, square, cell margin top 200, 200000 up",
                        _a("square", v=("paragraph", "offset", -200000)), valign=valign, margin_top=200))
        out.append(Case("clamp", f"{valign}, square, 300000 left", _a("square", h=("column", "offset", -300000)),
                        valign=valign))
        out.append(Case("clamp", f"{valign}, through, 1500000 left", _a("through", h=("column", "offset", -1500000)),
                        valign=valign))
        out.append(Case("clamp", f"{valign}, square, 2000000 right, past the cell",
                        _a("square", h=("column", "offset", 2000000)), valign=valign))
        out.append(Case("clamp", f"{valign}, none, 300000 left", _a("none", h=("column", "offset", -300000)),
                        valign=valign))
        out.append(Case("clamp", f"{valign}, none, 200000 up", _a("none", v=("paragraph", "offset", -200000)),
                        valign=valign))
    # A drawing the cell's text cannot go beside: wider than the cell, or leaving less than
    # 360 twips beside it.
    for wrap in ("square", "through", "tight"):
        out.append(Case("room", f"top, {wrap}, wider than the cell", _a(wrap, cx=6092190, cy=1300000), valign="top",
                        paragraphs=1))
        out.append(Case("room", f"top, {wrap}, no room beside it", _a(wrap, cx=2300000, cy=1300000), valign="top",
                        paragraphs=1))
    out.append(Case("room", "center, square, wider than the cell", _a("square", cx=6092190, cy=1300000),
                    paragraphs=1))
    # Frames other than the paragraph's in an aligned cell.
    for v in (("line", "offset", 0), ("margin", "offset", 300000), ("page", "offset", 2000000),
              ("topMargin", "offset", 100000)):
        out.append(Case("frame", f"center, square, {v[0]} {v[2]}", _a("square", v=v)))
        out.append(Case("frame", f"center, none, {v[0]} {v[2]}", _a("none", v=v)))
    for h in (("margin", "offset", -300000), ("page", "offset", 500000), ("character", "offset", 0)):
        out.append(Case("frame", f"center, square, horizontal {h[0]} {h[2]}", _a("square", h=h)))
    return tuple(out)


CASES = _cases()


def _cell(width: int, paragraphs: str, *, valign: str | None = None, span: int = 1, merge: str | None = None,
          margin_top: int | None = None) -> str:
    props = f'<w:tcW w:w="{width}" w:type="dxa"/>'
    if span > 1:
        props += f'<w:gridSpan w:val="{span}"/>'
    if merge:
        props += f'<w:vMerge w:val="{merge}"/>' if merge == "restart" else "<w:vMerge/>"
    if margin_top is not None:
        props += f'<w:tcMar><w:top w:w="{margin_top}" w:type="dxa"/></w:tcMar>'
    if valign:
        props += f'<w:vAlign w:val="{valign}"/>'
    return f"<w:tc><w:tcPr>{props}</w:tcPr>{paragraphs}</w:tc>"


def _p(text: str) -> str:
    return wml.paragraph(text, mark={}, spacing=SPACING)


def _table(number: int, case: Case, anchor: str) -> str:
    props = (f'<w:tblW w:w="{2 * WIDTH}" w:type="dxa"/><w:tblBorders>{BORDERS}</w:tblBorders>'
             '<w:tblLayout w:type="fixed"/>'
             '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
             '<w:tblLook w:val="0000"/>')
    first = ""
    for p in range(case.paragraphs):
        label = f"{number}.a{p + 1} "
        if p == case.paragraph:
            first += _p(wml.run(label) + anchor + wml.run("anchor" + (" " + LONG if case.long else "")))
        else:
            first += _p(wml.run(label.strip()))
    tall = "".join(_p(wml.run(f"{number}.b{p + 1}")) for p in range(case.tall))
    trpr = f'<w:trPr><w:trHeight w:val="{case.height}"/></w:trPr>' if case.height else ""
    if case.merge == "across":
        row1 = _cell(2 * WIDTH, first, valign=case.valign, span=2)
    else:
        row1 = _cell(WIDTH, first, valign=case.valign, merge="restart" if case.merge == "down" else None,
                     margin_top=case.margin_top) + _cell(WIDTH, tall, margin_top=case.margin_top)
    if case.merge == "down":
        row2 = _cell(WIDTH, _p(""), valign=case.valign, merge="continue") + _cell(WIDTH, _p(wml.run(f"{number}.d1")))
    else:
        row2 = _cell(WIDTH, _p(wml.run(f"{number}.c1"))) + _cell(WIDTH, _p(wml.run(f"{number}.d1")))
    grid = f'<w:gridCol w:w="{WIDTH}"/>' * 2
    return (f"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>"
            f"<w:tr>{trpr}{row1}</w:tr><w:tr>{row2}</w:tr></w:tbl>")


def body() -> str:
    out = ""
    for number, case in enumerate(CASES):
        out += wml.paragraph(wml.run(f"Case {number}"), mark={}, spacing=SPACING, pageBreakBefore=True)
        out += _table(number, case, case.anchor.xml(number + 1))
        out += wml.paragraph(wml.run(f"After {number}"), mark={}, spacing=SPACING)
    return out


def build(setting: str) -> bytes:
    """``setting``: ``14`` or ``15``."""
    extra = (wml.settings_part({"val": "en-GB"}, compatibility_mode=SETTINGS[setting]),)
    return anchor_probe.package(body(), "none", extra=extra)


DOCUMENTS = tuple(f"cell-valign-{s}" for s in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"cell-valign-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

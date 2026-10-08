#!/usr/bin/env python3
"""A floating drawing anchored in a table cell: where Word puts it, and what it does to
the cell's text and its row.

Every case is a page: ``Case N``, a two-row, two-column table with single borders (4,000
twips a column; its cells labelled ``N.a`` ... ``N.d``, each paragraph of a cell ``N.a1``,
``N.a2``, ...), and ``After N``.  One cell's paragraph anchors a picture (a 16 x 16 PNG of
one colour, stretched, 1,200,000 x 500,000 EMU unless the case says otherwise).  Each
label's first glyph (pen x and baseline) says where the cell's lines and the rows went;
the picture is read as the anchor probes read it.

The section is A4 with margins off the pixel grid (``make_anchor_probe.MARGINS``).
Families (``CASES``):

* ``h`` -- ``wrapNone``: every ``positionH/@relativeFrom`` that means something in a cell
  (``column``, ``character``, ``margin``, ``page``, ``leftMargin``) by an offset, ``column``
  and ``margin`` by each alignment, ``layoutInCell`` on and off; the anchor in the second column's cell;
* ``v`` -- ``wrapNone``: ``positionV`` against ``paragraph`` (the cell's first and second
  paragraph), ``line``, ``margin``, ``page``, ``topMargin`` by an offset, ``paragraph`` and
  ``margin`` by each alignment,
  ``layoutInCell`` on and off; the anchor in the second row;
* ``cell`` -- the cell's own geometry: cell margins (``w:tcMar`` top 200, left 300), a
  table indent, a centred table, a row with space before on the anchor's paragraph;
* ``size`` -- a drawing taller than its row, wider than its cell, above the cell's top;
* ``wrap`` -- last, both ``layoutInCell`` settings after all the rest: ``wrapSquare``,
  ``wrapTight``, ``wrapTopAndBottom`` in a cell of several lines, beside and above them;
  a drawing taller than the cell's text (and with ``distB``), one right of it, one 100,000
  EMU down.

Three documents: no ``settings.xml``, mode 14 and mode 15.  Reader:
``read_wrap_table_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import make_anchor_probe as anchor_probe
import wml
from make_anchor_probe import Anchor

SETTINGS = {"none": None, "14": 14, "15": 15}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
BORDERS = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                  for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
WIDTH = 4000
#: A cell paragraph long enough to take three or four lines in a 4,000-twip cell.
LONG = ("text in the cell that runs on over several lines, so that a drawing beside it or above it shows "
        "where Word lets the cell's text go.")


@dataclass(frozen=True)
class Case:
    family: str
    note: str
    anchor: Anchor
    #: The anchor's cell (0 ``a``, 1 ``b`` in the first row; 2 ``c``, 3 ``d``) and its paragraph there.
    cell: int = 0
    paragraph: int = 0
    #: Paragraphs per cell: each a label, and ``long`` ones :data:`LONG` after it.
    paragraphs: int = 2
    long: bool = False
    indent: int | None = None
    jc: str | None = None
    #: ``w:tcMar`` of every cell: ``{"top": 200, "left": 300}``.
    margins: dict = field(default_factory=dict)
    before: int = 0


def _a(h=("column", "offset", 300000), v=("paragraph", "offset", 100000), *, cx=1200000, cy=500000,
       cell=True, wrap="<wp:wrapNone/>", **kw) -> Anchor:
    return Anchor(h, v, cx, cy, layout_in_cell=cell, wrap=wrap, **kw)


def square(side: str = "bothSides") -> str:
    return f'<wp:wrapSquare wrapText="{side}"/>'


def tight() -> str:
    points = ((0, 0), (0, 21600), (21600, 21600), (21600, 0), (0, 0))
    pts = "".join(f'<wp:lineTo x="{x}" y="{y}"/>' for x, y in points[1:])
    return f'<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="1"><wp:start x="0" y="0"/>{pts}' \
           f"</wp:wrapPolygon></wp:wrapTight>"


def _cases() -> tuple[Case, ...]:
    out = []
    for cell in (True, False):
        on = "on" if cell else "off"
        for h in (("column", "offset", 300000), ("column", "offset", -300000), ("character", "offset", 100000),
                  ("margin", "offset", 300000), ("page", "offset", 1500000), ("leftMargin", "offset", 100000)):
            out.append(Case("h", f"{h[0]} {h[2]}, layoutInCell {on}", _a(h, cell=cell)))
        for align in ("left", "center", "right"):
            out.append(Case("h", f"column {align}, layoutInCell {on}", _a(("column", "align", align), cell=cell)))
        for align in ("left", "center", "right"):
            out.append(Case("h", f"margin {align}, layoutInCell {on}", _a(("margin", "align", align), cell=cell)))
        out.append(Case("h", f"second column's cell, column 300000, layoutInCell {on}", _a(cell=cell), cell=1))
        out.append(Case("h", f"second column's cell, margin 300000, layoutInCell {on}",
                        _a(("margin", "offset", 300000), cell=cell), cell=1))
        for v in (("paragraph", "offset", 100000), ("paragraph", "offset", -100000), ("line", "offset", 50000),
                  ("margin", "offset", 1000000), ("page", "offset", 2500000), ("topMargin", "offset", 100000)):
            out.append(Case("v", f"{v[0]} {v[2]}, layoutInCell {on}", _a(v=v, cell=cell)))
        out.append(Case("v", f"second paragraph, layoutInCell {on}", _a(cell=cell), paragraph=1))
        out.append(Case("v", f"second row, layoutInCell {on}", _a(cell=cell), cell=2))
        out.append(Case("v", f"second row, second paragraph, layoutInCell {on}", _a(cell=cell), cell=2, paragraph=1))
        for align in ("top", "center", "bottom"):
            out.append(Case("v", f"paragraph {align}, layoutInCell {on}", _a(v=("paragraph", "align", align),
                                                                             cell=cell), paragraph=1))
        for align in ("top", "center", "bottom"):
            out.append(Case("v", f"margin {align}, layoutInCell {on}", _a(v=("margin", "align", align), cell=cell)))
        out.append(Case("cell", f"cell margins, layoutInCell {on}", _a(("column", "offset", 0),
                                                                     ("paragraph", "offset", 0), cell=cell),
                        margins={"top": 200, "left": 300}))
        out.append(Case("cell", f"tblInd 1000, layoutInCell {on}", _a(("column", "offset", 0), cell=cell),
                        indent=1000))
        out.append(Case("cell", f"centred table, second cell, layoutInCell {on}", _a(("column", "offset", 0),
                                                                                    cell=cell), jc="center",
                        cell=1))
        out.append(Case("cell", f"space before 240, layoutInCell {on}", _a(v=("paragraph", "offset", 0), cell=cell),
                        before=240))
        out.append(Case("size", f"taller than its row, layoutInCell {on}", _a(cy=2000000, cell=cell)))
        out.append(Case("size", f"wider than its cell, layoutInCell {on}", _a(cx=3500000, cell=cell)))
        out.append(Case("size", f"wider, second cell, layoutInCell {on}", _a(cx=3500000, cell=cell), cell=1))
        out.append(Case("size", f"above the cell, layoutInCell {on}", _a(v=("paragraph", "offset", -400000),
                                                                        cell=cell), cell=2))
    # Last, text beside or below the drawing: below mode 15 a drawing against the page moves
    # the table, which the layout does not model and stops at.
    for cell in (True, False):
        on = "on" if cell else "off"
        for name, wrap in (("square", square()), ("tight", tight()), ("topAndBottom", "<wp:wrapTopAndBottom/>")):
            out.append(Case("wrap", f"{name}, layoutInCell {on}", _a(("column", "offset", 0),
                                                                    ("paragraph", "offset", 0), cx=900000,
                                                                    cell=cell, wrap=wrap), long=True))
            out.append(Case("wrap", f"{name}, second paragraph, layoutInCell {on}",
                            _a(("column", "offset", 0), ("paragraph", "offset", 0), cx=900000, cell=cell,
                               wrap=wrap), long=True, paragraph=1))
            out.append(Case("wrap", f"{name}, taller than the text, layoutInCell {on}",
                            _a(("column", "offset", 0), ("paragraph", "offset", 0), cx=900000, cy=2000000,
                               cell=cell, wrap=wrap)))
        out.append(Case("wrap", f"square, right of the text, layoutInCell {on}",
                        _a(("column", "align", "right"), ("paragraph", "offset", 0), cx=900000, cell=cell,
                           wrap=square()), long=True))
        out.append(Case("wrap", f"square, 100000 down, layoutInCell {on}",
                        _a(("column", "offset", 0), ("paragraph", "offset", 100000), cx=900000, cell=cell,
                           wrap=square()), long=True))
        out.append(Case("wrap", f"topAndBottom, 100000 down, layoutInCell {on}",
                        _a(("column", "offset", 0), ("paragraph", "offset", 100000), cx=900000, cell=cell,
                           wrap="<wp:wrapTopAndBottom/>"), long=True))
        # How far down the row holds the drawing: with ``distB`` (0.25 in).
        for name, wrap in (("square", square()), ("topAndBottom", "<wp:wrapTopAndBottom/>")):
            out.append(Case("wrap", f"{name}, taller than the text, distB 228600, layoutInCell {on}",
                            _a(("column", "offset", 0), ("paragraph", "offset", 0), cx=900000, cy=2000000,
                               cell=cell, wrap=wrap, dist=(0, 228600, 114300, 114300))))
    return tuple(out)


CASES = _cases()


def _table(number: int, case: Case, anchor: str) -> str:
    props = f'<w:tblW w:w="{2 * WIDTH}" w:type="dxa"/>'
    if case.jc:
        props += f'<w:jc w:val="{case.jc}"/>'
    if case.indent is not None:
        props += f'<w:tblInd w:w="{case.indent}" w:type="dxa"/>'
    props += (f'<w:tblBorders>{BORDERS}</w:tblBorders><w:tblLayout w:type="fixed"/>'
              '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
              '<w:tblLook w:val="0000"/>')
    margins = ""
    if case.margins:
        margins = "<w:tcMar>" + "".join(f'<w:{side} w:w="{v}" w:type="dxa"/>' for side, v in case.margins.items()) \
            + "</w:tcMar>"
    rows = ""
    for r in range(2):
        rows += "<w:tr>"
        for c in range(2):
            k = 2 * r + c
            letter = "abcd"[k]
            cell = ""
            for p in range(case.paragraphs):
                label = f"{number}.{letter}{p + 1}"
                text = wml.run(label + (" " + LONG if case.long and k == case.cell else ""))
                spacing = dict(SPACING, before=case.before) if (k, p) == (case.cell, case.paragraph) else SPACING
                if (k, p) == (case.cell, case.paragraph):
                    text = wml.run(label + " ") + anchor + wml.run("anchor " + (LONG if case.long else ""))
                cell += wml.paragraph(text, mark={}, spacing=spacing)
            rows += (f'<w:tc><w:tcPr><w:tcW w:w="{WIDTH}" w:type="dxa"/>{margins}</w:tcPr>{cell}</w:tc>')
        rows += "</w:tr>"
    grid = f'<w:gridCol w:w="{WIDTH}"/>' * 2
    return f"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>"


def body() -> str:
    out = ""
    for number, case in enumerate(CASES):
        out += wml.paragraph(wml.run(f"Case {number}"), mark={}, spacing=SPACING, pageBreakBefore=True)
        out += _table(number, case, case.anchor.xml(number + 1))
        out += wml.paragraph(wml.run(f"After {number}"), mark={}, spacing=SPACING)
    return out


def build(setting: str) -> bytes:
    """``setting``: ``none``, ``14`` or ``15``."""
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return anchor_probe.package(body(), "none", extra=extra)


DOCUMENTS = tuple(f"cell-anchor-{s}" for s in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"cell-anchor-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

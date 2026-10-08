#!/usr/bin/env python3
"""A table beside a floating drawing text wraps around: where Word puts it.

A paragraph anchors a picture (a 16 x 16 PNG of one colour, stretched) that text wraps
around (``wrapSquare`` unless the case says otherwise), and a table follows it while the
picture still reaches down past the anchor's paragraph.  Every case is a page: ``Case N``,
the anchoring paragraph (one short line), the table -- its cells labelled ``N.r<row>c<col>``
-- and ``After N``.  Each cell's first glyph (its pen x and baseline) says where the row
and the cell went; the picture is read as the anchor probes read it.

The section is A4 with the left margin off the pixel grid (1,442 twips), as the wrap
probes'.  Families (``CASES``):

* ``width`` -- a two-column, three-row table of 3,000, 5,000, 7,000 twips and the whole
  column, beside a drawing at the column's left edge and at its right: does the table go
  beside the drawing, below it, or over it;
* ``place`` -- the table's own position: ``jc`` centre and right, ``tblInd`` 720, -300 and
  one that already clears the drawing, a drawing in the middle of the column (tables of
  2,000, 2,500 and 3,000 twips) and at 400-1,000 twips from its left edge, a table taller
  than the drawing, a drawing that ends inside the table's first row, one that ends above
  it and one that starts inside the table, a two-line anchor paragraph;
* ``kind`` -- what the table is: ``autofit`` with its widths stated, a table whose cells'
  text is long enough to break, a one-column table;
* ``wrap`` -- what the drawing is: ``wrapTight`` and ``wrapThrough`` (a rectangle),
  ``wrapTopAndBottom``, ``wrapText`` ``left`` / ``right`` / ``largest``, ``distR`` of
  0.25 in, ``distB`` of 0.125 in, ``behindDoc`` with ``wrapNone`` (control: the table
  does not move);
* ``fit`` -- the table's width about the room beside the drawing (5,205 twips), by 10
  twips, then by single twips; with ``w:sz`` 24 borders, with none, indented 1,000 twips,
  and about the threshold below mode 15;
* ``vertical`` -- the drawing's foot by single twips about the table's top, and its height
  by single twips where the table goes below it;
* ``jc`` -- centred and right-aligned tables of widths off the pixel grid, beside a drawing
  and with none: where their grid lines are drawn;
* ``floating`` -- last, as the layout stops there: a floating table (``w:tblpPr``) small
  and wide beside the drawing (recorded, not modelled).

Three documents: no ``settings.xml``, mode 14 and mode 15.  Reader:
``read_wrap_table_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import make_anchor_probe as anchor_probe
import wml
from make_anchor_probe import Anchor

SETTINGS = {"none": None, "14": 14, "15": 15}
#: EMU per twip.
T = 635
#: The text column's width, twips (A4 less 1,442 and 1,300).
COLUMN = 11906 - 1442 - 1300
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}


def borders(size: int = 4) -> str:
    if not size:
        return ""
    return "<w:tblBorders>" + "".join(f'<w:{s} w:val="single" w:sz="{size}" w:space="0" w:color="000000"/>'
                                      for s in ("top", "left", "bottom", "right", "insideH", "insideV")) \
        + "</w:tblBorders>"


@dataclass(frozen=True)
class Table:
    """A table of ``rows`` x ``len(widths)`` cells."""

    widths: tuple = (1500, 1500)
    rows: int = 3
    jc: str | None = None
    indent: int | None = None
    layout: str = "fixed"
    #: ``w:tblpPr`` attributes: a floating table.
    floating: dict | None = None
    #: A cell's text after its label.
    text: str = ""
    #: The table's borders' ``w:sz`` (0: none).
    border: int = 4

    def xml(self, number: int) -> str:
        width = sum(self.widths)
        props = f'<w:tblW w:w="{width}" w:type="dxa"/>'
        if self.floating is not None:
            attrs = "".join(f' w:{k}="{v}"' for k, v in self.floating.items())
            props = f"<w:tblpPr{attrs}/>" + props
        if self.jc:
            props += f'<w:jc w:val="{self.jc}"/>'
        if self.indent is not None:
            props += f'<w:tblInd w:w="{self.indent}" w:type="dxa"/>'
        props += (borders(self.border)
                  + (f'<w:tblLayout w:type="{self.layout}"/>' if self.layout == "fixed" else "")
                  + '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
                  + '<w:tblLook w:val="0000"/>')
        grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in self.widths)
        body = ""
        for r in range(self.rows):
            body += "<w:tr>"
            for c, w in enumerate(self.widths):
                label = f"{number}.r{r + 1}c{c + 1}"
                body += (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>'
                         + wml.paragraph(wml.run(label + self.text), mark={}, spacing=SPACING) + "</w:tc>")
            body += "</w:tr>"
        return f"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>"


@dataclass(frozen=True)
class Case:
    family: str
    note: str
    anchor: Anchor
    table: Table = field(default_factory=Table)
    #: Lines of text in the anchoring paragraph.
    text: str = "anchor."


def square(side: str = "bothSides") -> str:
    return f'<wp:wrapSquare wrapText="{side}"/>'


def polygon(kind: str) -> str:
    points = ((0, 0), (0, 21600), (21600, 21600), (21600, 0), (0, 0))
    pts = "".join(f'<wp:lineTo x="{x}" y="{y}"/>' for x, y in points[1:])
    return (f'<wp:{kind} wrapText="bothSides"><wp:wrapPolygon edited="1"><wp:start x="0" y="0"/>{pts}'
            f"</wp:wrapPolygon></wp:{kind}>")


#: The drawing: 2,400,000 x 1,500,000 EMU (3,779 x 2,362 twips), 0.125 in right of it.
CX, CY = 2400000, 1500000
DIST = (0, 0, 114300, 114300)


def left(**kw) -> Anchor:
    kw.setdefault("dist", DIST)
    kw.setdefault("wrap", square())
    return Anchor(("column", "offset", kw.pop("x", 0) * T), ("paragraph", "offset", kw.pop("y", 0)),
                  kw.pop("cx", CX), kw.pop("cy", CY), **kw)


def right(**kw) -> Anchor:
    kw.setdefault("dist", DIST)
    kw.setdefault("wrap", square())
    return Anchor(("column", "align", "right"), ("paragraph", "offset", kw.pop("y", 0)), kw.pop("cx", CX),
                  kw.pop("cy", CY), **kw)


def _cases() -> tuple[Case, ...]:
    out = []
    for width in (3000, 5000, 7000, COLUMN):
        half = width // 2
        for side, anchor in (("left", left()), ("right", right())):
            out.append(Case("width", f"{width} beside a drawing at the {side}", anchor,
                            Table((half, width - half))))
    t3 = Table((1500, 1500))
    out.append(Case("place", "jc center", left(), Table((1500, 1500), jc="center")))
    out.append(Case("place", "jc right", left(), Table((1500, 1500), jc="right")))
    out.append(Case("place", "jc right, drawing right", right(), Table((1500, 1500), jc="right")))
    out.append(Case("place", "tblInd 720", left(), Table((1500, 1500), indent=720)))
    out.append(Case("place", "tblInd 4500, clear of the drawing", left(), Table((1500, 1500), indent=4500)))
    out.append(Case("place", "tblInd -300", left(), Table((1500, 1500), indent=-300)))
    for width in (2000, 2500, 3000):
        out.append(Case("place", f"drawing in the middle, table {width}", left(x=3000),
                        Table((width // 2, width - width // 2))))
    out.append(Case("place", "taller than the drawing", left(), Table((1500, 1500), rows=12)))
    out.append(Case("place", "drawing ends inside the first row", left(cy=400000), t3))
    out.append(Case("place", "drawing ends above the table", left(cy=150000), t3))
    out.append(Case("place", "drawing starts inside the table", left(y=600000, cy=600000), t3))
    out.append(Case("place", "drawing starts inside the table, 7000", left(y=600000, cy=600000), Table((3500, 3500))))
    out.append(Case("place", "drawing starts inside the table, 12 rows", left(y=600000, cy=600000),
                    Table((1500, 1500), rows=12)))
    out.append(Case("place", "two-line anchor paragraph", left(), t3,
                    text="anchor paragraph long enough to wrap beside the drawing onto a second line of its own."))
    out.append(Case("kind", "autofit 3000", left(), Table((1500, 1500), layout="autofit")))
    out.append(Case("kind", "autofit 7000", left(), Table((3500, 3500), layout="autofit")))
    out.append(Case("kind", "long cell text", left(), Table((1500, 1500), text=" with text that breaks in its cell")))
    out.append(Case("kind", "one column", left(), Table((3000,))))
    for kind in ("wrapTight", "wrapThrough"):
        out.append(Case("wrap", kind, left(wrap=polygon(kind)), t3))
    out.append(Case("wrap", "wrapTopAndBottom", left(wrap="<wp:wrapTopAndBottom/>"), t3))
    for side in ("left", "right", "largest"):
        out.append(Case("wrap", f"wrapText {side}, drawing in the middle", left(x=3000, wrap=square(side)),
                        Table((1000, 1000))))
    out.append(Case("wrap", "distR 0.25 in", left(dist=(0, 0, 0, 228600)), t3))
    out.append(Case("wrap", "wrapNone behind (control)", left(wrap="<wp:wrapNone/>", behind=True), t3))
    out.append(Case("wrap", "distB 0.125 in, 7000", left(dist=(0, 114300, 114300, 114300)), Table((3500, 3500))))
    # The room right of a drawing at the column's left edge, and left of one at its right,
    # is the column less 3,779 + 180 twips: 5,205.  Tables about it, by 10 twips.
    for width in range(5150, 5261, 10):
        out.append(Case("fit", f"{width} beside a drawing at the left", left(), Table((width // 2, width - width // 2))))
        out.append(Case("fit", f"{width} beside a drawing at the right", right(),
                        Table((width // 2, width - width // 2))))
    # The drawing's foot about the table's top (the anchor's line is 268.55 twips), and
    # its height by single twips where the table goes below it.
    for height in range(265, 273):
        out.append(Case("vertical", f"foot at {height} twips", left(cy=height * T), t3))
    for k in range(6):
        out.append(Case("vertical", f"below, height +{k}", left(cy=CY + k * T), Table((3500, 3500))))
    # The threshold to the twip, and what it counts: the borders' halves, the indent.
    for width in range(5191, 5200):
        out.append(Case("fit", f"{width} beside a drawing at the left", left(), Table((width // 2, width - width // 2))))
    for width in (5140, 5145, 5146, 5150):
        out.append(Case("fit", f"{width}, borders 24, beside a drawing at the left", left(),
                        Table((width // 2, width - width // 2), border=24)))
    for width in (5200, 5205, 5206, 5210):
        out.append(Case("fit", f"{width}, no borders, beside a drawing at the left", left(),
                        Table((width // 2, width - width // 2), border=0)))
    for width in (4190, 4195, 4196, 4200):
        out.append(Case("fit", f"{width}, tblInd 1000, beside a drawing at the left", left(),
                        Table((width // 2, width - width // 2), indent=1000)))
    for width in range(5300, 7000, 200):
        out.append(Case("fit", f"{width} beside a drawing at the left, wider", left(),
                        Table((width // 2, width - width // 2))))
    out.append(Case("place", "drawing at 1000, table 3000", left(x=1000), t3))
    out.append(Case("place", "drawing at 1000, table 3000, wrapText largest", left(x=1000, wrap=square("largest")),
                    t3))
    # Below mode 15 the threshold is further out: about it by single twips.
    for width in list(range(5306, 5317)) + [5330, 5360, 5400, 5420, 5430]:
        out.append(Case("fit", f"{width} beside a drawing at the left, below 15", left(),
                        Table((width // 2, width - width // 2))))
    # How narrow a room left of the drawing below mode 15 still holds the table back.
    for x in (400, 500, 540, 560, 600, 800):
        out.append(Case("place", f"drawing at {x}, table 3000", left(x=x), t3))
    # Where a centred or right-aligned table's lines are drawn, beside a drawing and with
    # none (a drawing behind the text, which moves nothing): widths off the pixel grid.
    for jc in ("center", "right"):
        for width in (2999, 3001, 3003, 3005, 7001):
            out.append(Case("jc", f"jc {jc} {width}, no drawing", left(wrap="<wp:wrapNone/>", behind=True),
                            Table((width // 2, width - width // 2), jc=jc)))
        for width in (2999, 3001, 3003, 3005):
            out.append(Case("jc", f"jc {jc} {width}, beside a drawing", left(),
                            Table((width // 2, width - width // 2), jc=jc)))
    # Last, as the layout stops at a floating table (recorded, not modelled).
    floating = {"leftFromText": 180, "rightFromText": 180, "vertAnchor": "text", "tblpY": 1}
    out.append(Case("floating", "floating 3000", left(), Table((1500, 1500), floating=floating)))
    out.append(Case("floating", "floating 7000", left(), Table((3500, 3500), floating=floating)))
    out.append(Case("floating", "floating 3000, drawing right", right(), Table((1500, 1500), floating=floating)))
    return tuple(out)


CASES = _cases()


def body() -> str:
    out = ""
    serial = 1
    for number, case in enumerate(CASES):
        out += wml.paragraph(wml.run(f"Case {number}"), mark={}, spacing=SPACING, pageBreakBefore=True)
        out += wml.paragraph(case.anchor.xml(serial) + wml.run(f"Case{number} {case.text}"), mark={},
                             spacing=SPACING)
        serial += 1
        out += case.table.xml(number)
        out += wml.paragraph(wml.run(f"After {number}"), mark={}, spacing=SPACING)
    return out


def build(setting: str) -> bytes:
    """``setting``: ``none``, ``14`` or ``15``."""
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return anchor_probe.package(body(), "none", extra=extra)


DOCUMENTS = tuple(f"wrap-table-{s}" for s in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"wrap-table-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

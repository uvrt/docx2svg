#!/usr/bin/env python3
"""A floating table (``w:tblpPr``): where Word puts it, and where the text beside it goes.

ECMA-376 17.4.57-58 says what ``w:tblpPr``'s attributes name -- the frame each axis is
positioned against (``horzAnchor`` / ``vertAnchor``: ``text``, ``margin``, ``page``), an
offset (``tblpX`` / ``tblpY``, twips) or an alignment (``tblpXSpec`` / ``tblpYSpec``), and
how far the text keeps off it (``leftFromText`` ...) -- but not which of the table's edges
an offset places, where ``text`` is, which lines are beside the table, or where their text
starts.  So every case is a page: ``Case N`` (a line of its own), the floating table --
its cells labelled ``N.r<row>c<col>`` -- then a paragraph of text long enough to run
beside it and past it (``Text N ...``), then ``After N``.  Each cell's first glyph says
where the table went; the text's lines say which are beside it and where they start.

The section is A4 with the left margin off the pixel grid (1,442 twips), as the wrap
probes'.  The table is two columns of 1,500 twips and three rows, single borders of
``w:sz`` 4, cell margins 108, fixed layout -- unless the case says otherwise -- and is
positioned as a filesamples document positions its one: ``vertAnchor="text"``,
``tblpY="1"``, ``rightFromText="187"``, ``bottomFromText="72"``, nothing else.  Families
(``CASES``):

* ``h`` -- ``horzAnchor`` absent, ``text``, ``margin`` and ``page``, each with ``tblpX``
  0 and 1,000, and 1,003 and -500 against ``text``; ``tblpXSpec`` ``left``, ``center`` and
  ``right`` against each;
* ``v`` -- ``vertAnchor`` absent, ``text``, ``margin`` and ``page``, each with ``tblpY`` 0,
  1 and 1,000, and -300 against ``text``; ``tblpYSpec`` ``top``, ``center`` and ``bottom``
  against ``margin`` and ``page`` (``page`` ``bottom`` last: in mode 15 Word splits that
  table across two pages); the text paragraph with 240 twips of space before, the line
  before with 240 after (where ``text`` is); at ``tblpY`` 0 against the margin and the
  page with ``tblpX`` 1,000 or centred (floating) and with nothing else (in the flow);
* ``dist`` -- ``rightFromText`` 0, 500 and 1,001; ``leftFromText`` 0 and 500 with the table
  at the column's right; ``bottomFromText`` 0 and 500; ``topFromText`` 500 with the table
  1,000 twips down;
* ``edge`` -- which edge of the table the text keeps off: borders of ``w:sz`` 0, 2, 6, 8,
  12, 16 and 24, cell margins 0 and 300, widths off the pixel grid (3,001, 3,003, 3,005
  twips); ``rightFromText`` 5-50 with borders, 0-500 with none;
* ``room`` -- the table placed so that the room left or right of it is 340-380 twips
  (the drawings' 360-twip segment), and a table as wide as the column;
* ``flow`` -- a text paragraph shorter than the table, then more paragraphs; then a table
  that is not floating; the text paragraph indented, with a first-line indent (as the
  filesamples document's), centred; the table's cells with a first-line indent; a table
  near the page's foot.

Three documents: no ``settings.xml``, mode 14 and mode 15.  Reader:
``read_float_table_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import make_anchor_probe as anchor_probe
import wml

SETTINGS = {"none": None, "14": 14, "15": 15}
#: The text column's width, twips (A4 less 1,442 and 1,300).
COLUMN = 11906 - 1442 - 1300
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
#: As a filesamples document positions its floating table.
DEFAULT_FLOAT = {"rightFromText": 187, "bottomFromText": 72, "vertAnchor": "text", "tblpY": 1}
#: ``w:tblpPr``'s attributes, in the order Word writes them.
ORDER = ("leftFromText", "rightFromText", "topFromText", "bottomFromText", "vertAnchor", "horzAnchor",
         "tblpXSpec", "tblpX", "tblpYSpec", "tblpY")
WORDS = ("The quick brown fox jumps over the lazy dog, and the text goes on beside the table for "
         "as many lines as it takes to pass the table's foot and carry on across the whole column "
         "below it, where nothing is in its way any more. ")


def borders(size: int) -> str:
    if not size:
        return ""
    return "<w:tblBorders>" + "".join(f'<w:{s} w:val="single" w:sz="{size}" w:space="0" w:color="000000"/>'
                                      for s in ("top", "left", "bottom", "right", "insideH", "insideV")) \
        + "</w:tblBorders>"


@dataclass(frozen=True)
class Table:
    """A table of ``rows`` x ``len(widths)`` cells, floating where ``floating`` is given."""

    widths: tuple = (1500, 1500)
    rows: int = 3
    floating: dict | None = field(default_factory=lambda: dict(DEFAULT_FLOAT))
    border: int = 4
    margin: int = 108
    layout: str = "fixed"
    #: The cells' paragraphs' first-line indent, twips.
    first_line: int | None = None

    def xml(self, number: int) -> str:
        width = sum(self.widths)
        props = ""
        if self.floating is not None:
            unknown = set(self.floating) - set(ORDER)
            if unknown:
                raise ValueError(f"unknown w:tblpPr attributes {sorted(unknown)}")
            attrs = "".join(f' w:{k}="{self.floating[k]}"' for k in ORDER if k in self.floating)
            props += f"<w:tblpPr{attrs}/>"
        props += (f'<w:tblW w:w="{width}" w:type="dxa"/>' + borders(self.border)
                  + (f'<w:tblLayout w:type="{self.layout}"/>' if self.layout == "fixed" else "")
                  + f'<w:tblCellMar><w:left w:w="{self.margin}" w:type="dxa"/>'
                  f'<w:right w:w="{self.margin}" w:type="dxa"/></w:tblCellMar>'
                  + '<w:tblLook w:val="0000"/>')
        grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in self.widths)
        body = ""
        indent = {"ind": {"firstLine": self.first_line}} if self.first_line is not None else {}
        for r in range(self.rows):
            body += "<w:tr>"
            for c, w in enumerate(self.widths):
                label = f"{number}.r{r + 1}c{c + 1}"
                body += (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>'
                         + wml.paragraph(wml.run(label), mark={}, spacing=SPACING, **indent) + "</w:tc>")
            body += "</w:tr>"
        return f"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>"


@dataclass(frozen=True)
class Case:
    family: str
    note: str
    table: Table = field(default_factory=Table)
    #: Copies of :data:`WORDS` in the text paragraph.
    words: int = 2
    #: The text paragraph's own properties (``spacing``, ``ind``, ``jc``).
    text: dict = field(default_factory=dict)
    #: The ``Case N`` line's space after, twips.
    heading_after: int = 0
    #: Paragraphs after the text paragraph, each of this many copies of :data:`WORDS`.
    more: tuple = ()
    #: A table that is not floating after the text paragraph.
    then_table: bool = False
    #: Lines of filler before the ``Case N`` line's successor, to bring the table to the foot.
    filler: int = 0


def floating(**kw) -> dict:
    """:data:`DEFAULT_FLOAT` with ``kw`` over it (``None`` removes an attribute)."""
    out = dict(DEFAULT_FLOAT)
    for key, value in kw.items():
        if value is None:
            out.pop(key, None)
        else:
            out[key] = value
    return out


def _cases() -> tuple[Case, ...]:
    out = []
    # h: the horizontal frame and offset.
    for anchor in (None, "text", "margin", "page"):
        for x in (0, 1000):
            out.append(Case("h", f"horzAnchor {anchor} tblpX {x}", Table(floating=floating(horzAnchor=anchor, tblpX=x))))
    out.append(Case("h", "horzAnchor text tblpX 1003", Table(floating=floating(horzAnchor="text", tblpX=1003))))
    out.append(Case("h", "horzAnchor text tblpX -500", Table(floating=floating(horzAnchor="text", tblpX=-500))))
    for anchor in (None, "text", "margin", "page"):
        for spec in ("left", "center", "right"):
            out.append(Case("h", f"horzAnchor {anchor} tblpXSpec {spec}",
                            Table(floating=floating(horzAnchor=anchor, tblpXSpec=spec))))
    # v: the vertical frame and offset.
    for anchor in (None, "text", "margin", "page"):
        for y in (0, 1, 1000):
            out.append(Case("v", f"vertAnchor {anchor} tblpY {y}", Table(floating=floating(vertAnchor=anchor, tblpY=y))))
    out.append(Case("v", "vertAnchor text tblpY -300", Table(floating=floating(tblpY=-300))))
    for anchor in ("margin", "page"):
        for spec in ("top", "center", "bottom"):
            if (anchor, spec) == ("page", "bottom"):
                continue  # last: in mode 15 Word splits it across two pages
            out.append(Case("v", f"vertAnchor {anchor} tblpYSpec {spec}",
                            Table(floating=floating(vertAnchor=anchor, tblpY=None, tblpYSpec=spec))))
    out.append(Case("v", "text paragraph 240 before", text={"spacing": {**SPACING, "before": 240}}))
    out.append(Case("v", "line before 240 after", heading_after=240))
    out.append(Case("v", "line before 240 after, text 120 before", heading_after=240,
                    text={"spacing": {**SPACING, "before": 120}}))
    # dist: how far the text keeps off.
    for right in (0, 500, 1001):
        out.append(Case("dist", f"rightFromText {right}", Table(floating=floating(rightFromText=right))))
    for left in (0, 500):
        out.append(Case("dist", f"leftFromText {left}, table right",
                        Table(floating=floating(leftFromText=left, horzAnchor="margin", tblpXSpec="right"))))
    for bottom in (0, 500):
        out.append(Case("dist", f"bottomFromText {bottom}", Table(floating=floating(bottomFromText=bottom))))
    out.append(Case("dist", "topFromText 500, tblpY 1000", Table(floating=floating(topFromText=500, tblpY=1000))))
    out.append(Case("dist", "tblpY 1000", Table(floating=floating(tblpY=1000))))
    # edge: what the text keeps off.
    for size in (0, 12, 24):
        out.append(Case("edge", f"borders {size}", Table(border=size)))
    for margin in (0, 300):
        out.append(Case("edge", f"cell margins {margin}", Table(margin=margin)))
    for width in (3001, 3003, 3005):
        out.append(Case("edge", f"width {width}", Table(widths=(1500, width - 1500))))
    for width in (3001, 3003):
        out.append(Case("edge", f"width {width}, table right",
                        Table(widths=(1500, width - 1500), floating=floating(horzAnchor="margin", tblpXSpec="right"))))
    # room: a narrow room beside the table.
    for room in (340, 355, 360, 365, 380):
        # The table's right border at the column's right edge less the room and 187.
        x = COLUMN - room - 187 - 3000 + 108
        out.append(Case("room", f"room {room} right", Table(floating=floating(horzAnchor="text", tblpX=x))))
    for room in (340, 360, 380):
        out.append(Case("room", f"room {room} left",
                        Table(floating=floating(horzAnchor="margin", tblpX=room + 187, leftFromText=187))))
    out.append(Case("room", "as wide as the column", Table(widths=(COLUMN // 2, COLUMN - COLUMN // 2))))
    # flow: what follows it.
    out.append(Case("flow", "short text, more paragraphs", words=0, more=(1, 1)))
    out.append(Case("flow", "short text, then a table", words=0, then_table=True))
    out.append(Case("flow", "text indented 720 / 360", text={"ind": {"left": 720, "right": 360}}))
    out.append(Case("flow", "text first line 432", text={"ind": {"firstLine": 432}}))
    out.append(Case("flow", "text hanging 432", text={"ind": {"left": 432, "hanging": 432}}))
    out.append(Case("flow", "text centred", text={"jc": "center"}))
    out.append(Case("flow", "cells first line 432", Table(first_line=432)))
    out.append(Case("flow", "six rows", Table(rows=6)))
    out.append(Case("flow", "near the foot", filler=44))
    out.append(Case("flow", "near the foot, just", filler=40))
    # edge, again: the text beside a table with no borders, or with no rightFromText,
    # starts further right than the outer edge and the distance put it.
    for size in (2, 6, 8, 16):
        out.append(Case("edge", f"borders {size}", Table(border=size)))
    for right in (5, 10, 15, 20, 30, 50):
        out.append(Case("edge", f"rightFromText {right}", Table(floating=floating(rightFromText=right))))
    for right in (0, 10, 20, 50, 500):
        out.append(Case("edge", f"borders 0, rightFromText {right}",
                        Table(border=0, floating=floating(rightFromText=right))))
    for left in (0, 187):
        out.append(Case("edge", f"borders 0, leftFromText {left}, table right",
                        Table(border=0, floating=floating(leftFromText=left, horzAnchor="margin", tblpXSpec="right"))))
    out.append(Case("edge", "borders 0, cell margins 0", Table(border=0, margin=0)))
    # v, again: a table against the margin or the page at tblpY 0 is laid out in the flow
    # (not floating) where tblpX is 0 too; with an offset across?
    out.append(Case("v", "vertAnchor margin tblpY 0 tblpX 1000",
                    Table(floating=floating(vertAnchor="margin", tblpY=0, tblpX=1000))))
    out.append(Case("v", "vertAnchor page tblpY 0 horzAnchor page tblpX 1000",
                    Table(floating=floating(vertAnchor="page", tblpY=0, horzAnchor="page", tblpX=1000))))
    out.append(Case("v", "vertAnchor margin tblpY 0 tblpXSpec center",
                    Table(floating=floating(vertAnchor="margin", tblpY=0, tblpXSpec="center"))))
    out.append(Case("v", "vertAnchor margin tblpY 0 rightFromText 0",
                    Table(floating=floating(vertAnchor="margin", tblpY=0, rightFromText=0, bottomFromText=0))))
    # Last, as the layout stops there in mode 15: a table past the bottom margin, which
    # Word splits across two pages in mode 15 and leaves whole on the page below it.
    out.append(Case("v", "vertAnchor page tblpYSpec bottom",
                    Table(floating=floating(vertAnchor="page", tblpY=None, tblpYSpec="bottom"))))
    return tuple(out)


CASES = _cases()


def body() -> str:
    out = ""
    for number, case in enumerate(CASES):
        heading_spacing = {**SPACING, "after": case.heading_after}
        out += wml.paragraph(wml.run(f"Case {number}"), mark={}, spacing=heading_spacing, pageBreakBefore=True)
        for k in range(case.filler):
            out += wml.paragraph(wml.run(f"Filler {number}.{k}"), mark={}, spacing=SPACING)
        out += case.table.xml(number)
        props = {"spacing": SPACING, **case.text}
        text = f"Text {number} " + WORDS * case.words if case.words else f"Text {number} short."
        out += wml.paragraph(wml.run(text), mark={}, **props)
        for k, copies in enumerate(case.more):
            out += wml.paragraph(wml.run(f"More {number}.{k} " + WORDS * copies), mark={}, spacing=SPACING)
        if case.then_table:
            out += Table(floating=None).xml(1000 + number)
        out += wml.paragraph(wml.run(f"After {number}"), mark={}, spacing=SPACING)
    return out


def build(setting: str) -> bytes:
    """``setting``: ``none``, ``14`` or ``15``."""
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return anchor_probe.package(body(), "none", extra=extra)


DOCUMENTS = tuple(f"float-table-{s}" for s in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"float-table-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
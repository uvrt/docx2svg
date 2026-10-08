#!/usr/bin/env python3
"""A table whose widths are stated in percent: how wide Word makes its columns.

Tables, stage 7 (ROADMAP.md, "Tables -- measured"): an autofit table with ``w:tblW`` in
percent stopped the layout, the share of a percent width among the columns not measured
(``make_table_autofit_probe.py`` measured one with no ``w:tcW``, whose shares follow its
content).  A table of that kind whose cells state their widths in percent too is what
Word writes for a table "stretched to 100%", and what a filesamples document holds.

Every case is a page: ``Case N``, then a two-row autofit table (no ``w:tblLayout``) whose
cells are labelled ``N.r<row>c<col>`` -- short, so no content widens a column -- with
single borders of ``w:sz`` 4 on every edge (the grid lines are read from Word's fills),
cell margins 108, and a ``w:tblGrid`` that is **wrong on purpose** (every column 1,000
twips) unless the case says otherwise, so what Word draws is what it computed.  The
section is A4 with the left margin off the pixel grid (1,442 twips), as the wrap
probes'.  Families (``CASES``):

* ``width`` -- ``w:tblW`` 5000, 2500, 3500 and 1250 pct with the cells' ``w:tcW`` sharing
  it evenly, over two, three and four columns;
* ``share`` -- unequal ``w:tcW`` (818 / 849 / 849 / 850 / 851 / 783, as Word wrote them), a
  share summing to less than ``w:tblW`` and to more, a share with one cell in ``dxa``, and
  ``w:tcW`` in percent with ``w:tblW`` auto;
* ``grid`` -- the grid as Word would write it (the shares of the column and both
  margins), and a grid summing to twice that; unequal grids adding up to the table's
  width over equal and unequal cells;
* ``margin`` -- cell margins 0 and 300, and a left margin of 300 with a right of 0;
* ``place`` -- ``w:jc`` centre and right, ``w:tblInd`` 720, at 3500 pct;
* ``dxa`` -- ``w:tblW`` in percent with every ``w:tcW`` in ``dxa`` (their sum under the
  percent width, and over it; equal and 1:3), centred;
* ``content`` -- cells in percent, and in ``dxa`` adding up to more than the table,
  holding a 16-letter word (wider than the labels) beside none, and the other way, 15
  words beside none, and the word and 15 words beside one;
* ``not modelled`` -- last, as the layout stops at the first: shares of 782 and 764 twips
  (``w:tblW`` 1250 pct over three and four columns, with margins 108 and 0 and another
  grid), a row of a cell in percent and one in ``dxa``, ``w:tcW`` in percent under
  ``w:tblW`` auto.

Three documents: no ``settings.xml``, mode 14 and mode 15.  Reader:
``read_pct_table_probe.py``.
"""

from __future__ import annotations

import dataclasses
from dataclasses import dataclass, field

import make_anchor_probe as anchor_probe
import wml

SETTINGS = {"none": None, "14": 14, "15": 15}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}


@dataclass(frozen=True)
class Table:
    """A two-row table: ``cells`` is each column's ``w:tcW`` as ``(w, type)``."""

    cells: tuple = ((2500, "pct"), (2500, "pct"))
    width: tuple = (5000, "pct")
    #: ``w:tblGrid``; ``None``: every column 1,000 twips.
    grid: tuple | None = None
    margins: tuple = (108, 108)
    jc: str | None = None
    indent: int | None = None
    rows: int = 2
    #: Text after each column's label (``None``: the label alone).
    texts: tuple | None = None

    def xml(self, number: int) -> str:
        borders = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                          for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
        props = f'<w:tblW w:w="{self.width[0]}" w:type="{self.width[1]}"/>'
        if self.jc:
            props += f'<w:jc w:val="{self.jc}"/>'
        if self.indent is not None:
            props += f'<w:tblInd w:w="{self.indent}" w:type="dxa"/>'
        left, right = self.margins
        props += (f"<w:tblBorders>{borders}</w:tblBorders><w:tblCellMar><w:left w:w=\"{left}\" w:type=\"dxa\"/>"
                  f"<w:right w:w=\"{right}\" w:type=\"dxa\"/></w:tblCellMar><w:tblLook w:val=\"0000\"/>")
        grid = self.grid if self.grid is not None else tuple(1000 for _ in self.cells)
        body = ""
        for r in range(self.rows):
            body += "<w:tr>"
            for c, (w, kind) in enumerate(self.cells):
                label = f"{number}.r{r + 1}c{c + 1}"
                if self.texts and self.texts[c]:
                    label += " " + self.texts[c]
                body += (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="{kind}"/></w:tcPr>'
                         + wml.paragraph(wml.run(label), mark={}, spacing=SPACING) + "</w:tc>")
            body += "</w:tr>"
        columns = "".join(f'<w:gridCol w:w="{w}"/>' for w in grid)
        return f"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid>{columns}</w:tblGrid>{body}</w:tbl>"


@dataclass(frozen=True)
class Case:
    family: str
    note: str
    table: Table = field(default_factory=Table)


#: The text column, twips (A4 less 1,442 and 1,300).
COLUMN = 11906 - 1442 - 1300


def even(width: int, count: int) -> tuple:
    share = width // count
    return tuple((share if k < count - 1 else width - share * (count - 1), "pct") for k in range(count))


def _cases() -> tuple[Case, ...]:
    out = []
    last = []  # what the model does not lay out: after everything else, as it stops there
    for width in (5000, 2500, 3500, 1250):
        for count in (2, 3, 4):
            case = Case("width", f"tblW {width} pct, {count} even", Table(even(width, count), (width, "pct")))
            (last if width == 1250 and count > 2 else out).append(case)
    shares = (818, 849, 849, 850, 851, 783)
    out.append(Case("share", "six unequal", Table(tuple((s, "pct") for s in shares))))
    out.append(Case("share", "sum under tblW", Table(((1500, "pct"), (1500, "pct")))))
    out.append(Case("share", "sum over tblW", Table(((3500, "pct"), (3500, "pct")))))
    last.append(Case("share", "one cell dxa", Table(((2500, "pct"), (3000, "dxa")))))
    last.append(Case("share", "tblW auto", Table(((1500, "pct"), (2500, "pct")), (0, "auto"))))
    whole = COLUMN + 216
    out.append(Case("grid", "grid as Word writes it", Table(grid=(whole // 2, whole - whole // 2))))
    out.append(Case("grid", "grid twice that", Table(grid=(whole, whole))))
    out.append(Case("grid", "six unequal, grid as Word writes it",
                    Table(tuple((s, "pct") for s in shares),
                          grid=tuple(round(s * whole / 5000) for s in shares[:-1])
                          + (whole - sum(round(s * whole / 5000) for s in shares[:-1]),))))
    for margins in ((0, 0), (300, 300), (300, 0)):
        out.append(Case("margin", f"cell margins {margins[0]} / {margins[1]}", Table(margins=margins)))
    out.append(Case("place", "3500 pct centred", Table(even(3500, 2), (3500, "pct"), jc="center")))
    out.append(Case("place", "3500 pct right", Table(even(3500, 2), (3500, "pct"), jc="right")))
    out.append(Case("place", "3500 pct tblInd 720", Table(even(3500, 2), (3500, "pct"), indent=720)))
    out.append(Case("dxa", "3500 pct, tcW 2000 dxa", Table(((2000, "dxa"), (2000, "dxa")), (3500, "pct"))))
    out.append(Case("dxa", "3500 pct, tcW 4788 dxa", Table(((4788, "dxa"), (4788, "dxa")), (3500, "pct"))))
    out.append(Case("dxa", "3500 pct centred, tcW 4788 dxa",
                     Table(((4788, "dxa"), (4788, "dxa")), (3500, "pct"), jc="center")))
    out.append(Case("dxa", "5000 pct, tcW 2000 dxa", Table(((2000, "dxa"), (2000, "dxa")), (5000, "pct"))))
    out.append(Case("dxa", "3500 pct, tcW 1000 / 3000 dxa", Table(((1000, "dxa"), (3000, "dxa")), (3500, "pct"))))
    out.append(Case("dxa", "3500 pct, tcW 3000 / 1000 dxa", Table(((3000, "dxa"), (1000, "dxa")), (3500, "pct"))))
    out.append(Case("share", "1000 / 2000 under 5000", Table(((1000, "pct"), (2000, "pct")))))
    out.append(Case("share", "3000 / 1000 under 5000", Table(((3000, "pct"), (1000, "pct")))))
    out.append(Case("share", "1000 / 4500 over 5000", Table(((1000, "pct"), (4500, "pct")))))
    for width in (1500, 1750, 2000, 2250):
        out.append(Case("width", f"tblW {width} pct, 3 even", Table(even(width, 3), (width, "pct"))))
    # A grid that adds up to the table's width in percent, unlike the cells' widths:
    # which does Word draw?
    table = COLUMN + 216
    part = table * 3500 // 5000
    out.append(Case("grid", "3500 pct, tcW 4788 dxa, a grid of the width, unequal",
                    Table(((4788, "dxa"), (4788, "dxa")), (3500, "pct"), grid=(3416, part - 3416))))
    out.append(Case("grid", "3500 pct, tcW 1000 / 3000 dxa, a grid of the width, unequal",
                    Table(((1000, "dxa"), (3000, "dxa")), (3500, "pct"), grid=(3416, part - 3416))))
    out.append(Case("grid", "5000 pct, tcW 2500 pct, a grid of the width, unequal",
                    Table(grid=(3000, table - 3000))))
    # Content: cells in percent share the table as they say whatever is in them; cells in
    # dxa under a width in percent that want more than it do not: each gives up its excess
    # by its width less its narrowest content.
    long_word = "x" * 16
    words = " ".join(["lorem"] * 15)
    for texts, note in (((long_word, ""), "a long word / none"), (("", long_word), "none / a long word"),
                        ((words, ""), "15 words / none"), ((long_word + " " + words, "lorem"),
                                                           "a long word and 15 words / one")):
        out.append(Case("content", f"3500 pct, tcW 2500 pct, {note}", Table(even(3500, 2), (3500, "pct"),
                                                                              texts=texts)))
        out.append(Case("content", f"3500 pct, tcW 4788 dxa, {note}",
                        Table(((4788, "dxa"), (4788, "dxa")), (3500, "pct"), texts=texts)))
    # Where between 782 and 936 twips a column's share stops being what Word draws.
    last.append(Case("width", "tblW 1250 pct, 3 even, margins 0", Table(even(1250, 3), (1250, "pct"),
                                                                        margins=(0, 0))))
    last.append(Case("width", "tblW 1250 pct, 3 even, grid 500", Table(even(1250, 3), (1250, "pct"),
                                                                       grid=(500,) * 3)))
    return tuple(out + [dataclasses.replace(case, family="not modelled") for case in last])


CASES = _cases()


def body() -> str:
    out = ""
    for number, case in enumerate(CASES):
        out += wml.paragraph(wml.run(f"Case {number}"), mark={}, spacing=SPACING, pageBreakBefore=True)
        out += case.table.xml(number)
        out += wml.paragraph(wml.run(f"After {number}"), mark={}, spacing=SPACING)
    return out


def build(setting: str) -> bytes:
    """``setting``: ``none``, ``14`` or ``15``."""
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return anchor_probe.package(body(), "none", extra=extra)


DOCUMENTS = tuple(f"pct-table-{s}" for s in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"pct-table-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
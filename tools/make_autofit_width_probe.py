#!/usr/bin/env python3
"""Autofit tables sized from their content: how wide Word makes each column.

Tables, stage 7 (ROADMAP.md, "Tables -- measured") recorded that Word sizes an autofit
table's columns from their content (``make_table_autofit_probe.py``) and the model stopped
there.  This probe measures the rules: a cell's narrowest and widest content, what a
first-line indent does to them, how a stated width combines with content wider than it,
how cells with no width share the text column, and a table width in ``dxa`` or ``pct``.

Every case is a page: ``Case N``, then a table with borders of ``w:sz`` 4 on every edge
(the grid lines are read from Word's fills), cell margins 108, no ``w:tblLayout``
(autofit) and a ``w:tblGrid`` that is **wrong on purpose** (every column 1,000 twips)
unless the case says otherwise, so what Word draws is what it computed.  Cells hold short
labels (``N.a``) and words of ``x`` -- one ``x`` is 95.3 twips in Calibri 11.  A4, the
left margin off the pixel grid (1,442 twips; the text column 9,164 twips).  Families
(``CASES``):

* ``indent`` -- a first-line indent in a cell whose stated width holds the word but not
  the word and the indent (a filesamples document's calendar: its Normal style indents
  every first line); left and right aligned, a hanging indent, a left indent (on every
  line) that the word and it do not fit, a cell with no ``w:tcW``;
* ``auto`` -- cells with no ``w:tcW``: one to three columns of words that fit, the widest
  of two rows, words that wrap and ``w:noWrap``, cells with their own margins, empty
  cells, a left indent;
* ``widen`` -- cells in ``dxa`` with a word wider than the width;
* ``share`` -- content wider than the text column: cells with no width and with one, many
  words each;
* ``tblW`` -- ``w:tblW`` in ``dxa`` over cells in ``dxa`` (adding up to less and to more)
  and over cells with none; in ``pct`` over cells with none;
* ``span`` -- last, as the layout stops there: a cell across two columns wider than they
  are, with no widths and in ``dxa``.

Three documents: no ``settings.xml``, mode 14 and mode 15.  Reader:
``read_autofit_width_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import make_anchor_probe as anchor_probe
import wml

SETTINGS = {"none": None, "14": 14, "15": 15}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}


@dataclass(frozen=True)
class Cell:
    """``width``: ``w:tcW`` as ``(w, type)`` or ``None``; ``texts``: its paragraphs."""

    texts: tuple = ("",)
    width: tuple | None = None
    span: int = 1
    no_wrap: bool = False
    margins: tuple | None = None
    ind: dict | None = None
    jc: str | None = None

    def xml(self) -> str:
        props = f'<w:tcW w:w="{self.width[0]}" w:type="{self.width[1]}"/>' if self.width else ""
        if self.span > 1:
            props += f'<w:gridSpan w:val="{self.span}"/>'
        if self.no_wrap:
            props += "<w:noWrap/>"
        if self.margins:
            left, right = self.margins
            props += (f'<w:tcMar><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/>'
                      "</w:tcMar>")
        paragraph = {"spacing": SPACING}
        if self.ind:
            paragraph["ind"] = self.ind
        if self.jc:
            paragraph["jc"] = self.jc
        body = "".join(wml.paragraph(wml.run(text) if text else "", mark={}, **paragraph) for text in self.texts)
        return f"<w:tc><w:tcPr>{props}</w:tcPr>{body}</w:tc>"


@dataclass(frozen=True)
class Table:
    rows: tuple = ()
    width: tuple = (0, "auto")
    #: ``w:tblGrid``; ``None``: every column 1,000 twips.
    grid: tuple | None = None
    #: The borders' ``w:sz``; 0: none.
    border: int = 4

    def columns(self) -> int:
        return max(sum(cell.span for cell in row) for row in self.rows)

    def xml(self) -> str:
        borders = "".join((f'<w:{s} w:val="single" w:sz="{self.border}" w:space="0" w:color="000000"/>'
                           if self.border else f'<w:{s} w:val="nil"/>')
                          for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
        props = (f'<w:tblW w:w="{self.width[0]}" w:type="{self.width[1]}"/><w:tblBorders>{borders}</w:tblBorders>'
                 '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
                 '<w:tblLook w:val="0000"/>')
        grid = self.grid if self.grid is not None else (1000,) * self.columns()
        columns = "".join(f'<w:gridCol w:w="{w}"/>' for w in grid)
        body = "".join("<w:tr>" + "".join(cell.xml() for cell in row) + "</w:tr>" for row in self.rows)
        return f"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid>{columns}</w:tblGrid>{body}</w:tbl>"


@dataclass(frozen=True)
class Case:
    family: str
    note: str
    table: Table = field(default_factory=Table)


def x(n: int) -> str:
    return "x" * n


def dxa(w: int) -> tuple:
    return (w, "dxa")


def words(n: int, size: int = 4) -> str:
    return " ".join([x(size)] * n)


def _cases() -> tuple[Case, ...]:
    out: list[Case] = []

    def add(family: str, note: str, *rows, width=(0, "auto"), grid=None, border=4) -> None:
        number = len(out)
        labelled = []
        for r, row in enumerate(rows):
            cells = []
            for c, cell in enumerate(row):
                label = f"{number}.{'abcdefgh'[c]}{r + 1}"
                texts = tuple("" if text is None else f"{label} {text}".strip() if k == 0 else text
                              for k, text in enumerate(cell.texts))
                cells.append(Cell(texts, cell.width, cell.span, cell.no_wrap, cell.margins, cell.ind, cell.jc))
            labelled.append(tuple(cells))
        out.append(Case(family, note, Table(tuple(labelled), width, grid, border)))

    first = {"firstLine": 720}
    # indent: 1,440-twip cells (1,224 of text); x(10) is 953 twips, x(10) and 720 of indent
    # are not.
    for jc in (None, "right"):
        add("indent", f"firstLine 720, a word that fits without it, {jc or 'left'}",
            (Cell((x(10),), dxa(1440), ind=first, jc=jc), Cell(("",), dxa(1440), ind=first, jc=jc)),
            (Cell(("",), dxa(1440), ind=first, jc=jc), Cell((x(10),), dxa(1440), ind=first, jc=jc)))
    add("indent", "firstLine 720, noWrap and margins 43, right",
        (Cell((x(10),), dxa(1440), no_wrap=True, margins=(43, 43), ind=first, jc="right"),
         Cell((x(3),), dxa(300), no_wrap=True, margins=(43, 43), ind=first, jc="right")))
    add("indent", "hanging 720 (left 720), a word that fits the first line",
        (Cell((x(10),), dxa(1440), ind={"left": 720, "hanging": 720}), Cell(("",), dxa(1440))))
    add("indent", "left 720 on every line: the word and it do not fit",
        (Cell((x(10),), dxa(1440), ind={"left": 720}), Cell(("",), dxa(1440))))
    add("indent", "firstLine 720, no tcW",
        (Cell((x(10),), ind=first), Cell((x(3),))))
    add("indent", "firstLine 1440 wider than the text",
        (Cell((x(6),), dxa(1440), ind={"firstLine": 1440}), Cell(("",), dxa(1440))))
    # auto: cells with no w:tcW whose content fits the column.
    add("auto", "one column", (Cell((x(5),)),))
    add("auto", "three columns", (Cell((x(1),)), Cell((x(12),)), Cell((words(3),))))
    add("auto", "widest of two rows", (Cell((x(2),)), Cell((x(9),))), (Cell((x(12),)), Cell((x(1),))))
    add("auto", "noWrap words", (Cell((words(6),), no_wrap=True), Cell((x(3),))))
    add("auto", "two paragraphs", (Cell((x(3), words(4))), Cell((x(3),))))
    add("auto", "cell margins 43 and 300", (Cell((x(6),), margins=(43, 43)), Cell((x(6),), margins=(300, 300))))
    add("auto", "empty cells", (Cell((None,)), Cell((None,)), Cell((x(4),))))
    add("auto", "left and right indents", (Cell((x(6),), ind={"left": 360, "right": 240}), Cell((x(3),))))
    add("auto", "one cell in dxa", (Cell((x(3),), dxa(2000)), Cell((x(6),))))
    # widen: a word wider than the stated width.
    add("widen", "middle of three, 1440 each", (Cell((x(2),), dxa(1440)), Cell((x(20),), dxa(1440)),
                                                Cell((x(2),), dxa(1440))))
    add("widen", "two rows, the word in the second", (Cell((x(2),), dxa(1440)), Cell((x(2),), dxa(1440))),
        (Cell((x(24),), dxa(1440)), Cell((x(2),), dxa(1440))))
    add("widen", "narrow empty spacers kept", (Cell((x(2),), dxa(864)), Cell((None,), dxa(222)),
                                               Cell((x(2),), dxa(864))))
    add("widen", "spacer narrower than its margins", (Cell((x(2),), dxa(864)), Cell((None,), dxa(150)),
                                                      Cell((x(2),), dxa(864))))
    add("widen", "wider than the column with the others", (Cell((x(2),), dxa(4000)), Cell((x(60),), dxa(2000)),
                                                           Cell((x(2),), dxa(3000))))
    # share: content wider than the text column.
    add("share", "two of 60 words", (Cell((words(60),)), Cell((words(60),))))
    add("share", "60 words and 10", (Cell((words(60),)), Cell((words(10),))))
    add("share", "60 words and 20 words", (Cell((words(60),)), Cell((words(20),))))
    add("share", "60 words and a long word", (Cell((words(60),)), Cell((x(25),))))
    add("share", "60 words, three columns", (Cell((words(60),)), Cell((words(15),)), Cell((x(3),))))
    add("share", "60 words beside 3000 dxa of 60", (Cell((words(60),), dxa(3000)), Cell((words(60),))))
    add("share", "dxa adding up to more than the column", (Cell((x(3),), dxa(4000)), Cell((x(3),), dxa(4000)),
                                                           Cell((x(3),), dxa(4000))))
    # tblW
    add("tblW", "6000 dxa over 2000 / 2000 dxa", (Cell((x(2),), dxa(2000)), Cell((x(2),), dxa(2000))),
        width=dxa(6000))
    add("tblW", "6000 dxa over 1000 / 3000 dxa", (Cell((x(2),), dxa(1000)), Cell((x(2),), dxa(3000))),
        width=dxa(6000))
    add("tblW", "3000 dxa over 2000 / 2000 dxa", (Cell((x(2),), dxa(2000)), Cell((x(2),), dxa(2000))),
        width=dxa(3000))
    add("tblW", "6000 dxa over no widths", (Cell((x(2),)), Cell((x(8),))), width=dxa(6000))
    add("tblW", "6000 dxa over 1000 dxa and none", (Cell((x(2),), dxa(1000)), Cell((x(2),))), width=dxa(6000))
    add("tblW", "5000 pct over no widths", (Cell((x(2),)), Cell((x(8),))), width=(5000, "pct"))
    add("tblW", "2500 pct over no widths", (Cell((x(2),)), Cell((x(8),))), width=(2500, "pct"))
    add("tblW", "5000 pct over 1000 dxa and none", (Cell((x(2),), dxa(1000)), Cell((x(2),))), width=(5000, "pct"))
    # span: last, as the layout stops at a cell across columns wider than they are.
    add("span", "a cell across two, wider than they are", (Cell((x(30),), span=2), Cell((x(2),))),
        (Cell((x(2),)), Cell((x(4),)), Cell((x(2),))))
    add("span", "in dxa, across two of dxa", (Cell((x(30),), dxa(2000), span=2), Cell((x(2),), dxa(1000))),
        (Cell((x(2),), dxa(1000)), Cell((x(2),), dxa(1000)), Cell((x(2),), dxa(1000))))
    # Round two of span: what the uneven share follows -- the borders, the columns' content,
    # how many columns, which, how much wider.
    add("span", "across two of equal content", (Cell((x(30),), span=2), Cell((x(2),))),
        (Cell((x(3),)), Cell((x(3),)), Cell((x(2),))))
    add("span", "borders of 24", (Cell((x(30),), span=2), Cell((x(2),))),
        (Cell((x(2),)), Cell((x(4),)), Cell((x(2),))), border=24)
    add("span", "no borders", (Cell((x(30),), span=2), Cell((x(2),))),
        (Cell((x(2),)), Cell((x(4),)), Cell((x(2),))), border=0)
    add("span", "across three", (Cell((x(40),), span=3), Cell((x(2),))),
        (Cell((x(2),)), Cell((x(4),)), Cell((x(6),)), Cell((x(2),))))
    add("span", "across the second and third", (Cell((x(2),)), Cell((x(30),), span=2)),
        (Cell((x(2),)), Cell((x(4),)), Cell((x(2),))))
    add("span", "much wider", (Cell((x(45),), span=2), Cell((x(2),))),
        (Cell((x(2),)), Cell((x(4),)), Cell((x(2),))))
    add("span", "in dxa, across 1000 and 2000", (Cell((x(30),), dxa(3000), span=2), Cell((x(2),), dxa(1000))),
        (Cell((x(2),), dxa(1000)), Cell((x(2),), dxa(2000)), Cell((x(2),), dxa(1000))))
    return tuple(out)


CASES = _cases()


def body() -> str:
    out = ""
    for number, case in enumerate(CASES):
        out += wml.paragraph(wml.run(f"Case {number}"), mark={}, spacing=SPACING, pageBreakBefore=True)
        out += case.table.xml()
        out += wml.paragraph(wml.run(f"After {number}"), mark={}, spacing=SPACING)
    return out


def build(setting: str) -> bytes:
    """``setting``: ``none``, ``14`` or ``15``."""
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return anchor_probe.package(body(), "none", extra=extra)


DOCUMENTS = tuple(f"autofit-width-{s}" for s in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"autofit-width-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

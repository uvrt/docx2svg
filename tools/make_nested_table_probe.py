#!/usr/bin/env python3
"""A table nested in a table cell: where Word puts it, and how tall it makes the row.

A table in a cell stopped the layout (Tables, "What stays an obstacle").  Every case is a
page: ``Case N``, then an outer table of one row (or two) and two columns of 4,500 twips
-- single borders of ``w:sz`` 4, cell margins 108 -- whose first cell holds the nested
table between the paragraphs the case asks for (a cell must end in a paragraph, so one
always follows the nested table: ``N.b``), and whose second cell holds ``N.r1c2``; then
``After N``.  The nested table is two rows of two columns, labelled ``N.n.r<row>c<col>``,
borders of ``w:sz`` 4, margins 108, 1,500-twip columns in ``dxa`` unless the case says
otherwise.  Each cell's first glyph says where the nested table and its rows went, and the
line after the outer table how tall its row was.  The section is A4 with the left margin
off the pixel grid (1,442 twips), as the wrap probes'.  Families (``CASES``):

* ``place`` -- the nested table first in its cell, after a paragraph, after two; the
  paragraph after it with text and empty; ``w:tblInd`` 360 and -200; ``w:jc`` centre and
  right; the outer cell's margins 300; the nested table's margins 0; no borders on it;
* ``space`` -- 240 twips of space before and of space after on the paragraphs around it,
  and on the nested table's own paragraphs;
* ``width`` -- the nested table in percent: 5000 pct, and 4000 pct centred, over cells in
  percent;
* ``row`` -- the nested table taller than the other cell, and shorter; the outer row with
  ``w:trHeight`` 1,079 twips (``atLeast``); two outer rows, the nested table in the second;
  the nested table with a cell merged down two rows (``w:vMerge``);
* ``sample`` -- last, as the layout stops at the first: the nested table 5000 and 4000 pct
  (centred) over cells in ``dxa``, and as a filesamples document nests one: the outer
  table 3500 pct centred over
  cells of 4,788 twips, ``w:trHeight`` 1,079, the nested table 4000 pct centred over cells
  of 1,678 and 1,629 twips with a cell merged down, every paragraph 240 before and no
  first-line indent, and the empty paragraph after it.

Three documents: no ``settings.xml``, mode 14 and mode 15.  Reader:
``read_nested_table_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import make_anchor_probe as anchor_probe
import wml

SETTINGS = {"none": None, "14": 14, "15": 15}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}


def borders(size: int) -> str:
    if not size:
        return ""
    return "<w:tblBorders>" + "".join(f'<w:{s} w:val="single" w:sz="{size}" w:space="0" w:color="000000"/>'
                                      for s in ("top", "left", "bottom", "right", "insideH", "insideV")) \
        + "</w:tblBorders>"


def paragraph(text: str, spacing: dict | None = None) -> str:
    return wml.paragraph(wml.run(text) if text else "", mark={}, spacing=spacing or SPACING)


def table_properties(width: tuple, *, jc: str | None = None, indent: int | None = None, margins: int = 108,
                     border: int = 4) -> str:
    props = f'<w:tblW w:w="{width[0]}" w:type="{width[1]}"/>'
    if jc:
        props += f'<w:jc w:val="{jc}"/>'
    if indent is not None:
        props += f'<w:tblInd w:w="{indent}" w:type="dxa"/>'
    return (props + borders(border) + f'<w:tblCellMar><w:left w:w="{margins}" w:type="dxa"/>'
            f'<w:right w:w="{margins}" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/>')


@dataclass(frozen=True)
class Nested:
    """The nested table: ``cells`` each column's ``w:tcW``."""

    cells: tuple = ((1500, "dxa"), (1500, "dxa"))
    width: tuple = (0, "auto")
    jc: str | None = None
    indent: int | None = None
    margins: int = 108
    border: int = 4
    rows: int = 2
    #: Its paragraphs' spacing.
    spacing: dict | None = None
    #: The first column's cell merged down every row.
    merged: bool = False

    def xml(self, number: int) -> str:
        grid = "".join(f'<w:gridCol w:w="{w if kind == "dxa" else 1500}"/>' for w, kind in self.cells)
        body = ""
        for r in range(self.rows):
            body += "<w:tr>"
            for c, (w, kind) in enumerate(self.cells):
                merge = ""
                if self.merged and c == 0:
                    merge = '<w:vMerge w:val="restart"/>' if r == 0 else "<w:vMerge/>"
                text = f"{number}.n.r{r + 1}c{c + 1}" if not (self.merged and c == 0 and r) else ""
                body += (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="{kind}"/>{merge}</w:tcPr>'
                         + paragraph(text, self.spacing) + "</w:tc>")
            body += "</w:tr>"
        props = table_properties(self.width, jc=self.jc, indent=self.indent, margins=self.margins, border=self.border)
        return f"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>"


@dataclass(frozen=True)
class Case:
    family: str
    note: str
    nested: Nested = field(default_factory=Nested)
    #: Paragraphs before the nested table in its cell.
    before: int = 0
    #: The paragraph after it: its text ("" empty).
    after: str = "b"
    #: Spacing of the paragraphs around it in the cell.
    around: dict | None = None
    outer_margins: int = 108
    outer_width: tuple = (0, "auto")
    outer_cells: tuple = ((4500, "dxa"), (4500, "dxa"))
    outer_jc: str | None = None
    #: The outer row's w:trHeight (atLeast).
    height: int | None = None
    #: Rows of the outer table; the nested table is in the last.
    outer_rows: int = 1
    #: Lines of text in the outer table's second cell.
    lines: int = 1
    #: The outer table's w:tblGrid (``None``: its cells' widths).
    outer_grid: tuple | None = None


def _cases() -> tuple[Case, ...]:
    out = []
    out.append(Case("place", "first in its cell"))
    out.append(Case("place", "after a paragraph", before=1))
    out.append(Case("place", "after two paragraphs", before=2))
    out.append(Case("place", "the paragraph after it empty", after=""))
    out.append(Case("place", "tblInd 360", Nested(indent=360)))
    out.append(Case("place", "tblInd -200", Nested(indent=-200)))
    out.append(Case("place", "jc center", Nested(jc="center")))
    out.append(Case("place", "jc right", Nested(jc="right")))
    out.append(Case("place", "outer margins 300", outer_margins=300))
    out.append(Case("place", "nested margins 0", Nested(margins=0)))
    out.append(Case("place", "nested without borders", Nested(border=0)))
    before = {**SPACING, "before": 240}
    after = {**SPACING, "after": 240}
    out.append(Case("space", "240 before the paragraphs around it", before=1, around=before))
    out.append(Case("space", "240 after the paragraphs around it", before=1, around=after))
    out.append(Case("space", "240 before its own paragraphs", Nested(spacing=before), before=1))
    out.append(Case("space", "240 after its own paragraphs", Nested(spacing=after), before=1))
    out.append(Case("space", "240 before everything", Nested(spacing=before), before=1, around=before, after=""))
    out.append(Case("width", "5000 pct over cells in percent", Nested(((2500, "pct"), (2500, "pct")),
                                                                     (5000, "pct"))))
    out.append(Case("width", "4000 pct centred over cells in percent", Nested(((2000, "pct"), (2000, "pct")),
                                                                             (4000, "pct"), jc="center")))
    out.append(Case("row", "shorter than the other cell", lines=6))
    out.append(Case("row", "taller than the other cell", Nested(rows=4)))
    out.append(Case("row", "trHeight 1079", height=1079))
    out.append(Case("row", "trHeight 2500", height=2500))
    out.append(Case("row", "in the second of two rows", outer_rows=2))
    out.append(Case("row", "a cell merged down", Nested(merged=True)))
    # Last, as the layout stops at the first: widths in percent over cells in dxa, which
    # Word shares by their content (Tables, stage 7a).
    out.append(Case("sample", "5000 pct over cells in dxa", Nested(width=(5000, "pct"))))
    out.append(Case("sample", "4000 pct centred over cells in dxa", Nested(width=(4000, "pct"), jc="center")))
    out.append(Case("sample", "4000 pct centred, cells 1678 / 1629", Nested(((1678, "dxa"), (1629, "dxa")),
                                                                          (4000, "pct"), jc="center")))
    sample = {**SPACING, "before": 240}
    out.append(Case("sample", "as a filesamples document nests one",
                    Nested(((1678, "dxa"), (1629, "dxa")), (4000, "pct"), jc="center", spacing=sample, merged=True),
                    after="", around=sample, outer_width=(3500, "pct"), outer_cells=((4788, "dxa"), (4788, "dxa")),
                    outer_jc="center", height=1079))
    # The outer table's grid as Word writes it for that table (adding up to its width in
    # percent, 3500 pct of the column and both margins), unequal: Word draws what?
    out.append(Case("sample", "as a filesamples document nests one, its grid of the width",
                    Nested(((1678, "dxa"), (1629, "dxa")), (4000, "pct"), jc="center", spacing=sample, merged=True),
                    after="", around=sample, outer_width=(3500, "pct"), outer_cells=((4788, "dxa"), (4788, "dxa")),
                    outer_jc="center", height=1079, outer_grid=(3416, 3150)))
    out.append(Case("sample", "outer 3500 pct over 4788 dxa, grid of the width, a nested table 3000 dxa",
                    outer_width=(3500, "pct"), outer_cells=((4788, "dxa"), (4788, "dxa")), outer_jc="center",
                    outer_grid=(3416, 3150)))
    return tuple(out)


CASES = _cases()


def outer(number: int, case: Case) -> str:
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in (case.outer_grid or [w for w, _kind in case.outer_cells]))
    props = table_properties(case.outer_width, jc=case.outer_jc, margins=case.outer_margins)
    rows = ""
    for r in range(case.outer_rows):
        last = r == case.outer_rows - 1
        height = f'<w:trPr><w:trHeight w:val="{case.height}"/></w:trPr>' if case.height and last else ""
        first = ""
        if last:
            first = "".join(paragraph(f"{number}.a{k + 1}", case.around) for k in range(case.before))
            first += case.nested.xml(number) + paragraph(f"{number}.{case.after}" if case.after else "", case.around)
        else:
            first = paragraph(f"{number}.r{r + 1}c1")
        second = "".join(paragraph(f"{number}.r{r + 1}c2" + (f" line {k + 1}" if k else ""), case.around)
                         for k in range(case.lines if last else 1))
        (w1, k1), (w2, k2) = case.outer_cells
        rows += (f"<w:tr>{height}<w:tc><w:tcPr><w:tcW w:w=\"{w1}\" w:type=\"{k1}\"/></w:tcPr>{first}</w:tc>"
                 f"<w:tc><w:tcPr><w:tcW w:w=\"{w2}\" w:type=\"{k2}\"/></w:tcPr>{second}</w:tc></w:tr>")
    return f"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>"


def body() -> str:
    out = ""
    for number, case in enumerate(CASES):
        out += wml.paragraph(wml.run(f"Case {number}"), mark={}, spacing=SPACING, pageBreakBefore=True)
        out += outer(number, case)
        out += wml.paragraph(wml.run(f"After {number}"), mark={}, spacing=SPACING)
    return out


def build(setting: str) -> bytes:
    """``setting``: ``none``, ``14`` or ``15``."""
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return anchor_probe.package(body(), "none", extra=extra)


DOCUMENTS = tuple(f"nested-table-{s}" for s in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"nested-table-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)
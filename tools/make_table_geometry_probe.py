#!/usr/bin/env python3
"""Where Word puts a table's cells: its grid, indent, alignment, margins and borders.

Tables, stage 1 (ROADMAP.md, "Tables -- measured"): fixed layout, one row per case, so
nothing but horizontal geometry varies.  Every cell holds two paragraphs: a
left-aligned word ``L<case>.<cell>`` -- whose first glyph's pen x is where the cell's
text starts -- and a right-aligned word ``R<case>.<cell>`` -- whose last glyph ends where
a line of the cell may end, so the right edge of the line budget is read off the
drawing as well.  Borders are drawn, so Word's fills show the grid lines.

Each case varies one thing from the base (fixed layout, three 2000-twip columns, direct
``w:tblCellMar`` of 108 left and right, single ``w:sz`` 4 borders on every side, no
indent, no alignment):

* ``margins`` -- left/right 0, 10, 50, 200, and 50/150; none declared at all;
* ``border`` -- ``w:sz`` 2, 3, 5, 6, 12, 18, 24, 27 on every side; none; a thick outer
  left, inner vertical or outer right border alone;
* ``indent`` -- ``w:tblInd`` 0, 100, 333, -200 and 720; 333 with no borders and no
  margins; declared by the table's style only, and overridden by the table;
* ``jc`` -- centred and right-aligned, with and without an indent, with no borders and
  with thick ones;
* ``widths`` -- 1000/2500/3333, six of 1111, 1001/1999, and a table wider than the
  column;
* ``tcMar`` -- the middle cell's own margins (300 left, 20 right);
* ``tcW`` -- cells stating 1000/3000/2000 over a grid of 2000s; and **autofit**
  (``w:tblLayout`` left out) -- as the base, with those cells, and with an indent: the
  real documents' tables are all autofit, with every cell's width stated.

One document per compatibility setting: no ``settings.xml``, and modes 12, 14 and 15.
Calibri 11 pt everywhere, no paragraph spacing.  The reader is
``read_table_geometry_probe.py``.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}
SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")


def side(name: str, sz: int, val: str = "single") -> str:
    return f'<w:{name} w:val="{val}" w:sz="{sz}" w:space="0" w:color="000000"/>'


def borders(sz: int | dict) -> str:
    """Every side at ``sz``, or per side from a dict (a side left out: ``w:sz`` 4)."""
    sizes = sz if isinstance(sz, dict) else {name: sz for name in SIDES}
    return "".join(side(name, sizes.get(name, 4)) for name in SIDES if sizes.get(name, 4))


def margins(left: int, right: int) -> str:
    return (f'<w:top w:w="0" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/>'
            f'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/>')


BASE = {"widths": (2000, 2000, 2000), "margins": (108, 108), "border": 4, "indent": None, "jc": None,
        "style": None, "cell_margins": None, "layout": "fixed", "tcw": None}


def _cases() -> list[tuple[str, dict]]:
    out = [("base", {})]
    for left, right in ((0, 0), (10, 10), (50, 50), (200, 200), (50, 150)):
        out.append((f"margins {left}/{right}", {"margins": (left, right)}))
    out.append(("margins none", {"margins": None}))
    out.append(("margins none, no borders", {"margins": None, "border": 0}))
    for sz in (2, 3, 5, 6, 12, 18, 24, 27):
        out.append((f"border {sz}", {"border": sz}))
    out.append(("border none", {"border": 0}))
    out.append(("border outer left 24", {"border": {"left": 24}}))
    out.append(("border inside 24", {"border": {"insideV": 24}}))
    out.append(("border outer right 24", {"border": {"right": 24}}))
    for indent in (0, 100, 333, -200, 720):
        out.append((f"indent {indent}", {"indent": indent}))
    out.append(("indent 333, no borders, no margins", {"indent": 333, "border": 0, "margins": (0, 0)}))
    out.append(("indent by style", {"style": "Indented"}))
    out.append(("indent by style, table 400", {"style": "Indented", "indent": 400}))
    for jc in ("center", "right"):
        out.append((f"jc {jc}", {"jc": jc}))
        out.append((f"jc {jc}, indent 333", {"jc": jc, "indent": 333}))
        out.append((f"jc {jc}, no borders", {"jc": jc, "border": 0}))
        out.append((f"jc {jc}, border 24", {"jc": jc, "border": 24}))
    out.append(("widths 1000/2500/3333", {"widths": (1000, 2500, 3333)}))
    out.append(("widths six 1111", {"widths": (1111,) * 6}))
    out.append(("widths 1001/1999", {"widths": (1001, 1999)}))
    out.append(("widths wider than the column", {"widths": (3000, 3000, 3000, 3000)}))
    out.append(("tcMar middle 300/20", {"cell_margins": (1, 300, 20)}))
    out.append(("tcW 1000/3000/2000 over the grid", {"tcw": (1000, 3000, 2000)}))
    out.append(("autofit", {"layout": None}))
    out.append(("autofit, tcW 1000/3000/2000 over the grid", {"layout": None, "tcw": (1000, 3000, 2000)}))
    out.append(("autofit, indent 333, jc none", {"layout": None, "indent": 333}))
    return out


CASES = _cases()


def _paragraph(text: str, jc: str | None = None) -> str:
    props = {"jc": jc} if jc else {}
    return wml.paragraph(wml.run(text), **props)


def table(number: int, case: dict) -> str:
    c = dict(BASE, **case)
    widths = c["widths"]
    pr = f'<w:tblStyle w:val="{c["style"]}"/>' if c["style"] else ""
    pr += f'<w:tblW w:w="{sum(widths)}" w:type="dxa"/>'
    if c["jc"]:
        pr += f'<w:jc w:val="{c["jc"]}"/>'
    if c["indent"] is not None:
        pr += f'<w:tblInd w:w="{c["indent"]}" w:type="dxa"/>'
    if c["border"]:
        pr += f"<w:tblBorders>{borders(c['border'])}</w:tblBorders>"
    if c["layout"]:
        pr += f'<w:tblLayout w:type="{c["layout"]}"/>'
    if c["margins"] is not None:
        pr += f"<w:tblCellMar>{margins(*c['margins'])}</w:tblCellMar>"
    pr += ('<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
           ' w:noHBand="1" w:noVBand="1"/>')
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    cells = ""
    for index, width in enumerate(c["tcw"] or widths):
        own = ""
        if c["cell_margins"] and c["cell_margins"][0] == index:
            _, left, right = c["cell_margins"]
            own = f'<w:tcMar><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>'
        cells += (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{own}</w:tcPr>'
                  + _paragraph(f"L{number}.{index}") + _paragraph(f"R{number}.{index}", "right") + "</w:tc>")
    return f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>"


#: Cases to a page: a page break before every twelfth keeps every table on one page.
PER_PAGE = 12


def body() -> str:
    out = ""
    for number, (_, case) in enumerate(CASES):
        props = {"pageBreakBefore": True} if number and number % PER_PAGE == 0 else {}
        out += wml.paragraph(wml.run(f"Case {number}"), **props) + table(number, case)
    return out + _paragraph("End")


def styles() -> str:
    indented = (
        '<w:style w:type="table" w:styleId="Indented"><w:name w:val="Indented"/>'
        '<w:tblPr><w:tblInd w:w="200" w:type="dxa"/><w:tblCellMar>' + margins(60, 60)
        + "</w:tblCellMar></w:tblPr></w:style>")
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True), indented],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package(body(), styles=styles(), extra_parts=extra)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"table-geometry-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

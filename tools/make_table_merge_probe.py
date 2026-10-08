#!/usr/bin/env python3
"""Merged cells: a cell spanning grid columns (``w:gridSpan``), cells merged down rows
(``w:vMerge``), and rows that start or end off the grid (``w:gridBefore`` /
``w:gridAfter``).

Tables, stage 4 (ROADMAP.md, "Tables -- measured").  Every cell holds a left-aligned
word and a right-aligned one, so its text start and line end are both on the drawing;
cells beside a merge hold one or three lines, so the rows' heights are known.  Cases:

* ``span`` -- a cell over two of three columns beside a single one, and under it three
  single cells; a cell over all three; spans in both rows, offset;
* ``vMerge`` -- the first column merged down three rows of three-line cells, holding one
  line (top, centred and bottom); holding seven lines, taller than the three rows (the
  rows grow -- which?); merged down two of three rows;
* ``vMerge and span`` -- a cell over two columns merged down two rows;
* ``grid before / after`` -- a row starting one column in (``w:gridBefore`` with its
  ``w:wBefore``), one ending a column short, under full rows.

Fixed layout, three 2000-twip columns (four for ``grid before``), borders ``w:sz`` 4,
margins 108; one document per compatibility setting (none, 12, 14, 15).  Reader:
``read_table_probes.py``.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}
PER_PAGE = 3
WIDTH = 2000


def _side(name: str) -> str:
    return f'<w:{name} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'


BORDERS = "".join(_side(name) for name in ("top", "left", "bottom", "right", "insideH", "insideV"))


def _cell(name: str, span: int = 1, merge: str | None = None, lines: int = 1, align: str | None = None) -> str:
    """A cell of ``span`` columns: ``lines`` left-aligned words and a right-aligned one."""
    props = f'<w:tcW w:w="{WIDTH * span}" w:type="dxa"/>'
    if span > 1:
        props += f'<w:gridSpan w:val="{span}"/>'
    if merge == "restart":
        props += '<w:vMerge w:val="restart"/>'
    elif merge == "continue":
        props += "<w:vMerge/>"
    if align:
        props += f'<w:vAlign w:val="{align}"/>'
    if merge == "continue":
        content = wml.paragraph("")
    else:
        content = "".join(wml.paragraph(wml.run(f"{name}L{k}")) for k in range(lines))
        content += wml.paragraph(wml.run(f"{name}R"), jc="right")
    return f"<w:tc><w:tcPr>{props}</w:tcPr>{content}</w:tc>"


def _row(cells: str, before: int = 0, after: int = 0) -> str:
    pr = ""
    if before:
        pr += f'<w:gridBefore w:val="{before}"/><w:wBefore w:w="{WIDTH * before}" w:type="dxa"/>'
    if after:
        pr += f'<w:gridAfter w:val="{after}"/><w:wAfter w:w="{WIDTH * after}" w:type="dxa"/>'
    return f"<w:tr>{f'<w:trPr>{pr}</w:trPr>' if pr else ''}{cells}</w:tr>"


def _table(rows: list[str], columns: int = 3) -> str:
    pr = (f'<w:tblW w:w="{WIDTH * columns}" w:type="dxa"/><w:tblBorders>{BORDERS}</w:tblBorders>'
          '<w:tblLayout w:type="fixed"/>'
          '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
          '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
          '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
          ' w:noHBand="1" w:noVBand="1"/>')
    grid = f'<w:gridCol w:w="{WIDTH}"/>' * columns
    return f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{''.join(rows)}</w:tbl>"


def _cases() -> list[tuple[str, str]]:
    out = []
    t = "A"
    out.append(("span 2 + 1", _table([
        _row(_cell(f"{t}a", 2) + _cell(f"{t}b")),
        _row(_cell(f"{t}c") + _cell(f"{t}d") + _cell(f"{t}e"))])))
    t = "B"
    out.append(("span 3", _table([
        _row(_cell(f"{t}a", 3)),
        _row(_cell(f"{t}b") + _cell(f"{t}c") + _cell(f"{t}d"))])))
    t = "C"
    out.append(("span offset", _table([
        _row(_cell(f"{t}a") + _cell(f"{t}b", 2)),
        _row(_cell(f"{t}c", 2) + _cell(f"{t}d"))])))
    for align in (None, "center", "bottom"):
        t = {"center": "E", "bottom": "F"}.get(align, "D")
        out.append((f"vMerge 3 rows, one line{', ' + align if align else ''}", _table([
            _row(_cell(f"{t}a", merge="restart", align=align) + _cell(f"{t}b", lines=3) + _cell(f"{t}c")),
            _row(_cell(f"{t}x", merge="continue", align=align) + _cell(f"{t}d", lines=3) + _cell(f"{t}e")),
            _row(_cell(f"{t}y", merge="continue", align=align) + _cell(f"{t}f", lines=3) + _cell(f"{t}g"))])))
    t = "G"
    out.append(("vMerge 3 rows, taller than they are", _table([
        _row(_cell(f"{t}a", merge="restart", lines=7) + _cell(f"{t}b") + _cell(f"{t}c")),
        _row(_cell(f"{t}x", merge="continue") + _cell(f"{t}d") + _cell(f"{t}e")),
        _row(_cell(f"{t}y", merge="continue") + _cell(f"{t}f") + _cell(f"{t}g"))])))
    t = "H"
    out.append(("vMerge 2 of 3 rows", _table([
        _row(_cell(f"{t}a", merge="restart", lines=4) + _cell(f"{t}b") + _cell(f"{t}c")),
        _row(_cell(f"{t}x", merge="continue") + _cell(f"{t}d") + _cell(f"{t}e")),
        _row(_cell(f"{t}f") + _cell(f"{t}g") + _cell(f"{t}h"))])))
    t = "I"
    out.append(("vMerge and span", _table([
        _row(_cell(f"{t}a", 2, merge="restart") + _cell(f"{t}b", lines=2)),
        _row(_cell(f"{t}x", 2, merge="continue") + _cell(f"{t}c", lines=2)),
        _row(_cell(f"{t}d") + _cell(f"{t}e") + _cell(f"{t}f"))])))
    t = "J"
    out.append(("grid before / after", _table([
        _row(_cell(f"{t}a") + _cell(f"{t}b") + _cell(f"{t}c") + _cell(f"{t}d")),
        _row(_cell(f"{t}e") + _cell(f"{t}f") + _cell(f"{t}g"), before=1),
        _row(_cell(f"{t}h") + _cell(f"{t}i"), after=2)], columns=4)))
    return out


CASES = _cases()


def body() -> str:
    out = ""
    for number, (_, table) in enumerate(CASES):
        props = {"pageBreakBefore": True} if number and number % PER_PAGE == 0 else {}
        out += wml.paragraph(wml.run(f"Case {number}"), **props) + table + wml.paragraph(wml.run(f"After {number}"))
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

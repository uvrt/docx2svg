#!/usr/bin/env python3
"""Rows that do not fit their ``w:tblGrid``: more cells (or more columns spanned) than the
grid has, and fewer.

An edited file can hold such a row -- a row copied from a table with another grid, a
cell inserted without a grid column -- and Word opens it.  What Word draws is measured
here (ROADMAP.md, "Tables -- measured", "a row off its grid").  Every cell holds a
left-aligned word and a right-aligned one, so its text start and line end are both on
the drawing.  Cases, three 2000-twip grid columns unless said:

* ``one more`` -- a row of four cells of 2000 between rows of three;
* ``one more, narrow`` -- the fourth cell 1200 wide;
* ``span past`` -- a row of 2000, a cell over two columns (4000) and 2000: four columns'
  worth;
* ``no width`` -- the fourth cell with no ``w:tcW``;
* ``first row`` -- the first row has the fourth cell, the rows under it three;
* ``fewer`` -- a row of two cells under rows of three (no ``w:gridAfter``);
* ``two more`` -- a row of five cells of 2000;
* ``autofit`` -- ``one more`` with no ``w:tblLayout`` (autofit);
* ``one more, narrow`` with cell margins 0, and 300, with borders of 3 pt, and none (margins 108
  and 0); ``no width`` with margins 0;
* ``one more`` with no table width, 9000 twips, and 100 %;
* ``grid of four`` -- a four-column grid that fits, the table width 6000 (the control);
* ``one more, narrow`` autofit with no table width;
* ``tracked merge`` -- a merge tracked as Word tracks one (the cell merged away kept, with
  ``w:cellDel``), fixed in 6,000 twips and autofit with no width: Word's final view draws
  the deleted cell's room (its text hidden) on a grid widened to the row.

Borders ``w:sz`` 4, margins 108; one document per compatibility setting (none, 15).
Reader: ``read_grid_mismatch_probe.py``.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "15": 15}
PER_PAGE = 3
WIDTH = 2000


def _side(name: str, size: int = 4) -> str:
    return f'<w:{name} w:val="single" w:sz="{size}" w:space="0" w:color="000000"/>'


def _borders(size: int = 4) -> str:
    if not size:
        return "".join(f'<w:{name} w:val="nil"/>' for name in ("top", "left", "bottom", "right", "insideH", "insideV"))
    return "".join(_side(name, size) for name in ("top", "left", "bottom", "right", "insideH", "insideV"))


BORDERS = _borders()


#: A tracked change's attributes, as Word writes them.
TRACKED = 'w:author="Probe" w:date="2026-11-12T09:00:00Z"'


def _cell(name: str, width: int | None = WIDTH, span: int = 1, deleted: int = 0) -> str:
    props = f'<w:tcW w:w="{width}" w:type="dxa"/>' if width is not None else ""
    if span > 1:
        props += f'<w:gridSpan w:val="{span}"/>'
    if deleted:
        # A cell merged away, tracked (Word's own form): w:cellDel.
        props += f'<w:cellDel w:id="{deleted}" {TRACKED}/>'
    content = wml.paragraph(wml.run(f"{name}L")) + wml.paragraph(wml.run(f"{name}R"), jc="right")
    return f"<w:tc><w:tcPr>{props}</w:tcPr>{content}</w:tc>"


def _row(cells: str) -> str:
    return f"<w:tr>{cells}</w:tr>"


def _table(rows: list[str], columns: int = 3, fixed: bool = True, width: tuple | None = (WIDTH * 3, "dxa"),
           margin: int = 108, border: int = 4) -> str:
    stated = f'<w:tblW w:w="{width[0]}" w:type="{width[1]}"/>' if width else ""
    pr = (f'{stated}<w:tblBorders>{_borders(border)}</w:tblBorders>'
          + ('<w:tblLayout w:type="fixed"/>' if fixed else "")
          + f'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="{margin}" w:type="dxa"/>'
          f'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="{margin}" w:type="dxa"/></w:tblCellMar>'
          '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
          ' w:noHBand="1" w:noVBand="1"/>')
    grid = f'<w:gridCol w:w="{WIDTH}"/>' * columns
    return f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{''.join(rows)}</w:tbl>"


def _three(t: str, k: str) -> str:
    return _row(_cell(f"{t}{k}1") + _cell(f"{t}{k}2") + _cell(f"{t}{k}3"))


def _cases() -> list[tuple[str, str]]:
    out = []
    t = "A"
    out.append(("one more", _table([_three(t, "a"), _row("".join(_cell(f"{t}b{k}") for k in range(1, 5))),
                                    _three(t, "c")])))
    t = "B"
    out.append(("one more, narrow", _table([_three(t, "a"), _row(
        _cell(f"{t}b1") + _cell(f"{t}b2") + _cell(f"{t}b3") + _cell(f"{t}b4", 1200)), _three(t, "c")])))
    t = "C"
    out.append(("span past", _table([_three(t, "a"), _row(
        _cell(f"{t}b1") + _cell(f"{t}b2", 2 * WIDTH, span=2) + _cell(f"{t}b3")), _three(t, "c")])))
    t = "D"
    out.append(("no width", _table([_three(t, "a"), _row(
        _cell(f"{t}b1") + _cell(f"{t}b2") + _cell(f"{t}b3") + _cell(f"{t}b4", None)), _three(t, "c")])))
    t = "E"
    out.append(("first row", _table([_row("".join(_cell(f"{t}a{k}") for k in range(1, 5))), _three(t, "b"),
                                     _three(t, "c")])))
    t = "F"
    out.append(("fewer", _table([_three(t, "a"), _row(_cell(f"{t}b1") + _cell(f"{t}b2")), _three(t, "c")])))
    t = "G"
    out.append(("two more", _table([_three(t, "a"), _row("".join(_cell(f"{t}b{k}") for k in range(1, 6))),
                                    _three(t, "c")])))
    t = "H"
    out.append(("autofit", _table([_three(t, "a"), _row("".join(_cell(f"{t}b{k}") for k in range(1, 5))),
                                   _three(t, "c")], fixed=False)))

    def four(t: str) -> list[str]:
        return [_three(t, "a"), _row("".join(_cell(f"{t}b{k}") for k in range(1, 5))), _three(t, "c")]

    def narrow(t: str) -> list[str]:
        return [_three(t, "a"), _row(_cell(f"{t}b1") + _cell(f"{t}b2") + _cell(f"{t}b3") + _cell(f"{t}b4", 1200)),
                _three(t, "c")]

    out.append(("one more, narrow, margins 0", _table(narrow("I"), margin=0)))
    out.append(("one more, no table width", _table(four("J"), width=None)))
    out.append(("one more, table width 9000", _table(four("K"), width=(9000, "dxa"))))
    out.append(("one more, table width 100%", _table(four("L"), width=(5000, "pct"))))
    out.append(("no width, margins 0", _table([_three("M", "a"), _row(
        _cell("Mb1") + _cell("Mb2") + _cell("Mb3") + _cell("Mb4", None)), _three("M", "c")], margin=0)))
    out.append(("grid of four, table width 6000", _table(four("N"), columns=4)))
    out.append(("one more, narrow, autofit, no table width", _table(narrow("O"), fixed=False, width=None)))
    out.append(("one more, narrow, margins 300", _table(narrow("P"), margin=300)))
    out.append(("one more, narrow, borders 3 pt", _table(narrow("Q"), border=24)))
    out.append(("one more, narrow, no borders", _table(narrow("R"), border=0)))
    out.append(("one more, narrow, margins 0, no borders", _table(narrow("S"), margin=0, border=0)))
    # A merge tracked as Word tracks one: the first cell over two columns, the cell merged
    # away still there with w:cellDel -- four columns' worth over a grid of three.
    tracked = [_three("T", "a"), _row(_cell("Tb1", 2 * WIDTH, span=2) + _cell("Tb2", deleted=901) + _cell("Tb3")),
               _three("T", "c")]
    out.append(("tracked merge", _table(tracked)))
    out.append(("tracked merge, autofit, no table width",
                _table([_three("U", "a"), _row(_cell("Ub1", 2 * WIDTH, span=2) + _cell("Ub2", deleted=902)
                                               + _cell("Ub3")), _three("U", "c")], fixed=False, width=None)))
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


if __name__ == "__main__":
    import sys
    from pathlib import Path

    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    for setting in SETTINGS:
        (out / f"grid-mismatch-{setting}.docx").write_bytes(build(setting))

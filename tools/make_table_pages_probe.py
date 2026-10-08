#!/usr/bin/env python3
"""Tables across pages: when a row still fits at a page's foot, how a row that does not is
split, what may not be split, and what the next page starts with.

Tables, stage 6 (ROADMAP.md, "Tables -- measured").  Every case starts a page.  Two
families:

* ``fit`` -- an anchor line, a filler paragraph of one ``exact`` line of 12,000 twips
  with a space after swept by single twips, and a table of one-line rows whose second
  row ends near the page's foot.  Variants: borders of ``w:sz`` 4, 24 and none; a cell
  bottom margin of 100 twips; the row the table's last (two rows) or a middle one (three
  rows).  The sweep brackets, by single twips, the space after at which Word first moves
  the row to the next page (found beforehand at ten-twip steps).
* ``split`` -- body lines filling the page to a few lines from its foot, then a table
  whose second row does not fit: cells of one-line paragraphs (5 and 3 lines) with room
  for 0-3 of them; a six-line paragraph with room for 1-3 lines (widow control); ``w:sz``
  24 borders with cell margins of 100 above and below; paragraphs with 240 twips before
  (what a line starting the next page keeps); one and two header rows
  (``w:tblHeader``); a ``w:cantSplit`` row, and one taller than a page; rows of one line.

One document per compatibility setting (none, 12, 14, 15).  Calibri 11, no paragraph
spacing but where said.  Reader: ``read_table_probes.py``.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}
#: Fit variants: (``w:sz``, cell bottom margin, rows) -> the space after (twips) at which,
#: at ten-twip steps, the second row last stayed on the page.
FIT = {(4, 0, 2): 1113, (4, 0, 3): 1113, (24, 0, 2): 963, (24, 0, 3): 963, (0, 0, 2): 1143,
       (0, 0, 3): 1143, (4, 100, 2): 913, (4, 100, 3): 913, (0, 100, 2): 943, (0, 100, 3): 943}
FIT_STEPS = range(-1, 11)


def side(name: str, sz: int) -> str:
    return f'<w:{name} w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/>'


def borders(sz: int) -> str:
    return "".join(side(name, sz) for name in ("top", "left", "bottom", "right", "insideH", "insideV"))


def _margins(top: int, bottom: int) -> str:
    return (f'<w:top w:w="{top}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
            f'<w:bottom w:w="{bottom}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>')


def _table(rows: str, widths: tuple, sz: int = 4, top: int = 0, bottom: int = 0) -> str:
    pr = (f'<w:tblW w:w="{sum(widths)}" w:type="dxa"/>' + (f"<w:tblBorders>{borders(sz)}</w:tblBorders>" if sz else "")
          + f'<w:tblLayout w:type="fixed"/><w:tblCellMar>{_margins(top, bottom)}</w:tblCellMar>'
          '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
          ' w:noHBand="1" w:noVBand="1"/>')
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    return f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>"


def _p(text: str, before: int = 0) -> str:
    return wml.paragraph(wml.run(text), spacing={"before": before, "after": 0, "line": 240, "lineRule": "auto"})


def _cell(content: str, width: int = 3000) -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{content}</w:tc>'


def _lines(name: str, count: int, before: int = 0) -> str:
    return "".join(_p(f"{name}{k}", before) for k in range(count))


def _row(cells: str, props: str = "") -> str:
    return f"<w:tr>{f'<w:trPr>{props}</w:trPr>' if props else ''}{cells}</w:tr>"


def _fit_cases() -> list[tuple[str, str]]:
    out = []
    for (sz, bottom, rows), last in FIT.items():
        for step in FIT_STEPS:
            after = last + step
            name = f"fit sz {sz} bottom {bottom} rows {rows} after {after}"
            filler = wml.paragraph(wml.run("Filler"), spacing={"before": 0, "after": after, "line": 12000,
                                                               "lineRule": "exact"})
            body = "".join(_row(_cell(_p(f"V{len(out)}r{r}"), 4000)) for r in range(rows))
            out.append((name, filler + _table(body, (4000,), sz, 0, bottom)))
    return out


def _split_cases() -> list[tuple[str, str]]:
    out = []

    def two(a: str, b: str) -> str:
        return _cell(a) + _cell(b)

    for fill in (46, 47, 48, 49):
        rows = (_row(two(_lines("Ha", 1), _lines("Hb", 1))) + _row(two(_lines("Sa", 5), _lines("Sb", 3)))
                + _row(two(_lines("Ea", 1), _lines("Eb", 1))))
        out.append((f"split lines, fill {fill}", _lines("F", fill) + _table(rows, (3000, 3000))))
    words = " ".join(f"Pw{k}" for k in range(30))
    for fill in (46, 47, 48):
        rows = (_row(two(_lines("Ha", 1), _lines("Hb", 1))) + _row(two(_p(words), _lines("Q", 1)))
                + _row(two(_lines("Ea", 1), _lines("Eb", 1))))
        out.append((f"split paragraph, fill {fill}", _lines("F", fill) + _table(rows, (3000, 3000))))
    rows = (_row(two(_lines("Ha", 1), _lines("Hb", 1))) + _row(two(_lines("Sa", 5), _lines("Sb", 3)))
            + _row(two(_lines("Ea", 1), _lines("Eb", 1))))
    out.append(("split, borders 24, margins 100/100", _lines("F", 45) + _table(rows, (3000, 3000), 24, 100, 100)))
    rows = (_row(two(_lines("Ha", 1), _lines("Hb", 1))) + _row(two(_lines("Sa", 5, 240), _lines("Sb", 3, 240)))
            + _row(two(_lines("Ea", 1), _lines("Eb", 1))))
    out.append(("split, 240 before every line", _lines("F", 44) + _table(rows, (3000, 3000))))
    rows = (_row(two(_lines("Ha", 1), _lines("Hb", 1)), "<w:tblHeader/>")
            + "".join(_row(two(_lines(f"R{r}a", 2), _lines(f"R{r}b", 2))) for r in range(1, 5)))
    out.append(("one header row", _lines("F", 46) + _table(rows, (3000, 3000))))
    rows = (_row(two(_lines("Ha", 1), _lines("Hb", 1)), "<w:tblHeader/>")
            + _row(two(_lines("Ga", 1), _lines("Gb", 1)), "<w:tblHeader/>")
            + _row(two(_lines("Sa", 5), _lines("Sb", 3))) + _row(two(_lines("Ea", 1), _lines("Eb", 1))))
    out.append(("two header rows", _lines("F", 44) + _table(rows, (3000, 3000))))
    rows = (_row(two(_lines("Ha", 1), _lines("Hb", 1))) + _row(two(_lines("Sa", 3), _lines("Sb", 5)), "<w:cantSplit/>")
            + _row(two(_lines("Ea", 1), _lines("Eb", 1))))
    out.append(("cantSplit", _lines("F", 48) + _table(rows, (3000, 3000))))
    rows = (_row(two(_lines("Ha", 1), _lines("Hb", 1))) + _row(two(_lines("Ta", 60), _lines("Tb", 3)), "<w:cantSplit/>")
            + _row(two(_lines("Ea", 1), _lines("Eb", 1))))
    out.append(("cantSplit, taller than a page", _lines("F", 2) + _table(rows, (3000, 3000))))
    rows = "".join(_row(two(_lines(f"O{r}a", 1), _lines(f"O{r}b", 1))) for r in range(6))
    out.append(("one-line rows", _lines("F", 49) + _table(rows, (3000, 3000))))
    return out


CASES = _fit_cases() + _split_cases()


def body() -> str:
    out = ""
    for number, (_, content) in enumerate(CASES):
        out += wml.paragraph(wml.run(f"Case {number}"), pageBreakBefore=bool(number)) + content
        out += _p(f"After {number}")
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

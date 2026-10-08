#!/usr/bin/env python3
"""Table rows at a page's foot: a vertically merged cell, a header row, a paragraph kept
whole.

Tables, stage 6 measured how an ordinary row splits at the foot
(``make_table_pages_probe.py``); a row holding a cell merged down (``w:vMerge``) was left
unmeasured and moved whole, and docx-agent saw Word keep a two-line header row whole where
docx2svg split it.  Every case starts a page: body lines fill the page to a few lines
from its foot (the fill swept a line at a time), then a table of two 3000-twip columns.

* ``merge`` -- a one-line first row, then three rows whose first cell is merged down
  over all three (``restart``, ``continue``, ``continue``) beside two-line cells; the
  merged cell holds one line (``merge 1``) or seven, more than its rows (``merge 7``).
* ``merge row`` -- one row whose first cell starts a merge of two rows and holds five
  lines, beside a cell of three, then the row it merges into: the first row itself must
  split.
* ``header`` -- a header row (``w:tblHeader``) of two lines in each cell -- two one-line
  paragraphs (``lines``), or one paragraph that wraps to two (``wrapped``) -- then rows of
  one line; the same row not a header (``plain``); and a header row of a style that keeps
  its lines together (``keepLines``, as Word's headings do).
* ``keep`` -- an ordinary row of wrapped two-line paragraphs under ``w:keepLines``.

Two documents: no ``settings.xml`` and mode 15.  Calibri 11, no paragraph spacing.
Reader: ``read_table_foot_probe.py``.
"""

from __future__ import annotations

import probe_docx
import wml
from make_table_pages_probe import FACE, _cell, _lines, _p, _row, _table

SETTINGS = {"none": None, "15": 15}

#: Body lines before the table: the sweep that puts the interesting row at the foot.
FILLS = (44, 45, 46, 47, 48, 49)


def _merged(content: str, kind: str) -> str:
    return (f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="{kind}"/></w:tcPr>'
            f"{content}</w:tc>")


def _wrapped(name: str, keep: bool = False, style: str | None = None) -> str:
    """One paragraph that wraps to two lines in a 3000-twip cell."""
    words = " ".join(f"{name}{k}" for k in range(9))
    props = {"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}}
    if keep:
        props["keepLines"] = True
    if style:
        props["pStyle"] = style
    return wml.paragraph(wml.run(words), **props)


def _cases() -> list[tuple[str, str]]:
    out = []
    for lines in (1, 7):
        for fill in FILLS:
            rows = (_row(_cell(_lines("Ha", 1)) + _cell(_lines("Hb", 1)))
                    + _row(_merged(_lines("M", lines), "restart") + _cell(_lines("Ba", 2)))
                    + _row(_merged(_p(""), "continue") + _cell(_lines("Bb", 2)))
                    + _row(_merged(_p(""), "continue") + _cell(_lines("Bc", 2)))
                    + _row(_cell(_lines("Ea", 1)) + _cell(_lines("Eb", 1))))
            out.append((f"merge {lines}, fill {fill}", _lines("F", fill) + _table(rows, (3000, 3000))))
    for fill in FILLS:
        rows = (_row(_cell(_lines("Ha", 1)) + _cell(_lines("Hb", 1)))
                + _row(_merged(_lines("M", 5), "restart") + _cell(_lines("Ba", 3)))
                + _row(_merged(_p(""), "continue") + _cell(_lines("Bb", 1)))
                + _row(_cell(_lines("Ea", 1)) + _cell(_lines("Eb", 1))))
        out.append((f"merge row, fill {fill}", _lines("F", fill) + _table(rows, (3000, 3000))))
    for kind in ("lines", "wrapped", "plain", "keepLines"):
        for fill in (47, 48, 49, 50):
            if kind == "lines":
                first = _cell(_lines("Ha", 2)) + _cell(_lines("Hb", 2))
            elif kind == "keepLines":
                first = _cell(_wrapped("Ha", keep=True)) + _cell(_wrapped("Hb", keep=True))
            else:
                first = _cell(_wrapped("Ha")) + _cell(_wrapped("Hb"))
            header = "" if kind == "plain" else "<w:tblHeader/>"
            rows = _row(first, header) + "".join(_row(_cell(_lines(f"R{r}a", 1)) + _cell(_lines(f"R{r}b", 1)))
                                                 for r in range(1, 4))
            out.append((f"header {kind}, fill {fill}", _lines("F", fill) + _table(rows, (3000, 3000))))
    for fill in (46, 47, 48):
        rows = (_row(_cell(_lines("Ha", 1)) + _cell(_lines("Hb", 1)))
                + _row(_cell(_wrapped("Ka", keep=True)) + _cell(_wrapped("Kb", keep=True)))
                + _row(_cell(_lines("Ea", 1)) + _cell(_lines("Eb", 1))))
        out.append((f"keep, fill {fill}", _lines("F", fill) + _table(rows, (3000, 3000))))
    return out


CASES = _cases()


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

#!/usr/bin/env python3
"""Table rows under ``w:keepNext``: which rows a page's foot keeps together, and with what.

A paragraph's ``keepNext`` keeps it on the page of the next paragraph; a table row is
made of paragraphs, and Word lets a row's paragraphs keep the row with the row after it
(and the table's last row with the paragraph after the table).  Every case starts a
page: an anchor line, body lines filling the page to a few rows from its foot, then a
table of two 3000-twip columns of one-line rows (``R<k>a``/``R<k>b``) and a paragraph
after it.  With no keeps the table's first six rows stay (``plain``).  Cases:

* ``all`` -- every paragraph of every row ``keepNext``; ``but last`` -- every row but
  the last; ``row 5`` and ``rows 4-5`` -- only the page's last row(s); ``row 5 cell a``
  -- only one cell's paragraph of that row; ``row 5 first`` / ``row 5 last`` -- cells of
  two paragraphs, only the first / only the last of each ``keepNext``;
* ``foot`` -- the table's ten rows end at the page's foot and the paragraph after does
  not fit: ``plain``, ``last row`` (only the last row ``keepNext``), ``all``;
* ``long`` -- 70 rows, more than a page, every one ``keepNext`` (``all``) or every one
  but the last (``but last``), and the same under a header row (``w:tblHeader``);
* ``before`` -- the paragraph before the table ``keepNext``, the rows plain or all
  ``keepNext``;
* ``cantSplit`` -- the row after the page's last a two-line ``w:cantSplit`` row that
  does not fit, the row before it plain or ``keepNext``;
* ``header`` -- a header row (``w:tblHeader``) and every row ``keepNext``;
* ``nested`` -- row 5's first cell holding a one-row table: its paragraph ``keepNext``
  and the outer cell's not (``inner``), or the reverse (``outer``);
* ``in a cell`` -- a one-row table whose cell holds a 10-row table of ``keepNext`` rows
  across the page's foot (``all``), and the same rows plain;
* ``page break after`` -- the paragraph after the table starts a page
  (``w:pageBreakBefore``): the last row, or every row, ``keepNext``.

Two documents: no ``settings.xml`` and mode 15.  Calibri 11, no paragraph spacing.
Reader: ``read_table_keep_probe.py``.
"""

from __future__ import annotations

import probe_docx
import wml
from make_table_pages_probe import FACE, _cell, _table

SETTINGS = {"none": None, "15": 15}
#: Body lines before the table: six of its one-line rows fit on the page after them.
FILL = 44
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}


def _p(text: str, keep: bool = False) -> str:
    return wml.paragraph(wml.run(text), **({"keepNext": True} if keep else {}), spacing=SPACING)


def _fill(count: int) -> str:
    return "".join(_p(f"F{k}") for k in range(count))


def _row(name: str, keep=(False, False), lines: int = 1, props: str = "") -> str:
    """A row of two cells, each ``lines`` one-line paragraphs; ``keep`` per cell: a bool
    (every paragraph ``keepNext``) or a tuple (per paragraph)."""
    cells = ""
    for side, kept in zip("ab", keep):
        flags = kept if isinstance(kept, tuple) else (kept,) * lines
        cells += _cell("".join(_p(f"{name}{side}{k}" if lines > 1 else f"{name}{side}", flag)
                               for k, flag in enumerate(flags)))
    return f"<w:tr>{f'<w:trPr>{props}</w:trPr>' if props else ''}{cells}</w:tr>"


def _rows(count: int, kept=lambda r: False, prefix: str = "R", **special) -> str:
    out = ""
    for r in range(count):
        if r in special.get("rows", {}):
            out += special["rows"][r]
        else:
            out += _row(f"{prefix}{r}", (kept(r), kept(r)))
    return out


def _two(rows: str) -> str:
    return _table(rows, (3000, 3000))


def _nested(name: str, keep: bool) -> str:
    """A one-row, one-cell table of 2800 twips (a cell's content: a table, then a paragraph)."""
    rows = f"<w:tr>{_cell(_p(name, keep), 2800)}</w:tr>"
    return _table(rows, (2800,))


def _cases() -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    every = lambda r: True  # noqa: E731
    out.append(("plain", _fill(FILL) + _two(_rows(10))))
    out.append(("all", _fill(FILL) + _two(_rows(10, every))))
    out.append(("but last", _fill(FILL) + _two(_rows(10, lambda r: r < 9))))
    out.append(("row 5", _fill(FILL) + _two(_rows(10, lambda r: r == 5))))
    out.append(("rows 4-5", _fill(FILL) + _two(_rows(10, lambda r: r in (4, 5)))))
    out.append(("row 4", _fill(FILL) + _two(_rows(10, lambda r: r == 4))))
    out.append(("row 5 cell a", _fill(FILL) + _two(_rows(10, rows={5: _row("R5", (True, False))}))))
    out.append(("row 5 cell b", _fill(FILL) + _two(_rows(10, rows={5: _row("R5", (False, True))}))))
    out.append(("row 5 plain, two lines", _fill(FILL - 1) + _two(_rows(10, rows={5: _row("R5", lines=2)}))))
    out.append(("row 5 first", _fill(FILL - 1) + _two(_rows(10, rows={
        5: _row("R5", ((True, False), (True, False)), lines=2)}))))
    out.append(("row 5 last", _fill(FILL - 1) + _two(_rows(10, rows={
        5: _row("R5", ((False, True), (False, True)), lines=2)}))))
    out.append(("foot plain", _fill(FILL - 4) + _two(_rows(10))))
    out.append(("foot last row", _fill(FILL - 4) + _two(_rows(10, lambda r: r == 9))))
    out.append(("foot all", _fill(FILL - 4) + _two(_rows(10, every))))
    out.append(("foot but last", _fill(FILL - 4) + _two(_rows(10, lambda r: r < 9))))
    out.append(("long all", _fill(10) + _two(_rows(70, every))))
    out.append(("long but last", _fill(10) + _two(_rows(70, lambda r: r < 69))))
    header = _row("H", (True, True), props="<w:tblHeader/>")
    out.append(("long header all", _fill(10) + _two(_rows(70, every, rows={0: header}))))
    out.append(("before plain", _fill(FILL - 1) + _p("Lead", True) + _two(_rows(10))))
    out.append(("before all", _fill(FILL - 1) + _p("Lead", True) + _two(_rows(10, every))))
    split = _row("R6", lines=2, props="<w:cantSplit/>")
    out.append(("cantSplit plain", _fill(FILL) + _two(_rows(10, rows={6: split}))))
    out.append(("cantSplit row 5", _fill(FILL) + _two(_rows(10, lambda r: r == 5, rows={6: split}))))
    out.append(("header all", _fill(FILL) + _two(_rows(10, every, rows={0: header}))))
    inner = f"<w:tr>{_cell(_nested('N5', True) + _p('R5a'))}{_cell(_p('R5b'))}</w:tr>"
    outer = f"<w:tr>{_cell(_nested('N5', False) + _p('R5a', True))}{_cell(_p('R5b', True))}</w:tr>"
    out.append(("nested inner", _fill(FILL - 1) + _two(_rows(10, rows={5: inner}))))
    out.append(("nested outer", _fill(FILL - 1) + _two(_rows(10, rows={5: outer}))))
    for name, kept in (("in a cell plain", lambda r: False), ("in a cell all", every)):
        held = _table(_rows(10, kept, prefix="C").replace('w:w="3000"', 'w:w="1400"'), (1400, 1400))
        out.append((name, _fill(FILL) + _table(f"<w:tr>{_cell(held + _p('Ca'))}{_cell(_p('Cb'))}</w:tr>",
                                               (3000, 3000))))
    out.append(("row 5 a first only", _fill(FILL - 1) + _two(_rows(10, rows={
        5: _row("R5", ((True, False), (False, False)), lines=2)}))))
    out.append(("row 5 b first only", _fill(FILL - 1) + _two(_rows(10, rows={
        5: _row("R5", ((False, False), (True, False)), lines=2)}))))
    styled = _rows(10, rows={5: _row("R5").replace("<w:pPr>", '<w:pPr><w:pStyle w:val="Kept"/>')})
    out.append(("row 5 style", _fill(FILL) + _two(styled)))
    four = _row("R6", lines=4)
    out.append(("row 6 splits", _fill(FILL - 1) + _two(_rows(10, rows={6: four}))))
    out.append(("row 5, row 6 splits", _fill(FILL - 1) + _two(_rows(10, lambda r: r == 5, rows={6: four}))))
    words = " ".join(f"W{k}" for k in range(24))
    wrapped = (f"<w:tr>{_cell(_p(words, True))}{_cell(_p('R5b', True))}</w:tr>")
    out.append(("row 5 wrapped, cantSplit", _fill(FILL - 3) + _two(_rows(10, rows={5: wrapped, 6: split}))))
    plain_wrapped = (f"<w:tr>{_cell(_p(words))}{_cell(_p('R5b'))}</w:tr>")
    out.append(("row 5 wrapped plain", _fill(FILL - 3) + _two(_rows(10, rows={5: plain_wrapped, 6: split}))))
    broken = wml.paragraph(wml.run("Broken"), pageBreakBefore=True, spacing=SPACING)
    out.append(("last row, page break after", _fill(FILL) + _two(_rows(10, lambda r: r == 9)) + broken))
    out.append(("all, page break after", _fill(FILL) + _two(_rows(10, every)) + broken))
    out.extend(_split_cases())
    out.extend(_chain_cases())
    return out


def _kept_lead(name: str, lines: int, keep: bool, props: str = "", tall: bool = False) -> str:
    """A row of ``lines`` one-line paragraphs per cell, only the first cell's first one
    ``keepNext`` when ``keep``; ``tall``: its markers ``T<row><cell><line>`` (rows of ten
    lines or more)."""
    flags = (keep,) + (False,) * (lines - 1)
    if not tall:
        return _row(name, (flags, (False,) * lines), lines=lines, props=props)
    cells = "".join(_cell("".join(_p(f"T{name[1:]}{side}{k}", flags[k] and side == "a") for k in range(lines)))
                    for side in "ab")
    return f"<w:tr>{f'<w:trPr>{props}</w:trPr>' if props else ''}{cells}</w:tr>"


def _lines(name: str, count: int) -> str:
    return "".join(_p(f"{name}{k}") for k in range(count))


def _split_cases() -> list[tuple[str, str]]:
    """Round 36: a row that keeps with the next and does not fit at the page's foot --
    split there, or moved whole as a ``w:cantSplit`` row?  Each beside its plain twin:

    * ``kept row 6 of N`` -- row 6 of N lines (2, 4, 8) at the foot, only its first
      cell's first paragraph ``keepNext``; ``kept row 6 of 4, all`` every paragraph of it;
    * ``kept last row`` -- the table's last row, of 4 lines, at the foot;
    * ``kept rows of 4`` -- every row but the last ``keepNext``, rows of 4 lines;
    * ``kept tall row`` -- row 1 of 70 lines (taller than a page) after 20 body lines;
      ``... at the top`` right after the case's anchor line; ``cantSplit tall row`` the
      plain row ``w:cantSplit``;
    * ``kept nested row`` -- row 6's first cell starts with a one-row table whose
      paragraph keeps, then 4 lines;
    * ``header, kept row 6`` -- the same under a ``w:tblHeader`` row;
    * ``long kept rows of 4`` -- 20 rows of 4 lines (more than a page), every one but
      the last ``keepNext``, with and without a header row."""
    out: list[tuple[str, str]] = []
    for lines, fill in ((2, FILL), (4, FILL - 1), (8, FILL - 1)):
        for keep in (True, False):
            name = f"{'kept' if keep else 'plain'} row 6 of {lines}"
            out.append((name, _fill(fill) + _two(_rows(10, rows={6: _kept_lead("R6", lines, keep)}))))
    out.append(("kept row 6 of 4, all", _fill(FILL - 1) + _two(_rows(10, rows={6: _row("R6", (True, True),
                                                                                         lines=4)}))))
    for keep in (True, False):
        word = "kept" if keep else "plain"
        out.append((f"{word} last row", _fill(FILL - 6) + _two(_rows(10, rows={9: _kept_lead("R9", 4, keep)}))))
        four = {r: _kept_lead(f"R{r}", 4, keep and r < 9) for r in range(10)}
        out.append((f"{word} rows of 4", _fill(FILL - 14) + _two(_rows(10, rows=four))))
        tall = {1: _kept_lead("R1", 70, keep, tall=True)}
        out.append((f"{word} tall row", _fill(20) + _two(_rows(4, rows=tall))))
        out.append((f"{word} tall row at the top", _two(_rows(4, rows=tall))))
        nested = f"<w:tr>{_cell(_nested('N6', keep) + _lines('R6a', 4))}{_cell(_lines('R6b', 4))}</w:tr>"
        out.append((f"{word} nested row", _fill(FILL - 1) + _two(_rows(10, rows={6: nested}))))
    header = _row("H", props="<w:tblHeader/>")
    out.append(("header, kept row 6", _fill(FILL - 1) + _two(_rows(10, rows={0: header,
                                                                            6: _kept_lead("R6", 4, True)}))))
    split = _kept_lead("R1", 70, False, props="<w:cantSplit/>", tall=True)
    out.append(("cantSplit tall row", _fill(20) + _two(_rows(4, rows={1: split}))))
    for with_header in (False, True):
        rows = {r: _kept_lead(f"R{r}", 4, r < 19) for r in range(20)}
        if with_header:
            rows[0] = _row("H", (True, True), props="<w:tblHeader/>")
        name = f"long kept rows of 4{', header' if with_header else ''}"
        out.append((name, _fill(10) + _two(_rows(20, rows=rows))))
    return out


def _chain_cases() -> list[tuple[str, str]]:
    """Round 36: a chain of keeps taller than a page with a table in it -- ``Head``,
    ``Par`` (60 words over several lines, ``P<k>``), ``Mid`` paragraphs ``keepNext``,
    ``Tail`` not, ``K<k>`` one-line ``keepNext`` paragraphs, a table of 60 one-line rows
    (more than a page):

    * ``chain head par table`` -- ``Head``, ``Par``, the rows every one ``keepNext``,
      ``Tail``; ``... at the top`` right after the anchor line; ``... plain rows`` the
      rows not kept;
    * ``chain head table`` -- ``Head``, the table, ``Tail``;
    * ``chain table par`` -- ``Head``, the table, ``Par``, ``Mid``, ``Tail``;
    * ``chain two tables`` -- ``Head``, a 10-row kept table (``C<k>``), ``Mid``, the
      60-row table, ``Tail``;
    * ``chain paragraphs`` -- 60 ``K<k>`` paragraphs, ``Tail``."""
    every = lambda r: True  # noqa: E731
    words = " ".join(f"P{k}" for k in range(60))
    head, par, mid, tail = _p("Head", True), _p(words, True), _p("Mid", True), _p("Tail")
    long = _two(_rows(60, every))
    out: list[tuple[str, str]] = []
    out.append(("chain head par table", _fill(10) + head + par + long + tail))
    out.append(("chain head par table at the top", head + par + long + tail))
    out.append(("chain head par table plain rows", _fill(10) + head + par + _two(_rows(60)) + tail))
    out.append(("chain head table", _fill(10) + head + long + tail))
    out.append(("chain table par", _fill(10) + head + long + par + mid + tail))
    out.append(("chain two tables", _fill(10) + head + _two(_rows(10, every, prefix="C")) + mid + long + tail))
    out.append(("chain paragraphs", _fill(10) + "".join(_p(f"K{k}", True) for k in range(60)) + tail))
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
        [wml.style("paragraph", "Normal", default=True),
         wml.style("paragraph", "Kept", based_on="Normal", ppr_={"keepNext": True})],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    )


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package(body(), styles=styles(), extra_parts=extra)

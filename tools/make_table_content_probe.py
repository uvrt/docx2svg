#!/usr/bin/env python3
"""What a table cell's text does inside the cell: its line budget, spacing, line rules,
indents, alignment, tabs, lists and sizes -- and the size Word gives it by style.

Tables, stage 2 (ROADMAP.md, "Tables -- measured").  Stage 1 placed a cell's text start
and line end; here each cell holds paragraphs that wrap, and the question is whether
they lay out as the body's do, at the cell's width, with the existing line breaker and
vertical model.  Every table is fixed layout, one row, borders ``w:sz`` 4, margins 108
left and right (Word's own default table style's), Calibri 11 pt, no spacing unless a
case says so.  Families, one case each:

* ``wrap δ`` -- a cell whose line budget is exactly ``Hnio <word>`` less δ layout units
  (the cell's width chosen so, in multiples of 5 twips less its margins; the word
  composed of letters of known advance, as ``make_wrap_budget_probe.py`` composes its
  words), followed by `` Hnnn``: under Phase 3's rule the word stays on the first line
  iff δ >= 0.  δ = -1024, -2, 0, 2, 1024; and with six trailing spaces instead of one;
* ``spacing`` -- three paragraphs per cell, space before/after 120/0, 0/120, 120/240,
  240/120, 100/100; contextual spacing -- stated on Normal paragraphs (the table between
  Normal ones), and by a style ``Ctx`` whose paragraphs sit in the cell with Normal ones
  outside, and with ``Ctx`` ones outside; Word's own ``Normal.dotm`` spacing (after 160,
  line 259 auto);
* ``line`` -- two-line paragraphs under ``exact`` 300, ``atLeast`` 400, ``auto`` 360 and
  ``auto`` 200;
* ``indent`` / ``jc`` -- two-line paragraphs indented 200 left, 300 right, 360 first
  line, 360 hanging; centred, right-aligned and justified;
* ``tabs`` -- default stops, and custom left, centre and right stops;
* ``list`` -- a numbered list and a bulleted one in a cell;
* ``size`` -- 9, 14 and 20 pt in cells of one row, and a cell of mixed sizes.

One document per compatibility setting (none, 12, 14, 15).  **Finding 6** -- the size of
a cell's text when the Normal style's size differs from ``w:docDefaults`` -- is
``make_table_style_size_probe.py``'s.  Reader: ``read_table_probes.py``.
"""

from __future__ import annotations

import functools

import make_wrap_budget_probe as budget
import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}
MARGIN = 108
PER_PAGE = 7
W1 = budget.W1
TAIL = "Hnnn"
TEXT = ("Every cell of this row holds the same few words so that each of them breaks into two "
        "lines at its width")


def _side(name: str) -> str:
    return f'<w:{name} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'


BORDERS = "".join(_side(name) for name in ("top", "left", "bottom", "right", "insideH", "insideV"))


def _table(cells: list[str], widths: list[int]) -> str:
    pr = (f'<w:tblW w:w="{sum(widths)}" w:type="dxa"/><w:tblBorders>{BORDERS}</w:tblBorders>'
          '<w:tblLayout w:type="fixed"/>'
          f'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="{MARGIN}" w:type="dxa"/>'
          f'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="{MARGIN}" w:type="dxa"/></w:tblCellMar>'
          '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
          ' w:noHBand="1" w:noVBand="1"/>')
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    row = "".join(f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{cell}</w:tc>'
                  for w, cell in zip(widths, cells))
    return f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid><w:tr>{row}</w:tr></w:tbl>"


def _p(text: str, *, spacing: tuple = (0, 0, 240, "auto"), **props) -> str:
    before, after, line, rule = spacing
    props["spacing"] = {"before": before, "after": after, "line": line, "lineRule": rule}
    runs = "".join(wml.run(part) if part != "\t" else wml.run("\t") for part in _split_tabs(text))
    return wml.paragraph(runs, **props)


def _split_tabs(text: str) -> list[str]:
    out, current = [], ""
    for char in text:
        if char == "\t":
            if current:
                out.append(current)
            out.append("\t")
            current = ""
        else:
            current += char
    if current:
        out.append(current)
    return out


@functools.lru_cache(maxsize=None)
def _wrap(delta: int, target: int, trailing: int) -> tuple[int, str]:
    """``(cell width, text)``: a cell whose line budget less ``Hnio <word>`` is exactly
    ``delta`` units, the width a multiple of 5 twips once its margins are taken off."""
    hp = 22
    sums = budget._sums("Calibri")
    fixed = budget._units("Calibri", W1 + " ", hp)
    for step in range(0, 400):
        for column in (target + 5 * step, target - 5 * step):
            units = column * 1024 // 5
            want = units - delta - fixed
            if want > 0 and want % hp == 0 and want // hp < len(sums) and sums[want // hp]:
                return column + 2 * MARGIN, f"{W1} {sums[want // hp]}" + " " * trailing + TAIL
    raise ValueError(delta)


def _cases() -> list[tuple[str, str]]:
    """``(name, table)``."""
    out = []
    for delta in (-1024, -2, 0, 2, 1024):
        width, text = _wrap(delta, 2400, 1)
        out.append((f"wrap {delta}", _table([_p(text)], [width])))
    for delta in (-2, 0):
        width, text = _wrap(delta, 2600, 6)
        out.append((f"wrap {delta}, trailing spaces", _table([_p(text)], [width])))
    for before, after in ((120, 0), (0, 120), (120, 240), (240, 120), (100, 100)):
        cell = "".join(_p(f"Spacing {before}/{after} paragraph {k}", spacing=(before, after, 240, "auto"))
                       for k in range(3))
        out.append((f"spacing {before}/{after}", _table([cell], [4000])))
    cell = "".join(_p(f"Contextual paragraph {k}", spacing=(120, 240, 240, "auto"), contextualSpacing=True)
                   for k in range(3))
    out.append(("spacing contextual", _table([cell], [4000])))
    out.append(("spacing contextual, another style outside",
                _table([cell], [4000]) + wml.paragraph(wml.run("Other after"), pStyle="Other")))
    cell = "".join(wml.paragraph(wml.run(f"Ctx paragraph {k}"), pStyle="Ctx") for k in range(3))
    out.append(("spacing contextual style, Normal outside", _table([cell], [4000])))
    out.append(("spacing contextual style, the same outside",
                _table([cell], [4000]) + wml.paragraph(wml.run("Ctx after"), pStyle="Ctx")))
    # Contextual spacing at a cell's edges: one contextual paragraph (120 before, 240 after)
    # in a one-cell table, every pairing of its style and the styles around the table.
    for cell_style in ("Normal", "Ctx"):
        for before_style in ("Normal", "Ctx", "Other"):
            for after_style in ("Normal", "Ctx", "Other"):
                for direct in (True, False):
                    if cell_style == "Ctx" and not direct:
                        continue
                    props = {"contextualSpacing": True} if direct else {}
                    if cell_style != "Normal":
                        props["pStyle"] = cell_style
                    cell = wml.paragraph(wml.run("Edge cell"), spacing={"before": 120, "after": 240, "line": 240,
                                                                         "lineRule": "auto"}, **props)
                    after = wml.paragraph(wml.run("Edge after"), **({"pStyle": after_style}
                                                                    if after_style != "Normal" else {}))
                    out.append((f"edge {cell_style}{' (direct)' if direct else ''} between {before_style}"
                                f" and {after_style}", _table([cell], [4000]) + after))
    cell = "".join(_p(f"Normal.dotm spacing paragraph {k}", spacing=(0, 160, 259, "auto")) for k in range(3))
    out.append(("spacing Normal.dotm", _table([cell], [4000])))
    for name, spacing in (("exact 300", (0, 0, 300, "exact")), ("atLeast 400", (0, 0, 400, "atLeast")),
                          ("auto 360", (0, 0, 360, "auto")), ("auto 200", (0, 0, 200, "auto"))):
        out.append((f"line {name}", _table([_p(TEXT, spacing=spacing), _p(TEXT, spacing=spacing)], [3000, 3000])))
    for name, props in (("left 200", {"ind": {"left": 200}}), ("right 300", {"ind": {"right": 300}}),
                        ("first line 360", {"ind": {"firstLine": 360}}),
                        ("hanging 360", {"ind": {"left": 360, "hanging": 360}})):
        out.append((f"indent {name}", _table([_p(TEXT, **props), _p(TEXT, **props)], [3000, 3000])))
    for jc in ("center", "right", "both"):
        out.append((f"jc {jc}", _table([_p(TEXT, jc=jc), _p(TEXT, jc=jc)], [3000, 3000])))
    out.append(("tabs default", _table([_p("A\tBb\tCcc\tD"), _p("Tab\tstops")], [4500, 2000])))
    stops = ('<w:tab w:val="left" w:pos="1000"/><w:tab w:val="center" w:pos="2000"/>'
             '<w:tab w:val="right" w:pos="3800"/>')
    out.append(("tabs custom", _table([_p("A\tBbb\tCc\tDddd", tabs=stops), _p("x")],
                                      [4500, 2000])))
    numbered = "".join(_p(f"Numbered item {k}", numPr="<w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/>")
                       for k in range(3))
    bulleted = "".join(_p(f"Bulleted item {k}", numPr="<w:ilvl w:val=\"1\"/><w:numId w:val=\"1\"/>")
                       for k in range(2))
    out.append(("list", _table([numbered, bulleted], [3000, 3000])))
    sized = [wml.paragraph(wml.run(f"Size {hp / 2:g} pt", sz=hp, szCs=hp), mark={"sz": hp, "szCs": hp})
             for hp in (18, 28, 40)]
    out.append(("size 9 / 14 / 20", _table(sized, [2000, 2000, 3000])))
    mixed = wml.paragraph(wml.run("Mixed ") + wml.run("sizes", sz=36, szCs=36) + wml.run(" here", sz=16, szCs=16))
    out.append(("size mixed", _table([mixed, _p("Plain")], [3000, 3000])))
    return out


CASES = _cases()


def body() -> str:
    out = ""
    for number, (name, table) in enumerate(CASES):
        props = {"pageBreakBefore": True} if number and number % PER_PAGE == 0 else {}
        if name in ("spacing contextual style, the same outside", "spacing contextual, another style outside"):
            style = "Ctx" if name.startswith("spacing contextual style") else "Other"
            out += wml.paragraph(wml.run(f"Case {number}"), pStyle=style, **props) + table
            continue
        if name.startswith("edge "):
            style = name.split(" between ")[1].split(" and ")[0]
            if style != "Normal":
                props["pStyle"] = style
            out += wml.paragraph(wml.run(f"Case {number}"), **props) + table
            continue
        out += _p(f"Case {number}", **props) + table
    return out + _p("End")


def styles() -> str:
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True),
         wml.style("paragraph", "Other", based_on="Normal"),
         wml.style("paragraph", "Ctx", based_on="Normal",
                   ppr_={"spacing": {"before": 120, "after": 240, "line": 240, "lineRule": "auto"},
                         "contextualSpacing": True})],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    parts = [wml.numbering_part([("decimal", "%1.", {"ind": {"left": 360, "hanging": 360}}, {}),
                                 ("bullet", "\u2013", {"ind": {"left": 720, "hanging": 360}}, {})])]
    if mode is not None:
        parts.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=mode))
    return probe_docx.package(body(), styles=styles(), extra_parts=tuple(parts))

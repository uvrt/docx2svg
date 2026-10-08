#!/usr/bin/env python3
"""Where Word puts a table's border lines and its cells' shading, to the device pixel.

Phase 5.13's one-rasteriser instrument showed two real differences on every table page
of the committed documents: the table's left outer border drawn a pixel right of Word's,
and the cell shading painted up to a border where Word leaves a pixel column white.
Tables stage 5 had recorded both from the filled rectangles and left them unprobed.  The
cases there sat on few pixel fractions; this probe sweeps them.

Every case is a two-row, three-column table of **empty** cells (so no text start sits on
a tie and the glyph check has nothing to disagree with), the first row's cells shaded
(``w:shd``, yellow), borders black.  A table is placed by ``w:tblInd`` stepped by single
twips over 24 of them, so its left grid line takes every fraction of a pixel a twip can
reach (5/24 px each) -- and the ties, a grid line at exactly *n*.5 px -- and its columns
are 1500-1523 twips wide, varied per case, so the inner and right grid lines do too.

Families (``CASES``):

* ``right of the margin`` -- ``w:sz`` 2, 4, 6, 8, 12 and 24 on every side (1-12 px),
  ``w:tblInd`` 240-263: every grid line to the right of the text column's left edge;
* ``left of the margin`` -- ``w:sz`` 4 and 8, ``w:tblInd`` -240 to -217: the left grid
  line left of the text column's edge, the others right of it;
* ``all left of the margin`` -- ``w:sz`` 4 and 8, ``w:tblInd`` -1300 to -1277 and
  columns of 300 twips: every grid line left of the text column's edge;
* ``no vertical borders`` -- top, bottom and ``insideH`` only (``w:sz`` 8): the shading
  against grid lines with no line on them;
* ``no borders`` -- the shading alone.

Six documents: no ``settings.xml`` and mode 15, each with the page's left margin at 1440
twips (300 px, on the pixel grid), 1442 (300.4167 px) and 1438 (299.5833 px), so a rule
that rounds against the text column's edge and one that rounds against the page's can be
told apart, and how the edge itself rounds.  Reader: ``read_table_line_probe.py``.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
#: setting -> (compatibility mode, the page's left margin in twips).
SETTINGS = {"none": (None, 1440), "15": (15, 1440), "none-m1442": (None, 1442), "15-m1442": (15, 1442),
            "none-m1438": (None, 1438), "15-m1438": (15, 1438)}
SHADING = "FFFF00"
BORDER = "000000"
PER_PAGE = 14
SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")


def side(name: str, sz: int) -> str:
    return f'<w:{name} w:val="single" w:sz="{sz}" w:space="0" w:color="{BORDER}"/>'


def _cases() -> list[tuple[str, dict]]:
    out = []

    def widths(j: int, base: int = 1500) -> tuple[int, int, int]:
        return (base + (7 * j) % 24, base + (11 * j) % 24, base + (13 * j) % 24)

    for sz in (2, 4, 6, 8, 12, 24):
        for j in range(24):
            out.append((f"right of the margin, sz {sz}, indent {240 + j}",
                        {"sz": sz, "indent": 240 + j, "widths": widths(j), "sides": SIDES}))
    for sz in (4, 8):
        for j in range(24):
            out.append((f"left of the margin, sz {sz}, indent {-240 + j}",
                        {"sz": sz, "indent": -240 + j, "widths": widths(j), "sides": SIDES}))
    for sz in (4, 8):
        for j in range(24):
            out.append((f"all left of the margin, sz {sz}, indent {-1300 + j}",
                        {"sz": sz, "indent": -1300 + j, "widths": widths(j, 300), "sides": SIDES}))
    for j in range(24):
        out.append((f"no vertical borders, indent {240 + j}",
                    {"sz": 8, "indent": 240 + j, "widths": widths(j), "sides": ("top", "bottom", "insideH")}))
    for j in range(24):
        out.append((f"no borders, indent {240 + j}", {"sz": 0, "indent": 240 + j, "widths": widths(j), "sides": ()}))
    return out


CASES = _cases()


def _cell(width: int, shaded: bool) -> str:
    props = f'<w:tcW w:w="{width}" w:type="dxa"/>'
    if shaded:
        props += f'<w:shd w:val="clear" w:color="auto" w:fill="{SHADING}"/>'
    return f"<w:tc><w:tcPr>{props}</w:tcPr>{wml.paragraph('')}</w:tc>"


def table(case: dict) -> str:
    widths = case["widths"]
    pr = f'<w:tblW w:w="{sum(widths)}" w:type="dxa"/><w:tblInd w:w="{case["indent"]}" w:type="dxa"/>'
    if case["sides"]:
        pr += "<w:tblBorders>" + "".join(side(name, case["sz"]) for name in case["sides"]) + "</w:tblBorders>"
    pr += ('<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
           '<w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
           '</w:tblCellMar><w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0"'
           ' w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>')
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    rows = "".join("<w:tr>" + "".join(_cell(w, r == 0) for w in widths) + "</w:tr>" for r in range(2))
    return f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>"


def body() -> str:
    out = ""
    for number, (_, case) in enumerate(CASES):
        props = {"pageBreakBefore": True} if number and number % PER_PAGE == 0 else {}
        out += wml.paragraph(wml.run(f"Case {number}"), **props) + table(case)
    return out + wml.paragraph(wml.run("End"))


def styles() -> str:
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )


def build(setting: str) -> bytes:
    mode, left = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package(body(), styles=styles(), extra_parts=extra,
                              final_section=probe_docx.section(left=left))


if __name__ == "__main__":
    import sys
    from pathlib import Path

    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    for name in SETTINGS:
        (out / f"table-line-{name}.docx").write_bytes(build(name))
        print(out / f"table-line-{name}.docx")

#!/usr/bin/env python3
"""Table borders and shading: which border Word draws where two disagree, how much
room it takes, and where the lines and fills go.

Tables, stage 5 (ROADMAP.md, "Tables -- measured").  Each case is a two-row, two-column
table (each cell a left-aligned and a right-aligned word, so a border's room on either
side of a grid line shows in the text), with borders set on the table (``w:tblBorders``)
and on single cells (``w:tcBorders``) so that two borders meet on one edge:

* ``vertical`` -- the first cell's right border against the second cell's left: widths
  4/12, 12/4, the same width in two colours, ``single`` against ``double``, a border
  against ``nil``, against none stated; and with no cell margins, where each cell's text
  starts at the inner half of the border between them;
* ``horizontal`` -- the first row's bottom border against the second row's top, likewise;
* ``cell over table`` -- a thin cell border under a thick ``insideV`` / ``insideH``, a
  ``nil`` cell border under a table border;
* ``double`` and ``dotted`` lines, and ``w:sz`` 2 to 24;
* ``shading`` -- a cell's ``w:shd`` fill, the table's, with borders and without;
* ``style`` -- the table style's conditional borders (a ``firstRow`` bottom of ``w:sz``
  18 against ``insideH`` 8, as the real documents' "Light Grid" has).

Colours differ per border (red, blue, green...) so Word's fills say which border it
drew.  One document per compatibility setting (none, 12, 14, 15).  Reader:
``read_table_probes.py``.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}
PER_PAGE = 6
#: 3001 twips, not 3000: the second cell's text then starts clear of a half pixel (1440 +
#: 3000 + 108 twips is 947.5 px exactly, a tie Word breaks by the borders in a way not
#: settled -- ROADMAP.md, "Tables -- measured").
WIDTH = 3001
RED, BLUE, GREEN, ORANGE, GREY = "FF0000", "0000FF", "00A000", "FF8000", "808080"


def side(name: str, sz: int, color: str = "000000", val: str = "single") -> str:
    return f'<w:{name} w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'


def table_borders(sz: int = 4, color: str = "000000", **overrides) -> str:
    out = ""
    for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        out += overrides.get(name, side(name, sz, color))
    return out


def _cell(text: str, own: str = "", shading: str | None = None) -> str:
    props = f'<w:tcW w:w="{WIDTH}" w:type="dxa"/>'
    if own:
        props += f"<w:tcBorders>{own}</w:tcBorders>"
    if shading:
        props += f'<w:shd w:val="clear" w:color="auto" w:fill="{shading}"/>'
    return (f"<w:tc><w:tcPr>{props}</w:tcPr>{wml.paragraph(wml.run(text + 'L'))}"
            f"{wml.paragraph(wml.run(text + 'R'), jc='right')}</w:tc>")


def _table(number: int, borders: str, cells: dict | None = None, shading: dict | None = None,
           table_shading: str | None = None, style: str | None = None, margin: int = 108) -> str:
    cells = cells or {}
    shading = shading or {}
    pr = f'<w:tblStyle w:val="{style}"/>' if style else ""
    pr += f'<w:tblW w:w="{2 * WIDTH}" w:type="dxa"/>'
    if borders:
        pr += f"<w:tblBorders>{borders}</w:tblBorders>"
    if table_shading:
        pr += f'<w:shd w:val="clear" w:color="auto" w:fill="{table_shading}"/>'
    pr += ('<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
           f'<w:left w:w="{margin}" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
           f'<w:right w:w="{margin}" w:type="dxa"/></w:tblCellMar>')
    look = ('<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0"'
            ' w:noHBand="0" w:noVBand="1"/>') if style else (
        '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
        ' w:noHBand="1" w:noVBand="1"/>')
    pr += look
    rows = ""
    for r in range(2):
        rows += "<w:tr>" + "".join(_cell(f"T{number}.{r}{c}", cells.get((r, c), ""), shading.get((r, c)))
                                   for c in range(2)) + "</w:tr>"
    grid = f'<w:gridCol w:w="{WIDTH}"/>' * 2
    return f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>"


def _cases() -> list[tuple[str, dict]]:
    base = table_borders(4)
    out = [("plain", {"borders": base})]
    for a, b in ((4, 12), (12, 4), (8, 8), (2, 24), (24, 2), (6, 18)):
        out.append((f"vertical {a} red / {b} blue",
                    {"borders": base, "cells": {(0, 0): side("right", a, RED), (0, 1): side("left", b, BLUE)}}))
        out.append((f"horizontal {a} red / {b} blue",
                    {"borders": base, "cells": {(0, 0): side("bottom", a, RED), (1, 0): side("top", b, BLUE)}}))
    for a, b in ((4, 24), (24, 4), (12, 12), (6, 18)):
        out.append((f"vertical {a} red / {b} blue, no margins",
                    {"borders": base, "margin": 0,
                     "cells": {(0, 0): side("right", a, RED), (0, 1): side("left", b, BLUE)}}))
    out.append(("vertical 24 red / nil, no margins", {"borders": base, "margin": 0, "cells": {
        (0, 0): side("right", 24, RED), (0, 1): side("left", 0, BLUE, "nil")}}))
    out.append(("cell 4 under insideV 24, no margins", {"borders": table_borders(4, insideV=side("insideV", 24, GREEN)),
                                                          "margin": 0, "cells": {(0, 0): side("right", 4, RED)}}))
    out.append(("vertical single / double",
                {"borders": base, "cells": {(0, 0): side("right", 8, RED), (0, 1): side("left", 8, BLUE, "double")}}))
    out.append(("vertical double / single",
                {"borders": base, "cells": {(0, 0): side("right", 8, RED, "double"), (0, 1): side("left", 8, BLUE)}}))
    out.append(("vertical 12 / nil", {"borders": base, "cells": {(0, 0): side("right", 12, RED),
                                                                   (0, 1): side("left", 0, BLUE, "nil")}}))
    out.append(("vertical nil / 12", {"borders": base, "cells": {(0, 0): side("right", 0, RED, "nil"),
                                                                   (0, 1): side("left", 12, BLUE)}}))
    out.append(("cell 4 under insideV 24", {"borders": table_borders(4, insideV=side("insideV", 24, GREEN)),
                                              "cells": {(0, 0): side("right", 4, RED)}}))
    out.append(("cell nil under insideV 24", {"borders": table_borders(4, insideV=side("insideV", 24, GREEN)),
                                                "cells": {(0, 0): side("right", 0, RED, "nil")}}))
    out.append(("cell 4 under insideH 24", {"borders": table_borders(4, insideH=side("insideH", 24, GREEN)),
                                              "cells": {(0, 0): side("bottom", 4, RED)}}))
    out.append(("cell nil under insideH 24", {"borders": table_borders(4, insideH=side("insideH", 24, GREEN)),
                                                "cells": {(0, 0): side("bottom", 0, RED, "nil")}}))
    for sz in (2, 6, 12, 24):
        out.append((f"double {sz}", {"borders": "".join(side(n, sz, ORANGE, "double")
                                                        for n in ("top", "left", "bottom", "right", "insideH",
                                                                  "insideV"))}))
    out.append(("dotted 8", {"borders": "".join(side(n, 8, GREY, "dotted")
                                                for n in ("top", "left", "bottom", "right", "insideH", "insideV"))}))
    out.append(("shading cell", {"borders": base, "shading": {(0, 1): "FFFF00"}}))
    out.append(("shading cell, no borders", {"borders": "", "shading": {(0, 1): "FFFF00", (1, 0): "00FFFF"}}))
    out.append(("shading table", {"borders": base, "table_shading": "DDDDDD"}))
    out.append(("shading table and cell", {"borders": base, "table_shading": "DDDDDD", "shading": {(1, 1): "FFFF00"}}))
    out.append(("style: first row bottom 18 over insideH 8", {"borders": "", "style": "Grid"}))
    return out


CASES = _cases()


def body() -> str:
    out = ""
    for number, (_, case) in enumerate(CASES):
        props = {"pageBreakBefore": True} if number and number % PER_PAGE == 0 else {}
        out += wml.paragraph(wml.run(f"Case {number}"), **props) + _table(number, **case)
        out += wml.paragraph(wml.run(f"After {number}"))
    return out


def styles() -> str:
    grid = (
        '<w:style w:type="table" w:styleId="Grid"><w:name w:val="Grid"/>'
        '<w:tblPr><w:tblBorders>' + table_borders(8, BLUE) + "</w:tblBorders></w:tblPr>"
        '<w:tblStylePr w:type="firstRow"><w:tcPr><w:tcBorders>'
        + side("top", 8, BLUE) + side("left", 8, BLUE) + side("bottom", 18, RED) + side("right", 8, BLUE)
        + '<w:insideH w:val="nil"/>' + side("insideV", 8, BLUE)
        + "</w:tcBorders></w:tcPr></w:tblStylePr>"
        '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D3DFEE"/></w:tcPr>'
        "</w:tblStylePr></w:style>")
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True), grid],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package(body(), styles=styles(), extra_parts=extra)

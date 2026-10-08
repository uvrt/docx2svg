#!/usr/bin/env python3
"""How Word draws dotted, dashed and double border lines.

The model draws every border line solid (Phase 5.2, Tables stage 5); a ``dotted`` table
border drew 8,844 px Word does not, and ``double`` lines' corners and joins differ.  Each
case is one style at one width:

* ``box`` -- a paragraph with all four borders (``w:space`` 4), each side its own colour
  so the fills say which side drew them, at three right indents (so the pattern's phase
  and its ends show over three lengths);
* ``table`` -- a two-row, two-column table with every border (table and inside) in the
  style, the outer ones in one colour and the inside ones in another.

Styles ``dotted``, ``dashed``, ``dotDash``, ``dotDotDash``, ``dashSmallGap`` and
``double``, at ``w:sz`` 4, 8, 12 and 24.  Two documents: no ``settings.xml`` and mode
15.  Reader: ``read_border_pattern_probe.py``.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "15": 15}
STYLES = ("dotted", "dashed", "dotDash", "dotDotDash", "dashSmallGap", "double")
SIZES = (4, 8, 12, 24)
SIDE_COLOURS = {"top": "C00000", "left": "0070C0", "bottom": "00B050", "right": "7030A0"}
OUTER, INSIDE = "FF8000", "808080"


def side(name: str, val: str, sz: int, space: int, color: str) -> str:
    return f'<w:{name} w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'


def table(val: str, sz: int) -> str:
    borders = "".join(side(n, val, sz, 0, OUTER) for n in ("top", "left", "bottom", "right"))
    borders += "".join(side(n, val, sz, 0, INSIDE) for n in ("insideH", "insideV"))
    pr = (f'<w:tblW w:w="6002" w:type="dxa"/><w:tblInd w:w="240" w:type="dxa"/><w:tblBorders>{borders}'
          '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/>'
          '<w:right w:w="108" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0"'
          ' w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>')
    cell = '<w:tc><w:tcPr><w:tcW w:w="3001" w:type="dxa"/></w:tcPr>' + wml.paragraph(wml.run("Cell")) + "</w:tc>"
    rows = ("<w:tr>" + cell * 2 + "</w:tr>") * 2
    return (f'<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid><w:gridCol w:w="3001"/><w:gridCol w:w="3001"/>'
            f"</w:tblGrid>{rows}</w:tbl>")


def body() -> str:
    out = ""
    for val in STYLES:
        out += wml.paragraph(wml.run(f"Style {val}"), pageBreakBefore=True)
        for sz in SIZES:
            for right in (0, 1003, 2507):
                sides = "".join(side(n, val, sz, 4, c) for n, c in SIDE_COLOURS.items())
                out += wml.paragraph(wml.run(f"Boxed {val} {sz} {right}"), pBdr=sides,
                                     ind={"left": 240, "right": right})
                out += wml.paragraph(wml.run("Between"))
        for sz in SIZES:
            out += wml.paragraph(wml.run(f"Table {val} {sz}")) + table(val, sz)
    return out + wml.paragraph(wml.run("End"))


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

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"border-pattern-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

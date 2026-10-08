#!/usr/bin/env python3
"""How tall Word makes a table row, and where a cell's lines sit in it.

Tables, stage 3 (ROADMAP.md, "Tables -- measured").  A row is as tall as its tallest
cell -- margins and paragraphs -- unless ``w:trHeight`` says otherwise; each horizontal
border between rows takes its width once; ``w:vAlign`` moves a cell's lines down in a row
taller than they are.  Each case is a three-row, two-column table (a one-line cell
beside a three-line one, so a row's height is the second cell's, and the first cell has
room to move), then a paragraph whose baseline says where the table ended.  Families:

* ``atLeast`` -- ``w:trHeight`` 200 (less than the content), 1000, 1001, 1003 and 1203
  twips (heights off the pixel grid), and with no ``w:hRule`` (which means ``atLeast``);
* ``exact`` -- 400 (less than the content: it is cut), 800, 801, 1203;
* ``borders`` -- the same rows under a top border of ``w:sz`` 24, ``insideH`` 4 and a
  bottom border of 12, with ``atLeast`` 1000 and ``exact`` 800: which of a row's borders
  its stated height holds;
* ``vAlign`` -- ``center`` and ``bottom``, alone, with ``atLeast`` 1000, with ``exact``
  800, and with top and bottom cell margins (100 and 50);
* ``margins`` -- top and bottom cell margins with ``atLeast`` 1000.

One document per compatibility setting (none, 12, 14, 15); Calibri 11, no paragraph
spacing.  Reader: ``read_table_probes.py``.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}
PER_PAGE = 3
WIDTH = 3000


def _side(name: str, sz: int) -> str:
    return f'<w:{name} w:val="single" w:sz="{sz}" w:space="0" w:color="000000"/>'


def _borders(top: int = 4, inside: int = 4, bottom: int = 4) -> str:
    return (_side("top", top) + _side("left", 4) + _side("bottom", bottom) + _side("right", 4)
            + _side("insideH", inside) + _side("insideV", 4))


def _table(number: int, *, height: tuple | None = None, align: str | None = None, margins: tuple = (0, 0),
           borders: str | None = None) -> str:
    top, bottom = margins
    pr = (f'<w:tblW w:w="{2 * WIDTH}" w:type="dxa"/><w:tblBorders>{borders or _borders()}</w:tblBorders>'
          '<w:tblLayout w:type="fixed"/>'
          f'<w:tblCellMar><w:top w:w="{top}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
          f'<w:bottom w:w="{bottom}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
          '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
          ' w:noHBand="1" w:noVBand="1"/>')
    rows = ""
    for r in range(3):
        tr = ""
        if height is not None:
            value, rule = height
            tr = f'<w:trPr><w:trHeight w:val="{value}"' + (f' w:hRule="{rule}"' if rule else "") + "/></w:trPr>"
        cells = ""
        for c in range(2):
            own = f'<w:vAlign w:val="{align}"/>' if align else ""
            if c == 0:
                content = wml.paragraph(wml.run(f"R{number}.{r}a"))
            else:
                content = "".join(wml.paragraph(wml.run(f"R{number}.{r}b{k}")) for k in range(3))
            cells += f'<w:tc><w:tcPr><w:tcW w:w="{WIDTH}" w:type="dxa"/>{own}</w:tcPr>{content}</w:tc>'
        rows += f"<w:tr>{tr}{cells}</w:tr>"
    grid = f'<w:gridCol w:w="{WIDTH}"/>' * 2
    return f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>"


def _cases() -> list[tuple[str, dict]]:
    out = [("plain", {})]
    for value in (200, 1000, 1001, 1003, 1203):
        out.append((f"atLeast {value}", {"height": (value, "atLeast")}))
    out.append(("atLeast 1000, no hRule", {"height": (1000, None)}))
    for value in (400, 800, 801, 1203):
        out.append((f"exact {value}", {"height": (value, "exact")}))
    thick = _borders(24, 4, 12)
    out.append(("borders 24/4/12", {"borders": thick}))
    out.append(("borders 24/4/12, atLeast 1000", {"borders": thick, "height": (1000, "atLeast")}))
    out.append(("borders 24/4/12, exact 800", {"borders": thick, "height": (800, "exact")}))
    out.append(("borders 4/24/4, exact 800", {"borders": _borders(4, 24, 4), "height": (800, "exact")}))
    for align in ("center", "bottom"):
        out.append((f"vAlign {align}", {"align": align}))
        out.append((f"vAlign {align}, atLeast 1000", {"align": align, "height": (1000, "atLeast")}))
        out.append((f"vAlign {align}, exact 800", {"align": align, "height": (800, "exact")}))
        out.append((f"vAlign {align}, margins 100/50", {"align": align, "margins": (100, 50),
                                                         "height": (1000, "atLeast")}))
    out.append(("margins 100/50, atLeast 1000", {"margins": (100, 50), "height": (1000, "atLeast")}))
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
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package(body(), styles=styles(), extra_parts=extra)

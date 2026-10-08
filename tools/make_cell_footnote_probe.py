#!/usr/bin/env python3
"""Footnotes referenced in a table cell: where Word sets the note, how it numbers it, and
what the row does at a page's foot when its note does not fit.

4.13 measured footnotes referenced in the body's paragraphs; one referenced in a cell
took no room and was not drawn (``footnotes-not-drawn``).  One document per setting (no
``settings.xml``, mode 15), every case starting a page, the notes and separators as
``make_footnote_draw_probe.py`` writes them (``FootnoteText`` Calibri 10, single):

* ``short`` -- a body reference, a table of two rows -- a reference in a cell of the
  first, one in each cell of the second -- and a body reference after it;
* ``foot N`` -- body lines filling the page to a few lines from its foot, then a table of
  one-line rows whose third row holds a reference to a note of three lines: swept so the
  row, its note, or both fit;
* ``split`` -- body lines filling the page, then a table whose second row holds a cell of
  six one-line paragraphs, the fifth with a reference, the row split at the foot.

Reader: ``read_cell_footnote_probe.py``.
"""

from __future__ import annotations

import probe_docx
import wml
from make_footnote_draw_probe import (FACE, FOOTNOTES_CONTENT_TYPE, FOOTNOTES_REL, SPACING, Note, _p, reference,
                                      section, separators, text, words)

SETTINGS = {"none": None, "15": 15}
FILLS = (42, 43, 44, 45, 46, 47, 48)


def _cell(content: str, width: int = 3000) -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{content}</w:tc>'


def _table(rows: list[list[str]]) -> str:
    border = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    borders = "".join(border.format(side) for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    pr = ('<w:tblW w:w="6000" w:type="dxa"/>' f"<w:tblBorders>{borders}</w:tblBorders>"
          '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/>'
          '<w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
          '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
          ' w:noHBand="1" w:noVBand="1"/>')
    body = "".join("<w:tr>" + "".join(_cell(cell) for cell in row) + "</w:tr>" for row in rows)
    return (f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/>"
            f"</w:tblGrid>{body}</w:tbl>")


def content() -> tuple[str, list[Note]]:
    notes: list[Note] = []

    def note(value: Note) -> str:
        notes.append(value)
        return reference(len(notes))

    body = _p(text("Short page, body reference") + note(Note(("body before " + words(0, 5),))) + text("."))
    body += _table([[_p(text("Cell one") + note(Note(("cell one " + words(1, 6),)))), _p(text("Cell two"))],
                    [_p(text("Cell three") + note(Note(("cell three " + words(2, 4),)))),
                     _p(text("Cell four") + note(Note(("cell four " + words(3, 30),))))]])
    body += _p(text("After the table") + note(Note(("body after " + words(4, 5),))) + text("."))
    for fill in FILLS:
        body += _p(text(f"Foot {fill}"), pageBreakBefore=True)
        body += "".join(_p(text(f"F{k} " + words(k, 6))) for k in range(fill))
        body += _table([[_p(text("R1a")), _p(text("R1b"))], [_p(text("R2a")), _p(text("R2b"))],
                        [_p(text("R3a") + note(Note((f"foot {fill} " + words(fill, 40),)))), _p(text("R3b"))],
                        [_p(text("R4a")), _p(text("R4b"))]])
        body += _p(text(f"After foot {fill}"))
    for fill in (44, 45, 46, 47, 48):
        body += _p(text(f"Split {fill}"), pageBreakBefore=True)
        body += "".join(_p(text(f"S{k} " + words(k, 6))) for k in range(fill))
        cell = "".join(_p(text(f"C{k}") + (note(Note((f"split {fill} " + words(k, 8),))) if k == 4 else ""))
                       for k in range(6))
        body += _table([[_p(text("H1a")), _p(text("H1b"))], [cell, _p(text("D1"))], [_p(text("E1a")), _p(text("E1b"))]])
        body += _p(text(f"After split {fill}"))
    return body, notes


def build(setting: str) -> bytes:
    body, notes = content()
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True),
         wml.style("paragraph", "FootnoteText", based_on="Normal", ppr_={"spacing": SPACING}, rpr_={"sz": 20, "szCs": 20}),
         wml.style("character", "FootnoteReference", rpr_={"vertAlign": "superscript"})],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    )
    footnotes = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 f'<w:footnotes xmlns:w="{probe_docx.W_NS}">{separators()}'
                 + "".join(n.xml(k + 1) for k, n in enumerate(notes)) + "</w:footnotes>")
    extra = [("word/footnotes.xml", FOOTNOTES_CONTENT_TYPE, FOOTNOTES_REL, footnotes)]
    mode = SETTINGS[setting]
    if mode is not None:
        extra.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=mode))
    return probe_docx.package(body, final_section=section(), styles=styles, extra_parts=tuple(extra))

#!/usr/bin/env python3
"""Section breaks and the pages they make: ``nextPage``, ``continuous``, ``oddPage`` and
``evenPage``, and a change of page size and margins at each.

The pagination probe's sections document happened to start its ``oddPage`` section on an
odd page, so it could not show whether Word adds a blank page when the parity is wrong.
Here every section is short (a paragraph of a few lines, or of 70 to run over a page) and
the sequence is chosen so that each kind of break meets each parity: a section that ends
on an odd page followed by an ``oddPage`` one, and so on.  Page sizes and margins change
at most breaks.  Which page each section's lines are on is read from Word's export.
"""

from __future__ import annotations

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "14": 14, "15": 15}
SIZES = {"letter": (12240, 15840), "a4": (11906, 16838), "a5": (8391, 11906)}

#: (how the section starts, page size, top/bottom margin, lines in its paragraph)
SECTIONS = (
    (None, "letter", 1440, 5),
    ("oddPage", "a4", 1440, 5),      # after page 1 (odd): a blank page 2, then page 3
    ("oddPage", "a5", 720, 5),       # after page 3: a blank page 4
    ("evenPage", "letter", 1440, 5),  # after page 5: page 6
    ("evenPage", "a4", 1080, 70),    # after page 6: a blank 7, then 8 and 9
    ("oddPage", "letter", 1440, 5),  # after page 9: a blank 10
    ("nextPage", "a5", 1440, 70),    # page 12 onwards
    ("continuous", "a5", 720, 5),    # on the same page
    ("evenPage", "letter", 1440, 5),
    ("nextPage", "a4", 1440, 5),
    ("oddPage", "a4", 1440, 5),
    ("continuous", "letter", 1440, 70),  # a size change at a continuous break
    ("evenPage", "a5", 1440, 5),
)


def _section(kind, size, margin) -> str:
    width, height = SIZES[size]
    return ("<w:sectPr>" + (f'<w:type w:val="{kind}"/>' if kind else "")
            + f'<w:pgSz w:w="{width}" w:h="{height}"/>'
            f'<w:pgMar w:top="{margin}" w:right="1440" w:bottom="{margin}" w:left="1440"'
            ' w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="708"/></w:sectPr>')


def blocks() -> list[str]:
    run = {"rFonts": FACE, "sz": 22, "szCs": 22}
    out = []
    for number, (kind, size, margin, lines) in enumerate(SECTIONS):
        text = f"<w:r>{wml.rpr(**run)}" + "<w:br/>".join(
            f'<w:t xml:space="preserve">Section {number} l{k}</w:t>' for k in range(lines)) + "</w:r>"
        sect = _section(kind, size, margin) if number < len(SECTIONS) - 1 else ""
        out.append(wml.paragraph(text, mark=run, sect=sect,
                                 spacing={"before": 0, "after": 0, "line": 240, "lineRule": "auto"}))
    return out


def build(setting: str) -> bytes:
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package("".join(blocks()), final_section=_section(*SECTIONS[-1][:3]), styles=styles,
                              extra_parts=extra)

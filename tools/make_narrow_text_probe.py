#!/usr/bin/env python3
"""Lines with little or no room: a paragraph whose indents leave its text no width, or
less than a word's.

docx-agent (round 21) saw Word set a cell paragraph whose indents run past its cell one
word to a line where docx2svg broke inside the words.  Every case is a page: ``Case N``,
then a paragraph of words of 2 to 13 letters (``ab``, ``Word``, ... ``Extraordinary``),
in a text column of 6940 twips (a two-column table's cells 3470, margins 108):

* ``fixed`` -- a fixed-layout table, the first cell's paragraph indented to leave its
  text ``-2000``, ``-746``, ``0``, ``100`` or ``400`` twips;
* ``body`` -- the same widths in the body (left and right indents);
* ``autofit, one cell`` -- an autofit table (``w:tblW`` 6940 ``dxa``), the first cell
  indented: its content fits the room;
* ``autofit, both`` -- both cells indented alike (the second holding one word): their
  narrowest content is wider than the room (``make_autofit_over_probe.py``).

One document per compatibility setting (none, 15).  Calibri 11, no paragraph spacing.
Reader: ``read_narrow_text_probe.py``.
"""

from __future__ import annotations

import probe_docx
import wml
from make_table_pages_probe import FACE, _cell, _table

SETTINGS = {"none": None, "15": 15}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
WORDS = "ab Word cde Narrow fghij Extraordinary kl Mnopq Room rstuvwx yz"
#: The page's margins: a text column of 6940 twips.
LEFT, RIGHT = 1418, 3548
COLUMN = 11906 - LEFT - RIGHT
#: A cell's text width: 3470 less its margins.
CELL = 3470 - 2 * 108


def _p(text: str, **ind) -> str:
    return wml.paragraph(wml.run(text), spacing=SPACING, ind=dict(ind))


def _autofit(rows: str) -> str:
    return _table(rows, (3470, 3470)).replace('<w:tblLayout w:type="fixed"/>', "")


def _cases() -> list[tuple[str, str]]:
    out = []
    for width in (-2000, -746, 0, 100, 400):
        rows = f"<w:tr>{_cell(_p(WORDS, left=CELL - width), 3470)}{_cell(_p('Beside'), 3470)}</w:tr>"
        out.append((f"fixed, width {width}", _table(rows, (3470, 3470))))
        half = (COLUMN - width) // 2
        out.append((f"body, width {width}", _p(WORDS, left=half, right=COLUMN - width - half)))
    for width in (-2000, 0, 400):
        rows = f"<w:tr>{_cell(_p(WORDS, left=CELL - width), 3470)}{_cell(_p('Beside'), 3470)}</w:tr>"
        out.append((f"autofit, one cell, width {width}", _autofit(rows)))
    for width in (-2000, -746, -100, 0, 50, 100, 400):
        rows = (f"<w:tr>{_cell(_p(WORDS, left=CELL - width), 3470)}"
                f"{_cell(_p('Beside', left=CELL - width), 3470)}</w:tr>")
        out.append((f"autofit, both, width {width}", _autofit(rows)))
    return out


CASES = _cases()


def body() -> str:
    out = ""
    for number, (_, content) in enumerate(CASES):
        out += wml.paragraph(wml.run(f"Case {number}"), pageBreakBefore=bool(number)) + content
        out += wml.paragraph(wml.run(f"After {number}"), spacing=SPACING)
    return out


def styles() -> str:
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    )


def section() -> str:
    return ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="1440" w:right="{RIGHT}" w:bottom="1440" w:left="{LEFT}" w:header="720" '
            'w:footer="720" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>')


def build(setting: str) -> bytes:
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return probe_docx.package(body(), final_section=section(), styles=styles(), extra_parts=extra)

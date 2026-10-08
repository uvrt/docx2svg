#!/usr/bin/env python3
"""Autofit tables narrower than their narrowest content: what Word does with them.

Stage 7b (ROADMAP.md, "Tables -- measured") left one case of an autofit table sized from
its content as a stop: the columns' narrowest content -- each its widest word with the
cell's margins -- adding up to more than the room (``table.autofit_room``).  A table of
nine narrow columns headed by two-word labels on a portrait page is one.  This probe
measures it: whether Word keeps each column at its narrowest content and lets the table
run past the margin, or keeps it in the room and breaks words.

The tables are ``make_autofit_width_probe.py``'s: borders of ``w:sz`` 4, margins 108, a
``w:tblGrid`` wrong on purpose (every column 1,000 twips), words of ``x`` (95.3 twips
each in Calibri 11), A4 with the text column 9,164 twips.  Family ``over``:

* every column in ``dxa``, each holding one word wider than its width, together wider
  than the room -- three columns, five, and nine as a review table has them;
* columns with no width, the same words, and with many words beside the long ones;
* a column in ``dxa`` beside columns with no width;
* a single column whose one word is wider than the room;
* a table width in ``dxa`` and in percent over such content.

Three documents: no ``settings.xml``, mode 14 and mode 15.  Reader:
``read_autofit_over_probe.py``.
"""

from __future__ import annotations

import make_autofit_width_probe as base
from make_autofit_width_probe import SETTINGS, SPACING, Case, Cell, Table, dxa, words, x

import make_anchor_probe as anchor_probe
import wml


def _cases() -> tuple[Case, ...]:
    out: list[Case] = []

    def add(note: str, *rows, width=(0, "auto")) -> None:
        number = len(out)
        labelled = []
        for r, row in enumerate(rows):
            cells = []
            for c, cell in enumerate(row):
                label = f"{number}.{'abcdefghijk'[c]}{r + 1}"
                texts = tuple(f"{label} {text}".strip() if k == 0 else text for k, text in enumerate(cell.texts))
                cells.append(Cell(texts, cell.width, cell.span, cell.no_wrap, cell.margins, cell.ind, cell.jc))
            labelled.append(tuple(cells))
        out.append(Case("over", note, Table(tuple(labelled), width)))

    add("three in dxa of 2000, words of 3,336", *[(Cell((x(35),), dxa(2000)),) * 3])
    add("three in dxa of 4000, words of 3,336 and 2,000", (Cell((x(35),), dxa(4000)), Cell((x(35),), dxa(4000)),
                                                          Cell((x(21),), dxa(4000))))
    add("five in dxa of 1000, words of 2,001", *[(Cell((x(21),), dxa(1000)),) * 5])
    add("nine in dxa, two-word labels", (
        Cell((f"{x(14)} {x(4)}",), dxa(1400)), Cell((f"{x(11)} {x(3)}",), dxa(1080)),
        Cell((f"{x(12)} {x(2)}",), dxa(940)), Cell((f"{x(11)} {x(3)}",), dxa(940)),
        Cell((f"{x(10)} {x(4)}",), dxa(880)), Cell((f"{x(12)} {x(2)}",), dxa(940)),
        Cell((f"{x(11)} {x(3)}",), dxa(880)), Cell((f"{x(15)} {x(2)}",), dxa(1040)),
        Cell((f"{x(13)} {x(3)}",), dxa(920))),
        tuple(Cell((x(3),), dxa(w)) for w in (1400, 1080, 940, 940, 880, 940, 880, 1040, 920)))
    add("three with no width, words of 3,336", *[(Cell((x(35),)),) * 3])
    add("three with no width, long words and many words", (Cell((f"{x(35)} {words(30)}",)),
                                                          Cell((f"{x(35)} {words(10)}",)), Cell((x(35),))))
    add("dxa 2000 beside two with no width", (Cell((x(35),), dxa(2000)), Cell((x(35),)), Cell((x(35),))))
    add("one column, a word wider than the room", (Cell((x(110),)),))
    add("one column in dxa 3000, a word wider than the room", (Cell((x(110),), dxa(3000)),))
    add("tblW 6000 dxa over three in dxa of 2000", *[(Cell((x(35),), dxa(2000)),) * 3], width=dxa(6000))
    add("tblW 5000 pct over three with no width", *[(Cell((x(35),)),) * 3], width=(5000, "pct"))
    return tuple(out)


CASES = _cases()


def body() -> str:
    out = ""
    for number, case in enumerate(CASES):
        out += wml.paragraph(wml.run(f"Case {number}"), mark={}, spacing=SPACING, pageBreakBefore=True)
        out += case.table.xml()
        out += wml.paragraph(wml.run(f"After {number}"), mark={}, spacing=SPACING)
    return out


def build(setting: str) -> bytes:
    """``setting``: ``none``, ``14`` or ``15``."""
    mode = SETTINGS[setting]
    extra = () if mode is None else (wml.settings_part({"val": "en-GB"}, compatibility_mode=mode),)
    return anchor_probe.package(body(), "none", extra=extra)


DOCUMENTS = tuple(f"autofit-over-{s}" for s in SETTINGS)

assert base.SETTINGS is SETTINGS

if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"autofit-over-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

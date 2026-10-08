#!/usr/bin/env python3
"""A line break that ends a paragraph: the paragraph's mark goes on a line of its own.

Found by the local corpora's text boxes, whose list paragraphs end in ``<w:br/>``: Word
draws an empty line after the break -- the mark's -- and the model drew none, so every
line after it was a line high.  Here each case is a paragraph ending in one, two or three
breaks, followed by a line of text whose baseline says how tall the empty lines are:
under ``auto`` 240 and 360, ``exact`` 400 and ``atLeast`` 480, with space after, with
a larger mark, in a list, in a table cell and after a break followed by a space.

Two documents: no ``settings.xml`` and mode 15.  Reader: ``read_break_end_probe.py``.
"""

from __future__ import annotations

import make_anchor_probe as anchor_probe
import wml

SETTINGS = {"none": "none", "15": "15"}
BREAK = "<w:r><w:br/></w:r>"


def _cases() -> list[tuple[str, str]]:
    out = []
    for breaks in (1, 2, 3):
        for line, rule in ((240, "auto"), (360, "auto"), (400, "exact"), (480, "atLeast")):
            spacing = {"before": 0, "after": 0, "line": line, "lineRule": rule}
            out.append((f"{breaks} break(s), {rule} {line}",
                        wml.paragraph(wml.run(f"Ends in {breaks} break(s) under {rule} {line}.") + BREAK * breaks,
                                      mark={}, spacing=spacing)))
    out.append(("space after", wml.paragraph(wml.run("Ends in a break, with space after.") + BREAK, mark={},
                                             spacing={"before": 0, "after": 240, "line": 240, "lineRule": "auto"})))
    out.append(("a larger mark", wml.paragraph(wml.run("Ends in a break, its mark 20 pt.") + BREAK,
                                               mark={"sz": 40, "szCs": 40})))
    out.append(("a space after the break", wml.paragraph(wml.run("Ends in a break and a space.") + BREAK
                                                         + wml.run(" "), mark={})))
    out.append(("a table cell", wml.table([[wml.paragraph(wml.run("A cell's paragraph ending in a break.") + BREAK,
                                                          mark={}) + wml.paragraph(wml.run("Next in the cell."),
                                                                                   mark={})]])))
    return out


CASES = _cases()


def body() -> str:
    out = ""
    for number, (note, block) in enumerate(CASES):
        out += anchor_probe._p(f"Case {number} {note}", pageBreakBefore=True)
        out += block
        out += anchor_probe._p(f"Case {number} after it.")
    return out


def build(setting: str) -> bytes:
    return anchor_probe.package(body(), setting)

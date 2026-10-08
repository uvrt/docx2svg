#!/usr/bin/env python3
"""The line-advance accumulator: one face, one size and one line rule per page.

Phase 0.6 measured the *ratio* -- Calibri 11 pt advances by ``(asc + desc + gap) / upm``
em, 55.949 px at 300 dpi -- and could not reproduce the *positions*: rounding
``origin + n * a`` gets 61 of 70 baselines and is 1 px off on the other nine.  The
quantity is an integer (every baseline is a whole device pixel), so the model has to
produce integers, and the only way to find which integers is to vary what the rounding
could depend on, one thing at a time:

* **size** -- every half point from 6 to 12 pt and then up to 24 pt, so the fractional
  part of the advance in device pixels takes many values and a rounding rule that is
  right by coincidence at one size is wrong at the next;
* **face** -- Calibri and Aptos have identical ``asc + desc`` (2500/2048) and would hide
  any rule that depends on how the total splits into ascent and descent; Arial carries a
  non-zero ``hhea.lineGap``; Times New Roman and Cambria split differently again;
  Verdana is tall;
* **line rule** -- ``auto`` (a multiple of the font's height), ``exact`` and ``atLeast``,
  each in its own group, with values chosen off the device grid.

Each group starts on a new page (``w:pageBreakBefore``), so every group's first baseline
is measured from the same top margin and every group is a clean arithmetic sequence of
single-line paragraphs.  Each paragraph is the single glyph ``H`` with a paragraph mark of
the same size: the mark takes part in the line height, so leaving it at the default size
would make a 24 pt line also carry an 11 pt mark.

Throwaway: generated, exported through ``tools/oracle.py``, never committed.  Measure
with ``python tools/read_line_advance_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx as w

FACES = ("Calibri", "Aptos", "Arial", "Times New Roman", "Cambria", "Verdana")
HALF_POINTS = tuple(range(12, 25)) + (26, 28, 30, 32, 36, 40, 44, 48)

#: (rule, w:line) pairs for the rule sweep.  ``auto`` values are 240ths of the single
#: line; ``exact`` / ``atLeast`` values are twips, chosen so that most are *not* whole
#: device pixels (a pixel is 4.8 twips).
RULES = (
    ("auto", 240),
    ("auto", 276),
    ("auto", 360),
    ("auto", 233),
    ("exact", 240),
    ("exact", 277),
    ("exact", 301),
    ("atLeast", 1),
    ("atLeast", 277),
    ("atLeast", 400),
)
RULE_FACES = ("Calibri", "Times New Roman", "Arial")
RULE_HALF_POINTS = (13, 17, 22, 29, 40)

#: The second, denser rule sweep.  The first showed that ``exact`` and ``atLeast`` put the
#: rounding slack in different places from single spacing, and that ``auto`` multiples
#: other than 240 follow neither -- so this one spreads the fractional part of the line
#: height across [0, 1) for each rule, which is what the anchoring depends on.
RULES_DENSE = (
    tuple(("exact", v) for v in (250, 263, 290, 310, 333, 350, 411, 457))
    + tuple(("atLeast", v) for v in (290, 333, 350, 411, 457, 500))
    + tuple(("auto", v) for v in (200, 220, 250, 264, 288, 300, 320, 400, 480))
)
DENSE_FACES = ("Calibri", "Times New Roman", "Cambria")
DENSE_HALF_POINTS = (13, 19, 22, 29, 40)

#: Usable column height of an A4 page with 1-inch margins, in device px (9.69 in).
_COLUMN_PX = (16838 - 2 * 1440) / 4.8


@dataclass(frozen=True)
class Group:
    face: str
    half_points: int
    rule: str = "auto"
    line: int = 240

    @property
    def lines(self) -> int:
        """Enough lines to fill most of a page but never to spill onto the next."""
        size_px = self.half_points / 2 * 300 / 72
        if self.rule == "auto":
            pitch = 1.35 * size_px * self.line / 240
        elif self.rule == "exact":
            pitch = self.line / 4.8
        else:
            pitch = max(1.35 * size_px, self.line / 4.8)
        return max(3, min(80, int(0.9 * _COLUMN_PX / pitch)))


def groups(sweep: str, face: str) -> list[Group]:
    """One document per (sweep, face).

    Per face rather than one document for everything: a 126-page, 5,478-paragraph
    probe ran Word past the export script's 180 s AppleEvent timeout twice (-1712),
    and a smaller document is also a smaller thing to re-export when one face changes.
    """
    if sweep == "sizes":
        return [Group(face, hp) for hp in HALF_POINTS]
    if sweep == "rules":
        return [Group(face, hp, rule, line) for rule, line in RULES for hp in RULE_HALF_POINTS]
    if sweep == "dense":
        return [Group(face, hp, rule, line) for rule, line in RULES_DENSE for hp in DENSE_HALF_POINTS]
    raise ValueError(sweep)


def build(sweep: str, face: str) -> bytes:
    body = []
    for group in groups(sweep, face):
        mark = w.rpr(group.face, group.half_points)
        for index in range(group.lines):
            body.append(
                w.paragraph(
                    w.run("H", group.face, group.half_points),
                    line=group.line,
                    line_rule=group.rule,
                    page_break_before=index == 0,
                    mark_rpr=mark,
                )
            )
    return w.package("".join(body))

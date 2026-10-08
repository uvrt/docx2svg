#!/usr/bin/env python3
"""Does Word kern, and does tracking (``w:rPr/w:spacing``) go before or after it?

The sibling measured that PowerPoint applies the OpenType ``kern`` feature to body text
and that its chart engine lays text out *unkerned* while drawing it kerned.  Neither
answer is inherited here; this probe asks Word 16.106 directly, one variable at a time:

* **kerning off vs on** -- ``w:kern`` is Word's switch ("kern fonts at or above N half
  points").  With it absent, does Word kern anyway?  With it present, which pairs?
* **tracking** -- ``w:spacing/@w:val`` in twips, alone and together with kerning, at a
  value that is not a whole number of 1/4096 pt (a twip is 204.8 such units), so any
  rounding of the tracking itself shows.
* **run boundaries** -- the same string once as one run and once with every glyph its
  own run (alternating colours, which also gives every glyph an exact pen position in
  the PDF; see ``quartz_pdf.py``).  A kern pair that straddles two runs is the case a
  shaper typically drops.

Strings are chosen for large, well-known kern pairs (``AV``, ``To``, ``Ty``, ``LT``,
``Yo``, ``WA``) with ``HH`` as the control: ``H`` kerns with nothing in these faces, so a
width change on the control line is tracking and nothing else.

Last, one Times New Roman string of Greek pairs (``ΤΩ``, ``Υμ``) that the macOS system
copy of the face kerns and the copy in Word's own bundle does not -- which of the two
Word lays out with is otherwise invisible, since their advances are identical.

Throwaway: exported through ``tools/oracle.py``, never committed.  Measure with
``python tools/read_kerning_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx as w

FACES = ("Calibri", "Arial", "Times New Roman")
HALF_POINTS = (23, 40)
STRINGS = ("HHHHHHHH", "AVAVAVAV", "ToTyToTy", "LTWAYoVA")
GREEK = "ΤΩΤΩΥμΥμ"
#: (label, kern half points or None, tracking twips or None)
VARIANTS = (
    ("plain", None, None),
    ("kern", 2, None),
    ("track7", None, 7),
    ("kern+track7", 2, 7),
    ("track-5", None, -5),
)


@dataclass(frozen=True)
class Item:
    face: str
    half_points: int
    text: str
    variant: str
    kern: int | None
    track: int | None
    per_glyph: bool


def items() -> list[Item]:
    out = []
    for face in FACES:
        for hp in HALF_POINTS:
            for text in STRINGS:
                for label, kern, track in VARIANTS:
                    for per_glyph in (False, True):
                        out.append(Item(face, hp, text, label, kern, track, per_glyph))
    for label, kern, track in VARIANTS[:2]:
        out.append(Item("Times New Roman", 40, GREEK, label, kern, track, False))
    return out


def _paragraph(item: Item) -> str:
    props = dict(kern_half_points=item.kern, spacing_twips=item.track)
    if item.per_glyph:
        runs = "".join(
            w.run(ch, item.face, item.half_points, color="C00000" if i % 2 else "000000", **props)
            for i, ch in enumerate(item.text)
        )
    else:
        runs = w.run(item.text, item.face, item.half_points, **props)
    return w.paragraph(runs, mark_rpr=w.rpr(item.face, item.half_points))


def build() -> bytes:
    return w.package("".join(_paragraph(item) for item in items()))

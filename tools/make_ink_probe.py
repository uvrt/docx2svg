#!/usr/bin/env python3
"""Ink or advances: at what size Word's raster draws a glyph's outline.

Word lays glyphs out at the exact size (11 pt is 45.83 device px) and its PDF scales
their outlines by the size rounded to whole device px (46: Quartz's ``Tm``; Phase 2).
The difference is 0.36% at 11 pt, too small for a real page to separate; so this probe
sets text at the sizes whose rounding moves most, one size per page, in both
directions:

* rounded **up**: 9.5 pt (39.58 -> 40 px, +1.05%), 10.5 pt (43.75 -> 44, +0.57%) and
  11 pt (45.83 -> 46, +0.36%);
* rounded **down**: 8.5 pt (35.42 -> 35, -1.2%), 7.5 pt (31.25 -> 31, -0.8%), 14.5 pt
  (60.42 -> 60, -0.7%) and 20.5 pt (85.42 -> 85, -0.5%);
* and 12 pt (50 px, no rounding at all) as the control, where the two must agree.

Each page is 24 lines of ``HOHOxoxo nm`` repeated, in Calibri, Times New Roman and Arial
(eight lines each).  ``tools/fidelity.py --ink`` renders ours at both glyph sizes and
scores each page against Word's.
"""

from __future__ import annotations

import probe_docx
import wml

#: Half points, and the device px they are and are drawn at.
SIZES = (17, 19, 21, 22, 24, 29, 41, 15)
FACES = ("Calibri", "Times New Roman", "Arial")
TEXT = "HOHOxoxo nm HOHOxoxo nm HOHOxoxo nm HOHOxoxo"


def fonts(face: str) -> dict:
    return {"ascii": face, "hAnsi": face, "eastAsia": face, "cs": face}


def build() -> bytes:
    body = []
    for size in SIZES:
        for index, face in enumerate(FACES * 8):
            text = {"rFonts": fonts(face), "sz": size, "szCs": size}
            words = TEXT if size < 30 else TEXT[:23]
            body.append(wml.paragraph(wml.run(words, **text), mark=text, pageBreakBefore=index == 0,
                                      spacing={"before": 0, "after": 0, "line": 240, "lineRule": "auto"}))
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": fonts("Calibri"), "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    return probe_docx.package("".join(body), styles=styles)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("ink.docx")
    target.write_bytes(build())
    print(target)

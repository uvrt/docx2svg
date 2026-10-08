#!/usr/bin/env python3
"""Where the baseline of a paragraph with a bottom border sits.

The one baseline of the real-world samples (``tests/fixtures/samplelib/``) that the
resolved cascade plus Phase 2's model got wrong is the Title: Calibri 26 pt, single
spaced, ``w:pBdr/w:bottom`` 1 pt at 4 pt, drawn 1 px lower than an unbordered line (in
all three documents).  The paragraph *after* it is exact, so the border's total height
(space + width) is right and what is wrong is where the text sits in the taller line.

One group per page: 11 pairs of (bordered paragraph, plain paragraph), so the bordered
line's baseline and the space the border takes are both read, at one size and border.
Sizes are chosen so the fractional pixel of the line pitch, with and without the border,
covers both sides of 1/2.

Throwaway: exported through ``tools/oracle.py``; measured by ``read_border_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import probe_docx
import wml

CALIBRI = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
PAIRS = 11


@dataclass(frozen=True)
class Group:
    half_points: int
    #: (width in eighths of a point, space in points)
    width: int = 8
    space: int = 4


GROUPS = tuple(
    Group(hp, width, space)
    for hp in (22, 29, 40, 52)
    for width, space in ((8, 4), (4, 1), (12, 0), (6, 7))
)


def build() -> bytes:
    body = []
    for group in GROUPS:
        rpr = {"rFonts": CALIBRI, "sz": group.half_points}
        border = (f'<w:bottom w:val="single" w:sz="{group.width}" w:space="{group.space}"'
                  ' w:color="000000"/>')
        for index in range(PAIRS):
            body.append(wml.paragraph(wml.run("Hxb", **rpr), mark=rpr,
                                      pageBreakBefore=index == 0, pBdr=f"{border}"))
            body.append(wml.paragraph(wml.run("Hxp", **rpr), mark=rpr))
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": CALIBRI, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    return probe_docx.package("".join(body), styles=styles)

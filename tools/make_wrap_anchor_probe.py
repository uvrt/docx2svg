#!/usr/bin/env python3
"""Text above and below a floating drawing: ``wrapTopAndBottom``.

A drawing whose text wraps top and bottom keeps every line off the band it occupies --
its box widened by ``distT`` above and ``distB`` below, across the column -- so the text
continues below it.  That moves lines, and so pages.  Every case is a page: a heading,
paragraphs of text, the anchoring paragraph, and more text after it; the picture (a
16 x 16 PNG of one colour, stretched) is read from Word's PDF, and every line's baseline
by the glyph check.  Families (``CASES``):

* ``para`` -- positioned against its paragraph at offsets 0, below the first line, above
  the paragraph (negative), and inside the paragraph's third line; the anchor in the
  first line and in the fourth;
* ``page`` -- positioned against the page and the margin, over lines of paragraphs before
  and after the anchor's;
* ``dist`` -- ``distT`` / ``distB`` of 0, 1 twip, 114,300 EMU and 1/4 inch;
* ``foot`` -- a drawing that does not fit above the bottom margin, anchored near the foot;
* ``size`` -- heights stepped so the band's bottom takes every fraction of a pixel.

Two documents: no ``settings.xml`` and mode 15.  Reader: ``read_wrap_anchor_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import make_anchor_probe as anchor_probe
import wml
from make_anchor_probe import Anchor

SETTINGS = {"none": "none", "15": "15"}
WORDS = ("A line of text that wraps above and below the drawing, as many times as it takes to fill "
         "a few lines of the column, so that the drawing has lines on both sides of it here.")
TOP_BOTTOM = "<wp:wrapTopAndBottom/>"


@dataclass(frozen=True)
class Case:
    family: str
    anchor: Anchor
    #: Paragraphs of text before the anchor's paragraph, and the anchor's place in it:
    #: ``first`` (before its text) or ``fourth`` (in its fourth line).
    before: int = 1
    where: str = "first"
    #: Paragraphs after it.
    after: int = 3


def _cases() -> tuple[Case, ...]:
    out = []

    def pic(h, v, k, **kw):
        return Anchor(h, v, 1500000 + k * 3701, 500000 + k * 37003, wrap=TOP_BOTTOM, dist=kw.pop("dist", (0, 0, 114300,
                                                                                                        114300)),
                      **kw)

    for k, (offset, where) in enumerate(((0, "first"), (300000, "first"), (-200000, "first"), (0, "fourth"),
                                         (400000, "fourth"), (-150000, "fourth"))):
        out.append(Case("para", pic(("column", "offset", 200000), ("paragraph", "offset", offset), k), where=where))
    for k, (v, before) in enumerate(((("page", "offset", 2500000), 3), (("margin", "offset", 1200000), 3),
                                     (("margin", "offset", 3000000), 1), (("page", "offset", 1500000), 0))):
        out.append(Case("page", pic(("column", "offset", 0), v, k + 6), before=before))
    for k, dist in enumerate(((0, 0, 0, 0), (635, 635, 0, 0), (114300, 114300, 0, 0), (228600, 0, 0, 0),
                              (0, 228600, 0, 0))):
        out.append(Case("dist", pic(("column", "offset", 100000), ("paragraph", "offset", 100000), k + 10, dist=dist)))
    for k, before in enumerate((14, 15)):
        out.append(Case("foot", pic(("column", "offset", 0), ("paragraph", "offset", 0), k + 15), before=before))
    for k in range(8):
        out.append(Case("size", Anchor(("column", "offset", 0), ("paragraph", "offset", 60000), 1200000,
                                       400000 + k * 1270, wrap=TOP_BOTTOM,
                                       dist=(0, 0, 0, 0))))
    return tuple(out)


CASES = _cases()


def _p(text: str) -> str:
    return wml.paragraph(wml.run(text), mark={}, spacing={"before": 0, "after": 120, "line": 259, "lineRule": "auto"})


def body() -> str:
    out = ""
    for number, case in enumerate(CASES):
        out += anchor_probe._p(f"Case {number} {case.family} heading", pageBreakBefore=True)
        for k in range(case.before):
            out += _p(f"Case {number} before {k}. {WORDS}")
        drawing = case.anchor.xml(number + 1)
        spacing = {"before": 0, "after": 120, "line": 259, "lineRule": "auto"}
        if case.where == "first":
            runs = drawing + wml.run(f"Case{number} anchor. {WORDS} {WORDS}")
        else:
            runs = wml.run(f"Case{number} anchor. {WORDS} {WORDS[:60]}") + drawing + wml.run(WORDS[60:])
        out += wml.paragraph(runs, mark={}, spacing=spacing)
        for k in range(case.after):
            out += _p(f"Case {number} after {k}. {WORDS}")
    return out


def build(setting: str) -> bytes:
    return anchor_probe.package(body(), setting)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"wrap-anchor-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

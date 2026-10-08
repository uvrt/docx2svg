#!/usr/bin/env python3
"""Floating drawings wrapped top and bottom, anchored in consecutive paragraphs: where
each goes, and which page it goes on.

``make_wrap_anchor_probe.py`` measured one ``wrapTopAndBottom`` drawing at a time: it is
positioned from the page as laid out without it, and keeps the text off its band.  A
drawing anchored in the paragraph after another's is positioned from where that paragraph
stands once the first drawing has pushed it down -- docx-agent saw Word stack such
drawings onto two pages where docx2svg drew them on one (ROADMAP.md, F.24).  Every case is
a page or two: a heading, then paragraphs each anchoring a picture against its paragraph
(``wrapTopAndBottom``), and text after them.  Cases (``CASES``):

* two pictures in consecutive paragraphs, at offset 0;
* the same with a paragraph of text between them;
* the same at offsets of 0.25 and 0.1 inch;
* three, the third too tall for what is left of the page: its paragraph and its picture
  go on to the next page;
* the same with the third anchored in its paragraph's second line;
* a picture positioned against the margin, then one against its paragraph below it;
* two pictures in one paragraph, at offsets 0 and 1.6 inch.

Two documents: no ``settings.xml`` and mode 15.  Reader: ``read_anchor_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import make_anchor_probe as anchor_probe
import wml
from make_anchor_probe import Anchor

SETTINGS = {"none": "none", "15": "15"}
WORDS = "Text after the picture, a line or two of it, so that each paragraph's text shows where the band ends."
INCH = 914400


@dataclass(frozen=True)
class Case:
    note: str
    #: Per paragraph: its pictures as ``(vertical offset, height)`` in EMU, against the
    #: paragraph (or ``("margin", offset, height)``), and where they are anchored --
    #: ``first`` (before the text) or ``second`` (after a first line of text).
    paragraphs: tuple
    where: str = "first"


def _cases() -> tuple[Case, ...]:
    return (
        Case("two in consecutive paragraphs", (((0, int(1.5 * INCH)),), ((0, INCH),), ())),
        Case("two, a paragraph between", (((0, int(1.5 * INCH)),), (), ((0, INCH),), ())),
        Case("two at offsets", (((INCH // 4, int(1.5 * INCH)),), ((INCH // 10, INCH),), ())),
        Case("three, the third past the foot", (((0, int(3.5 * INCH)),), ((0, int(2.5 * INCH)),),
                                               ((0, int(2.5 * INCH)),), ())),
        Case("three, the third past the foot, anchored after a line", (
            ((0, int(3.5 * INCH)),), ((0, int(2.5 * INCH)),), ((0, int(2.5 * INCH)),), ()), where="second"),
        Case("against the margin, then against its paragraph", (
            (("margin", INCH, INCH),), ((0, INCH),), ())),
        Case("two in one paragraph", (((0, INCH), (int(1.6 * INCH), INCH)), ())),
    )


CASES = _cases()


def body() -> str:
    out = ""
    number = 0
    for index, case in enumerate(CASES):
        out += anchor_probe._p(f"Case {index} {case.note}", pageBreakBefore=True)
        for k, pictures in enumerate(case.paragraphs):
            drawings = ""
            for picture in pictures:
                number += 1
                if picture[0] == "margin":
                    v = ("margin", "offset", picture[1])
                    height = picture[2]
                else:
                    v = ("paragraph", "offset", picture[0])
                    height = picture[1]
                drawings += Anchor(("column", "offset", 0), v, 2 * INCH, height, wrap="<wp:wrapTopAndBottom/>",
                                   dist=(0, 0, 114300, 114300), image=number % 4).xml(number)
            text = f"Case{index} paragraph {k}. {WORDS}"
            if not drawings:
                runs = wml.run(text)
            elif case.where == "first":
                runs = drawings + wml.run(text)
            else:
                runs = wml.run(f"Case{index} paragraph {k}, its first line.") + "<w:r><w:br/></w:r>" + drawings \
                    + wml.run(text)
            out += wml.paragraph(runs)
    return out


def build(setting: str) -> bytes:
    return anchor_probe.package(body(), setting)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"wrap-stack-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

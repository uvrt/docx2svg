#!/usr/bin/env python3
"""Text beside a floating drawing: ``wrapSquare``, ``wrapTight``, ``wrapThrough``.

How Word shortens and splits a line beside a drawing text wraps around (ROADMAP.md,
"Floating drawings -- measured", F.8).  Every case is a page: a heading, paragraphs of
text, the anchoring paragraph -- long enough to run past the drawing -- and paragraphs
after it; the picture (a 16 x 16 PNG of one colour, stretched) is read from Word's PDF,
and every glyph by the glyph check.  Two sets of documents (``DOCUMENTS``):

``wrap-side-<setting>`` (A4, left margin 1,442 twips: off the pixel grid), families
(``SIDE``):

* ``square`` -- the gap left or right of a drawing against the column's edge (100 to
  1,440 twips), long words beside a narrow gap, ``wrapText`` ``left`` / ``right`` /
  ``largest``, indents beside a drawing (left, first line, hanging, right), centred,
  right-aligned and justified text, two drawings side by side, stacked and overlapping,
  the drawing's top stepped through a line, ``distT`` / ``distB`` / ``distL`` / ``distR``,
  a drawing above its paragraph, one taller than its paragraph, one positioned against
  the margin and the page, a drawing wider than the column;
* ``min`` -- gaps of 340 to 400 twips either side, at 8, 11 and 20 pt: the narrowest
  segment that takes text;
* ``vertical`` -- the drawing's top and bottom stepped by 2 twips across a line's edges
  under ``auto`` 259 and 240, ``exact`` 320 and ``atLeast`` 400: which lines are beside it;
* ``polygon`` -- ``wrapTight`` and ``wrapThrough`` around a triangle, a diamond, a C and a
  polygon past the extent, with and without ``distL`` / ``distR``;
* ``more`` -- a paragraph's rest too long for any segment, ``wrapText`` ``left`` and
  ``right`` beside a narrow gap, first-line and hanging indents beside a drawing inside
  the indent, tabs right of a drawing, an effect extent, ``distT`` on a tight wrap, empty
  and centred paragraphs beside it, a justified paragraph, a drawing near the page's foot.

``wrap-edge-<setting>-<left margin>`` (left margins 1,442 and 1,440 twips), families
(``EDGE``):

* ``right`` -- the drawing's offset, extent and ``distR`` each swept by single twips, and
  an offset against the page: where the text right of it starts;
* ``left`` -- words of known width (``nnnn``, ``nni``) with the drawing's left edge swept
  by single twips about the end of the third, fifth and seventh: where the text left of it
  must end;
* ``tie`` -- right edges on a half pixel of the text column and of the page.

Settings: no ``settings.xml`` and mode 15.  Reader: ``read_wrap_side_probe.py``.
"""

from __future__ import annotations

import dataclasses
import math
from dataclasses import dataclass, field

import make_anchor_probe as anchor_probe
import wml
from make_anchor_probe import Anchor

SETTINGS = ("none", "15")
MARGINS = (1442, 1440)
#: EMU per twip.
T = 635
#: The text column's width, twips (A4 less 1,442 and 1,300).
COLUMN = 11906 - 1442 - 1300
#: 2,400,000 EMU, the pictures' usual width, in whole twips.
WIDTH = 3779
WORDS = ("Text that runs beside the drawing, line after line, shortened where the drawing is and whole "
         "where it is not, so that every line shows where Word lets the text go. ")
SPACING = {"before": 0, "after": 120, "line": 259, "lineRule": "auto"}
#: Word widths in the layout unit (1/4096 pt) at 11 pt Calibri, and the space's.
WIDTHS = {"nnnn": 4 * 1076 * 22, "nni": (2 * 1076 + 470) * 22}
SPACE = 463 * 22


def square(side: str = "bothSides") -> str:
    return f'<wp:wrapSquare wrapText="{side}"/>'


def polygon(kind: str, side: str, points) -> str:
    pts = "".join(f'<wp:lineTo x="{x}" y="{y}"/>' for x, y in points[1:])
    return (f'<wp:{kind} wrapText="{side}"><wp:wrapPolygon edited="1"><wp:start x="{points[0][0]}" '
            f'y="{points[0][1]}"/>{pts}</wp:wrapPolygon></wp:{kind}>')


TRIANGLE = ((0, 21600), (21600, 21600), (0, 0), (0, 21600))
DIAMOND = ((10800, 0), (21600, 10800), (10800, 21600), (0, 10800), (10800, 0))
C_SHAPE = ((0, 0), (21600, 0), (21600, 5400), (5400, 5400), (5400, 16200), (21600, 16200), (21600, 21600),
           (0, 21600), (0, 0))
OUTSIDE = ((-2000, -2000), (-2000, 23600), (23600, 23600), (23600, -2000), (-2000, -2000))
RECTANGLE = ((0, 0), (0, 21600), (21600, 21600), (21600, 0), (0, 0))


@dataclass(frozen=True)
class Case:
    family: str
    note: str
    anchors: tuple
    #: The anchoring paragraph's text (``None``: :data:`WORDS` six times), or its runs.
    text: str | None = None
    runs: str | None = None
    #: Paragraphs of text before and after the anchoring one.
    before: int = 1
    after: int = 1
    props: dict = field(default_factory=dict)


def _a(x: int, v: int = 200000, cx: int = 2400000, cy: int = 900000, wrap: str | None = None,
       dist=(0, 0, 0, 0), relative: str = "column", vertical: str = "paragraph", **kw) -> Anchor:
    """A picture ``x`` twips from ``relative``'s edge and ``v`` EMU from ``vertical``'s."""
    return Anchor((relative, "offset", x * T), (vertical, "offset", v), cx, cy, wrap=wrap or square(), dist=dist,
                  **kw)


def _aligned(align: str, wrap: str | None = None) -> Anchor:
    return Anchor(("column", "align", align), ("paragraph", "offset", 200000), 2400000, 900000,
                  wrap=wrap or square(), dist=(0, 0, 0, 0))


def _side() -> tuple[Case, ...]:
    out = []
    long_words = "Incomprehensibilities " * 40
    for g in (100, 200, 300, 400, 500, 600, 720, 900, 1080, 1440):
        out.append(Case("square", f"left gap {g}", (_a(g),)))
    for g in (100, 300, 500, 720, 1080, 1440):
        out.append(Case("square", f"right gap {g}", (_a(COLUMN - g - WIDTH),)))
    for g in (900, 1440, 2000):
        out.append(Case("square", f"left gap {g}, long words", (_a(g),), text=long_words))
    for side in ("left", "right", "largest"):
        out.append(Case("square", f"{side}, centred", (_aligned("center", square(side)),)))
    out.append(Case("square", "largest, left wider", (_a(3000, wrap=square("largest")),)))
    out.append(Case("square", "largest, right wider by 20", (_a((COLUMN - WIDTH) // 2 - 10,
                                                                 wrap=square("largest")),)))
    out.append(Case("square", "largest, right wider by 200", (_a((COLUMN - WIDTH) // 2 - 100,
                                                                  wrap=square("largest")),)))
    out.append(Case("square", "largest, left wider by the column, narrower by the indent",
                    (_a((COLUMN - WIDTH) // 2 + 100, wrap=square("largest")),), props={"ind": {"left": 1440}}))
    out.append(Case("square", "left at the left edge", (_aligned("left", square("left")),)))
    out.append(Case("square", "right at the right edge", (_aligned("right", square("right")),)))
    out.append(Case("square", "indent 720, first line 360, drawing left", (_aligned("left"),),
                    props={"ind": {"left": 720, "firstLine": 360}}))
    out.append(Case("square", "indent 720, hanging 360, drawing left", (_aligned("left"),),
                    props={"ind": {"left": 720, "hanging": 360}}))
    out.append(Case("square", "right indent 720, drawing right", (_aligned("right"),),
                    props={"ind": {"right": 720}}))
    out.append(Case("square", "indents 720, drawing in the middle", (_a(3000),),
                    props={"ind": {"left": 720, "right": 720}}))
    for jc in ("center", "right", "both"):
        out.append(Case("square", f"jc {jc}", (_a(3000),), props={"jc": jc}))
    out.append(Case("square", "two side by side", (_a(2000, cx=1200000), _a(5500, cx=1200000))))
    out.append(Case("square", "two stacked", (_a(1000, cy=600000), _a(5000, v=600000, cy=600000))))
    out.append(Case("square", "two overlapping", (_a(2000, cx=1500000), _a(3500, v=500000, cx=1500000))))
    for k in range(8):
        out.append(Case("square", f"top {150000 + k * 12000}", (_a(3000, v=150000 + k * 12000, cy=300000),)))
    for dist in ((114300, 0, 0, 0), (0, 114300, 0, 0), (0, 0, 114300, 0), (0, 0, 0, 114300)):
        out.append(Case("square", f"dist {dist}", (_a(3000, dist=dist),)))
    out.append(Case("square", "above its paragraph", (_a(3000, v=-400000, cy=1200000),), before=2))
    out.append(Case("square", "taller than its paragraph", (_a(3000, cy=3000000),), after=3,
                    text="Case short anchor paragraph. " + WORDS))
    out.append(Case("square", "against the margin and the page", (_a(3000, relative="margin", v=2000000,
                                                                        vertical="page"),), before=2))
    out.append(Case("square", "a segment too narrow for a word", (_a(700),),
                    text="Anextraordinarilylongwordthatcannotfit " * 20))
    out.append(Case("square", "wider than the column", (Anchor(("column", "offset", -200000),
                                                               ("paragraph", "offset", 200000), 6200000, 900000,
                                                               wrap=square()),)))
    short = "ab " * 400
    for g in range(340, 401, 10):
        out.append(Case("min", f"left gap {g}", (_a(g, v=0),), text=short, before=0))
    for g in range(340, 401, 10):
        out.append(Case("min", f"right gap {g}", (_a(COLUMN - g - WIDTH, v=0),), text=short, before=0))
    for size in (16, 40):
        for g in (330, 360, 390, 420):
            out.append(Case("min", f"left gap {g}, {size // 2} pt", (_a(g, v=0),),
                            runs=wml.run(short, sz=size, szCs=size), before=0))
    for rule, line, tops in (("auto", 259, range(264, 278, 2)), ("auto", 240, range(264, 274, 2)),
                             ("exact", 320, range(314, 326, 2)), ("atLeast", 400, range(394, 406, 2))):
        spacing = {"before": 0, "after": 120, "line": line, "lineRule": rule}
        for top in tops:
            out.append(Case("vertical", f"{rule} {line}, top {top}", (_a(3000, v=top * T, cy=300000),), before=0,
                            props={"spacing": spacing}))
    for height in range(284, 296, 2):
        out.append(Case("vertical", f"auto 259, bottom {height}", (_a(3000, v=0, cy=height * T),), before=0))
    for kind in ("wrapTight", "wrapThrough"):
        for name, points in (("triangle", TRIANGLE), ("diamond", DIAMOND), ("C", C_SHAPE), ("outside", OUTSIDE)):
            for dist in ((0, 0, 0, 0), (0, 0, 114300, 114300)):
                out.append(Case("polygon", f"{kind} {name}, dist {dist[2]}",
                                (_a(3000, v=100000, cy=2000000, wrap=polygon(kind, "bothSides", points), dist=dist),),
                                before=0))
    out.append(Case("polygon", "wrapSquare with no polygon", (_a(3000, v=100000, cy=2000000),), before=0))
    for cy in (900000, 1200000, 700000):
        out.append(Case("more", f"nothing fits beside it, height {cy}",
                        (_a(450, v=0, cx=(COLUMN - 900) * T, cy=cy),), text=long_words, before=0))
    out.append(Case("more", "nothing fits beside it, below the first line",
                    (_a(450, v=300000, cx=(COLUMN - 900) * T),), text=long_words, before=0))
    for g in (100, 400):
        out.append(Case("more", f"wrap left, gap {g}", (_a(g, v=0, wrap=square("left")),), before=0))
        out.append(Case("more", f"wrap right, gap {g}", (_a(COLUMN - g - WIDTH, v=0, wrap=square("right")),),
                        before=0))
    out.append(Case("more", "first line beside, drawing wider than the indent", (_a(0, v=0),), before=0,
                    props={"ind": {"left": 720, "firstLine": 360}}))
    out.append(Case("more", "first line beside, drawing inside the indent", (_a(0, v=0, cx=200000),), before=0,
                    props={"ind": {"left": 720, "firstLine": 360}}))
    for hanging in (180, 360, 540):
        out.append(Case("more", f"hanging {hanging} beside a drawing inside the indent", (_a(0, v=0, cx=200000),),
                        before=0, props={"ind": {"left": 720, "hanging": hanging}}))
    out.append(Case("more", "tabs right of it", (_a(0, v=0, cx=1500000),), before=0,
                    runs=(wml.run("Start") + wml.run("\t") + wml.run("after a tab ")) * 30))
    out.append(Case("more", "effect extent", (_a(3000, v=0, effect=(63500, 0, 63500, 0)),), before=0))
    out.append(Case("more", "wrapTight, distT and distB", (_a(3000, v=300000, cy=600000, wrap=polygon(
        "wrapTight", "bothSides", RECTANGLE), dist=(228600, 228600, 0, 0)),), before=0))
    out.append(Case("more", "wrapSquare, distT and distB", (_a(3000, v=300000, cy=600000,
                                                               dist=(228600, 228600, 0, 0)),), before=0))
    out.append(Case("more", "empty paragraphs beside it", (_a(3000, v=0, cy=1500000),), text="", after=4,
                    before=0))
    out.append(Case("more", "a centred short line beside it", (_a(3000, v=0),), text="Short centred.", after=3,
                    before=0, props={"jc": "center"}))
    out.append(Case("more", "justified", (_a(3000, v=0),), before=0, props={"jc": "both"}))
    out.append(Case("more", "near the page's foot", (_a(3000, v=0, cy=2500000),), before=11))
    out.append(Case("more", "near the page's foot, a line later", (_a(3000, v=0, cy=2500000),), before=12))
    return tuple(out)


def _edge() -> tuple[Case, ...]:
    out = []
    right = (0, 0, 114300, 114300)
    for j in range(12):
        out.append(Case("right", f"offset +{j}", (_a(3149 + j, dist=right, cx=1400000, cy=600000),)))
    for j in range(12):
        out.append(Case("right", f"extent +{j}", (_a(3149, cx=1400000 + j * T, cy=600000, dist=right),)))
    for j in range(12):
        out.append(Case("right", f"distR +{j}", (_a(3149, cx=1400000, cy=600000,
                                                   dist=(0, 0, 114300, 114300 + j * T)),)))
    for j in range(12):
        out.append(Case("right", f"page offset +{j}", (_a(4566 + j, relative="page", cx=1400000, cy=600000,
                                                          dist=right),)))
    for word, width_units in WIDTHS.items():
        for k in (3, 5, 7):
            end = k * (width_units + SPACE) - SPACE
            twips = math.floor(end / 204.8)
            for d in range(-2, 3):
                out.append(Case("left", f"{word} x {k}, edge {twips + d}", (_a(twips + d, v=0, cy=1500000),),
                                text=(word + " ") * 100, before=0, after=0))
    for k in range(6):
        out.append(Case("tie", f"text edge on a half pixel, {2017 + 24 * k}", (_a(2017 + 24 * k, v=0),), before=0))
    for k in range(6):
        out.append(Case("tie", f"page edge on a half pixel, {2015 + 24 * k}", (_a(2015 + 24 * k, v=0),), before=0))
    # Each page the anchoring paragraph alone, three times the words where it states none.
    return tuple(dataclasses.replace(case, before=0, after=0, text=case.text if case.text is not None else
                                     "Anchor. " + WORDS * 3) for case in out)


SIDE = _side()
EDGE = _edge()


def body(cases) -> str:
    out = ""
    serial = 1
    for number, case in enumerate(cases):
        out += anchor_probe._p(f"Case {number} heading", pageBreakBefore=True)
        for _ in range(case.before):
            out += wml.paragraph(wml.run(f"Case {number} before. {WORDS}"), mark={}, spacing=SPACING)
        drawings = ""
        for anchor in case.anchors:
            drawings += anchor.xml(serial)
            serial += 1
        props = {"spacing": SPACING, **case.props}
        runs = case.runs if case.runs is not None else wml.run(
            case.text if case.text is not None else f"Case{number} anchor. " + WORDS * 6)
        out += wml.paragraph(drawings + runs, mark={}, **props)
        for _ in range(case.after):
            out += wml.paragraph(wml.run(f"Case {number} after. {WORDS}"), mark={}, spacing=SPACING)
    return out


def build(name: str) -> bytes:
    """``wrap-side-<setting>`` or ``wrap-edge-<setting>-<left margin>``."""
    parts = name.split("-")
    setting = parts[2]
    if parts[1] == "side":
        return anchor_probe.package(body(SIDE), setting)
    margins = dict(anchor_probe.MARGINS, left=int(parts[3]))
    return anchor_probe.package(body(EDGE), setting, final_section=anchor_probe.section(margins))


DOCUMENTS = tuple([f"wrap-side-{s}" for s in SETTINGS] + [f"wrap-edge-{s}-{m}" for s in SETTINGS for m in MARGINS])


def cases_of(name: str) -> tuple[Case, ...]:
    return SIDE if name.startswith("wrap-side") else EDGE


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for name in DOCUMENTS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"{name}.docx"
        path.write_bytes(build(name))
        print(path)

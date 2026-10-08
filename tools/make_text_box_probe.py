#!/usr/bin/env python3
"""Where a text box's text starts: insets, the outline, the geometry, the anchor, the size.

``make_drawing_probe.py`` put 144 swept text boxes' first glyphs where Word drew them --
boxes with no outline and no style.  Real documents' text boxes state their outline as
``a:ln`` with ``a:noFill`` under a style whose line reference names a theme line of some
width, some sit in a ``roundRect``, and there their text stood a couple of pixels off
Word's.  Every box here holds a line of text labelled ``B<n>`` (a second paragraph,
where there is one, ``C<n>``); the reader compares the first glyph of each with the
model's.  Families (``CASES``):

* ``line`` -- the outline: none, drawn at 0.5 to 8 pt, stated with ``a:noFill`` at 4 and
  8 pt, the style's line reference 1, 2 and 3 under ``a:noFill`` (the documents' form)
  and drawn, reference 0; each at ``anchor`` ``t``, ``ctr`` and ``b``; and drawn with no
  width stated, alone and under a style;
* ``inset`` -- zero, the default and others, with and without the style's line; centred
  and right-aligned text;
* ``geometry`` -- a ``roundRect`` at its default and other adjustments, an ellipse, a
  triangle, a right triangle, an octagon, a custom geometry with a text rectangle of its
  own; anchors ``t`` and ``ctr``;
* ``wrap`` -- ``wrap="none"``, left and centred;
* ``size`` -- 8 to 28 pt, one and two lines, ``auto`` 276, ``exact``, space before;
  anchors ``t`` and ``b``;
* ``autofit`` -- ``spAutoFit`` and ``normAutofit`` (recorded);
* ``rotated`` -- ``bodyPr`` ``rot``, ``vert`` / ``vert270``, the shape turned 90 degrees
  (recorded: the model does not lay these out).

Boxes are positioned against the margin at offsets off the pixel grid, six to a page.
Two documents: no ``settings.xml`` and mode 15.  Reader: ``read_text_box_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import make_anchor_probe as anchor_probe
import make_dml_probe as dml_probe
import make_drawing_probe as drawing_probe
import wml
from make_anchor_probe import Anchor
from make_drawing_probe import NO_LINE, graphic, preset, solid, srgb, xfrm

SETTINGS = {"none": None, "15": 15}
PER_PAGE = 6
DEFAULT_INSETS = (91440, 45720, 91440, 45720)


@dataclass(frozen=True)
class Box:
    family: str
    note: str
    geometry: str = preset("rect")
    fill: str = "<a:noFill/>"
    line: str = NO_LINE
    style: str = ""
    insets: tuple = DEFAULT_INSETS
    anchor: str = "t"
    wrap: str = "square"
    fit: str = "<a:noAutofit/>"
    body_attrs: str = ""
    cx: int = 2000000
    cy: int = 700000
    rot: int = 0
    size: int = 22
    jc: str | None = None
    second: bool = False
    spacing: tuple = (0, 0, 240, "auto")
    words: str = "Sample text"

    @property
    def recorded_only(self) -> bool:
        return self.family in ("rotated", "autofit")


def style(ln: int) -> str:
    colour = '<a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr>'
    return (f'<wps:style><a:lnRef idx="{ln}">{colour}</a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/>'
            '</a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
            '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></wps:style>')


def drawn(width: int) -> str:
    return drawing_probe.outline(width, srgb("C00000"))


def stated_none(width: int) -> str:
    return f'<a:ln w="{width}"><a:noFill/></a:ln>'


def _cases() -> tuple[Box, ...]:
    out: list[Box] = []
    lines = (
        ("none", NO_LINE, ""),
        ("drawn 6350", drawn(6350), ""),
        ("drawn 12700", drawn(12700), ""),
        ("drawn 25400", drawn(25400), ""),
        ("drawn 50800", drawn(50800), ""),
        ("drawn 101600", drawn(101600), ""),
        ("noFill w 50800", stated_none(50800), ""),
        ("noFill w 101600", stated_none(101600), ""),
        ("lnRef 1 noFill", NO_LINE, style(1)),
        ("lnRef 2 noFill", NO_LINE, style(2)),
        ("lnRef 3 noFill", NO_LINE, style(3)),
        ("lnRef 2 drawn", "", style(2)),
        ("lnRef 3 drawn", "", style(3)),
        ("lnRef 0", "", style(0)),
        ("lnRef 2 own w 50800 noFill", stated_none(50800), style(2)),
    )
    for anchor in ("t", "ctr", "b"):
        for note, line, st in lines:
            out.append(Box("line", f"{note} {anchor}", line=line, style=st, anchor=anchor))
    # An outline that states no width: drawn (Word's default width) and under a style.
    no_width = f"<a:ln>{solid(srgb('C00000'))}</a:ln>"
    out.append(Box("line", "drawn no w t", line=no_width))
    out.append(Box("line", "drawn no w lnRef 3 t", line=no_width, style=style(3)))
    for insets in ((0, 0, 0, 0), (137160, 45720, 91440, 45720), (91440, 91440, 91440, 45720), (45720, 0, 0, 0)):
        for note, line, st in (("none", NO_LINE, ""), ("lnRef 2 noFill", NO_LINE, style(2)),
                               ("drawn 50800", drawn(50800), "")):
            out.append(Box("inset", f"{insets} {note}", insets=insets, line=line, style=st))
    for jc in ("center", "right"):
        for note, line, st in (("none", NO_LINE, ""), ("lnRef 2 noFill", NO_LINE, style(2)),
                               ("drawn 50800", drawn(50800), "")):
            out.append(Box("inset", f"jc {jc} {note}", jc=jc, line=line, style=st))
    geometries = (
        ("roundRect", preset("roundRect")),
        ("roundRect 7705", preset("roundRect", {"adj": 7705})),
        ("roundRect 30000", preset("roundRect", {"adj": 30000})),
        ("roundRect 50000", preset("roundRect", {"adj": 50000})),
        ("ellipse", preset("ellipse")),
        ("triangle", preset("triangle")),
        ("rtTriangle", preset("rtTriangle")),
        ("octagon", preset("octagon")),
        ("custom rect", '<a:custGeom><a:avLst/><a:gdLst><a:gd name="q" fmla="*/ w 1 4"/>'
                        '<a:gd name="v" fmla="*/ h 1 4"/></a:gdLst><a:ahLst/><a:cxnLst/>'
                        '<a:rect l="q" t="v" r="r" b="b"/><a:pathLst><a:path w="1000" h="1000"><a:moveTo>'
                        '<a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="1000" y="0"/></a:lnTo><a:lnTo>'
                        '<a:pt x="1000" y="1000"/></a:lnTo><a:lnTo><a:pt x="0" y="1000"/></a:lnTo><a:close/>'
                        '</a:path></a:pathLst></a:custGeom>'),
    )
    for anchor in ("t", "ctr"):
        for note, geometry in geometries:
            out.append(Box("geometry", f"{note} {anchor}", geometry=geometry, fill=solid(srgb("DEEBF7")),
                           anchor=anchor))
    out.append(Box("geometry", "roundRect 7705 lnRef 2 noFill ctr", geometry=preset("roundRect", {"adj": 7705}),
                   fill=solid(srgb("DEEBF7")), style=style(2), anchor="ctr", insets=(137160, 45720, 91440, 45720)))
    out.append(Box("geometry", "ellipse lnRef 2 noFill ctr", geometry=preset("ellipse"),
                   fill=solid(srgb("DEEBF7")), style=style(2), anchor="ctr"))
    for jc in (None, "center"):
        for note, line, st in (("none", NO_LINE, ""), ("lnRef 2 noFill", NO_LINE, style(2))):
            out.append(Box("wrap", f"none {jc or 'left'} {note}", wrap="none", jc=jc, line=line, style=st))
    for size in (16, 22, 28, 36, 48, 56):
        for anchor in ("t", "b"):
            out.append(Box("size", f"{size / 2:g} pt {anchor}", size=size, anchor=anchor, cy=900000))
    for note, spacing in (("auto 276", (0, 0, 276, "auto")), ("exact 400", (0, 0, 400, "exact")),
                          ("before 120 after 240", (120, 240, 240, "auto"))):
        for anchor in ("t", "b"):
            out.append(Box("size", f"{note} two {anchor}", spacing=spacing, second=True, anchor=anchor, cy=900000))
    out.append(Box("size", "two lnRef 2 noFill b", second=True, anchor="b", style=style(2), cy=900000))
    out.append(Box("autofit", "spAutoFit", fit="<a:spAutoFit/>", second=True, cy=300000))
    out.append(Box("autofit", "normAutofit", fit='<a:normAutofit fontScale="70000"/>', second=True, cy=300000))
    out.append(Box("rotated", "bodyPr rot 5400000", body_attrs=' rot="5400000"', cy=1600000))
    out.append(Box("rotated", "vert", body_attrs=' vert="vert"', cy=1600000))
    out.append(Box("rotated", "vert270", body_attrs=' vert="vert270"', cy=1600000))
    out.append(Box("rotated", "shape rot 90", rot=5400000))
    return tuple(out)


CASES = _cases()


def label(number: int) -> str:
    return f"B{number}"


def _paragraph(box: Box, text: str) -> str:
    before, after, line, rule = box.spacing
    props = {"spacing": {"before": before, "after": after, "line": line, "lineRule": rule}}
    if box.jc:
        props["jc"] = box.jc
    run = {"sz": box.size, "szCs": box.size}
    return wml.paragraph(wml.run(text, **run), mark=run, **props)


def shape(box: Box, number: int) -> str:
    text = _paragraph(box, f"{label(number)} {box.words}")
    if box.second:
        text += _paragraph(box, f"C{number} second")
    l, t, r, b = box.insets
    body = (f'<wps:bodyPr rot="0" vert="horz" wrap="{box.wrap}" lIns="{l}" tIns="{t}" rIns="{r}" bIns="{b}"'
            f' anchor="{box.anchor}" anchorCtr="0"{box.body_attrs}>{box.fit}</wps:bodyPr>')
    body = body.replace(' rot="0" vert="horz"', "") if box.body_attrs else body
    if box.body_attrs and "rot=" not in box.body_attrs:
        body = body.replace("<wps:bodyPr ", '<wps:bodyPr rot="0" ')
    if box.body_attrs and "vert=" not in box.body_attrs:
        body = body.replace("<wps:bodyPr ", '<wps:bodyPr vert="horz" ')
    return (f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>{xfrm(0, 0, box.cx, box.cy, rot=box.rot)}'
            f"{box.geometry}{box.fill}{box.line}</wps:spPr>{box.style}"
            f"<wps:txbx><w:txbxContent>{text}</w:txbxContent></wps:txbx>{body}</wps:wsp>")


def pages() -> list[list[int]]:
    return [list(range(k, min(k + PER_PAGE, len(CASES)))) for k in range(0, len(CASES), PER_PAGE)]


def body() -> str:
    out = ""
    serial = 1
    for page, numbers in enumerate(pages()):
        out += anchor_probe._p(f"Page {page} of text boxes", pageBreakBefore=True)
        drawings = ""
        for slot, number in enumerate(numbers):
            box = CASES[number]
            # Two columns, three rows; each box off the pixel grid by its own amount.
            x = (slot % 2) * 2900000 + number * 1693
            y = 200000 + (slot // 2) * 2200000 + number * 3217
            drawings += Anchor(("margin", "offset", x), ("margin", "offset", y), box.cx, box.cy,
                               graphic=graphic(shape(box, number), "wps")).xml(serial)
            serial += 1
        out += wml.paragraph(wml.run(f"Page {page} anchors.") + drawings, mark={},
                             spacing={"before": 0, "after": 0, "line": 240, "lineRule": "auto"})
    return out


def build(setting: str) -> bytes:
    key = "15" if SETTINGS[setting] == 15 else "none"
    return anchor_probe.package(body(), key, extra=(dml_probe.theme(),))


if __name__ == "__main__":
    print(len(CASES), "boxes on", len(pages()), "pages")

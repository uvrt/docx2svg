#!/usr/bin/env python3
"""What a floating drawing shows: DrawingML shapes, text boxes and groups, as Word draws
them.

``make_anchor_probe.py`` measured where an anchor goes with pictures.  The local
templates' anchors are mostly not pictures: shapes (``wps:wsp``) with a preset or a
custom geometry, a solid fill in an RGB or a theme colour with luminance transforms, an
outline or none, text boxes (``wps:txbx``) with insets and a vertical anchor, and groups
(``wpg:wgp``) of them.  Every case is a page with a heading line, the anchoring paragraph
and a line after it; each drawing is ``wrapNone`` in front of the text, positioned
against the margin, so its box is known from the anchor rules.  Families (``CASES``):

* ``fill`` -- rectangles filled in ``srgbClr`` and in ``schemeClr`` with ``lumMod`` /
  ``lumOff`` (the theme's colour scheme, ``tx1`` and ``bg1`` through the default map);
* ``line`` -- outlines of several widths, the ``line`` preset, flipped;
* ``preset`` -- ``roundRect`` (its default and another adjustment), ``ellipse``,
  ``triangle``, ``rtTriangle`` and flipped;
* ``custom`` -- ``a:custGeom`` paths: lines, a cubic, an arc;
* ``text`` -- text boxes: the default insets and others, ``anchor`` ``t`` / ``ctr`` /
  ``b``, two paragraphs, a centred one, a larger face, a filled box, text taller than
  its box, ``spAutoFit``;
* ``group`` -- groups of rectangles, unscaled and scaled (``a:chExt`` twice the extent),
  offset children, a nested group, a flipped child, a picture;
* ``other`` -- a rotated rectangle and a gradient fill, which are not drawn to the
  pixel.

Half the drawings are wrapped in ``mc:AlternateContent`` as Word writes them, half are
bare ``w:drawing``.  Two documents: no ``settings.xml`` and mode 15.  Reader:
``read_anchor_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import make_anchor_probe as anchor_probe
import wml
from make_anchor_probe import A_NS, PIC_NS, Anchor

SETTINGS = {"none": None, "15": 15}
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def srgb(value: str) -> str:
    return f'<a:srgbClr val="{value}"/>'


def scheme(value: str, *transforms: tuple[str, int]) -> str:
    inner = "".join(f'<a:{name} val="{amount}"/>' for name, amount in transforms)
    return f'<a:schemeClr val="{value}">{inner}</a:schemeClr>'


def solid(colour: str) -> str:
    return f"<a:solidFill>{colour}</a:solidFill>"


NO_LINE = "<a:ln><a:noFill/></a:ln>"


def outline(width: int, colour: str = srgb("000000")) -> str:
    return f'<a:ln w="{width}">{solid(colour)}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'


def preset(name: str, adjust: dict | None = None) -> str:
    guides = "".join(f'<a:gd name="{k}" fmla="val {v}"/>' for k, v in (adjust or {}).items())
    return f'<a:prstGeom prst="{name}"><a:avLst>{guides}</a:avLst></a:prstGeom>'


def custom(paths: str) -> str:
    return ('<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>'
            f"<a:pathLst>{paths}</a:pathLst></a:custGeom>")


def xfrm(x: int, y: int, cx: int, cy: int, *, rot: int = 0, flip_h: bool = False, flip_v: bool = False,
         child: tuple[int, int, int, int] | None = None) -> str:
    attrs = (f' rot="{rot}"' if rot else "") + (' flipH="1"' if flip_h else "") + (' flipV="1"' if flip_v else "")
    inner = f'<a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/>'
    if child is not None:
        inner += f'<a:chOff x="{child[0]}" y="{child[1]}"/><a:chExt cx="{child[2]}" cy="{child[3]}"/>'
    return f"<a:xfrm{attrs}>{inner}</a:xfrm>"


def body_pr(insets: tuple[int, int, int, int] = (91440, 45720, 91440, 45720), anchor: str = "t",
            fit: str = "<a:noAutofit/>") -> str:
    l, t, r, b = insets
    return (f'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="{l}" tIns="{t}" rIns="{r}" bIns="{b}"'
            f' anchor="{anchor}" anchorCtr="0">{fit}</wps:bodyPr>')


def text_paragraph(text: str, *, jc: str | None = None, size: int = 22, bold: bool = False,
                   before: int = 0, after: int = 0) -> str:
    props = {"spacing": {"before": before, "after": after, "line": 240, "lineRule": "auto"}}
    if jc:
        props["jc"] = jc
    run_props = {"sz": size, "szCs": size}
    if bold:
        run_props["b"] = True
    return wml.paragraph(wml.run(text, **run_props), mark={}, **props)


def shape(geometry: str, fill: str, line: str, *, x: int = 0, y: int = 0, cx: int, cy: int, rot: int = 0,
          flip_h: bool = False, flip_v: bool = False, text: str | None = None, body: str | None = None) -> str:
    """One ``wps:wsp`` (its ``a:xfrm`` in its parent's coordinates)."""
    box = f"<wps:txbx><w:txbxContent>{text}</w:txbxContent></wps:txbx>" if text is not None else ""
    marker = ' txBox="1"' if text is not None else ""
    return (f"<wps:wsp><wps:cNvSpPr{marker}/><wps:spPr>"
            + xfrm(x, y, cx, cy, rot=rot, flip_h=flip_h, flip_v=flip_v) + geometry + fill + line
            + f"</wps:spPr>{box}{body or body_pr()}</wps:wsp>")


def picture(x: int, y: int, cx: int, cy: int, image: int = 1) -> str:
    return (f'<pic:pic xmlns:pic="{PIC_NS}"><pic:nvPicPr><pic:cNvPr id="0" name="Picture"/><pic:cNvPicPr/>'
            f'</pic:nvPicPr><pic:blipFill><a:blip r:embed="rId{image + 2}"/><a:stretch><a:fillRect/></a:stretch>'
            f"</pic:blipFill><pic:spPr>{xfrm(x, y, cx, cy)}{preset('rect')}</pic:spPr></pic:pic>")


def group(cx: int, cy: int, children: str, *, child: tuple[int, int, int, int] | None = None,
          inner: bool = False, x: int = 0, y: int = 0) -> str:
    frame = xfrm(x, y, cx, cy, child=child or (0, 0, cx, cy))
    tag = "grpSp" if inner else "wgp"
    return f"<wpg:{tag}><wpg:cNvGrpSpPr/><wpg:grpSpPr>{frame}</wpg:grpSpPr>{children}</wpg:{tag}>"


def graphic(content: str, kind: str) -> str:
    uri = WPS_NS if kind == "wps" else WPG_NS
    return (f'<a:graphic xmlns:a="{A_NS}"><a:graphicData uri="{uri}" xmlns:wps="{WPS_NS}" xmlns:wpg="{WPG_NS}"'
            f' xmlns:w="{W_NS}">{content}</a:graphicData></a:graphic>')


@dataclass(frozen=True)
class Case:
    family: str
    content: str
    kind: str  # wps, wpg
    cx: int = 1800000
    cy: int = 900000
    note: str = ""
    behind: bool = False
    #: More drawings on the page: ``(content, cx, cy, x offset, y offset)``, EMU from the
    #: margin (the sweep of text boxes).
    more: tuple = ()


def _cases() -> tuple[Case, ...]:
    out: list[Case] = []
    rect = preset("rect")

    def wsp(geometry, fill, line, cx=1800000, cy=900000, note="", **kw):
        out.append(Case(out_family[0], shape(geometry, fill, line, cx=cx, cy=cy, **kw), "wps", cx, cy, note))

    out_family = ["fill"]
    for k, colour in enumerate((srgb("C6FDF8"), srgb("1F3864"), scheme("accent1"), scheme("accent1", ("lumMod", 75000)),
                                scheme("accent2", ("lumMod", 60000), ("lumOff", 40000)),
                                scheme("tx1", ("lumMod", 50000), ("lumOff", 50000)), scheme("bg1", ("lumMod", 85000)),
                                scheme("accent6", ("lumMod", 20000), ("lumOff", 80000)))):
        wsp(rect, solid(colour), NO_LINE, cx=1800000 + k * 3701, cy=900000 + k * 37003, note=f"fill {k}")
    out_family[0] = "line"
    for k, width in enumerate((6350, 12700, 28575, 63500)):
        wsp(rect, "<a:noFill/>", outline(width, srgb("C00000")), cx=1500000 + k * 3701, cy=700000, note=f"outline {width}")
    wsp(preset("line"), "", outline(63500, srgb("2E75B6")), cx=3000000, cy=0, note="line")
    wsp(preset("line"), "", outline(19050, srgb("2E75B6")), cx=1500000, cy=800000, flip_v=True, note="line flipV")
    out_family[0] = "preset"
    for name, adjust, flip in (("roundRect", None, False), ("roundRect", {"adj": 30000}, False), ("ellipse", None, False),
                               ("triangle", None, False), ("rtTriangle", None, False), ("rtTriangle", None, True)):
        wsp(preset(name, adjust), solid(srgb("70AD47")), NO_LINE, flip_h=flip, note=name + (" flipH" if flip else ""))
    out_family[0] = "custom"
    lines_path = ('<a:path w="1000" h="500"><a:moveTo><a:pt x="0" y="500"/></a:moveTo><a:lnTo><a:pt x="300" y="0"/>'
                  '</a:lnTo><a:lnTo><a:pt x="1000" y="120"/></a:lnTo><a:lnTo><a:pt x="800" y="500"/></a:lnTo>'
                  "<a:close/></a:path>")
    cubic = ('<a:path w="1000" h="1000"><a:moveTo><a:pt x="0" y="1000"/></a:moveTo><a:cubicBezTo><a:pt x="0" y="0"/>'
             '<a:pt x="1000" y="0"/><a:pt x="1000" y="1000"/></a:cubicBezTo><a:close/></a:path>')
    arc = ('<a:path w="1000" h="1000"><a:moveTo><a:pt x="500" y="0"/></a:moveTo>'
           '<a:arcTo wR="500" hR="500" stAng="16200000" swAng="10800000"/><a:lnTo><a:pt x="500" y="0"/></a:lnTo>'
           "<a:close/></a:path>")
    for paths, note in ((lines_path, "lines"), (cubic, "cubic"), (arc, "arc")):
        wsp(custom(paths), solid(srgb("ED7D31")), NO_LINE, note=note)
    out_family[0] = "text"
    words = "Text in a box, laid out at the box's inner width, which wraps onto a second line here."
    for k, (insets, anchor, paras, fill, fit, note) in enumerate((
            ((91440, 45720, 91440, 45720), "t", text_paragraph("One line."), "<a:noFill/>", None, "default insets"),
            ((91440, 45720, 91440, 45720), "ctr", text_paragraph("Centred box."), "<a:noFill/>", None, "anchor ctr"),
            ((91440, 45720, 91440, 45720), "b", text_paragraph("Bottom box."), "<a:noFill/>", None, "anchor b"),
            ((137160, 0, 91440, 0), "ctr", text_paragraph("Insets 137160 0."), "<a:noFill/>", None, "insets"),
            ((0, 0, 0, 0), "t", text_paragraph(words), "<a:noFill/>", None, "no insets, wrapped"),
            ((91440, 45720, 91440, 45720), "t", text_paragraph("First paragraph.")
             + text_paragraph("Second, centred.", jc="center"), "<a:noFill/>", None, "two paragraphs"),
            ((91440, 91440, 91440, 45720), "t", text_paragraph("Bold 16 pt title", size=32, bold=True),
             solid(srgb("FFF2CC")), None, "filled, larger"),
            ((91440, 45720, 91440, 45720), "ctr", text_paragraph(words) + text_paragraph(words),
             "<a:noFill/>", None, "taller than its box"),
            ((91440, 45720, 91440, 45720), "t", text_paragraph("Auto-fit box."), "<a:noFill/>",
             "<a:spAutoFit/>", "spAutoFit"),
            ((91440, 45720, 91440, 45720), "t", text_paragraph("Spaced.", before=120, after=240)
             + text_paragraph("After."), "<a:noFill/>", None, "spacing"),
            ((91440, 45720, 91440, 45720), "ctr", text_paragraph("Right.", jc="right"), solid(srgb("DEEBF7")), None,
             "right-aligned"))):
        body = body_pr(insets, anchor, fit or "<a:noAutofit/>")
        cx, cy = 2000000 + k * 3701, 500000 + k * 37003
        out.append(Case("text", shape(preset("rect"), fill, NO_LINE, cx=cx, cy=cy, text=paras, body=body), "wps",
                        cx, cy, note))
    out_family[0] = "group"
    blue, red = solid(srgb("4472C4")), solid(srgb("C00000"))
    out.append(Case("group", group(2000000, 1000000, shape(preset("rect"), blue, NO_LINE, x=0, y=0, cx=1000000,
                                                           cy=500000)
                                   + shape(preset("rect"), red, NO_LINE, x=1000000, y=500000, cx=1000000, cy=500000)),
                    "wpg", 2000000, 1000000, "unscaled"))
    out.append(Case("group", group(2000000, 1000000, shape(preset("rect"), blue, NO_LINE, x=0, y=0, cx=2000000,
                                                           cy=1000000)
                                   + shape(preset("rect"), red, NO_LINE, x=2000000, y=1000000, cx=2000000,
                                           cy=1000000), child=(0, 0, 4000000, 2000000)),
                    "wpg", 2000000, 1000000, "scaled by a half"))
    out.append(Case("group", group(2000003, 1000007, shape(preset("rect"), blue, NO_LINE, x=500000, y=300000,
                                                           cx=700001, cy=300003)
                                   + shape(preset("rect"), red, NO_LINE, x=1300000, y=700000, cx=900000,
                                           cy=400000), child=(500000, 300000, 1700000, 800000)),
                    "wpg", 2000003, 1000007, "offset and scaled"))
    nested = group(1000000, 500000, shape(preset("rect"), red, NO_LINE, x=0, y=0, cx=500000, cy=250000)
                   + shape(preset("rect"), blue, NO_LINE, x=500000, y=250000, cx=500000, cy=250000),
                   child=(0, 0, 1000000, 500000), inner=True, x=1000000, y=500000)
    out.append(Case("group", group(2000000, 1000000, shape(preset("rect"), solid(srgb("A9D18E")), NO_LINE, cx=1000000,
                                                           cy=500000) + nested), "wpg", 2000000, 1000000, "nested"))
    out.append(Case("group", group(2000000, 1000000, shape(preset("rtTriangle"), blue, NO_LINE, cx=1000000,
                                                           cy=1000000, flip_h=True)
                                   + shape(custom(lines_path), red, NO_LINE, x=1000000, y=0, cx=1000000,
                                           cy=1000000, flip_v=True)), "wpg", 2000000, 1000000, "flipped children"))
    out.append(Case("group", group(2000000, 1000000, picture(0, 0, 1000000, 1000000, 1)
                                   + shape(preset("rect"), blue, NO_LINE, x=1000000, y=0, cx=1000000, cy=1000000)),
                    "wpg", 2000000, 1000000, "picture in a group"))
    out.append(Case("group", group(2000000, 1000000, shape(
        preset("rect"), solid(srgb("FFF2CC")), NO_LINE, x=0, y=0, cx=2000000, cy=1000000,
        text=text_paragraph("Text in a group."), body=body_pr())), "wpg", 2000000, 1000000, "text box in a group"))
    out.append(Case("other", shape(preset("rect"), solid(srgb("7030A0")), NO_LINE, cx=1800000, cy=900000,
                                   rot=1800000), "wps", note="rotated 30 degrees"))
    gradient = ('<a:gradFill><a:gsLst><a:gs pos="0">' + srgb("FFFFFF") + '</a:gs><a:gs pos="100000">'
                + srgb("4472C4") + '</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>')
    out.append(Case("other", shape(preset("rect"), gradient, NO_LINE, cx=1800000, cy=900000), "wps",
                    note="gradient"))
    out.append(Case("other", shape(preset("rect"), solid(srgb("C6FDF8")), NO_LINE, cx=1800000, cy=900000),
                    "wps", note="behind the text", behind=True))
    # Text boxes swept down by single twips (0.21 px), eight to a page: where a box's
    # text starts, and how its first baseline rounds.
    variants = ((22, False, "t", 45720), (32, True, "t", 45720), (22, False, "t", 91440), (32, True, "t", 91440),
                (22, False, "ctr", 45720), (32, True, "ctr", 45720))
    for size, bold, anchor, top in variants:
        for page in range(3):
            boxes = []
            for k in range(8):
                step = page * 8 + k
                content = shape(preset("rect"), "<a:noFill/>", NO_LINE, cx=1100000, cy=600000,
                                text=text_paragraph(f"Sweep {step}", size=size, bold=bold),
                                body=body_pr((91440, top, 91440, 45720), anchor))
                boxes.append((content, 1100000, 600000, (k % 4) * 1250000 + step * 127,
                              300000 + (k // 4) * 1000000 + step * 635))
            out.append(Case("sweep", "", "wps", note=f"{size} {'bold ' if bold else ''}{anchor} {top}",
                            more=tuple(boxes)))
    return tuple(out)


CASES = _cases()


def drawing_xml(case: Case, number: int, wrapped: bool) -> str:
    """The case's anchor, as a bare ``w:drawing`` or inside ``mc:AlternateContent``."""
    run = Anchor(("margin", "offset", 700000 + number * 1001), ("margin", "offset", 400000 + number * 10007),
                 case.cx, case.cy, behind=case.behind,
                 graphic=graphic(case.content, case.kind)).xml(number + 1)
    if not wrapped:
        return run
    start = run.index("<w:drawing")
    end = run.index("</w:drawing>") + len("</w:drawing>")
    requires = "wps" if case.kind == "wps" else "wpg"
    return (run[:start] + f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:wps="{WPS_NS}" xmlns:wpg="{WPG_NS}">'
            f'<mc:Choice Requires="{requires}">' + run[start:end] + "</mc:Choice></mc:AlternateContent>"
            + run[end:])


def body() -> str:
    out = ""
    serial = 1000
    for number, case in enumerate(CASES):
        out += anchor_probe._p(f"Case {number} {case.family} {case.note}", pageBreakBefore=True)
        drawings = drawing_xml(case, number, number % 2 == 0) if case.content else ""
        for content, cx, cy, x, y in case.more:
            drawings += Anchor(("margin", "offset", x), ("margin", "offset", y), cx, cy,
                               graphic=graphic(content, "wps")).xml(serial)
            serial += 1
        out += wml.paragraph(wml.run(f"Case{number} anchors a drawing here: ") + drawings
                             + wml.run("and the line goes on."), mark={},
                             spacing={"before": 0, "after": 0, "line": 240, "lineRule": "auto"})
        out += anchor_probe._p(f"Case {number} after the anchor's paragraph.")
    return out


THEME = wml.theme_part({"latin": "Calibri Light"}, {"latin": "Calibri"})


def build(setting: str) -> bytes:
    key = "15" if SETTINGS[setting] == 15 else "none"
    return anchor_probe.package(body(), key, extra=(THEME,))


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"drawing-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

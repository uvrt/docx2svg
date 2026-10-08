#!/usr/bin/env python3
"""How Word paints DrawingML: colour transforms, gradients, patterns, dashes, caps, joins,
arrowheads, theme style references, group fills, effects and WordArt text.

``make_drawing_probe.py`` measured where a shape goes and its solid fill.  What a shape's
fill and outline *look like* beyond that is drawn by ``ooxml-common``'s DrawingML
renderers, which were measured against PowerPoint; this probe measures Word on each of
them, so that where Word differs the difference is a measured parameter, not a guess.
Every case is a page: a heading line, the anchoring paragraph, a line after it, and one
drawing (``wrapNone``, in front, against the margin) -- often a group, so that one page
holds a family's variations side by side in a known order.  Families (``CASES``):

* ``colour`` -- swatches of every colour transform Word may meet: ``tint``, ``shade``,
  ``satMod``, ``lumMod`` / ``lumOff`` in both orders, the theme's gradient stops'
  combinations, ``alpha``, ``sysClr``, ``prstClr``, ``hslClr``, ``scrgbClr``;
* ``gradient`` -- linear at several angles, ``scaled`` on and off, on a 2:1 box; three
  stops; ``path`` ``circle`` / ``rect`` / ``shape`` with centred and cornered
  ``fillToRect``; ``rotWithShape`` on a rotated shape; a flipped shape; an outline;
* ``pattern`` -- ``a:pattFill`` presets, and one off the 8-point lattice;
* ``dash`` -- every ``prstDash`` at each cap, a ``custDash``, two widths;
* ``join`` / ``arrow`` / ``compound`` -- joins on a thick outline, arrowheads of every
  type and size at two widths, ``cmpd`` lines;
* ``style`` -- shapes that state nothing and take their fill, line and effect from the
  theme's format scheme (``wps:style``), which here is Office's own;
* ``group`` -- ``a:grpFill`` under a solid and a gradient group fill;
* ``effect`` -- an outer and inner shadow, glow, soft edge;
* ``auto`` -- text boxes whose text's colour is automatic, on greys and colours, and
  styled shapes whose ``a:fontRef`` names a colour, on dark and light fills;
* inline -- a gradient rectangle and a dotted ellipse as ``wp:inline`` drawings;
* ``wordart`` -- text with ``w14:textFill`` / ``w14:textOutline``, in a WordArt text box
  and in the body.

One document, mode 15.  Reader: ``read_dml_probe.py``.
"""

from __future__ import annotations

from dataclasses import dataclass

import make_anchor_probe as anchor_probe
import make_drawing_probe as drawing_probe
import wml
from make_anchor_probe import A_NS, Anchor
from make_drawing_probe import NO_LINE, custom, graphic, group, preset, scheme, shape, solid, srgb

W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
SETTINGS = {"15": 15}


def colour(kind: str, value: str, *transforms: tuple[str, int]) -> str:
    inner = "".join(f'<a:{name} val="{amount}"/>' for name, amount in transforms)
    return f'<a:{kind} val="{value}">{inner}</a:{kind}>'


def gradient(stops, *, lin: tuple[int, int] | None = (0, 0), path: str | None = None,
             to_rect: tuple[int, int, int, int] = (50000, 50000, 50000, 50000), rot_with_shape: bool | None = None,
             flip: str | None = None) -> str:
    attrs = ""
    if flip:
        attrs += f' flip="{flip}"'
    if rot_with_shape is not None:
        attrs += f' rotWithShape="{int(rot_with_shape)}"'
    shade = ""
    if path:
        l, t, r, b = to_rect
        shade = f'<a:path path="{path}"><a:fillToRect l="{l}" t="{t}" r="{r}" b="{b}"/></a:path>'
    elif lin is not None:
        shade = f'<a:lin ang="{lin[0]}" scaled="{lin[1]}"/>'
    gs = "".join(f'<a:gs pos="{pos}">{c}</a:gs>' for pos, c in stops)
    return f"<a:gradFill{attrs}><a:gsLst>{gs}</a:gsLst>{shade}</a:gradFill>"


RED, BLUE, YELLOW = srgb("C00000"), srgb("0070C0"), srgb("FFD966")
TWO = ((0, RED), (100000, BLUE))


def line(width: int, fill: str, *, dash: str | None = "solid", cap: str | None = None, join: str = "<a:miter lim=\"800000\"/>",
         head: str = "", tail: str = "", cmpd: str | None = None, cust: str = "") -> str:
    attrs = f' w="{width}"'
    if cap:
        attrs += f' cap="{cap}"'
    if cmpd:
        attrs += f' cmpd="{cmpd}"'
    dash_xml = f'<a:prstDash val="{dash}"/>' if dash else ""
    if cust:
        dash_xml = f"<a:custDash>{cust}</a:custDash>"
    return f"<a:ln{attrs}>{fill}{dash_xml}{join}{head}{tail}</a:ln>"


def end(kind: str, which: str, w: str = "med", length: str = "med") -> str:
    return f'<a:{which}End type="{kind}" w="{w}" len="{length}"/>'


def wsp_style(ln: int, fill: int, effect: int, colour_xml: str) -> str:
    return (f"<wps:style><a:lnRef idx=\"{ln}\">{colour_xml.replace('X', '')}</a:lnRef>"
            f"<a:fillRef idx=\"{fill}\">{colour_xml}</a:fillRef><a:effectRef idx=\"{effect}\">{colour_xml}</a:effectRef>"
            f'<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>')


def styled_shape(geometry: str, style: str, *, x: int = 0, y: int = 0, cx: int, cy: int, text: str | None = None) -> str:
    """A ``wps:wsp`` whose ``spPr`` states only its geometry: everything else is ``style``'s."""
    box = f"<wps:txbx><w:txbxContent>{text}</w:txbxContent></wps:txbx>" if text is not None else ""
    return (f"<wps:wsp><wps:cNvSpPr/><wps:spPr>{drawing_probe.xfrm(x, y, cx, cy)}{geometry}</wps:spPr>{style}"
            f"{box}{drawing_probe.body_pr(anchor='ctr')}</wps:wsp>")


@dataclass(frozen=True)
class Case:
    family: str
    content: str
    kind: str = "wps"
    cx: int = 1800000
    cy: int = 900000
    note: str = ""
    #: Drawn inline (``wp:inline``, a character of its line) rather than floating.
    inline: bool = False


def swatches(colours, *, cx: int = 5400000, cy: int = 1600000, per_row: int = 6) -> tuple[str, int, int]:
    """A group of swatches, left to right and top to bottom, each a rectangle of one colour."""
    rows = (len(colours) + per_row - 1) // per_row
    w = cx // per_row
    h = cy // max(rows, 1)
    members = "".join(shape(preset("rect"), solid(c), NO_LINE, x=(k % per_row) * w, y=(k // per_row) * h,
                            cx=w - 20000, cy=h - 20000) for k, c in enumerate(colours))
    return group(cx, rows * h, members), cx, rows * h


def row(members, *, cx: int, cy: int) -> str:
    return group(cx, cy, "".join(members))


def _cases() -> tuple[Case, ...]:
    out: list[Case] = []
    rect = preset("rect")

    # -- colour ---------------------------------------------------------------------
    families = {
        "tint": [colour("srgbClr", "4472C4", ("tint", t)) for t in (10000, 25000, 40000, 50000, 60000, 75000, 90000)]
        + [colour("srgbClr", "C00000", ("tint", t)) for t in (20000, 50000, 80000)]
        + [colour("schemeClr", "accent2", ("tint", 40000)), colour("srgbClr", "808080", ("tint", 50000))],
        "shade": [colour("srgbClr", "4472C4", ("shade", t)) for t in (10000, 25000, 40000, 50000, 60000, 75000, 90000)]
        + [colour("srgbClr", "FFC000", ("shade", t)) for t in (20000, 50000, 80000)]
        + [colour("schemeClr", "accent6", ("shade", 50000)), colour("srgbClr", "808080", ("shade", 50000))],
        "sat": [colour("srgbClr", "4472C4", ("satMod", t)) for t in (25000, 50000, 150000, 200000, 300000)]
        + [colour("srgbClr", "ED7D31", ("satOff", 20000)), colour("srgbClr", "ED7D31", ("satOff", -20000)),
           colour("srgbClr", "70AD47", ("satMod", 175000)), colour("srgbClr", "808080", ("satMod", 200000)),
           colour("srgbClr", "4472C4", ("hueMod", 50000)), colour("srgbClr", "4472C4", ("hueOff", 3600000)),
           colour("srgbClr", "4472C4", ("lumOff", 20000))],
        "order": [colour("srgbClr", "4472C4", ("lumMod", 60000), ("lumOff", 40000)),
                  colour("srgbClr", "4472C4", ("lumOff", 40000), ("lumMod", 60000)),
                  colour("srgbClr", "ED7D31", ("tint", 50000), ("satMod", 300000)),
                  colour("srgbClr", "ED7D31", ("satMod", 300000), ("tint", 50000)),
                  colour("srgbClr", "4472C4", ("shade", 30000), ("satMod", 115000)),
                  colour("srgbClr", "4472C4", ("satMod", 115000), ("shade", 30000)),
                  # Office's theme gradient stops.
                  colour("srgbClr", "4472C4", ("lumMod", 110000), ("satMod", 105000), ("tint", 67000)),
                  colour("srgbClr", "4472C4", ("lumMod", 105000), ("satMod", 103000), ("tint", 73000)),
                  colour("srgbClr", "4472C4", ("lumMod", 105000), ("satMod", 109000), ("tint", 81000)),
                  colour("srgbClr", "4472C4", ("satMod", 103000), ("lumMod", 102000), ("tint", 94000)),
                  colour("srgbClr", "4472C4", ("satMod", 110000), ("lumMod", 100000), ("shade", 100000)),
                  colour("srgbClr", "4472C4", ("lumMod", 99000), ("satMod", 120000), ("shade", 78000))],
        "kinds": ['<a:sysClr val="windowText" lastClr="000000"/>', '<a:sysClr val="window" lastClr="FFFFFF"/>',
                  '<a:sysClr val="highlight" lastClr="0078D7"/>', '<a:prstClr val="red"/>',
                  '<a:prstClr val="darkSeaGreen"/>', '<a:hslClr hue="14400000" sat="100000" lum="50000"/>',
                  '<a:scrgbClr r="50000" g="20000" b="0"/>', colour("srgbClr", "4472C4", ("alpha", 50000)),
                  colour("schemeClr", "tx2"), colour("schemeClr", "bg2"),
                  colour("srgbClr", "4472C4", ("gray", 0)), colour("srgbClr", "4472C4", ("inv", 0)),
                  colour("srgbClr", "4472C4", ("comp", 0))],
    }
    for note, colours in families.items():
        content, cx, cy = swatches(colours)
        out.append(Case("colour", content, "wpg", cx, cy, note))

    # -- gradient -------------------------------------------------------------------
    three = ((0, RED), (30000, YELLOW), (100000, BLUE))
    for note, fill, geometry, extra in (
            ("lin 0", gradient(TWO, lin=(0, 0)), rect, {}),
            ("lin 90", gradient(TWO, lin=(5400000, 0)), rect, {}),
            ("lin 45 scaled 0", gradient(TWO, lin=(2700000, 0)), rect, {}),
            ("lin 45 scaled 1", gradient(TWO, lin=(2700000, 1)), rect, {}),
            ("lin 30 scaled 0", gradient(TWO, lin=(1800000, 0)), rect, {}),
            ("lin 30 scaled 1", gradient(TWO, lin=(1800000, 1)), rect, {}),
            ("lin 135 scaled 0", gradient(TWO, lin=(8100000, 0)), rect, {}),
            ("lin 300 scaled 0", gradient(TWO, lin=(18000000, 0)), rect, {}),
            ("three stops", gradient(three, lin=(0, 0)), rect, {}),
            ("unsorted stops", gradient(((100000, BLUE), (0, RED), (60000, YELLOW)), lin=(0, 0)), rect, {}),
            ("inset stops", gradient(((20000, RED), (70000, BLUE)), lin=(0, 0)), rect, {}),
            ("no lin", gradient(TWO, lin=None), rect, {}),
            ("circle centre", gradient(TWO, path="circle"), rect, {}),
            ("circle corner", gradient(TWO, path="circle", to_rect=(0, 0, 100000, 100000)), rect, {}),
            ("circle offset", gradient(TWO, path="circle", to_rect=(25000, 75000, 75000, 25000)), rect, {}),
            ("rect centre", gradient(TWO, path="rect"), rect, {}),
            ("rect corner", gradient(TWO, path="rect", to_rect=(100000, 100000, 0, 0)), rect, {}),
            ("shape centre", gradient(TWO, path="shape"), rect, {}),
            ("shape ellipse", gradient(TWO, path="shape"), preset("ellipse"), {}),
            ("circle on ellipse", gradient(TWO, path="circle"), preset("ellipse"), {}),
            ("lin 0 on ellipse", gradient(TWO, lin=(0, 0)), preset("ellipse"), {}),
            ("lin 0 on triangle", gradient(TWO, lin=(0, 0)), preset("triangle"), {}),
            ("alpha stop", gradient(((0, RED), (100000, colour("srgbClr", "0070C0", ("alpha", 20000)))),
                                    lin=(0, 0)), rect, {}),
            ("rotWithShape 1, rot 30", gradient(TWO, lin=(0, 0), rot_with_shape=True), rect, {"rot": 1800000}),
            ("rotWithShape 0, rot 30", gradient(TWO, lin=(0, 0), rot_with_shape=False), rect, {"rot": 1800000}),
            ("flipH lin 0", gradient(TWO, lin=(0, 0), rot_with_shape=True), rect, {"flip_h": True}),
            ("flipV lin 90", gradient(TWO, lin=(5400000, 0), rot_with_shape=True), rect, {"flip_v": True}),
            ("flipH rotWithShape 0", gradient(TWO, lin=(0, 0), rot_with_shape=False), rect, {"flip_h": True})):
        out.append(Case("gradient", shape(geometry, fill, NO_LINE, cx=1800000, cy=900000, **extra), note=note))
    out.append(Case("gradient", shape(rect, "<a:noFill/>", line(152400, gradient(TWO, lin=(0, 0))), cx=1800000,
                                      cy=900000), note="gradient outline"))

    # -- pattern --------------------------------------------------------------------
    patterns = ("pct5", "pct50", "pct90", "ltDnDiag", "dkHorz", "smGrid", "lgCheck", "cross", "dashDnDiag",
                "shingle", "weave", "sphere")
    members = [shape(rect, f'<a:pattFill prst="{p}"><a:fgClr>{RED}</a:fgClr><a:bgClr>{srgb("FFF2CC")}</a:bgClr>'
                           "</a:pattFill>", NO_LINE, x=(k % 4) * 1400000, y=(k // 4) * 900000, cx=1300000,
                     cy=800000) for k, p in enumerate(patterns)]
    out.append(Case("pattern", row(members, cx=5600000, cy=2700000), "wpg", 5600000, 2700000, "presets"))
    out.append(Case("pattern", shape(rect, f'<a:pattFill prst="smGrid"><a:fgClr>{BLUE}</a:fgClr><a:bgClr>'
                                           f'{srgb("FFFFFF")}</a:bgClr></a:pattFill>', NO_LINE, cx=1803701,
                                     cy=903701), note="off the lattice"))
    out.append(Case("pattern", shape(rect, f'<a:pattFill prst="dkVert"><a:fgClr>{BLUE}</a:fgClr><a:bgClr>'
                                           f'{srgb("FFFFFF")}</a:bgClr></a:pattFill>', NO_LINE, cx=1800000,
                                     cy=900000, rot=1800000), note="rotated"))

    # -- dash -----------------------------------------------------------------------
    dashes = ("solid", "dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot", "sysDash", "sysDot",
              "sysDashDot", "sysDashDotDot")
    for cap in (None, "rnd", "sq"):
        for width in (38100, 12700):
            members = [shape(preset("line"), "", line(width, solid(BLUE), dash=d, cap=cap), x=0, y=k * 250000,
                             cx=5000000, cy=0) for k, d in enumerate(dashes)]
            out.append(Case("dash", row(members, cx=5000000, cy=250000 * len(dashes)), "wpg", 5000000,
                            250000 * len(dashes), f"cap {cap or 'flat'} w {width}"))
    cust = '<a:ds d="300000" sp="100000"/><a:ds d="100000" sp="100000"/>'
    members = [shape(preset("line"), "", line(38100, solid(BLUE), cap=cap, cust=cust), x=0, y=k * 300000,
                     cx=5000000, cy=0) for k, cap in enumerate((None, "rnd", "sq"))]
    members += [shape(rect, "<a:noFill/>", line(38100, solid(RED), dash="dash"), x=0, y=1000000, cx=2000000,
                      cy=800000),
                shape(preset("ellipse"), "<a:noFill/>", line(38100, solid(RED), dash="sysDot", cap="rnd"),
                      x=2500000, y=1000000, cx=2000000, cy=800000)]
    out.append(Case("dash", row(members, cx=5000000, cy=1800000), "wpg", 5000000, 1800000, "custom, closed"))

    # -- join, arrow, compound ------------------------------------------------------
    joins = ('<a:miter lim="800000"/>', "<a:round/>", "<a:bevel/>", "")
    members = [shape(preset("triangle"), "<a:noFill/>", line(152400, solid(BLUE), join=j), x=k * 1300000,
                     y=100000, cx=1100000, cy=900000) for k, j in enumerate(joins)]
    out.append(Case("join", row(members, cx=5200000, cy=1100000), "wpg", 5200000, 1100000, "miter round bevel none"))
    for width in (12700, 38100):
        members = []
        types = ("triangle", "stealth", "diamond", "oval", "arrow")
        sizes = (("sm", "sm"), ("med", "med"), ("lg", "lg"), ("sm", "lg"), ("lg", "sm"))
        for k, kind in enumerate(types):
            for j, (w, length) in enumerate(sizes):
                members.append(shape(preset("line"), "", line(width, solid(BLUE), tail=end(kind, "tail", w, length),
                                                              head=end(kind, "head", w, length) if j == 1 else ""),
                                     x=j * 1100000, y=k * 450000, cx=800000, cy=0))
        out.append(Case("arrow", row(members, cx=5500000, cy=2250000), "wpg", 5500000, 2250000, f"w {width}"))
    members = [shape(preset("line"), "", line(76200, solid(BLUE), cmpd=c), x=0, y=k * 300000, cx=5000000, cy=0)
               for k, c in enumerate(("sng", "dbl", "thickThin", "thinThick", "tri"))]
    members.append(shape(rect, "<a:noFill/>", line(76200, solid(RED), cmpd="dbl"), x=0, y=1600000, cx=2000000,
                         cy=700000))
    out.append(Case("compound", row(members, cx=5000000, cy=2300000), "wpg", 5000000, 2300000, "cmpd"))

    # -- style ----------------------------------------------------------------------
    accent1, accent2 = scheme("accent1"), scheme("accent2")
    for note, style in (("Word's default shape", wsp_style(2, 1, 0, scheme("accent1").replace("/>", "/>"))),
                        ("fillRef 2", wsp_style(0, 2, 0, accent2)), ("fillRef 3", wsp_style(0, 3, 0, accent2)),
                        ("lnRef 1", wsp_style(1, 0, 0, accent1)), ("lnRef 3", wsp_style(3, 0, 0, accent1)),
                        ("effectRef 3", wsp_style(0, 1, 3, accent1)),
                        ("bgFillRef 1002", wsp_style(0, 1002, 0, accent1)),
                        ("lnRef shade", "<wps:style><a:lnRef idx=\"2\">" + scheme("accent1", ("shade", 50000))
                         + '</a:lnRef><a:fillRef idx="1">' + accent1 + '</a:fillRef><a:effectRef idx="0">' + accent1
                         + '</a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>')):
        out.append(Case("style", styled_shape(rect, style, cx=1800000, cy=900000), note=note))
    text = drawing_probe.text_paragraph("Style font colour")
    out.append(Case("style", styled_shape(rect, wsp_style(2, 1, 0, accent1), cx=1800000, cy=900000, text=text),
                    note="fontRef lt1 text"))
    out.append(Case("style", "<wps:wsp><wps:cNvCnPr/><wps:spPr>" + drawing_probe.xfrm(0, 0, 1800000, 0)
                    + preset("line") + "</wps:spPr><wps:style><a:lnRef idx=\"1\">" + accent1
                    + '</a:lnRef><a:fillRef idx="0">' + accent1 + '</a:fillRef><a:effectRef idx="0">' + accent1
                    + '</a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></wps:style>'
                    + drawing_probe.body_pr() + "</wps:wsp>", cy=0, note="connector lnRef 1"))

    # -- group fill -----------------------------------------------------------------
    def grp(fill: str, children: str, cx: int = 3600000, cy: int = 900000) -> str:
        frame = drawing_probe.xfrm(0, 0, cx, cy, child=(0, 0, cx, cy))
        return f"<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>{frame}{fill}</wpg:grpSpPr>{children}</wpg:wgp>"

    kids = (shape(rect, "<a:grpFill/>", NO_LINE, x=0, y=0, cx=1700000, cy=900000)
            + shape(preset("ellipse"), "<a:grpFill/>", NO_LINE, x=1900000, y=0, cx=1700000, cy=900000))
    out.append(Case("group", grp(solid(srgb("007A7A")), kids), "wpg", 3600000, 900000, "solid"))
    out.append(Case("group", grp(gradient(TWO, lin=(0, 0)), kids), "wpg", 3600000, 900000, "gradient"))
    out.append(Case("group", grp(f'<a:pattFill prst="ltDnDiag"><a:fgClr>{RED}</a:fgClr><a:bgClr>{srgb("FFFFFF")}'
                                 "</a:bgClr></a:pattFill>", kids), "wpg", 3600000, 900000, "pattern"))
    out.append(Case("group", grp("", kids), "wpg", 3600000, 900000, "none stated"))
    nested = ("<wpg:grpSp><wpg:cNvGrpSpPr/><wpg:grpSpPr>"
              + drawing_probe.xfrm(1900000, 0, 1700000, 900000, child=(0, 0, 1700000, 900000))
              + "</wpg:grpSpPr>" + shape(rect, "<a:grpFill/>", NO_LINE, cx=1700000, cy=900000) + "</wpg:grpSp>")
    out.append(Case("group", grp(solid(srgb("7030A0")), shape(rect, "<a:grpFill/>", NO_LINE, cx=1700000, cy=900000)
                                 + nested), "wpg", 3600000, 900000, "nested, inner states none"))
    glyph = ('<a:path w="1000" h="1000"><a:moveTo><a:pt x="0" y="1000"/></a:moveTo><a:lnTo><a:pt x="400" y="0"/>'
             '</a:lnTo><a:lnTo><a:pt x="600" y="0"/></a:lnTo><a:lnTo><a:pt x="1000" y="1000"/></a:lnTo>'
             '<a:lnTo><a:pt x="800" y="1000"/></a:lnTo><a:lnTo><a:pt x="500" y="250"/></a:lnTo>'
             '<a:lnTo><a:pt x="200" y="1000"/></a:lnTo><a:close/></a:path>')
    out.append(Case("group", grp(solid(srgb("003F5A")), shape(custom(glyph), "<a:grpFill/>", line(
        9525, "<a:noFill/>", cap="flat"), cx=900000, cy=900000) + shape(custom(glyph), "<a:grpFill/>", line(
            9525, "<a:noFill/>", cap="flat"), x=1000000, cx=900000, cy=900000)), "wpg", 3600000, 900000,
        "outlined letters"))

    # -- effect ---------------------------------------------------------------------
    for note, effect in (
            ("outer shadow", '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" '
                             'rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr>'
                             "</a:outerShdw></a:effectLst>"),
            ("hard shadow", '<a:effectLst><a:outerShdw dist="76200" dir="2700000" algn="tl" rotWithShape="0">'
                            + srgb("000000") + "</a:outerShdw></a:effectLst>"),
            ("inner shadow", '<a:effectLst><a:innerShdw blurRad="63500" dist="50800" dir="13500000">'
                             '<a:prstClr val="black"><a:alpha val="50000"/></a:prstClr></a:innerShdw></a:effectLst>'),
            ("glow", '<a:effectLst><a:glow rad="101600">' + colour("srgbClr", "FFC000", ("alpha", 60000))
                     + "</a:glow></a:effectLst>"),
            ("soft edge", '<a:effectLst><a:softEdge rad="63500"/></a:effectLst>')):
        out.append(Case("effect", shape(rect, solid(BLUE), NO_LINE, cx=1800000, cy=900000).replace(
            "</wps:spPr>", effect + "</wps:spPr>"), note=note))

    # -- automatic text colour -------------------------------------------------------
    # Text boxes whose text's colour is automatic, on fills from black to white and in
    # colours, and styled shapes whose font reference names a colour: which Word draws
    # the text in -- the fill's contrast, the reference's colour, or black.
    def text_row(fills, styles=None) -> str:
        members = []
        for k, fill in enumerate(fills):
            style = styles[k] if styles else ""
            text = drawing_probe.text_paragraph(f"Auto {k}")
            box = shape(rect, solid(fill), NO_LINE, x=k * 700000, cx=650000, cy=400000, text=text,
                        body=drawing_probe.body_pr(anchor="ctr"))
            if style:
                box = box.replace("<wps:txbx>", style + "<wps:txbx>")
            members.append(box)
        return row(members, cx=700000 * len(fills), cy=400000)

    greys = [srgb(f"{v:02X}{v:02X}{v:02X}") for v in (0x00, 0x30, 0x50, 0x60, 0x70, 0x80, 0x90, 0xA0)]
    out.append(Case("auto", text_row(greys), "wpg", 5600000, 400000, "greys 00-A0"))
    greys = [srgb(f"{v:02X}{v:02X}{v:02X}") for v in (0xA8, 0xB0, 0xB8, 0xC0, 0xC8, 0xD0, 0xE0, 0xFF)]
    out.append(Case("auto", text_row(greys), "wpg", 5600000, 400000, "greys A8-FF"))
    greys = [srgb(f"{v:02X}{v:02X}{v:02X}") for v in range(0x34, 0x54, 4)]
    out.append(Case("auto", text_row(greys), "wpg", 5600000, 400000, "greys 34-50"))
    colours = [srgb(v) for v in ("4472C4", "C00000", "70AD47", "FFC000", "0070C0", "7030A0", "ED7D31", "5B9BD5")]
    out.append(Case("auto", text_row(colours), "wpg", 5600000, 400000, "colours"))
    colours = [srgb(v) for v in ("D00000", "E00000", "A00000", "0000C0", "0000FF", "008000", "404000", "602060")]
    out.append(Case("auto", text_row(colours), "wpg", 5600000, 400000, "more colours"))

    def font_style(colour_xml: str) -> str:
        return (f'<wps:style><a:lnRef idx="0">{scheme("accent1")}</a:lnRef><a:fillRef idx="0">{scheme("accent1")}'
                f'</a:fillRef><a:effectRef idx="0">{scheme("accent1")}</a:effectRef><a:fontRef idx="minor">'
                f"{colour_xml}</a:fontRef></wps:style>")

    fills = [srgb("1F3864"), srgb("FFF2CC"), srgb("FFFFFF"), srgb("1F3864"), srgb("FFF2CC"), srgb("808080"),
             srgb("C00000"), srgb("000000")]
    styles = [font_style(scheme("tx1")), font_style(scheme("lt1")), font_style(scheme("accent2")),
              font_style(scheme("lt1")), font_style(scheme("tx1")), font_style(scheme("accent6")),
              font_style(srgb("FFFF00")), font_style(scheme("tx1"))]
    out.append(Case("auto", text_row(fills, styles), "wpg", 5600000, 400000, "font references"))

    # -- inline ---------------------------------------------------------------------
    out.append(Case("gradient", shape(rect, gradient(TWO, lin=(0, 0)), NO_LINE, cx=1800000, cy=900000),
                    note="inline", inline=True))
    out.append(Case("dash", shape(preset("ellipse"), "<a:noFill/>", line(38100, solid(RED), dash="sysDot", cap="rnd"),
                                  cx=1800000, cy=900000), note="inline, rounded dots", inline=True))

    # -- wordart --------------------------------------------------------------------
    def w14_run(text: str, fill: str, outline: str = "", size: int = 72) -> str:
        return (f'<w:r><w:rPr><w:b/><w:color w:val="4472C4"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
                f"{outline}{fill}</w:rPr><w:t xml:space=\"preserve\">{text}</w:t></w:r>")

    w14_solid = '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="C00000"/></w14:solidFill></w14:textFill>'
    w14_grad = ('<w14:textFill><w14:gradFill><w14:gsLst><w14:gs w14:pos="0"><w14:srgbClr w14:val="C00000"/></w14:gs>'
                '<w14:gs w14:pos="100000"><w14:srgbClr w14:val="0070C0"/></w14:gs></w14:gsLst>'
                '<w14:lin w14:ang="0" w14:scaled="0"/></w14:gradFill></w14:textFill>')
    w14_line = ('<w14:textOutline w14:w="19050" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill>'
                '<w14:srgbClr w14:val="000000"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/>'
                "</w14:textOutline>")
    w14_hollow = '<w14:textFill><w14:noFill/></w14:textFill>'
    runs = {"solid fill": w14_run("Solid", w14_solid), "gradient fill": w14_run("Gradient", w14_grad),
            "outline": w14_run("Outlined", w14_solid, w14_line), "hollow": w14_run("Hollow", w14_hollow, w14_line)}
    for note, run in runs.items():
        paragraph = f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>{run}</w:p>'
        body = drawing_probe.body_pr((91440, 45720, 91440, 45720), "t", "<a:spAutoFit/>").replace(
            'anchorCtr="0"', 'anchorCtr="0" fromWordArt="1"').replace(
            "<a:spAutoFit/>", '<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:spAutoFit/>')
        out.append(Case("wordart", shape(rect, "<a:noFill/>", NO_LINE, cx=4000000, cy=900000, text=paragraph,
                                         body=body), cx=4000000, note=note))
    return tuple(out)


CASES = _cases()

#: Office's own format scheme (Office 2013-2022 theme), which is what documents made in
#: Word carry: the style references resolve through it.
FORMAT_SCHEME = (
    '<a:fmtScheme name="Office"><a:fillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/>'
    '</a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/>'
    '<a:tint val="73000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/>'
    '<a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/>'
    '</a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/>'
    '<a:shade val="100000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/>'
    '<a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
    "</a:fillStyleLst><a:lnStyleLst>"
    + "".join(f'<a:ln w="{w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
              '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' for w in (6350, 12700, 19050))
    + "</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>"
    "<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst>"
    '<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000">'
    '<a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>'
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/>'
    '<a:lumMod val="102000"/></a:schemeClr></a:gs>'
    '<a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/>'
    '<a:lumMod val="103000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>'
)


def theme() -> tuple[str, str, str, str]:
    name, content_type, rel, xml = wml.theme_part({"latin": "Calibri Light"}, {"latin": "Calibri"})
    start = xml.index("<a:fmtScheme")
    stop = xml.index("</a:fmtScheme>") + len("</a:fmtScheme>")
    return name, content_type, rel, xml[:start] + FORMAT_SCHEME + xml[stop:]


def body() -> str:
    out = ""
    for number, case in enumerate(CASES):
        out += anchor_probe._p(f"Case {number} {case.family} {case.note}", pageBreakBefore=True)
        run = Anchor(("margin", "offset", 300000), ("margin", "offset", 700000), case.cx, case.cy,
                     graphic=graphic(case.content, case.kind)).xml(number + 1)
        if case.inline:
            start, end = run.index("<wp:anchor"), run.index("</wp:anchor>") + len("</wp:anchor>")
            run = (run[:start] + f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{case.cx}" '
                   f'cy="{case.cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{number + 1}" '
                   f'name="Drawing {number + 1}"/><wp:cNvGraphicFramePr/>{graphic(case.content, case.kind)}'
                   "</wp:inline>" + run[end:])
        run = run.replace("<w:drawing ", f'<w:drawing xmlns:w14="{W14_NS}" ', 1)
        out += wml.paragraph(wml.run(f"Case{number} anchors a drawing here.") + run, mark={},
                             spacing={"before": 0, "after": 0, "line": 240, "lineRule": "auto"})
        out += anchor_probe._p(f"Case {number} after the anchor's paragraph.")
    return out


def build(setting: str = "15") -> bytes:
    data = anchor_probe.package(body(), setting, extra=(theme(),))
    return data


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"dml-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

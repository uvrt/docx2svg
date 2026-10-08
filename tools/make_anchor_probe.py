#!/usr/bin/env python3
"""Where Word puts a floating drawing that text does not wrap around, and in what order it
paints it: the anchors every local template's first paragraph holds.

A floating drawing (``wp:anchor``) is positioned against something -- the page, a margin,
the column, its paragraph, its line, its character -- by an offset or an alignment, on
each axis on its own (ECMA-376 20.4.2.10-11, 20.4.3.1-2).  None of that was measured.
Every case here is a page: a heading line, then the paragraph that anchors a picture
(a 1 x 1 PNG stretched, so Word's image box is the drawing's box), then a line of text
after it.  With ``wrapNone`` the text is where it would be without the picture, which the
glyph check holds; the picture's box is read from Word's PDF.

The section is off the pixel grid on purpose: A4, margins top 1500, right 1300, bottom
1600, left 1442 twips (300.42 px), header 700, footer 650.  Families (``CASES``):

* ``h`` -- every ``positionH/@relativeFrom`` (``page``, ``margin``, ``column``,
  ``character``, ``leftMargin``, ``rightMargin``, ``insideMargin``, ``outsideMargin``)
  with an offset, and with each alignment (``left``, ``center``, ``right``; ``inside``
  and ``outside`` against the page and the margin), negative offsets;
* ``v`` -- every ``positionV/@relativeFrom`` (``page``, ``margin``, ``paragraph``,
  ``line``, ``topMargin``, ``bottomMargin``, ``insideMargin``, ``outsideMargin``) with an
  offset and each alignment (``top``, ``center``, ``bottom``, ``inside``, ``outside``);
  each ``v`` case is paired with an ``h`` case -- the two axes are read apart;
* ``where`` -- the anchor paragraph after space before, second on its page, the anchor on
  the second line of a paragraph and after text on a centred line (``line``,
  ``character``, ``paragraph``);
* ``simple`` -- ``simplePos="1"`` with ``wp:simplePos``;
* ``effect`` -- a ``wp:effectExtent`` on each side;
* ``z`` -- overlapping pictures in four colours: ``relativeHeight`` against document
  order, ``behindDoc`` against the text, a shaded paragraph and a shaded table cell;
  ``allowOverlap="0"``, ``locked="1"``;
* ``cell`` -- a picture anchored in a table cell, ``layoutInCell`` on and off.

Extents step by 3,701 EMU across and 37,003 down, so their fractions of a pixel differ.
Three documents: no ``settings.xml``, mode 15, and mode 15 with ``w:mirrorMargins``
(inside and outside).  Reader: ``read_anchor_probe.py``.
"""

from __future__ import annotations

import io
import struct
import zipfile
import zlib
from dataclasses import dataclass, field

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
#: ``(compatibility mode, mirror margins)``.
SETTINGS = {"none": (None, False), "15": (15, False)}
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
#: The pictures' colours; image ``k`` is ``rId{k + 2}`` (styles are ``rId1``).
COLOURS = ((0x80, 0x80, 0x80), (0xE0, 0x30, 0x30), (0x30, 0xA0, 0x30), (0x30, 0x50, 0xE0))
#: The section: A4, and margins off the pixel grid.
PAGE = {"w": 11906, "h": 16838}
MARGINS = {"top": 1500, "right": 1300, "bottom": 1600, "left": 1442, "header": 700, "footer": 650, "gutter": 0}


#: The pictures' size in pixels.  Not 1 x 1: Word's PDF export drops a one-pixel picture
#: that the page's edge crops, where it draws a larger one cropped (``overflow``).
PNG_SIZE = 16


def png(rgb: tuple[int, int, int]) -> bytes:
    """A :data:`PNG_SIZE`-square PNG of one colour."""
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    rows = (b"\x00" + bytes(rgb) * PNG_SIZE) * PNG_SIZE
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", PNG_SIZE, PNG_SIZE, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


def position(axis: str, relative: str, kind: str, value) -> str:
    """``wp:positionH`` / ``wp:positionV``: ``kind`` ``offset`` (EMU) or ``align``."""
    inner = f"<wp:posOffset>{value}</wp:posOffset>" if kind == "offset" else f"<wp:align>{value}</wp:align>"
    return f'<wp:position{axis} relativeFrom="{relative}">{inner}</wp:position{axis}>'


def pic_graphic(number: int, cx: int, cy: int, image: int = 0) -> str:
    return (
        f'<a:graphic xmlns:a="{A_NS}"><a:graphicData uri="{PIC_NS}"><pic:pic xmlns:pic="{PIC_NS}">'
        f'<pic:nvPicPr><pic:cNvPr id="{number}" name="Picture {number}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="rId{image + 2}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
    )


@dataclass(frozen=True)
class Anchor:
    """One ``wp:anchor``: where, how big, how it stacks, and what it shows."""

    h: tuple = ("column", "offset", 0)
    v: tuple = ("paragraph", "offset", 0)
    cx: int = 1200000
    cy: int = 700000
    effect: tuple[int, int, int, int] = (0, 0, 0, 0)
    simple: tuple[int, int] | None = None
    behind: bool = False
    height: int = 251659264
    allow_overlap: bool = True
    layout_in_cell: bool = True
    locked: bool = False
    wrap: str = "<wp:wrapNone/>"
    dist: tuple[int, int, int, int] = (0, 0, 114300, 114300)  # t, b, l, r
    image: int = 0
    #: The ``a:graphic`` to draw instead of a picture (a shape, a group).
    graphic: str | None = None

    def xml(self, number: int) -> str:
        t, b, l, r = self.dist
        simple = self.simple or (0, 0)
        el, et, er, eb = self.effect
        graphic = self.graphic if self.graphic is not None else pic_graphic(number, self.cx, self.cy, self.image)
        return (
            f'<w:r><w:rPr><w:noProof/></w:rPr><w:drawing xmlns:wp="{WP_NS}" xmlns:r="{R_NS}">'
            f'<wp:anchor distT="{t}" distB="{b}" distL="{l}" distR="{r}" simplePos="{int(self.simple is not None)}"'
            f' relativeHeight="{self.height}" behindDoc="{int(self.behind)}" locked="{int(self.locked)}"'
            f' layoutInCell="{int(self.layout_in_cell)}" allowOverlap="{int(self.allow_overlap)}">'
            f'<wp:simplePos x="{simple[0]}" y="{simple[1]}"/>'
            + position("H", *self.h) + position("V", *self.v)
            + f'<wp:extent cx="{self.cx}" cy="{self.cy}"/><wp:effectExtent l="{el}" t="{et}" r="{er}" b="{eb}"/>'
            + self.wrap
            + f'<wp:docPr id="{number}" name="Drawing {number}"/><wp:cNvGraphicFramePr/>'
            + graphic + "</wp:anchor></w:drawing></w:r>"
        )


@dataclass(frozen=True)
class Case:
    family: str
    anchors: tuple[Anchor, ...]
    #: ``alone`` (a paragraph of the anchor only), ``start`` (before the text), ``mid``
    #: (after four words), ``line2`` (on the paragraph's second line), ``cell``.
    where: str = "mid"
    before: int = 0
    #: A paragraph of text between the heading and the anchor's paragraph.
    filler: bool = False
    jc: str | None = None
    #: The text paragraph's shading, for paint order.
    shaded: bool = False
    #: The anchor is in the page's first paragraph (the heading), with this space before.
    heading: int | None = None
    #: In the second section, whose header and footer distances differ.
    second: bool = False
    note: str = ""
    extra: dict = field(default_factory=dict)


def _size(k: int) -> tuple[int, int]:
    return 1200000 + k * 3701, 700000 + k * 37003


def _cases() -> tuple[Case, ...]:
    h_cases = []
    for relative in ("page", "margin", "column", "character", "leftMargin", "rightMargin", "insideMargin",
                     "outsideMargin"):
        h_cases.append((relative, "offset", 400003))
    for relative in ("page", "margin", "column", "character", "leftMargin", "rightMargin", "insideMargin",
                     "outsideMargin"):
        for align in ("left", "center", "right"):
            h_cases.append((relative, "align", align))
    for relative in ("page", "margin"):
        for align in ("inside", "outside"):
            h_cases.append((relative, "align", align))
    h_cases += [("column", "offset", -250001), ("page", "offset", 3000017), ("margin", "offset", 1)]
    v_cases = []
    for relative in ("page", "margin", "paragraph", "line", "topMargin", "bottomMargin", "insideMargin",
                     "outsideMargin"):
        v_cases.append((relative, "offset", 300007))
    for relative in ("page", "margin", "paragraph", "line", "topMargin", "bottomMargin", "insideMargin",
                     "outsideMargin"):
        for align in ("top", "center", "bottom"):
            v_cases.append((relative, "align", align))
    for relative in ("page", "margin"):
        for align in ("inside", "outside"):
            v_cases.append((relative, "align", align))
    v_cases += [("paragraph", "offset", -200003), ("page", "offset", 5000011), ("line", "offset", -1)]
    out = []
    count = max(len(h_cases), len(v_cases))
    for k in range(count):
        cx, cy = _size(k)
        h = h_cases[k % len(h_cases)]
        v = v_cases[k % len(v_cases)]
        out.append(Case("hv", (Anchor(h, v, cx, cy),), where=("mid", "start", "alone", "line2")[k % 4]))
    # The same inside/outside cases again, so each falls on a page of the other parity.
    for k, (h, v) in enumerate(zip([c for c in h_cases if c[1] == "align" and c[2] in ("inside", "outside")]
                                   + [c for c in h_cases if c[0] in ("insideMargin", "outsideMargin")],
                                   [c for c in v_cases if c[1] == "align" and c[2] in ("inside", "outside")]
                                   + [c for c in v_cases if c[0] in ("insideMargin", "outsideMargin")]
                                   + [("page", "offset", 0)] * 20)):
        cx, cy = _size(k + 7)
        out.append(Case("parity", (Anchor(h, v, cx, cy),), filler=bool(k % 2)))
    # Where the anchor is: its paragraph after space before, second on its page, on its
    # second line, on a centred line, right-aligned.
    for k, (where, before, filler, jc) in enumerate((
            ("mid", 240, False, None), ("mid", 137, True, None), ("line2", 0, True, None),
            ("mid", 0, False, "center"), ("line2", 0, False, "right"), ("alone", 300, True, None),
            ("start", 0, True, "center"))):
        cx, cy = _size(k + 3)
        for h, v in ((("character", "offset", 123457), ("line", "offset", 76543)),
                     (("column", "offset", 654321), ("paragraph", "offset", 45679))):
            out.append(Case("where", (Anchor(h, v, cx, cy),), where=where, before=before, filler=filler, jc=jc))
    # ``simplePos``: the position from the page's top left corner, whatever else is said.
    for k, simple in enumerate(((914400, 1828800), (1000003, 3000007))):
        cx, cy = _size(k + 11)
        out.append(Case("simple", (Anchor(("column", "offset", 100000), ("paragraph", "offset", 100000), cx, cy,
                                          simple=simple),)))
    # An effect extent on each side, with an offset and with an alignment.
    for k, effect in enumerate(((63500, 0, 0, 0), (0, 63500, 0, 0), (0, 0, 63500, 0), (0, 0, 0, 63500),
                                (12701, 25402, 38103, 50804))):
        cx, cy = _size(k + 17)
        out.append(Case("effect", (Anchor(("margin", "offset", 500000), ("paragraph", "offset", 200000), cx, cy,
                                          effect=effect),)))
        out.append(Case("effect", (Anchor(("margin", "align", "right"), ("margin", "align", "bottom"), cx, cy,
                                          effect=effect),)))
    # Stacking: pictures overlapping one another and the text.
    def overlap(k: int, **kwargs) -> Anchor:
        return Anchor(("column", "offset", 200000 + k * 300000), ("paragraph", "offset", -100000 + k * 120000),
                      1500000, 600000, image=k % 4, **kwargs)
    out.append(Case("z", (overlap(0, height=3), overlap(1, height=1), overlap(2, height=2)), note="heights"))
    out.append(Case("z", (overlap(0, height=5, behind=True), overlap(1, height=4, behind=True),
                          overlap(2, height=6), overlap(3, height=7, behind=True)), note="behind"))
    out.append(Case("z", (overlap(0, height=2, behind=True), overlap(1, height=1)), shaded=True, note="shaded"))
    out.append(Case("z", (overlap(0, height=2, allow_overlap=False), overlap(1, height=1, allow_overlap=False)),
                    note="allowOverlap 0"))
    out.append(Case("z", (overlap(0, height=2, locked=True), overlap(1, height=1, locked=True)), note="locked"))
    # The anchor in the page's first paragraph: the local templates' title.
    for k, before in enumerate((0, 240, 480, 0)):
        cx, cy = _size(k + 29)
        for h, v in ((("column", "offset", 4039870), ("paragraph", "offset", -83820)),
                     (("character", "offset", 23456), ("line", "offset", 34567)),
                     (("column", "align", "center"), ("line", "align", "bottom"))):
            out.append(Case("first", (Anchor(h, v, cx, cy),), where=("mid", "start", "alone", "line2")[k],
                            heading=before))
    # What the character is: after a letter, after a tab, first on a line after a break.
    for k, where in enumerate(("glued", "tab", "break", "big")):
        cx, cy = _size(k + 41)
        out.append(Case("char", (Anchor(("character", "offset", 0), ("line", "offset", 0), cx, cy),), where=where))
        out.append(Case("char", (Anchor(("character", "align", "right"), ("line", "align", "center"), cx, cy),),
                        where=where))
    # Past the page's edges: the bottom by a pixel and by more, the right, the top.
    for k, (h, v) in enumerate((
            (("page", "offset", 100000), ("page", "offset", 9991725)),
            (("page", "offset", 100000), ("page", "offset", 9991725 + 6350)),
            (("page", "offset", 100000), ("page", "offset", 9991725 + 635000)),
            (("page", "offset", 7000000), ("page", "offset", 1000000)),
            (("page", "offset", 7400000), ("page", "offset", 1000000)),
            (("page", "offset", 100000), ("page", "offset", -500000)),
            (("column", "offset", 100000), ("paragraph", "offset", 9000000)),
            (("column", "offset", 100000), ("margin", "align", "bottom")))):
        out.append(Case("overflow", (Anchor(h, v, 1500000, 700000),), where="mid"))
    # Stacking with heights as Word writes them.
    base = 251658240
    def stacked(k: int, height: int, **kwargs) -> Anchor:
        return Anchor(("column", "offset", 200000 + k * 300000), ("paragraph", "offset", -100000 + k * 120000),
                      1500000, 600000, image=k % 4, height=base + height, **kwargs)
    out.append(Case("z", (stacked(0, 3072), stacked(1, 1024), stacked(2, 2048)), note="heights, real"))
    out.append(Case("z", (stacked(0, 1024), stacked(1, 3072), stacked(2, 2048)), note="heights, real 2"))
    out.append(Case("z", (stacked(0, 1024), stacked(1, 1024), stacked(2, 1024)), note="heights, tied"))
    out.append(Case("z", (stacked(0, 5120, behind=True), stacked(1, 4096, behind=True), stacked(2, 6144),
                          stacked(3, 1024)), note="behind, real"))
    # Effect extents with alignments to the left, the centre and the top.
    for k, effect in enumerate(((63500, 25400, 12700, 38100), (12701, 25402, 38103, 50804))):
        cx, cy = _size(k + 51)
        for h, v in ((("margin", "align", "left"), ("margin", "align", "top")),
                     (("margin", "align", "center"), ("margin", "align", "center")),
                     (("character", "align", "center"), ("line", "align", "center"))):
            out.append(Case("effect", (Anchor(h, v, cx, cy, effect=effect),)))
    # The line's box with space before, and a taller line.
    for k, (before, where) in enumerate(((240, "start"), (240, "alone"), (0, "big"))):
        cx, cy = _size(k + 57)
        for v in (("line", "align", "center"), ("line", "align", "bottom"), ("paragraph", "align", "bottom")):
            out.append(Case("line", (Anchor(("column", "offset", 0), v, cx, cy),), where=where, before=before))
    # The page's inside and outside against other header and footer distances.
    for k, (h, v) in enumerate(((("page", "align", "inside"), ("page", "align", "inside")),
                                (("page", "align", "outside"), ("page", "align", "outside")),
                                (("page", "align", "inside"), ("page", "align", "inside")),
                                (("page", "align", "outside"), ("page", "align", "outside")))):
        cx, cy = _size(k + 61)
        out.append(Case("second", (Anchor(h, v, cx, cy),), second=True))
    # Anchored in a table cell: last, as a table holding one stops the layout.
    out.append(Case("z", (overlap(0, height=9, behind=True),), where="cell", note="under a shaded cell",
                    second=True))
    for k, in_cell in enumerate((True, False, True, False)):
        cx, cy = _size(k + 23)
        h = ("column", "offset", 300000) if k < 2 else ("page", "offset", 1000000)
        v = ("paragraph", "offset", 100000) if k < 2 else ("page", "offset", 2000000)
        out.append(Case("cell", (Anchor(h, v, cx, cy, layout_in_cell=in_cell),), where="cell", second=True))
    return tuple(out)


CASES = _cases()
#: The second section's margins: other header and footer distances.
SECOND = {**MARGINS, "header": 900, "footer": 500}

WORDS = ("Anchored drawings float over this text, which Word lays out as if the drawing were not "
         "there at all, line after line, until the paragraph ends here.")


def _p(text: str, **props) -> str:
    props.setdefault("spacing", {"before": 0, "after": 0, "line": 240, "lineRule": "auto"})
    return wml.paragraph(wml.run(text), mark={}, **props)


def _cell_table(runs: str, shaded: bool) -> str:
    shading = '<w:shd w:val="clear" w:color="auto" w:fill="FFE0A0"/>' if shaded else ""
    cell = (f'<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>{shading}</w:tcPr>'
            f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            f'{runs}</w:p><w:p><w:r><w:t>Second line of the cell.</w:t></w:r></w:p></w:tc>')
    other = ('<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
             '<w:p><w:r><w:t>Next cell.</w:t></w:r></w:p></w:tc>')
    borders = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                      for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    return (f'<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblInd w:w="400" w:type="dxa"/>'
            f'<w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
            '<w:tblLook w:val="0000"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/>'
            f'</w:tblGrid><w:tr>{cell}{other}</w:tr></w:tbl>')


def body() -> str:
    out = ""
    serial = 1
    for number, case in enumerate(CASES):
        anchors = ""
        for anchor in case.anchors:
            anchors += anchor.xml(serial)
            serial += 1
        if case.heading is not None:
            spacing = {"before": case.heading, "after": 0, "line": 240, "lineRule": "auto"}
            heading = f"Case {number} {case.family} heading"
            runs = {"mid": wml.run(heading[:8]) + anchors + wml.run(heading[8:]), "start": anchors + wml.run(heading),
                    "alone": anchors, "line2": wml.run(f"Case{number} " + WORDS[:118]) + anchors
                    + wml.run(WORDS[118:])}[case.where]
            out += wml.paragraph(runs, mark={}, pageBreakBefore=True, spacing=spacing)
            out += _p(f"Case {number} after the anchor's paragraph, a line of text.")
            continue
        if case.second and not CASES[number - 1].second:
            # The paragraph that ends the first section carries its properties.
            out += wml.paragraph(wml.run(f"Case {number} ends the first section."), mark={},
                                 spacing={"before": 0, "after": 0, "line": 240, "lineRule": "auto"},
                                 sect=section(MARGINS))
        out += _p(f"Case {number} {case.family} heading", pageBreakBefore=True)
        if case.filler:
            out += _p(f"Case {number} filler line of text.", spacing={"before": 0, "after": 120, "line": 240,
                                                                      "lineRule": "auto"})
        spacing = {"before": case.before, "after": 0, "line": 240, "lineRule": "auto"}
        props = {"spacing": spacing}
        if case.jc:
            props["jc"] = case.jc
        if case.shaded:
            props["shd"] = {"val": "clear", "color": "auto", "fill": "C0E0FF"}
        if case.where == "alone":
            out += wml.paragraph(anchors, mark={}, **props)
        elif case.where == "start":
            out += wml.paragraph(anchors + wml.run(f"Case{number} " + WORDS), mark={}, **props)
        elif case.where == "mid":
            out += wml.paragraph(wml.run(f"Case{number} text before the ") + anchors + wml.run("anchor. " + WORDS),
                                 mark={}, **props)
        elif case.where == "line2":
            out += wml.paragraph(wml.run(f"Case{number} " + WORDS[:118]) + anchors + wml.run(WORDS[118:]),
                                 mark={}, **props)
        elif case.where == "glued":
            out += wml.paragraph(wml.run(f"Case{number} text before the") + anchors + wml.run("anchor. " + WORDS),
                                 mark={}, **props)
        elif case.where == "tab":
            out += wml.paragraph(wml.run(f"Case{number} text") + wml.run("\t") + anchors + wml.run("after a tab."),
                                 mark={}, **props)
        elif case.where == "break":
            out += wml.paragraph(wml.run(f"Case{number} text before a break.") + '<w:r><w:br/></w:r>' + anchors
                                 + wml.run("After the break."), mark={}, **props)
        elif case.where == "big":
            out += wml.paragraph(wml.run(f"Case{number} text before the ") + anchors
                                 + wml.run("big", sz=48, szCs=48) + wml.run(" anchor. " + WORDS), mark={}, **props)
        elif case.where == "cell":
            out += _cell_table(wml.run(f"Case{number} in a ") + anchors + wml.run("cell."),
                               shaded=case.note == "under a shaded cell")
        out += _p(f"Case {number} after the anchor's paragraph, a line of text.")
    return out


def _settings(mode: int | None, mirror: bool) -> tuple[str, str, str, str] | None:
    if mode is None and not mirror:
        return None
    name, content_type, rel, xml = wml.settings_part({"val": "en-GB"}, compatibility_mode=mode)
    if mirror:
        xml = xml.replace('<w:settings xmlns:w="' + probe_docx.W_NS + '">',
                          '<w:settings xmlns:w="' + probe_docx.W_NS + '"><w:mirrorMargins/>')
    return name, content_type, rel, xml


def section(margins: dict | None = None) -> str:
    m = margins or MARGINS
    return (f'<w:sectPr><w:pgSz w:w="{PAGE["w"]}" w:h="{PAGE["h"]}"/>'
            f'<w:pgMar w:top="{m["top"]}" w:right="{m["right"]}" w:bottom="{m["bottom"]}" w:left="{m["left"]}"'
            f' w:header="{m["header"]}" w:footer="{m["footer"]}" w:gutter="{m["gutter"]}"/></w:sectPr>')


def package(body_xml: str, setting: str, *, extra: tuple = (), final_section: str | None = None) -> bytes:
    """A probe package: ``body_xml``, the four colour images, the setting's settings part."""
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    mode, mirror = SETTINGS[setting]
    parts = [(f"word/media/image{k + 1}.png", "image/png", IMAGE_REL, "") for k in range(len(COLOURS))]
    found = _settings(mode, mirror)
    if found is not None:
        parts.append(found)
    parts += list(extra)
    data = probe_docx.package(body_xml, styles=styles, extra_parts=tuple(parts),
                              final_section=final_section if final_section is not None else section())
    source = zipfile.ZipFile(io.BytesIO(data))
    buffer = io.BytesIO()
    images = {f"word/media/image{k + 1}.png": png(colour) for k, colour in enumerate(COLOURS)}
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for info in source.infolist():
            archive.writestr(info, images.get(info.filename) or source.read(info.filename))
    return buffer.getvalue()


def anchor_paths(data: bytes) -> list[list[str]]:
    """Per case, the body path of each drawing's anchor paragraph (``w:body/w:p[k]``,
    ``w:body/w:tbl[k]/w:tr[1]/w:tc[1]/w:p[1]``), read back from the built package."""
    from xml.etree import ElementTree

    w = "{" + probe_docx.W_NS + "}"
    root = ElementTree.fromstring(zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml"))
    where: dict[str, str] = {}
    counts: dict[str, int] = {}
    for node in root.find(w + "body"):
        name = node.tag[len(w):]
        counts[name] = counts.get(name, 0) + 1
        here = f"w:body/w:{name}[{counts[name]}]"
        paragraphs = [(here, node)] if name == "p" else [
            (f"{here}/w:tr[1]/w:tc[1]/w:p[1]", node.find(f"{w}tr/{w}tc/{w}p"))] if name == "tbl" else []
        for path, paragraph in paragraphs:
            for element in paragraph.iter("{" + WP_NS + "}docPr"):
                where[element.get("id")] = path
    out, serial = [], 1
    for case in CASES:
        out.append([where[str(serial + k)] for k in range(len(case.anchors))])
        serial += len(case.anchors)
    return out


def build(setting: str) -> bytes:
    return package(body(), setting, final_section=section(SECOND) if any(c.second for c in CASES) else None)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"anchor-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

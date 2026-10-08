#!/usr/bin/env python3
"""Picture bullets: how large Word draws a list label that is a picture, and where.

A list level may show a picture instead of its text (``w:lvlPicBulletId``, naming a
``w:numPicBullet`` of ``word/numbering.xml``, which holds a VML ``w:pict`` -- what every
Word writes -- or a DrawingML ``w:drawing``).  Nothing of it was measured: the layout
stopped at such a paragraph (``scope.paragraph_features``: ``drawing (picture bullet)``).

Every case is a page: ``Case N``, then a list paragraph whose label is the picture and
whose text is a label (``N.a``) and a word of ``x``, then ``After N``.  The picture is a
PNG of our own (:func:`make_anchor_probe.png`, one colour, 16 x 16 px unless the case says
otherwise), so Word's image box is the bullet's box.  A4, margins off the pixel grid
(left 1,442 twips); Calibri 11 unless the case says otherwise; the level indented 720
hanging 360, ``w:lvlJc`` left, a tab after the label, ``w:lvlText`` a Symbol bullet as
Word writes it.  Families (``CASES``):

* ``size`` -- the label's size (the paragraph mark's, which the text shares): 8, 11, 16,
  24, 40 pt, a 9 pt square picture;
* ``natural`` -- the picture's own size (``v:shape/@style``): 4.5, 18, 36 and 72 pt square
  at 11 pt, and a 16 x 16 px picture stated 9 pt against one of 64 x 64 px;
* ``aspect`` -- 18 x 9 pt and 9 x 18 pt, at 11 and 24 pt, and a 32 x 16 px image stated
  square;
* ``level`` -- the level's ``w:rPr`` size (20 pt) over an 11 pt mark and text; the mark 20 pt
  under 11 pt text; the level naming Calibri instead of Symbol; bold;
* ``face`` -- the text and mark in Times New Roman, Courier New and Cambria at 11 pt;
* ``indent`` -- hanging 720, hanging 0, a first-line indent, left 0 hanging 360, ``w:lvlJc``
  center and right, a label wider than its hanging indent (40 pt);
* ``suffix`` -- ``w:suff`` space and nothing;
* ``lines`` -- three lines of text after the label, ``exact`` 400 and ``atLeast`` 600
  twips, ``auto`` 276, space before 120;
* ``position`` -- the level's ``w:position`` +6 and -6 half points, and ``w:vertAlign``
  superscript;
* ``drawing`` -- the picture as a DrawingML ``w:drawing`` (``wp:inline``) of 9 pt, and of
  18 x 9 pt;
* ``pixels`` -- (round two, after the first showed the picture's pixels count) square
  pictures of 8 to 128 px, 16 x 32 and 64 x 16 px, a ``pHYs`` of 300 and 72 dpi, at 11 and
  24 pt;
* ``sizes`` -- the label's size from 9 to 72 pt, and at 9.5 to 12.5 pt;
* ``suffix`` -- a space after the label at 8 and 24 pt, in a level in Calibri, after a
  64 px picture;
* ``jc`` -- (round three) ``w:lvlJc`` center and right at 24 pt and with nothing after the
  label, and a text label (``%1.``) as a control.

Three documents: no ``settings.xml``, mode 14 and mode 15.  Reader:
``read_picture_bullet_probe.py``.
"""

from __future__ import annotations

import io
import struct
import zipfile
import zlib
from dataclasses import dataclass, field

import probe_docx
import wml

SETTINGS = {"none": None, "14": 14, "15": 15}
FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SYMBOL = {"ascii": "Symbol", "hAnsi": "Symbol", "hint": "default"}
SPACING = {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}
PAGE = {"w": 11906, "h": 16838}
MARGINS = {"top": 1500, "right": 1300, "bottom": 1600, "left": 1442, "header": 700, "footer": 650, "gutter": 0}
IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
V_NS = "urn:schemas-microsoft-com:vml"
O_NS = "urn:schemas-microsoft-com:office:office"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
#: The images, by name: (width px, height px, colour, dots per inch stated in a ``pHYs``
#: chunk or ``None`` for none).
IMAGES = {"square": (16, 16, (0xE0, 0x30, 0x30), None), "big": (64, 64, (0x30, 0xA0, 0x30), None),
          "wide": (32, 16, (0x30, 0x50, 0xE0), None),
          **{f"px{n}": (n, n, (0x80, 0x40, 0xC0), None) for n in (8, 12, 24, 32, 48, 128)},
          "tall": (16, 32, (0x30, 0x50, 0xE0), None), "flat": (64, 16, (0x30, 0x50, 0xE0), None),
          "dpi300": (16, 16, (0xE0, 0x30, 0x30), 300), "dpi72": (64, 64, (0x30, 0xA0, 0x30), 72)}


def png(width: int, height: int, rgb: tuple[int, int, int], dpi: int | None = None) -> bytes:
    """A PNG of one colour, ``width`` x ``height`` px, with a ``pHYs`` of ``dpi`` if given."""
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    rows = (b"\x00" + bytes(rgb) * width) * height
    phys = b""
    if dpi is not None:
        per_metre = round(dpi / 0.0254)
        phys = chunk(b"pHYs", struct.pack(">IIB", per_metre, per_metre, 1))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + phys + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


@dataclass(frozen=True)
class Case:
    family: str
    note: str
    #: The picture's stated size in points (``v:shape/@style``, or the inline extent).
    size: tuple = (9, 9)
    image: str = "square"
    #: ``pict`` (VML) or ``drawing`` (DrawingML).
    kind: str = "pict"
    #: The mark's and the text's half points, and face.
    mark: int = 22
    text: int = 22
    face: str = "Calibri"
    #: The level's ``w:rPr`` beyond its Symbol font: ``sz``, ``b``, ``position``...
    level_rpr: dict = field(default_factory=dict)
    symbol: bool = True
    ind: dict = field(default_factory=lambda: {"left": 720, "hanging": 360})
    jc: str = "left"
    suffix: str | None = None
    words: int = 1
    spacing: dict = field(default_factory=lambda: dict(SPACING))
    #: ``False``: a text label (``w:lvlText`` ``%1.``, decimal), no picture -- a control.
    picture: bool = True


def _cases() -> tuple[Case, ...]:
    out: list[Case] = []
    add = out.append
    for pt in (8, 11, 16, 24, 40):
        add(Case("size", f"label and text {pt} pt", mark=pt * 2, text=pt * 2))
    for pt in (4.5, 18, 36, 72):
        add(Case("natural", f"picture {pt} pt square", size=(pt, pt)))
    add(Case("natural", "64 px image stated 9 pt", image="big"))
    for pt in (11, 24):
        add(Case("aspect", f"18 x 9 pt at {pt} pt", size=(18, 9), mark=pt * 2, text=pt * 2))
        add(Case("aspect", f"9 x 18 pt at {pt} pt", size=(9, 18), mark=pt * 2, text=pt * 2))
    add(Case("aspect", "32 x 16 px image stated square", image="wide"))
    add(Case("level", "level sz 40 over 11 pt", level_rpr={"sz": 40, "szCs": 40}))
    add(Case("level", "mark 20 pt under 11 pt text", mark=40))
    add(Case("level", "text 20 pt, mark 11 pt", text=40))
    add(Case("level", "level in Calibri, no Symbol", symbol=False))
    add(Case("level", "level bold", level_rpr={"b": True}))
    for face in ("Times New Roman", "Courier New", "Cambria"):
        add(Case("face", face, face=face))
    add(Case("indent", "hanging 720", ind={"left": 1440, "hanging": 720}))
    add(Case("indent", "hanging 0", ind={"left": 720, "hanging": 0}))
    add(Case("indent", "first line 360", ind={"left": 720, "firstLine": 360}))
    add(Case("indent", "left 0 hanging 360", ind={"left": 0, "hanging": 360}))
    add(Case("indent", "lvlJc center", jc="center"))
    add(Case("indent", "lvlJc right", jc="right"))
    add(Case("indent", "a 40 pt picture wider than the hanging indent", size=(40, 40)))
    add(Case("suffix", "suff space", suffix="space"))
    add(Case("suffix", "suff nothing", suffix="nothing"))
    add(Case("lines", "three lines", words=40))
    add(Case("lines", "exact 400", words=40, spacing={"before": 0, "after": 0, "line": 400, "lineRule": "exact"}))
    add(Case("lines", "atLeast 600", spacing={"before": 0, "after": 0, "line": 600, "lineRule": "atLeast"}))
    add(Case("lines", "auto 276", words=40, spacing={"before": 0, "after": 0, "line": 276, "lineRule": "auto"}))
    add(Case("lines", "before 120", spacing={"before": 120, "after": 0, "line": 240, "lineRule": "auto"}))
    add(Case("position", "level position +6", level_rpr={"position": 6}))
    add(Case("position", "level position -6", level_rpr={"position": -6}))
    add(Case("position", "level superscript", level_rpr={"vertAlign": "superscript"}))
    add(Case("drawing", "DrawingML 9 pt", kind="drawing"))
    add(Case("drawing", "DrawingML 18 x 9 pt", kind="drawing", size=(18, 9)))
    # Round two: what round one showed depends on the image's pixels and the label's size.
    for name in ("px8", "px12", "px24", "px32", "px48", "px128", "tall", "flat", "dpi300", "dpi72"):
        add(Case("pixels", f"image {name}", image=name))
    add(Case("pixels", "image px12 at 24 pt", image="px12", mark=48, text=48))
    add(Case("pixels", "image big at 24 pt", image="big", mark=48, text=48))
    add(Case("pixels", "image big stated 36 pt", image="big", size=(36, 36)))
    for pt in (9, 10, 12, 14, 20, 28, 36, 48, 72):
        add(Case("sizes", f"label and text {pt} pt", mark=pt * 2, text=pt * 2))
    for half in (21, 23, 25, 19):
        add(Case("sizes", f"label and text {half / 2:g} pt", mark=half, text=half))
    add(Case("suffix", "suff space at 8 pt", suffix="space", mark=16, text=16))
    add(Case("suffix", "suff space at 24 pt", suffix="space", mark=48, text=48))
    add(Case("suffix", "suff space, level in Calibri", suffix="space", symbol=False))
    add(Case("suffix", "suff space, image big", suffix="space", image="big"))
    # Round three: w:lvlJc, which the model did not read, on text labels as controls.
    for jc in ("center", "right"):
        add(Case("jc", f"text label, lvlJc {jc}", jc=jc, picture=False, symbol=False))
        add(Case("jc", f"lvlJc {jc} at 24 pt", jc=jc, mark=48, text=48))
        add(Case("jc", f"lvlJc {jc}, suff nothing", jc=jc, suffix="nothing"))
        add(Case("jc", f"text label, lvlJc {jc}, suff nothing", jc=jc, suffix="nothing", picture=False,
                 symbol=False))
    return tuple(out)


CASES = _cases()


def _fonts(face: str) -> dict:
    return {"ascii": face, "hAnsi": face, "eastAsia": face, "cs": face}


def _pt(value: float) -> str:
    return f"{value:g}pt"


#: The formulas of Word's picture frame shape type (``_x0000_t75``), as every Word writes
#: it: without them Word stops at the document with a dialog, and the export times out.
FORMULAS = "<v:formulas>" + "".join(f'<v:f eqn="{eqn}"/>' for eqn in (
    "if lineDrawn pixelLineWidth 0", "sum @0 1 0", "sum 0 0 @1", "prod @2 1 2", "prod @3 21600 pixelWidth",
    "prod @3 21600 pixelHeight", "sum @0 0 1", "prod @6 1 2", "prod @7 21600 pixelWidth", "sum @8 21600 0",
    "prod @7 21600 pixelHeight", "sum @10 21600 0")) + "</v:formulas>"


def picture_xml(number: int, case: Case) -> str:
    """The ``w:numPicBullet``'s content: VML as Word writes it, or an inline drawing."""
    rid = f"rIdImg{list(IMAGES).index(case.image) + 1}"
    width, height = case.size
    if case.kind == "pict":
        return (f'<w:pict><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t"'
                f' path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/>{FORMULAS}'
                '<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>'
                '<o:lock v:ext="edit" aspectratio="t"/></v:shapetype>'
                f'<v:shape id="_x0000_i10{number:02d}" type="#_x0000_t75" style="width:{_pt(width)};height:{_pt(height)}"'
                f' o:bullet="t"><v:imagedata r:id="{rid}" o:title=""/></v:shape></w:pict>')
    cx, cy = round(width * 12700), round(height * 12700)
    return (f'<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{number + 1}" name="Bullet {number}"/>'
            f'<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="{A_NS}" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            f'<a:graphic xmlns:a="{A_NS}"><a:graphicData uri="{PIC_NS}"><pic:pic xmlns:pic="{PIC_NS}">'
            f'<pic:nvPicPr><pic:cNvPr id="{number + 1}" name="Bullet {number}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
            '</wp:inline></w:drawing>')


def numbering_xml() -> str:
    """One ``w:numPicBullet``, ``w:abstractNum`` and ``w:num`` a case: ``w:num`` N + 1."""
    bullets = "".join(f'<w:numPicBullet w:numPicBulletId="{n}">{picture_xml(n, case)}</w:numPicBullet>'
                      for n, case in enumerate(CASES))
    abstracts = nums = ""
    for n, case in enumerate(CASES):
        rpr = {}
        if case.symbol:
            rpr["rFonts"] = SYMBOL
        rpr.update(case.level_rpr)
        suffix = f'<w:suff w:val="{case.suffix}"/>' if case.suffix else ""
        text = "\uf0b7" if case.symbol else "\u2022" if case.picture else "%1."
        picture = f'<w:lvlPicBulletId w:val="{n}"/>' if case.picture else ""
        abstracts += (
            f'<w:abstractNum w:abstractNumId="{n}"><w:multiLevelType w:val="hybridMultilevel"/>'
            f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="{"bullet" if case.picture else "decimal"}"/>'
            f'{suffix}<w:lvlText w:val="{text}"/>{picture}<w:lvlJc w:val="{case.jc}"/>'
            f'<w:pPr>{wml._ordered(wml.PPR_ORDER, {"ind": case.ind})}</w:pPr>{wml.rpr(**rpr)}</w:lvl>'
            "</w:abstractNum>")
        nums += f'<w:num w:numId="{n + 1}"><w:abstractNumId w:val="{n}"/></w:num>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:numbering xmlns:w="{probe_docx.W_NS}" xmlns:v="{V_NS}" xmlns:o="{O_NS}" xmlns:r="{R_NS}"'
            f' xmlns:wp="{WP_NS}">{bullets}{abstracts}{nums}</w:numbering>')


def x(n: int) -> str:
    return "x" * n


def body() -> str:
    out = ""
    for number, case in enumerate(CASES):
        out += wml.paragraph(wml.run(f"Case {number}"), mark={}, spacing=SPACING, pageBreakBefore=True)
        text = f"{number}.a " + " ".join([x(4)] * case.words) if case.words > 1 else f"{number}.a {x(6)}"
        fonts = _fonts(case.face)
        run = wml.run(text, rFonts=fonts, sz=case.text, szCs=case.text)
        out += wml.paragraph(run, mark={"rFonts": fonts, "sz": case.mark, "szCs": case.mark},
                             numPr=f'<w:ilvl w:val="0"/><w:numId w:val="{number + 1}"/>', spacing=case.spacing)
        out += wml.paragraph(wml.run(f"After {number}"), mark={}, spacing=SPACING)
    return out


def section() -> str:
    m = MARGINS
    return (f'<w:sectPr><w:pgSz w:w="{PAGE["w"]}" w:h="{PAGE["h"]}"/>'
            f'<w:pgMar w:top="{m["top"]}" w:right="{m["right"]}" w:bottom="{m["bottom"]}" w:left="{m["left"]}"'
            f' w:header="{m["header"]}" w:footer="{m["footer"]}" w:gutter="{m["gutter"]}"/></w:sectPr>')


def _entry(name: str) -> zipfile.ZipInfo:
    """A zip entry dated as ``probe_docx`` dates its own, so a build is the same bytes
    every time (the oracle caches by content)."""
    info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o600 << 16
    info.create_system = 3  # not the platform default (0 on Windows)
    return info


def build(setting: str) -> bytes:
    """``setting``: ``none``, ``14`` or ``15``."""
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": SPACING},
    )
    mode = SETTINGS[setting]
    extra = [("word/numbering.xml", wml.NUMBERING_CONTENT_TYPE, wml.NUMBERING_REL, numbering_xml())]
    if mode is not None:
        extra.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=mode))
    data = probe_docx.package(body(), styles=styles, extra_parts=tuple(extra), final_section=section())
    source = zipfile.ZipFile(io.BytesIO(data))
    buffer = io.BytesIO()
    rels = "".join(f'<Relationship Id="rIdImg{k + 1}" Type="{IMAGE_REL}" Target="media/bullet{k + 1}.png"/>'
                   for k in range(len(IMAGES)))
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for info in source.infolist():
            content = source.read(info.filename)
            if info.filename == "[Content_Types].xml":
                content = content.replace(b'<Default Extension="xml"',
                                          b'<Default Extension="png" ContentType="image/png"/><Default Extension="xml"')
            archive.writestr(info, content)
        archive.writestr(_entry("word/_rels/numbering.xml.rels"),
                         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                         f"{rels}</Relationships>")
        for k, (width, height, colour, dpi) in enumerate(IMAGES.values()):
            archive.writestr(_entry(f"word/media/bullet{k + 1}.png"), png(width, height, colour, dpi))
    return buffer.getvalue()


DOCUMENTS = tuple(f"picture-bullet-{s}" for s in SETTINGS)


if __name__ == "__main__":
    import sys
    from pathlib import Path

    for setting in SETTINGS:
        path = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / f"picture-bullet-{setting}.docx"
        path.write_bytes(build(setting))
        print(path)

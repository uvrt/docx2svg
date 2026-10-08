#!/usr/bin/env python3
"""How tall a line holding an inline picture is: the last thing ``sample4`` and
``samplelib/sample-long`` need before they can be paginated from the file alone.

Both documents put inline pictures (``wp:inline``, ``pic:pic``) in paragraphs of their
own.  A picture sits on the baseline, so the line reaches the picture's height above it;
what is below it, and what a line rule does with a picture, was never measured.  Every
case is a page: a plain anchor with ``w:pageBreakBefore``, a text line, the picture's
paragraph, and a text line after it; the distance between the two text lines' baselines
is the picture line's pitch (and its paragraph's spacing), to the pixel.  The heights are
swept in steps that are not whole pixels, so the fractional part of the pitch covers
[0, 1) under every rule.

Cases (``CASES``): heights 300,000 + k x 37,000 EMU (k = 0..29, 23.6 to 139 pt) under
``auto`` 240, 259, 276 and 360, ``exact`` 300 twips (less than the picture) and
``atLeast`` 300; half of them with ``wp:effectExtent b`` = 6350 (``sample4``'s); ten
with the paragraph mark at 20 pt; ten with ``sample4``'s spacing (``auto`` 259, 160
after) and two pictures side by side.

The picture is a 1 x 1 PNG made here, stretched.  One document per compatibility
setting: no ``settings.xml``, and modes 12, 14 and 15.
"""

from __future__ import annotations

import functools
import io
import struct
import zipfile
import zlib
from dataclasses import dataclass

import probe_docx
import wml

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
SETTINGS = {"none": None, "12": 12, "14": 14, "15": 15}
RULES = ((240, "auto"), (259, "auto"), (276, "auto"), (360, "auto"), (300, "exact"), (300, "atLeast"))
IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
#: The image part's relationship id: styles are rId1, the image rId2 (``probe_docx``).
IMAGE_ID = "rId2"


@dataclass(frozen=True)
class Case:
    height: int  # EMU
    line: int
    rule: str
    effect_bottom: int = 0
    mark: int = 22
    after: int = 0
    pictures: int = 1

    @property
    def key(self) -> str:
        return (f"{self.pictures}x{self.height} {self.rule} {self.line} b{self.effect_bottom} mark {self.mark}"
                f" after {self.after}")


def _cases() -> tuple[Case, ...]:
    out = []
    for line, rule in RULES:
        for k in range(30):
            out.append(Case(300000 + k * 37000, line, rule, 6350 if k % 2 else 0))
    for k in range(10):
        out.append(Case(300000 + k * 111000, 240, "auto", mark=40))
        out.append(Case(1000000 + k * 97000, 259, "auto", 6350 if k % 2 else 0, after=160, pictures=1 + k % 2))
    return tuple(out)


CASES = _cases()


def png() -> bytes:
    """A 1 x 1 grey PNG."""
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 0, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"\x00\x80")) + chunk(b"IEND", b""))


def picture(number: int, cx: int, cy: int, effect_bottom: int = 0) -> str:
    return (
        f'<w:r><w:drawing xmlns:wp="{WP_NS}" xmlns:r="{R_NS}"><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="{effect_bottom}"/>'
        f'<wp:docPr id="{number}" name="Picture {number}"/>'
        f'<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="{A_NS}" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="{A_NS}"><a:graphicData uri="{PIC_NS}"><pic:pic xmlns:pic="{PIC_NS}">'
        f'<pic:nvPicPr><pic:cNvPr id="{number}" name="Picture {number}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{IMAGE_ID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
        "</wp:inline></w:drawing></w:r>"
    )


def _p(text: str, **props) -> str:
    props.setdefault("spacing", {"before": 0, "after": 0, "line": 240, "lineRule": "auto"})
    return wml.paragraph(wml.run(text), mark={}, **props)


@functools.lru_cache(maxsize=1)
def blocks() -> tuple[tuple[str, str], ...]:
    """``(kind, w:p)``: ``anchor``, ``before``, ``picture``, ``after``."""
    out = []
    serial = 1
    for number, case in enumerate(CASES):
        out.append(("anchor", _p(f"Case {number} anchor", pageBreakBefore=True)))
        out.append(("before", _p(f"Case {number} before")))
        runs = ""
        for _ in range(case.pictures):
            runs += picture(serial, 1500000, case.height, case.effect_bottom)
            serial += 1
        mark = {"rFonts": FACE, "sz": case.mark, "szCs": case.mark}
        out.append(("picture", wml.paragraph(runs, mark=mark, spacing={
            "before": 0, "after": case.after, "line": case.line, "lineRule": case.rule})))
        out.append(("after", _p(f"Case {number} after")))
    return tuple(out)


def kinds() -> list[str]:
    return [kind for kind, _ in blocks()]


def build(setting: str) -> bytes:
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    mode = SETTINGS[setting]
    extra = [("word/media/image1.png", "image/png", IMAGE_REL, "")]
    if mode is not None:
        extra.append(wml.settings_part({"val": "en-GB"}, compatibility_mode=mode))
    data = probe_docx.package("".join(p for _, p in blocks()), styles=styles, extra_parts=tuple(extra))
    # probe_docx writes text parts; put the image's bytes in place of its empty part.
    source = zipfile.ZipFile(io.BytesIO(data))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for info in source.infolist():
            content = png() if info.filename == "word/media/image1.png" else source.read(info.filename)
            archive.writestr(info, content)
    return buffer.getvalue()

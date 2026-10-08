#!/usr/bin/env python3
"""How far Word raises a superscript and lowers a subscript -- with faces of our own.

Finding 5 left the offset recorded, not settled (ROADMAP.md, "Superscripts,
subscripts..."): for faces whose ``OS/2.ySuperscriptYOffset`` is at most 0.39 em it is
nearly always ``round(w:sz x offset / upm)`` half points, and for Calibri, Times New
Roman, Arial and Courier New (0.42-0.48 em) it is not, and not monotonic in the size.
Real faces change every field at once.  So this probe makes its own: a face of seven
rectangle glyphs (``H``, ``x``, ``h``, space and three digits), written by this script
with ``fontTools``, **ours and nobody else's** -- no outline comes from any existing font
-- in which one field moves at a time:

* ``sup`` -- ``ySuperscriptYOffset`` from 0 to 1400 units of 2048 (script size 1331,
  0.65 em; ascent 1900, descent 500), and ``sub`` -- ``ySubscriptYOffset`` likewise;
* ``size`` -- the script size 820, 1024, 1331 and 1600 at one offset (600);
* ``ascent`` -- hhea/typo/win ascent 1400, 1900 and 2400, and descent 300 and 800, at
  one superscript offset (900, past the 0.42 em where real faces stop following it) and
  one subscript offset (300);
* ``negative`` -- a negative subscript offset (Charter's ``-411``), with a positive
  superscript one.

Each face is **embedded** in the document (ECMA-376 17.8.1: obfuscated with its
``w:fontKey``), which Word lays out and draws with when no face of that name is
installed (``filesamples/sample1``).  Every face is swept over ``w:sz`` 8-96 in steps of
four half points, one line each: ``Hx``, a superscript ``Hxh``, `` Hx ``, a subscript
``Hxh``, `` Hx``, all at one size.  Then, for the two faces at either side of the
threshold, a script **beside a taller run** (text at 11 pt, a 36 pt run, then the
scripts at 11 pt) and a script **larger than its text** (text at 11 pt, scripts at
24 pt), to see whether the line's extent bounds the raise.

The generator is deterministic: font bytes, keys and the package are the same on every
run, so ``oracle.py``'s cache holds.
"""

from __future__ import annotations

import hashlib
import io
import zipfile
from dataclasses import dataclass

import probe_docx
import wml

UPM = 2048
TEXT = "Hx "
SCRIPT = "Hxh"
SIZES = tuple(range(8, 97, 4))


@dataclass(frozen=True)
class Face:
    name: str
    sup_offset: int = 600
    sub_offset: int = 300
    script_size: int = 1331
    ascent: int = 1900
    descent: int = 500


def _faces() -> tuple[Face, ...]:
    out = []
    for value in (0, 200, 400, 600, 700, 800, 860, 900, 960, 1000, 1100, 1200, 1400):
        out.append(Face(f"Dx Sup {value}", sup_offset=value))
    for value in (0, 100, 200, 300, 400, 500, 600, 800, 1000):
        out.append(Face(f"Dx Sub {value}", sub_offset=value))
    for size in (820, 1024, 1600):
        out.append(Face(f"Dx Size {size}", script_size=size))
    for ascent in (1400, 2400):
        out.append(Face(f"Dx Ascent {ascent}", sup_offset=900, ascent=ascent))
    for descent in (300, 800):
        out.append(Face(f"Dx Descent {descent}", sup_offset=900, descent=descent))
    out.append(Face("Dx Negative", sup_offset=600, sub_offset=-411))
    return tuple(out)


FACES = _faces()
#: Faces the taller-run and larger-script lines use: either side of where real faces
#: stop following their offset.
BESIDE = ("Dx Sup 600", "Dx Sup 900", "Dx Sup 1200")


# -- the font -----------------------------------------------------------------------------

#: Glyph name -> (advance, rectangles as (x0, y0, x1, y1) in font units).  Outlines of our
#: own: plain boxes, so that nothing here derives from anyone's design.
GLYPHS = {
    ".notdef": (1000, [(100, 0, 900, 1400)]),
    "space": (500, []),
    "H": (1300, [(150, 0, 350, 1400), (950, 0, 1150, 1400), (350, 600, 950, 800)]),
    "x": (1000, [(150, 0, 850, 900)]),
    "h": (1100, [(150, 0, 350, 1500), (350, 700, 950, 900), (750, 0, 950, 700)]),
    "one": (1050, [(450, 0, 650, 1400)]),
    "two": (1050, [(150, 0, 900, 200), (150, 600, 900, 800), (150, 1200, 900, 1400)]),
}
CMAP = {32: "space", 72: "H", 120: "x", 104: "h", 49: "one", 50: "two"}


def font_bytes(face: Face) -> bytes:
    """A TrueType face ``face.name``, Regular, whose metrics and ``OS/2`` script fields
    are ``face``'s.  Needs fontTools (the ``measure`` extra)."""
    from fontTools.fontBuilder import FontBuilder
    from fontTools.pens.ttGlyphPen import TTGlyphPen

    builder = FontBuilder(UPM, isTTF=True)
    order = list(GLYPHS)
    builder.setupGlyphOrder(order)
    builder.setupCharacterMap(CMAP)
    glyphs = {}
    for name, (_, boxes) in GLYPHS.items():
        pen = TTGlyphPen(None)
        for x0, y0, x1, y1 in boxes:
            pen.moveTo((x0, y0))
            pen.lineTo((x0, y1))
            pen.lineTo((x1, y1))
            pen.lineTo((x1, y0))
            pen.closePath()
        glyphs[name] = pen.glyph()
    builder.setupGlyf(glyphs)
    metrics = {name: (advance, min((b[0] for b in boxes), default=0)) for name, (advance, boxes) in GLYPHS.items()}
    builder.setupHorizontalMetrics(metrics)
    builder.setupHorizontalHeader(ascent=face.ascent, descent=-face.descent, lineGap=0)
    builder.setupNameTable({
        "familyName": face.name, "styleName": "Regular",
        "uniqueFontIdentifier": f"docx2svg probe: {face.name}",
        "fullName": face.name, "psName": face.name.replace(" ", ""),
        "version": "Version 1.0", "copyright": "Made by tools/make_script_offset_probe.py; public domain",
    })
    builder.setupOS2(
        sTypoAscender=face.ascent, sTypoDescender=-face.descent, sTypoLineGap=0,
        usWinAscent=face.ascent, usWinDescent=face.descent, fsType=0, achVendID="DXSV",
        ySubscriptXSize=face.script_size, ySubscriptYSize=face.script_size, ySubscriptXOffset=0,
        ySubscriptYOffset=face.sub_offset,
        ySuperscriptXSize=face.script_size, ySuperscriptYSize=face.script_size, ySuperscriptXOffset=0,
        ySuperscriptYOffset=face.sup_offset,
        yStrikeoutSize=102, yStrikeoutPosition=530, usWeightClass=400, fsSelection=0x40,
        ulUnicodeRange1=1, ulCodePageRange1=1,
    )
    builder.setupPost(underlinePosition=-200, underlineThickness=100)
    builder.setupHead(unitsPerEm=UPM, created=0, modified=0)
    buffer = io.BytesIO()
    builder.save(buffer)
    return buffer.getvalue()


def font_key(face: Face) -> str:
    """A GUID for ``w:fontKey``, derived from the face's name so the package is the same
    on every run."""
    digest = hashlib.sha256(face.name.encode()).hexdigest().upper()
    return f"{{{digest[:8]}-{digest[8:12]}-{digest[12:16]}-{digest[16:20]}-{digest[20:32]}}}"


def obfuscate(data: bytes, key: str) -> bytes:
    """ECMA-376 17.8.1: XOR the first 32 bytes with the key's 16 bytes, read from its last
    hex pair to its first (the inverse of ``docx2svg.fonts.embedded_faces``)."""
    digits = key.strip("{}").replace("-", "")
    mask = [int(digits[i:i + 2], 16) for i in range(30, -1, -2)]
    out = bytearray(data)
    for i in range(32):
        out[i] ^= mask[i % 16]
    return bytes(out)


# -- the document ---------------------------------------------------------------------------


def fonts(face: str) -> dict:
    return {"ascii": face, "hAnsi": face, "eastAsia": face, "cs": face}


def line(face: str, half_points: int) -> str:
    """``Hx``, a superscript, `` Hx ``, a subscript, `` Hx``: each script followed by a text
    object of its own, as ``make_script_probe.size_paragraph`` does."""
    text = {"rFonts": fonts(face), "sz": half_points, "szCs": half_points}
    runs = (wml.run("Hx ", **text) + wml.run(SCRIPT, **text, vertAlign="superscript")
            + wml.run(" Hx ", **text) + wml.run(SCRIPT, **text, vertAlign="subscript")
            + wml.run(" Hx", **text))
    return wml.paragraph(runs, mark=text, spacing={"before": 0, "after": 0, "line": 240, "lineRule": "auto"})


def beside(face: str, kind: str) -> str:
    """``tall``: 11 pt text, a 36 pt run, then the scripts at 11 pt; ``large``: 11 pt
    text and scripts at 24 pt."""
    text = {"rFonts": fonts(face), "sz": 22, "szCs": 22}
    script = dict(text) if kind == "tall" else {"rFonts": fonts(face), "sz": 48, "szCs": 48}
    runs = wml.run("Hx ", **text)
    if kind == "tall":
        runs += wml.run("H", rFonts=fonts(face), sz=72, szCs=72) + wml.run(" ", **text)
    runs += (wml.run(SCRIPT, **script, vertAlign="superscript") + wml.run(" Hx ", **text)
             + wml.run(SCRIPT, **script, vertAlign="subscript") + wml.run(" Hx", **text))
    return wml.paragraph(runs, mark=text, spacing={"before": 0, "after": 0, "line": 240, "lineRule": "auto"})


def cases() -> list[tuple[str, str, int]]:
    """``(kind, face, w:sz)`` for every paragraph in order: ``size`` lines, then
    ``tall`` and ``large`` lines (``w:sz`` 22)."""
    out = [("size", face.name, hp) for face in FACES for hp in SIZES]
    out += [(kind, name, 22) for name in BESIDE for kind in ("tall", "large")]
    return out


def build() -> bytes:
    body = "".join(line(face, hp) if kind == "size" else beside(face, kind) for kind, face, hp in cases())
    styles = wml.styles_part(
        [wml.style("paragraph", "Normal", default=True)],
        run_defaults={"rFonts": fonts("Calibri"), "sz": 22, "szCs": 22},
        paragraph_defaults={"spacing": {"before": 0, "after": 0, "line": 240, "lineRule": "auto"}},
    )
    settings = ("word/settings.xml", wml.SETTINGS_CONTENT_TYPE, wml.SETTINGS_REL,
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:settings xmlns:w="{probe_docx.W_NS}"><w:embedTrueTypeFonts/></w:settings>')
    table = ("word/fontTable.xml",
             "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
             "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
             font_table())
    package = probe_docx.package(body, styles=styles, extra_parts=(settings, table))
    return _add_fonts(package)


def font_table() -> str:
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    entries = "".join(
        f'<w:font w:name="{face.name}"><w:charset w:val="00"/><w:family w:val="swiss"/>'
        f'<w:pitch w:val="variable"/><w:embedRegular r:id="rIdF{index}" w:fontKey="{font_key(face)}"/></w:font>'
        for index, face in enumerate(FACES, start=1))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:fonts xmlns:w="{probe_docx.W_NS}" xmlns:r="{r_ns}">{entries}</w:fonts>')


def _add_fonts(package: bytes) -> bytes:
    """The package with every face as ``word/fonts/fontN.odttf``, related from the font
    table, and the ``odttf`` content type."""
    source = zipfile.ZipFile(io.BytesIO(package))
    parts = {name: source.read(name) for name in source.namelist()}
    types = parts["[Content_Types].xml"].decode()
    types = types.replace(
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>')
    parts["[Content_Types].xml"] = types.encode()
    rels = "".join(
        f'<Relationship Id="rIdF{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        f'relationships/font" Target="fonts/font{index}.odttf"/>' for index in range(1, len(FACES) + 1))
    parts["word/_rels/fontTable.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + rels + "</Relationships>").encode()
    for index, face in enumerate(FACES, start=1):
        parts[f"word/fonts/font{index}.odttf"] = obfuscate(font_bytes(face), font_key(face))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            archive.writestr(info, data)
    return buffer.getvalue()


if __name__ == "__main__":
    import sys
    from pathlib import Path

    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("script-offset.docx")
    target.write_bytes(build())
    print(target)

"""Vertical metrics of an installed face, by family name -- read locally, never shipped.

``docx2svg.vertical`` takes a face as four integers.  This finds those integers for a
family name the cascade resolved, the way Word finds the face: **by name, system copy
first**.  ROADMAP.md 2.4 measured that for Times New Roman Word lays out with the macOS
system copy's metrics although it embeds its own bundle's, so the search order is
``/System/Library/Fonts`` and ``/Library/Fonts`` before Word's ``DFonts``.

Needs ``fonttools`` (the ``measure`` extra).  Only the four integers leave this module;
they are facts about a face, and the file they came from is not read into anything that
is committed.
"""

from __future__ import annotations

import functools
import logging
from pathlib import Path

from docx2svg.vertical import FaceMetrics

logging.getLogger("fontTools").setLevel(logging.ERROR)

#: Families for which Word lays out with its *own bundle's* copy although macOS has one
#: of the same name.  Symbol: the bullet line of sample-resume.docx is 1.2251 em tall
#: (the bundle's SymbolMT, hhea 2059 + 450) and not 1.0 em (the system Symbol, 1436 +
#: 612), and the PDF embeds SymbolMT.  Times New Roman goes the other way (ROADMAP 2.4),
#: so "system copy first" is not a rule Word follows; it is two observations.
PREFER_BUNDLE = frozenset({"symbol"})
WORD_FONTS = Path("/Applications/Microsoft Word.app/Contents/Resources/DFonts")

FONT_DIRS = (
    Path("/System/Library/Fonts"),
    Path("/System/Library/Fonts/Supplemental"),
    Path("/Library/Fonts"),
    Path.home() / "Library" / "Fonts",
    WORD_FONTS,
)


@functools.lru_cache(maxsize=None)
def _index() -> dict[tuple[str, bool, bool], tuple[str, int]]:
    from fontTools.ttLib import TTCollection, TTFont

    index: dict[tuple[str, bool, bool], tuple[str, int]] = {}
    for directory in FONT_DIRS:
        if not directory.is_dir():
            continue
        for path in sorted(directory.iterdir()):
            suffix = path.suffix.lower()
            if suffix not in (".ttf", ".otf", ".ttc"):
                continue
            try:
                count = len(TTCollection(str(path), lazy=True).fonts) if suffix == ".ttc" else 1
            except Exception:
                continue
            for number in range(count):
                try:
                    font = TTFont(str(path), fontNumber=number, lazy=True)
                    names = font["name"]
                    style = font["OS/2"].fsSelection if "OS/2" in font else 0
                    head = font["head"].macStyle
                except Exception:
                    continue
                bold = bool(style & 0x20 or head & 1)
                italic = bool(style & 0x01 or head & 2)
                # The typographic family (nameID 16) only for a face whose typographic
                # subfamily (17) is one of the four styles Word asks for: Aptos-Light.ttf
                # is typographic family "Aptos", subfamily "Light", and taking it as
                # "Aptos" regular (it sorts before Aptos.ttf) measured the wrong face.
                # Its vertical metrics are Aptos's, so no baseline noticed; its advances
                # are 1-2% narrower (read_wrap_budget_probe.py found it).
                subfamily = {str(r.toUnicode()) for r in names.names if r.nameID == 17 and _decodable(r)}
                typographic = not subfamily or subfamily & {"Regular", "Bold", "Italic", "Bold Italic"}
                families = {
                    str(record.toUnicode()) for record in names.names
                    if (record.nameID == 1 or (record.nameID == 16 and typographic)) and _decodable(record)
                }
                for family in families:
                    # First found wins: system directories are searched first, except
                    # for the families Word takes from its own bundle.
                    key = (family.lower(), bold, italic)
                    if key[0] in PREFER_BUNDLE and directory == WORD_FONTS:
                        index[key] = (str(path), number)
                    else:
                        index.setdefault(key, (str(path), number))
    return index


def _decodable(record) -> bool:
    try:
        record.toUnicode()
        return True
    except Exception:
        return False


@functools.lru_cache(maxsize=None)
def metrics(family: str, bold: bool = False, italic: bool = False) -> FaceMetrics | None:
    """``hhea`` (or typo metrics under ``USE_TYPO_METRICS``) of the installed face."""
    from fontTools.ttLib import TTFont

    index = _index()
    found = (index.get((family.lower(), bold, italic))
             or index.get((family.lower(), bold, False))
             or index.get((family.lower(), False, False)))
    if found is None:
        return None
    return _face_metrics(TTFont(found[0], fontNumber=found[1], lazy=True))


def _face_metrics(font) -> FaceMetrics:
    """``hhea`` (typo metrics under ``USE_TYPO_METRICS``), and the ``OS/2`` script sizes."""
    upm = font["head"].unitsPerEm
    os2 = font["OS/2"] if "OS/2" in font else None
    scripts = (os2.ySuperscriptYSize, os2.ySubscriptYSize) if os2 is not None else (None, None)
    if os2 is not None and os2.fsSelection & 0x80:
        return FaceMetrics(upm, os2.sTypoAscender, -os2.sTypoDescender, os2.sTypoLineGap, *scripts)
    hhea = font["hhea"]
    return FaceMetrics(upm, hhea.ascent, -hhea.descent, hhea.lineGap, *scripts)


def embedded(docx: bytes) -> dict[tuple[str, bool, bool], FaceMetrics]:
    """The four integers of every face **embedded in the document** (``w:embedRegular``
    and its siblings in ``fontTable.xml``), keyed like :func:`metrics`.

    Word draws an embedded face it does not have installed: ``filesamples/sample1.docx``
    embeds Ubuntu, this machine has none, and the PDF draws ``Ubuntu-Regular``.  The
    part is an obfuscated TrueType file (ECMA-376 17.8.1): the first 32 bytes are XORed
    with the 16 bytes of ``w:fontKey``, read from the GUID's last hex pair to its first.
    It is read in memory and only the integers leave (the four vertical ones and the two
    ``OS/2`` script sizes); the file is not ours either.
    """
    import io
    import re
    import zipfile

    from fontTools.ttLib import TTFont

    package = zipfile.ZipFile(io.BytesIO(docx))
    names = set(package.namelist())
    if "word/fontTable.xml" not in names or "word/_rels/fontTable.xml.rels" not in names:
        return {}
    table = package.read("word/fontTable.xml").decode("utf-8")
    rels = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"',
                           package.read("word/_rels/fontTable.xml.rels").decode("utf-8")))
    rels.update({k: v for v, k in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"',
                                              package.read("word/_rels/fontTable.xml.rels").decode("utf-8"))})
    out: dict[tuple[str, bool, bool], FaceMetrics] = {}
    styles = {"Regular": (False, False), "Bold": (True, False), "Italic": (False, True),
              "BoldItalic": (True, True)}
    for name, body in re.findall(r'<w:font w:name="([^"]+)"(.*?)</w:font>', table, re.S):
        for style, rid, key in re.findall(r'<w:embed(\w+) r:id="([^"]+)" w:fontKey="\{([^}]+)\}"', body):
            target = rels.get(rid)
            if target is None or f"word/{target}" not in names:
                continue
            data = bytearray(package.read(f"word/{target}"))
            digits = key.replace("-", "")
            mask = [int(digits[i:i + 2], 16) for i in range(30, -1, -2)]
            for i in range(32):
                data[i] ^= mask[i % 16]
            found = _face_metrics(TTFont(io.BytesIO(bytes(data)), lazy=True))
            bold, italic = styles.get(style, (False, False))
            out[(name.lower(), bold, italic)] = found
    return out

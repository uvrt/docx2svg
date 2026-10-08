"""Advance widths and kern pairs of an installed face, by family name -- read locally.

The horizontal companion of ``face_metrics.py``, and found the same way: by name, in the
order Word finds a face (``face_metrics.FONT_DIRS``: the macOS system copy first, except
the families in ``face_metrics.PREFER_BUNDLE``).  ROADMAP.md 2.4 measured that for Times
New Roman Word lays out with the system copy; its advances are identical to the bundle's,
but its kern pairs are not.

Only numbers leave this module -- an advance per character asked for, a kern value per
pair -- and they leave only into recordings, which are measurements.  No font file is
read into anything that is committed.  Needs ``fonttools`` (the ``measure`` extra).
"""

from __future__ import annotations

import functools
import io
import logging
import re
import zipfile

logging.getLogger("fontTools").setLevel(logging.ERROR)


class FaceAdvances:
    """One face's ``hmtx`` advances (through its best ``cmap``) and legacy ``kern`` pairs."""

    def __init__(self, font) -> None:
        self.units_per_em = font["head"].unitsPerEm
        self._cmap = dict(font.getBestCmap() or {})
        # A symbol font (Wingdings, Symbol) maps its glyphs at U+F0xx in a (3, 0)
        # subtable, which is what a list label's w:lvlText names; getBestCmap skips it.
        for table in font["cmap"].tables:
            if (table.platformID, table.platEncID) == (3, 0):
                for code, name in table.cmap.items():
                    self._cmap.setdefault(code, name)
        self._hmtx = font["hmtx"]
        self._kern: dict = {}
        if "kern" in font:
            for table in font["kern"].kernTables:
                pairs = getattr(table, "kernTable", None) or {}
                for key, value in pairs.items():
                    self._kern.setdefault(key, value)

    def advance(self, char: str) -> int | None:
        name = self._cmap.get(ord(char))
        if name is None:
            return None
        return self._hmtx[name][0]

    def kern(self, left: str, right: str) -> int:
        a, b = self._cmap.get(ord(left)), self._cmap.get(ord(right))
        if a is None or b is None:
            return 0
        return self._kern.get((a, b), 0)


@functools.lru_cache(maxsize=None)
def advances(family: str, bold: bool = False, italic: bool = False) -> FaceAdvances | None:
    """The installed face Word lays ``family`` out with, or ``None``."""
    import face_metrics
    from fontTools.ttLib import TTFont

    index = face_metrics._index()
    found = (index.get((family.lower(), bold, italic))
             or index.get((family.lower(), bold, False))
             or index.get((family.lower(), False, False)))
    if found is None:
        return None
    return FaceAdvances(TTFont(found[0], fontNumber=found[1], lazy=True))


def embedded(docx: bytes) -> dict[tuple[str, bool, bool], FaceAdvances]:
    """Every face embedded in the document, keyed like :func:`advances`.

    The same de-obfuscation as ``face_metrics.embedded`` (ECMA-376 17.8.1), in memory.
    """
    from fontTools.ttLib import TTFont

    package = zipfile.ZipFile(io.BytesIO(docx))
    names = set(package.namelist())
    if "word/fontTable.xml" not in names or "word/_rels/fontTable.xml.rels" not in names:
        return {}
    table = package.read("word/fontTable.xml").decode("utf-8")
    rels_xml = package.read("word/_rels/fontTable.xml.rels").decode("utf-8")
    rels = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels_xml))
    rels.update({k: v for v, k in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels_xml)})
    styles = {"Regular": (False, False), "Bold": (True, False), "Italic": (False, True),
              "BoldItalic": (True, True)}
    out: dict[tuple[str, bool, bool], FaceAdvances] = {}
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
            bold, italic = styles.get(style, (False, False))
            out[(name.lower(), bold, italic)] = FaceAdvances(TTFont(io.BytesIO(bytes(data)), lazy=True))
    return out


class InstalledAdvances:
    """:class:`docx2svg.measure.Advances` from the installed faces, and from the faces
    embedded in the document where none is installed (which is when Word draws them)."""

    def __init__(self, docx: bytes | None = None) -> None:
        self.embedded = embedded(docx) if docx else {}

    def face(self, face: str, bold: bool, italic: bool) -> FaceAdvances | None:
        found = advances(face, bold, italic)
        if found is None:
            key = face.lower()
            found = (self.embedded.get((key, bold, italic)) or self.embedded.get((key, bold, False))
                     or self.embedded.get((key, False, False)))
        return found

    def advance(self, face, bold, italic, char):
        found = self.face(face, bold, italic)
        width = found.advance(char) if found is not None else None
        return None if width is None else (width, found.units_per_em)

    def kern(self, face, bold, italic, left, right):
        found = self.face(face, bold, italic)
        return None if found is None else found.kern(left, right)


class RecordingAdvances:
    """Answers like ``inner`` and remembers every answer, for a recording."""

    def __init__(self, inner, faces: dict | None = None) -> None:
        self.inner = inner
        self.faces: dict = faces if faces is not None else {}

    def _entry(self, face, bold, italic) -> dict:
        return self.faces.setdefault(f"{face}|{int(bold)}|{int(italic)}", {"upm": None, "advances": {}, "kern": {}})

    def advance(self, face, bold, italic, char):
        found = self.inner.advance(face, bold, italic, char)
        entry = self._entry(face, bold, italic)
        if found is not None:
            entry["upm"] = found[1]
            entry["advances"][char] = found[0]
        return found

    def kern(self, face, bold, italic, left, right):
        value = self.inner.kern(face, bold, italic, left, right)
        if value:
            self._entry(face, bold, italic)["kern"][left + right] = value
        return value


class RecordedAdvances:
    """:class:`docx2svg.measure.Advances` from a recording: an advance not recorded is
    unknown (``None``); a kern pair not recorded is zero."""

    def __init__(self, faces: dict) -> None:
        self.faces = faces

    def advance(self, face, bold, italic, char):
        entry = self.faces.get(f"{face}|{int(bold)}|{int(italic)}")
        if entry is None or char not in entry["advances"]:
            return None
        return entry["advances"][char], entry["upm"]

    def kern(self, face, bold, italic, left, right):
        entry = self.faces.get(f"{face}|{int(bold)}|{int(italic)}")
        if entry is None:
            return None
        return entry["kern"].get(left + right, 0)

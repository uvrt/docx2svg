"""The faces Word lays a document out with, read in place with the standard library.

Everything the layout needs from a face is a handful of integers: the four vertical ones
(:class:`~docx2svg.vertical.FaceMetrics`), an advance per character, the legacy ``kern``
table's pairs, and -- for drawing -- the ``OS/2`` script offsets and the underline and
strikeout geometry.  Until Phase 5 they came from ``tools/face_metrics.py`` and
``tools/face_advances.py``, which read the installed files with ``fontTools``; a
renderer whose output depends on a dev extra is not a renderer, so this module reads the
same tables with :mod:`struct` and nothing else.

**It finds a face the way those tools do, which is the way Word does** (ROADMAP.md 2.4,
"Which copy of a face Word lays out with"): by family name, the macOS system copy first
and Word's own bundle last -- except the families in :data:`PREFER_BUNDLE`, which Word
takes from its bundle -- with a typographic family (name ID 16) counted only for a face
whose typographic subfamily is one of the four styles Word asks for.  Then, where no
installed face answers, the faces the document itself embeds (ECMA-376 17.8.1), which
is when Word draws them.  ``tests/test_fonts.py`` holds this reader to every advance,
kern pair and metric integer the ``fontTools`` tools recorded.

**Office's cloud fonts too** (:data:`OFFICE_CLOUD_FONTS`): the faces Office for Mac
downloads on demand -- Aptos Display, Word 365's heading face, among them -- are searched
after every installed face, each family's folder read in place, as Word's bundle is.

**The locations, the order and the reader are shared** with pptx2svg, which reads the
same places in PowerPoint's order: :mod:`ooxml_common.fonts.office` has Word's bundle,
macOS's folders and the cloud cache, Word's measured order (``office.WORD``), the
standard-library table reader (:class:`ooxml_common.fonts.office.Face`) and resvg's CSS
matching (:func:`css_match`).  This module is Word's lookup over them -- by family and
style, the macOS copy first -- and what the layout and the rasteriser need on top.

**Nothing is copied.**  Files are read where they are installed, and only numbers leave
this module.  No font file enters the repository or any artefact it makes.

**Open substitutes, where Word's face is absent** (:data:`SUBSTITUTES`): a face neither
installed nor embedded is laid out and drawn with an installed open face whose advances
and line metrics were measured equal to it -- Carlito for Calibri, Liberation Sans, Serif
and Mono for Arial, Times New Roman and Courier New -- and every such use is recorded
(:attr:`InstalledFonts.substitutions`) and reported (``font-substituted``).  Never in
place of a face that is present: fidelity is measured against Word's own faces.  A face
with no measured clone (Aptos, Cambria, Calibri Light, Symbol, Wingdings...) is not
substituted unless the caller names one (``ConvertOptions.font_substitutes``), which is
then reported as approximate.
"""

from __future__ import annotations

import functools
import io
import os
import re
import struct
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from ooxml_common.fonts import office as _office
from ooxml_common.fonts.office import (  # noqa: F401  (re-exported: docx2svg.fonts' names)
    CSS_STRETCH,
    OFFICE_CLOUD_FONTS,
    WORD_FONTS,
    FontError,
    cloud_font_dirs,
    css_match,
)
from ooxml_common.fonts.office import decode_name as _decode_name  # noqa: F401
from ooxml_common.fonts.office import font_offsets as _offsets
from ooxml_common.fonts.office import name_records as _name_records
from ooxml_common.fonts.office import read_cmap as _cmap  # noqa: F401
from ooxml_common.fonts.office import read_legacy_kern as _kern  # noqa: F401
from ooxml_common.fonts.office import table_directory as _directory  # noqa: F401
from ooxml_common.fonts.office import weight_order as _weight_order  # noqa: F401

from .vertical import FaceMetrics

#: Families for which Word lays out with its *own bundle's* copy although macOS has one
#: of the same name (``tools/face_metrics.py`` records the evidence: the bullet line of
#: ``sample-resume.docx`` is SymbolMT's 1.2251 em, not the system Symbol's 1.0).  Word's
#: measured order, :data:`ooxml_common.fonts.office.WORD`, carries it.
PREFER_BUNDLE = _office.WORD.prefer_bundle

#: Where installed faces are looked for, in the order Word's choice was measured to follow,
#: then the faces Office has downloaded (:func:`cloud_font_dirs`): a face Word draws from
#: its cloud cache is found only where nothing installed answers to its name -- every face
#: there is one macOS and Word's bundle lack.  Word's layout order
#: (:func:`ooxml_common.fonts.office.font_dirs`): macOS's four folders, Word's bundle, the
#: cache's families.
FONT_DIRS: tuple[Path, ...] = _office.font_dirs(_office.WORD)

_FONT_SUFFIXES = _office.FONT_SUFFIXES


def _names(table: bytes) -> list[tuple[int, str]]:
    """``[(name ID, text)]`` of every decodable record."""
    return [(name_id, text) for _platform, _encoding, _language, name_id, text in _name_records(table)]


@dataclass(frozen=True)
class Decorations:
    """What drawing a run needs from its face beyond its metrics, in font units.

    ``OS/2``'s script offsets (``ySuperscriptYOffset`` positive up, ``ySubscriptYOffset``
    positive *down*, as the spec defines them), its strikeout, and ``post``'s underline.
    """

    units_per_em: int
    superscript_offset: int | None = None
    subscript_offset: int | None = None
    strikeout_position: int | None = None
    strikeout_size: int | None = None
    underline_position: int | None = None
    underline_thickness: int | None = None
    #: The ``hhea`` ascent and descent (positive), whatever the layout takes (typo
    #: metrics under ``USE_TYPO_METRICS``).
    hhea_ascent: int = 0
    hhea_descent: int = 0

    def integers(self) -> list:
        return [self.units_per_em, self.superscript_offset, self.subscript_offset, self.strikeout_position,
                self.strikeout_size, self.underline_position, self.underline_thickness, self.hhea_ascent,
                self.hhea_descent]


class Face(_office.Face):
    """One font, read lazily (:class:`ooxml_common.fonts.office.Face`): with what the
    layout's vertical model and the drawing need of it besides."""

    @functools.cached_property
    def metrics(self) -> FaceMetrics:
        """``hhea`` (typo metrics under ``USE_TYPO_METRICS``) and the ``OS/2`` script sizes."""
        os2 = self.table(b"OS/2")
        scripts: tuple = (None, None, None, None)
        if os2 is not None and len(os2) >= 26:
            # ySuperscriptYSize (offset 20), ySubscriptYSize (12), ySuperscriptYOffset (24)
            # and ySubscriptYOffset (16).
            scripts = tuple(struct.unpack_from(">h", os2, offset)[0] for offset in (20, 12, 24, 16))
        if os2 is not None and len(os2) >= 74 and struct.unpack_from(">H", os2, 62)[0] & 0x80:
            ascender, descender, gap = struct.unpack_from(">hhh", os2, 68)
            return FaceMetrics(self.units_per_em, ascender, -descender, gap, *scripts)
        ascender, descender, gap = struct.unpack_from(">hhh", self.table(b"hhea"), 4)
        return FaceMetrics(self.units_per_em, ascender, -descender, gap, *scripts)

    @functools.cached_property
    def decorations(self) -> Decorations:
        os2 = self.table(b"OS/2")
        post = self.table(b"post")
        hhea = self.table(b"hhea")
        ascender, descender = struct.unpack_from(">hh", hhea, 4)
        values: dict = {}
        if os2 is not None and len(os2) >= 30:
            values["subscript_offset"] = struct.unpack_from(">h", os2, 16)[0]
            values["superscript_offset"] = struct.unpack_from(">h", os2, 24)[0]
            values["strikeout_size"], values["strikeout_position"] = struct.unpack_from(">hh", os2, 26)
        if post is not None and len(post) >= 12:
            values["underline_position"], values["underline_thickness"] = struct.unpack_from(">hh", post, 8)
        return Decorations(self.units_per_em, hhea_ascent=ascender, hhea_descent=-descender, **values)

    @functools.cached_property
    def _gpos_kerning(self):
        from .outlines import PairKerning

        table = self.table(b"GPOS")
        try:
            found = PairKerning(table) if table else None
        except (struct.error, IndexError):
            found = None
        return found if found is not None and found.subtables else None

    def pair_kern(self, left: str, right: str) -> int:
        """The kerning between two characters a watermark's letters are set with, font
        units: the ``GPOS`` ``kern`` feature's pair adjustment where the face has one, else
        its legacy ``kern`` table's (:meth:`kern`) -- both measured in Word's watermarks
        (:data:`docx2svg.drawing.TEXT_PATH_INSETS`: Calibri's, Arial's and Aptos's agree;
        Minion Pro, CFF, has only ``GPOS``)."""
        gpos = self._gpos_kerning
        if gpos is None:
            return self.kern(left, right)
        a, b = self.glyph(left), self.glyph(right)
        return 0 if a is None or b is None else gpos.pair(a, b)

    def pens(self, text: str, *, kerned: bool = False) -> list[int] | None:
        """Where each character of ``text`` is set from pen 0, font units: at its advances,
        with :meth:`pair_kern` between neighbours when ``kerned``; ``None`` for a character
        the face lacks."""
        pens, pen = [], 0
        for i, char in enumerate(text):
            advance = self.advance(char)
            if advance is None or self.glyph(char) is None:
                return None
            pens.append(pen)
            pen += advance + (self.pair_kern(char, text[i + 1]) if kerned and i + 1 < len(text) else 0)
        return pens

    @functools.cached_property
    def _cff(self):
        from .outlines import CFF

        table = self.table(b"CFF ")
        try:
            return CFF(table) if table else None
        except (struct.error, IndexError, ValueError, KeyError):
            return None

    def glyph_box(self, glyph: int, slant: float = 0.0) -> tuple | None:
        """The box round glyph ``glyph``'s outline (:mod:`docx2svg.outlines`), font units, y
        up, slanted by ``slant`` (``x + slant * y``); ``()`` for an empty glyph, ``None`` for
        an outline not read (or a face with neither ``glyf`` nor ``CFF ``).  An upright
        ``glyf`` glyph's box is its header's."""
        from . import outlines

        cache = self.__dict__.setdefault("_glyph_boxes", {})
        key = (glyph, slant)
        if key in cache:
            return cache[key]
        found = None
        head, loca, glyf = self.table(b"head"), self.table(b"loca"), self.table(b"glyf")
        if head is not None and loca is not None and glyf is not None:
            long_offsets = struct.unpack_from(">h", head, 50)[0] == 1
            if slant:
                found = outlines.glyf_box(glyf, loca, long_offsets, glyph, slant)
            else:
                span = outlines.glyph_range(loca, long_offsets, glyph)
                if span is not None and span[0] == span[1]:
                    found = ()
                elif span is not None and span[0] + 10 <= span[1] <= len(glyf):
                    found = struct.unpack_from(">hhhh", glyf, span[0] + 2)
        elif self._cff is not None:
            found = self._cff.box(glyph, slant)
        cache[key] = found
        return found

    def ink(self, text: str, *, kerned: bool = False, slant: float = 0.0,
            pens: list | None = None) -> tuple | None:
        """``(x min, y min, x max, y max)``, font units, y up from the baseline: the box
        round the outlines of ``text`` set at :meth:`pens` (or the ``pens`` given) from pen
        0, each glyph's outline box (:meth:`glyph_box`, slanted by ``slant``) -- what a VML
        text path that fits its shape is stretched by (:func:`docx2svg.drawing.text_path_markup`).
        ``None`` for a character the face lacks, an outline not read, or no outline at all."""
        if pens is None:
            pens = self.pens(text, kerned=kerned)
        if pens is None:
            return None
        box: list | None = None
        for char, pen in zip(text, pens):
            glyph = self.glyph(char)
            found = self.glyph_box(glyph, slant)
            if found is None:
                return None
            if not found:
                continue  # an empty glyph (a space)
            found = (pen + found[0], found[1], pen + found[2], found[3])
            box = list(found) if box is None else [min(box[0], found[0]), min(box[1], found[1]),
                                                   max(box[2], found[2]), max(box[3], found[3])]
        return tuple(box) if box is not None and box[2] > box[0] and box[3] > box[1] else None


_read = _office.read_file


def _faces_in(path: Path) -> list[tuple[int, Face]]:
    try:
        data = _read(str(path))
        return [(number, Face(data, offset, source=str(path))) for number, offset in enumerate(_offsets(data))]
    except (OSError, struct.error, FontError):
        return []


#: The tables a folder scan reads of each face: its names, its style and its CSS keys
#: (:meth:`Face.families`, :attr:`Face.style`, :attr:`Face.css`,
#: :attr:`Face.rasteriser_families`).
_SCAN_TABLES = (b"name", b"OS/2", b"head")


class _ScannedFace(Face):
    """A face read for a folder scan: its table directory and :data:`_SCAN_TABLES` only,
    read from the file at their offsets -- a few kilobytes -- not the whole file.  Any other
    table is read from the file when first asked for; :attr:`data` reads the whole file.
    The same answers as a :class:`Face` over the file's bytes: a table past the file's end
    is cut short in both."""

    def __init__(self, path: str, offset: int, tables: dict, read: dict) -> None:
        self.source = path
        self.offset = offset
        self._tables = tables
        self._read = read

    @property
    def data(self) -> bytes:
        return _read(self.source)

    def table(self, tag: bytes) -> bytes | None:
        found = self._tables.get(tag)
        if found is None:
            return None
        if tag not in self._read:
            with open(self.source, "rb") as handle:
                self._read[tag] = _read_at(handle, *found)
        return self._read[tag]


def _read_at(handle, offset: int, length: int) -> bytes:
    handle.seek(offset)
    return handle.read(length)


def _scanned_faces(path: Path) -> list[tuple[int, Face]]:
    """:func:`_faces_in` for a folder scan: the same faces in the same order, failing the
    same way (none, for a file whose header or a face's table directory cannot be read),
    from each face's directory and :data:`_SCAN_TABLES` alone.  Reading every installed
    file whole -- 1.4 GB on a Mac with Office -- was most of a process's first layout, and
    kept the bytes in memory (:func:`ooxml_common.fonts.office.read_file` caches them)."""
    source = str(path)
    try:
        with open(source, "rb") as handle:
            head = handle.read(12)
            if head[:4] == b"ttcf":
                count = struct.unpack_from(">I", head, 8)[0]
                offsets = list(struct.unpack_from(f">{count}I", head + handle.read(4 * count), 12))
            else:
                offsets = [0]
            faces = []
            for number, offset in enumerate(offsets):
                header = _read_at(handle, offset, 12)
                if len(header) < 12:
                    raise FontError("too short for an offset table")
                tables_count = struct.unpack_from(">H", header, 4)[0]
                directory = _read_at(handle, offset + 12, 16 * tables_count)
                tables: dict[bytes, tuple[int, int]] = {}
                for i in range(tables_count):
                    tables[directory[16 * i:16 * i + 4]] = struct.unpack_from(">II", directory, 16 * i + 8)
                read = {tag: _read_at(handle, *tables[tag]) for tag in _SCAN_TABLES if tag in tables}
                faces.append((number, _ScannedFace(source, offset, tables, read)))
            return faces
    except (OSError, struct.error, FontError):
        return []


#: :func:`_scanned_entries`' answers, by path, with the file's ``(st_mtime_ns, st_size)``
#: before it was read (``None`` if it could not be stated): kept for the process, as the
#: index is.
_SCANNED: dict[str, tuple] = {}


def _scanned_entries(path: str) -> tuple:
    """``(number in the file, (bold, italic, family names) or None, rasteriser families or
    None)`` of every face of ``path`` a scan can read (:func:`_scanned_faces`), ``None``
    where reading that part of the face fails: what :func:`installed_index` (both orders)
    and :func:`_rasteriser_families_index` need, read once per process -- or taken from a
    :class:`FontScan` (:func:`load_scan`)."""
    found = _SCANNED.get(path)
    if found is not None:
        return found[1]
    try:
        stat = os.stat(path)
        stamp: tuple | None = (stat.st_mtime_ns, stat.st_size)
    except OSError:
        stamp = None
    out = []
    for number, face in _scanned_faces(Path(path)):
        try:
            bold, italic = face.style
            named: tuple | None = (bold, italic, tuple(face.families()))
        except (OSError, struct.error, TypeError, FontError):
            named = None
        try:
            rasteriser: tuple | None = tuple(face.rasteriser_families)
        except (OSError, struct.error, TypeError, FontError):
            rasteriser = None
        out.append((number, named, rasteriser))
    return _SCANNED.setdefault(path, (stamp, tuple(out)))[1]


#: What a :class:`FontScan` holds per file; a scan of another format is not loaded.
SCAN_FORMAT = 1


@dataclass(frozen=True)
class FontScan:
    """What a process read of the font folders' files to find faces by name: per file,
    its modification time and size and what each face in it is called, in which style --
    a few hundred bytes a file.  Picklable: a render worker saves the one
    :func:`docx2svg.warm` returns and hands it to the next worker's, which then reads only
    the files that changed (:func:`load_scan`).

    Keyed by each file's ``(st_mtime_ns, st_size)``
    (:func:`ooxml_common.fonts.rasterdb.file_stamps`): a file changed, replaced or removed
    is read afresh, a file added is read as it would be; the folders themselves are listed
    by every process, so the scan never decides which files there are."""

    #: path -> ``(st_mtime_ns, st_size, entries)`` (:func:`_scanned_entries`).
    files: dict
    #: The docx2svg version that read it: a scan from another version is not loaded.
    version: str = ""
    format: int = SCAN_FORMAT


def _listed(dirs) -> list[str]:
    """The font files :func:`installed_index` lists in ``dirs``, each once."""
    out: dict[str, None] = {}
    for directory in dirs:
        if not directory.is_dir():
            continue
        for path in sorted(directory.iterdir()):
            if path.suffix.lower() in _FONT_SUFFIXES:
                out.setdefault(str(path), None)
    return list(out)


def font_scan(dirs, version: str = "") -> FontScan:
    """The :class:`FontScan` of every font file in ``dirs``, from what this process has
    read (reading what it has not)."""
    files = {}
    for path in _listed(dirs):
        _scanned_entries(path)
        stamp, entries = _SCANNED[path]
        if stamp is not None:
            files[path] = (*stamp, entries)
    return FontScan(files, version)


def load_scan(scan: FontScan, version: str = "") -> int:
    """Take what ``scan`` read of each file whose modification time and size are still
    those it recorded, so this process does not read it; the number of files taken.  A
    scan of another version or format, or not a :class:`FontScan`, is not taken (0)."""
    from ooxml_common.fonts.rasterdb import file_stamps

    if not isinstance(scan, FontScan) or scan.format != SCAN_FORMAT or scan.version != version:
        return 0
    taken = 0
    for path, mtime, size in file_stamps(font_files=list(scan.files)):
        recorded = scan.files[path]
        if (recorded[0], recorded[1]) == (mtime, size):
            _SCANNED.setdefault(path, ((mtime, size), recorded[2]))
            taken += 1
    return taken


@functools.lru_cache(maxsize=None)
def installed_index(dirs: tuple[Path, ...] = FONT_DIRS, *,
                    bundle_first: bool = False) -> dict[tuple[str, bool, bool], tuple[str, int]]:
    """``(family lowercased, bold, italic) -> (path, number in the file)``, first found
    winning, the directories searched in order -- :data:`PREFER_BUNDLE` families from
    Word's bundle (every family, with ``bundle_first``: :func:`drawing_dirs`)."""
    index: dict[tuple[str, bool, bool], tuple[str, int]] = {}
    for directory in dirs:
        if not directory.is_dir():
            continue
        for path in sorted(directory.iterdir()):
            if path.suffix.lower() not in _FONT_SUFFIXES:
                continue
            for number, named, _ in _scanned_entries(str(path)):
                if named is None:
                    continue
                bold, italic, families = named
                for family in families:
                    key = (family.lower(), bold, italic)
                    if key[0] in PREFER_BUNDLE and directory == WORD_FONTS and not bundle_first:
                        index[key] = (str(path), number)
                    else:
                        index.setdefault(key, (str(path), number))
    return index


def drawing_dirs(dirs: tuple[Path, ...] = FONT_DIRS) -> tuple[Path, ...]:
    """The same directories, Word's bundle first: the copy Word *draws* a face with.

    Word lays a face out with the macOS system copy (ROADMAP.md 2.4) but embeds -- draws
    -- its own bundle's where it has one: the fonts embedded in its PDFs are Times New
    Roman 7.00, Arial 6.80 and Verdana 5.02 (the bundle's; the system's are 5.01), and
    Courier New and Georgia 5.00 (the system's: the bundle has none).  Their advances
    and vertical metrics agree; their outlines and kerning need not.  A rasteriser that
    should draw what Word draws is given these files.
    """
    return tuple(d for d in dirs if d == WORD_FONTS) + tuple(d for d in dirs if d != WORD_FONTS)


@functools.lru_cache(maxsize=None)
def _face_at(path: str, number: int) -> Face:
    data = _read(path)
    return Face(data, _offsets(data)[number], source=path)


@functools.lru_cache(maxsize=None)
def _rasteriser_families_index(dirs: tuple[Path, ...]) -> dict[str, list[tuple[str, int]]]:
    """Every installed face (as :func:`installed_index` finds them) by each family name
    a rasteriser files it under (:attr:`Face.rasteriser_families`), lowercased."""
    out: dict[str, list[tuple[str, int]]] = {}
    for path, number in sorted(set(installed_index(dirs, bundle_first=True).values())):
        names = next((found for at, _, found in _scanned_entries(path) if at == number), None)
        if names is None:
            continue
        for name in {name.lower() for name in names}:
            out.setdefault(name, []).append((path, number))
    return out


def _rasteriser_family(dirs: tuple[Path, ...], family: str) -> list[Face]:
    return [_face_at(path, number) for path, number in _rasteriser_families_index(dirs).get(family, [])]


def _lookup(table: dict, family: str, bold: bool, italic: bool):
    key = family.lower()
    return table.get((key, bold, italic)) or table.get((key, bold, False)) or table.get((key, False, False))


#: Word's face (lowercased) -> the open face laid out and drawn in its place when Word's is
#: neither installed nor embedded.  Only metric compatible pairs, each measured against
#: the copy Word lays out with (``installed_index`` over Word's folders), regular, bold,
#: italic and bold italic, by this module's own reader:
#:
#: * **Calibri -> Carlito** (1.1): the same advance for every character both map but
#:   three (U+0192, U+026A, U+0299: 2094 compared), and the same vertical metrics --
#:   2048 units, ascent 1950, descent 550, no line gap, the same script sizes and offsets.
#: * **Arial -> Liberation Sans, Times New Roman -> Liberation Serif, Courier New ->
#:   Liberation Mono** (2.1.5): every Latin, Greek and Cyrillic advance equal (2209 and
#:   2118 compared; the differences are in Hebrew, historic Cyrillic and, in Mono, the
#:   combining marks), the same vertical metrics, and the same legacy ``kern`` pairs
#:   wherever sampled.
#:
#: What differs is drawn only: outlines, and the underline and strikeout geometry.  And
#: Carlito has no legacy ``kern`` table where Calibri has one, so a run Word kerns
#: (``w:kern``) is laid out unkerned in it.  Not here, because measured *not* to be
#: compatible: Caladea for Cambria (advances 1.8 to 5.7% narrow, line 1.150 em against
#: 1.172), Carlito for Calibri Light (1.2 to 3.4% wide).  Aptos, Symbol and Wingdings have
#: no open clone.  A caller may name any of them (``ConvertOptions.font_substitutes``):
#: the substitution is then reported as approximate.
SUBSTITUTES: dict[str, str] = {
    "calibri": "Carlito",
    "arial": "Liberation Sans",
    "times new roman": "Liberation Serif",
    "courier new": "Liberation Mono",
}

#: The faces an absent face with no substitute of its own is laid out with, by its class
#: (:func:`face_class`), first installed winning: the open faces the renderers already
#: measure and draw with (the substitutes above, pptx2svg's Arimo, Tinos and Cousine,
#: Carlito and Caladea), then the Office and macOS faces they stand in for.  None of them
#: has the absent face's metrics: every use is reported as approximate, naming both.
GENERIC_FALLBACKS: dict[str, tuple[str, ...]] = {
    "sans-serif": ("Liberation Sans", "Arimo", "Arial", "Helvetica", "Carlito", "Calibri"),
    "serif": ("Liberation Serif", "Tinos", "Times New Roman", "Times", "Caladea", "Cambria"),
    "monospace": ("Liberation Mono", "Cousine", "Courier New", "Courier"),
}

#: ``w:font/w:family`` in ``word/fontTable.xml`` (ECMA-376 17.8.3.9, ST_FontFamily) -> the
#: class :data:`GENERIC_FALLBACKS` is keyed by; ``modern`` is fixed pitch.
_FONT_TABLE_CLASS = {"roman": "serif", "swiss": "sans-serif", "modern": "monospace"}


def face_class(family: str, font_table: dict | None = None) -> str:
    """``sans-serif``, ``serif`` or ``monospace`` for a face: what the document's font
    table says of it (``w:family``, ``w:pitch="fixed"``), else what its name says
    (:func:`ooxml_common.text.fontmap.generic_family`, the rule pptx2svg's SVG names its
    fallback by)."""
    from ooxml_common.text.fontmap import generic_family

    entry = (font_table or {}).get(family.lower())
    if entry is not None:
        kind, pitch = entry
        if pitch == "fixed":
            return "monospace"
        if kind in _FONT_TABLE_CLASS:
            return _FONT_TABLE_CLASS[kind]
    return generic_family(family)


def default_substitutes(family: str, font_table: dict | None = None) -> list[tuple[str, str]]:
    """The faces an absent ``family`` is laid out with when neither the caller nor
    :data:`SUBSTITUTES` names one, best first, each with its kind:

    * ``default`` -- the face pptx2svg measures and draws it with
      (:func:`ooxml_common.text.fontmap.substitution_for`: Carlito for Calibri Light and
      Aptos, Caladea for Cambria...), then the Office face that one is the measured clone
      of here (Calibri for Carlito, :data:`SUBSTITUTES`): the rule PowerPoint rendering
      already uses, shared rather than copied;
    * ``generic`` -- :data:`GENERIC_FALLBACKS` for its class (:func:`face_class`)."""
    from ooxml_common.text.fontmap import SUBSTITUTIONS, substitution_for

    out: list[tuple[str, str]] = []
    row = substitution_for(family)
    if row is not None and row.substitute.lower() != family.lower():
        out.append((row.substitute, "default"))
        out += [(other.office, "default") for other in SUBSTITUTIONS.values()
                if other.substitute == row.substitute and SUBSTITUTES.get(other.office.lower()) == row.substitute]
    out += [(name, "generic") for name in GENERIC_FALLBACKS[face_class(family, font_table)]]
    seen: set[str] = {family.lower()}
    kept = []
    for name, kind in out:
        if name.lower() not in seen:
            seen.add(name.lower())
            kept.append((name, kind))
    return kept


def font_table_classes(package_bytes: bytes | None) -> dict[str, tuple[str | None, str | None]]:
    """``word/fontTable.xml``'s ``w:family`` and ``w:pitch`` per face (lowercased name)."""
    if not package_bytes:
        return {}
    try:
        with zipfile.ZipFile(io.BytesIO(package_bytes)) as archive:
            table = archive.read("word/fontTable.xml").decode("utf-8", "replace")
    except (KeyError, zipfile.BadZipFile, OSError):
        return {}
    out: dict[str, tuple[str | None, str | None]] = {}
    for name, body in re.findall(r'<w:font w:name="([^"]+)"[^>]*?(?:/>|>(.*?)</w:font>)', table, re.S):
        kind = re.search(r'<w:family w:val="([^"]+)"', body)
        pitch = re.search(r'<w:pitch w:val="([^"]+)"', body)
        out[name.lower()] = (kind.group(1) if kind else None, pitch.group(1) if pitch else None)
    return out


#: What :attr:`Substitution.substitute` says for a face laid out from :mod:`docx2svg.recorded`.
RECORDED = "recorded metrics"


class RecordedFace:
    """A symbol face this machine lacks, answered from the integers recorded from Word's
    copy (:mod:`docx2svg.recorded`): what the layout asks of a :class:`Face`, and nothing
    a rasteriser could draw with."""

    def __init__(self, family: str) -> None:
        from . import recorded

        key = family.lower()
        self.family = family
        self.source = f"recorded:{family}"
        self.offset = 0
        self.style = (False, False)
        self.metrics = FaceMetrics(*recorded.METRICS[key])
        self.units_per_em = self.metrics.units_per_em
        self.decorations = Decorations(*recorded.DECORATIONS[key])
        self._advances = recorded.ADVANCES[key]
        #: What each character is drawn as (:data:`docx2svg.recorded.UNICODE`).
        self.unicode = {chr(code): text for code, text in recorded.UNICODE.get(key, {}).items()}

    def advance(self, char: str) -> int | None:
        return self._advances.get(ord(char)) if len(char) == 1 else None

    def kern(self, left: str, right: str) -> int:
        return 0  # neither recorded face has a ``kern`` table


def _recorded(family: str) -> RecordedFace | None:
    from . import recorded

    return RecordedFace(family) if family.lower() in recorded.METRICS else None


@dataclass(frozen=True)
class Substitution:
    """A face laid out with another because it is absent: what was asked for, what was
    used, and whether the two were measured metric compatible (:data:`SUBSTITUTES`)."""

    family: str
    substitute: str
    metric_compatible: bool
    #: How the substitute was chosen: ``named`` by the caller (``font_substitutes``, or
    #: :data:`SUBSTITUTES`), ``default`` (:func:`default_substitutes`: the face PowerPoint
    #: rendering uses for it) or ``generic`` (the first installed face of its class).
    kind: str = "named"
    #: For a ``generic`` one, the class it was chosen by (:func:`face_class`).
    face_class: str | None = None

    def message(self) -> str:
        if self.substitute == RECORDED:
            return (f"{self.family!r} is not installed: laid out with the metrics recorded from Word's copy "
                    "(docx2svg.recorded: the same advances and line metrics), and its symbols drawn as their "
                    "Unicode equivalents in a generic face")
        if self.metric_compatible:
            return (f"{self.family!r} is not installed: laid out and drawn with {self.substitute!r}, its open "
                    "metric compatible substitute (the same advances and line metrics; outlines, underline "
                    "geometry and kerning may differ)")
        if self.kind in ("default", "generic"):
            why = ("the face PowerPoint rendering uses for it" if self.kind == "default"
                   else f"the first installed {self.face_class or 'sans-serif'} face")
            return (f"{self.family!r} is not installed and no substitute was named: laid out and drawn with "
                    f"{self.substitute!r}, {why}, as an approximation (its advances and line metrics are not "
                    "the absent face's: line breaks and page breaks may differ; name a substitute in "
                    "ConvertOptions.font_substitutes, or install the face, to change it)")
        return (f"{self.family!r} is not installed: laid out and drawn with {self.substitute!r}, an approximate "
                "substitute the caller named (its advances and line metrics are not Word's: line breaks and "
                "page breaks may differ)")

    def as_dict(self) -> dict:
        return {"family": self.family, "substitute": self.substitute, "metric_compatible": self.metric_compatible,
                "kind": self.kind}


def embedded_faces(package_bytes: bytes) -> dict[tuple[str, bool, bool], Face]:
    """Every face the document embeds (``w:embedRegular`` and its siblings in
    ``word/fontTable.xml``), keyed like :func:`installed_index`.

    The part is an obfuscated font (ECMA-376 17.8.1): its first 32 bytes are XORed with
    the 16 bytes of ``w:fontKey``, read from the GUID's last hex pair to its first.  It
    is undone in memory.
    """
    try:
        archive = zipfile.ZipFile(io.BytesIO(package_bytes))
    except zipfile.BadZipFile:
        return {}
    names = set(archive.namelist())
    if "word/fontTable.xml" not in names or "word/_rels/fontTable.xml.rels" not in names:
        return {}
    table = archive.read("word/fontTable.xml").decode("utf-8")
    rels_xml = archive.read("word/_rels/fontTable.xml.rels").decode("utf-8")
    rels = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels_xml))
    rels.update({k: v for v, k in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels_xml)})
    styles = {"Regular": (False, False), "Bold": (True, False), "Italic": (False, True), "BoldItalic": (True, True)}
    out: dict[tuple[str, bool, bool], Face] = {}
    for name, body in re.findall(r'<w:font w:name="([^"]+)"(.*?)</w:font>', table, re.S):
        for style, rid, key in re.findall(r'<w:embed(\w+) r:id="([^"]+)" w:fontKey="\{([^}]+)\}"', body):
            target = rels.get(rid)
            if target is None or f"word/{target}" not in names:
                continue
            data = bytearray(archive.read(f"word/{target}"))
            digits = key.replace("-", "")
            mask = [int(digits[i:i + 2], 16) for i in range(30, -1, -2)]
            for i in range(min(32, len(data))):
                data[i] ^= mask[i % 16]
            try:
                face = Face(bytes(data), source=f"embedded:{name}:{style}")
                face.units_per_em  # noqa: B018 -- fail here, not later
            except (struct.error, FontError, TypeError):
                continue
            bold, italic = styles.get(style, (False, False))
            out[(name.lower(), bold, italic)] = face
    return out


@dataclass
class InstalledFonts:
    """Metrics, advances, kern pairs and decorations from the faces Word lays out with:
    installed faces by name (:func:`installed_index`), then the document's embedded ones.

    It answers the :class:`docx2svg.measure.Advances` protocol, and ``metrics(face, bold,
    italic)`` as the vertical model asks.  Where a face is neither installed nor embedded,
    its open substitute (``substitutes``: :data:`SUBSTITUTES` by default, ``{}`` for none)
    when that is installed, recorded in :attr:`substitutions`; else ``None``, recorded in
    :attr:`missing`, and the layout then reports the paragraph unmeasurable rather than
    guess.
    """

    package: bytes | None = None
    dirs: tuple[Path, ...] = FONT_DIRS
    #: Word's face (any case) -> the face used in its place when it is absent.
    substitutes: dict = field(default_factory=lambda: dict(SUBSTITUTES))
    #: Lay an absent symbol face out from :mod:`docx2svg.recorded` (Symbol, Wingdings).
    recorded_faces: bool = True
    _embedded: dict = field(default_factory=dict, repr=False)
    #: Every face laid out with a substitute, by the name the document gives it.
    substitutions: dict = field(default_factory=dict, repr=False)
    #: Every face asked for that was found nowhere and had no installed substitute.
    missing: set = field(default_factory=set, repr=False)
    #: Lay an absent face with no substitute named out with :func:`default_substitutes`'
    #: first installed face, as an approximation (``False``: it is missing).
    fallbacks: bool = True
    #: ``(face, character)`` -> the face its advance was taken from where the face lacks
    #: the character (:meth:`advance`), or ``.notdef`` (the face's missing glyph).
    glyph_fallbacks: dict = field(default_factory=dict, repr=False)
    _defaults: dict = field(default_factory=dict, repr=False)
    _glyphs: dict = field(default_factory=dict, repr=False)
    _font_table: dict | None = field(default=None, repr=False)

    def __post_init__(self) -> None:
        self._embedded = embedded_faces(self.package) if self.package else {}
        self.substitutes = {name.lower(): value for name, value in self.substitutes.items()}

    def face(self, family: str, bold: bool = False, italic: bool = False) -> Face | None:
        return self._find(installed_index(self.dirs), family, bold, italic, recorded=True)

    def drawing_face(self, family: str, bold: bool = False, italic: bool = False, *,
                     record: bool = True) -> Face | None:
        """The copy of ``family`` Word draws with (:func:`drawing_dirs`): its bundle's where
        it has one, then the installed ones, then the document's embedded face, then its
        substitute.  ``record=False`` leaves :attr:`substitutions` and :attr:`missing` as
        they are (a face asked for by the caller, not by the document)."""
        return self._find(installed_index(drawing_dirs(self.dirs), bundle_first=True), family, bold, italic,
                          record=record)

    def _find(self, index: dict, family: str, bold: bool, italic: bool, *, record: bool = True,
              recorded: bool = False) -> Face | None:
        if not family:
            return None
        found = _lookup(index, family, bold, italic)
        if found is not None:
            return _face_at(*found)
        embedded = _lookup(self._embedded, family, bold, italic)
        if embedded is not None:
            return embedded
        # Word's face is absent, in every style: the substitute named for it -- followed
        # through the names (a caller's Calibri Light -> Calibri is Carlito where Calibri
        # is absent too) -- if one is installed.
        found, substitution = self._named(index, family, bold, italic)
        if found is None:
            # A symbol face: laid out from the integers recorded from Word's copy, and drawn
            # by :meth:`drawing_name`'s Unicode equivalents (no file to draw it with).
            face = _recorded(family) if self.recorded_faces else None
            if face is not None:
                if not recorded:
                    return None  # drawing: no file (:meth:`drawing_name` names a generic face)
                if record and family not in self.substitutions:
                    self.substitutions[family] = Substitution(family, RECORDED, True)
                return face
            if self.fallbacks:
                # No substitute named and none installed: the face PowerPoint rendering
                # uses for it, else the first installed face of its class -- approximate,
                # and said so, rather than a document that stops at its first paragraph.
                found, substitution = self._default(index, family, bold, italic)
        if found is None:
            if record:
                self.missing.add(family)
            return None
        if record and family not in self.substitutions:
            self.substitutions[family] = substitution
        return _face_at(*found)

    def _named(self, index: dict, family: str, bold: bool, italic: bool):
        """The first installed face of the chain of substitutes named for ``family``."""
        seen = {family.lower()}
        name = self.substitutes.get(family.lower())
        while name and name.lower() not in seen:
            found = _lookup(index, name, bold, italic)
            if found is not None:
                compatible = SUBSTITUTES.get(family.lower(), "").lower() == name.lower()
                return found, Substitution(family, name, compatible)
            seen.add(name.lower())
            name = self.substitutes.get(name.lower())
        return None, None

    def _default(self, index: dict, family: str, bold: bool, italic: bool):
        """:func:`default_substitutes`' first installed face for ``family`` (or the
        substitute named for it), cached."""
        key = (id(index), family.lower(), bold, italic)
        if key not in self._defaults:
            result = None, None
            if self._font_table is None:
                self._font_table = font_table_classes(self.package)
            for name, kind in default_substitutes(family, self._font_table):
                found = _lookup(index, name, bold, italic)
                used = name
                if found is None:
                    found, named = self._named(index, name, bold, italic)
                    used = named.substitute if named is not None else name
                if found is not None:
                    result = found, Substitution(family, used, False, kind,
                                                 face_class(family, self._font_table) if kind == "generic" else None)
                    break
            self._defaults[key] = result
        return self._defaults[key]

    def drawing_name(self, family: str, bold: bool = False, italic: bool = False) -> DrawingName | None:
        """How the SVG names ``family`` so that a rasteriser finds the file Word draws it
        with (:class:`DrawingName`); ``None`` where the document's name and ``bold`` /
        ``italic`` already do -- the face answers to that name and has that style --
        and for a face this machine lacks."""
        face = self.drawing_face(family, bold, italic)
        if face is None:
            if family in self.substitutions and self.substitutions[family].substitute == RECORDED:
                return DrawingName("sans-serif", 700 if bold else 400, "normal", "italic" if italic else "normal",
                                   chars=RecordedFace(family).unicode)
            return None
        families = face.rasteriser_families
        if face.style == (bold, italic) and family.lower() in (f.lower() for f in families):
            return None
        weight_class, width, style = face.css
        pool = [face] + [other for other in _rasteriser_family(drawing_dirs(self.dirs), families[0].lower())
                         if (other.source, other.offset) != (face.source, face.offset)]
        pool += [other for other in self._embedded.values() if other is not face
                 and families[0].lower() in (f.lower() for f in other.rasteriser_families)]
        # The multiples of 100 nearest the face's own weight first (a tie to the lighter).
        candidates = sorted(range(100, 1000, 100), key=lambda w: (abs(w - weight_class), w))
        weight = next((w for w in candidates if css_match(pool, width, style, w) == [face]), candidates[0])
        return DrawingName(families[0], weight, CSS_STRETCH[width - 1], style)

    def metrics(self, family: str, bold: bool = False, italic: bool = False) -> FaceMetrics | None:
        face = self.face(family, bold, italic)
        return None if face is None else face.metrics

    def decorations(self, family: str, bold: bool = False, italic: bool = False) -> Decorations | None:
        face = self.face(family, bold, italic)
        return None if face is None else face.decorations

    def advance(self, family: str, bold: bool, italic: bool, char: str) -> tuple[int, int] | None:
        face = self.face(family, bold, italic)
        width = face.advance(char) if face is not None else None
        if width is None and face is not None and self.fallbacks and len(char) == 1:
            key = (family, bold, italic, char)
            if key not in self._glyphs:
                self._glyphs[key] = self._glyph_fallback(face, family, bold, italic, char)
            return self._glyphs[key]
        return None if width is None else (width, face.units_per_em)

    def _glyph_fallback(self, face, family: str, bold: bool, italic: bool, char: str) -> tuple[int, int] | None:
        """The advance of a character ``family`` lacks, from the first installed face that
        has it -- Times New Roman first, the face Word draws Hebrew in where the run's face
        has none (ROADMAP.md, "Where Word is not ECMA-376", cases f09, f13, f24), then
        :func:`default_substitutes`' faces and every class's :data:`GENERIC_FALLBACKS` --
        else the face's own missing glyph (``.notdef``).  Not measured: Word's font
        fallback for other characters is not known.  Recorded in :attr:`glyph_fallbacks`
        (``font-glyph-fallback``) rather than stopping the paragraph."""
        index = installed_index(self.dirs)
        names = ["Times New Roman"] + [name for name, _ in default_substitutes(family, self._font_table or {})]
        names += [name for group in GENERIC_FALLBACKS.values() for name in group]
        for name in dict.fromkeys(names):
            found = _lookup(index, name, bold, italic)
            if found is None:
                found, named = self._named(index, name, bold, italic)
                name = named.substitute if named is not None else name
            if found is None:
                continue
            other = _face_at(*found)
            width = other.advance(char)
            if width is not None:
                self.glyph_fallbacks.setdefault((family, char), name)
                return width, other.units_per_em
        advances = getattr(face, "advances", None)
        if not advances:
            return None
        self.glyph_fallbacks.setdefault((family, char), ".notdef")
        return advances[0], face.units_per_em

    def kern(self, family: str, bold: bool, italic: bool, left: str, right: str) -> int | None:
        face = self.face(family, bold, italic)
        return None if face is None else face.kern(left, right)

    def ink(self, family: str, bold: bool, italic: bool, text: str, oblique: float = 0.0) -> tuple | None:
        """``(units per em, (x min, y min, x max, y max), pens, slant)`` of ``text``'s
        outlines in the face it is drawn with (:meth:`drawing_face`, :meth:`Face.ink`),
        kerned (:meth:`Face.pair_kern`) -- a watermark's letters, as Word sets them -- or
        ``None``.  ``slant`` is ``oblique`` where italic is asked of a family with no italic
        face (the upright one is slanted), else 0."""
        face = self.drawing_face(family, bold, italic, record=False)
        if not isinstance(face, Face):
            return None
        slant = oblique if italic and not face.style[1] else 0.0
        pens = face.pens(text, kerned=True)
        box = face.ink(text, slant=slant, pens=pens) if pens is not None else None
        return None if box is None else (face.units_per_em, box, pens, slant)


_STRETCH_CLASS = {name: index + 1 for index, name in enumerate(CSS_STRETCH)}


@dataclass(frozen=True)
class DrawingName:
    """How an SVG names a face the document's own name does not reach in a rasteriser:
    ``font-family="<the document's name>, <family>"`` with this weight, stretch and style.

    resvg's font database files a face under its typographic family (name ID 16) and
    never then under name ID 1, so Calibri Light's file (``calibril.ttf``: "Calibri",
    "Light", weight 300) answers ``font-family="Calibri"`` at ``font-weight="300"`` and
    nothing at "Calibri Light".  The list keeps the document's name first -- a browser
    matching full names still finds it -- and falls through to the family the file is
    filed under.  ``weight`` is a multiple of 100 because resvg reads no other
    (``font-weight="350"`` is drawn as ``normal``, measured): the one nearest the face's
    ``usWeightClass`` that CSS font matching resolves to this face among the installed
    faces of its family (:func:`css_match`).
    """

    family: str
    weight: int
    stretch: str
    style: str
    #: Characters drawn as others: a symbol face laid out from recorded metrics
    #: (:class:`RecordedFace`) is drawn as its symbols' Unicode equivalents.
    chars: dict | None = field(default=None, compare=False)


def _query(face_name: str, name: "DrawingName | None", bold: bool, italic: bool):
    """``(families, usWidthClass, style, weight)`` the SVG asks for (:mod:`docx2svg.svg`)."""
    if name is None:
        return [face_name], 5, "italic" if italic else "normal", 700 if bold else 400
    families = [face_name] + ([name.family] if name.family.lower() != face_name.lower() else [])
    return families, _STRETCH_CLASS[name.stretch], name.style, name.weight


def resolves_to(face: Face, pool: list, face_name: str, name: "DrawingName | None", bold: bool,
                italic: bool) -> bool:
    """Whether a rasteriser filing ``pool`` as resvg's database does draws ``face`` for
    what the SVG asks: the first family of the list that any face answers to, then
    :func:`css_match` there, to one face alone."""
    families, stretch, style, weight = _query(face_name, name, bold, italic)
    for family in families:
        members = [other for other in pool if family.lower() in (f.lower() for f in other.rasteriser_families)]
        if members:
            found = css_match(members, stretch, style, weight)
            return len(found) == 1 and found[0] is face
    return False


def _is_microsoft(face: Face) -> bool:
    """A face whose copyright, trademark or manufacturer names Microsoft: never written
    to a file by this library, in any form."""
    return any("microsoft" in text.lower() for name_id, text in face.names if name_id in (0, 7, 8))


def rasteriser_files(fonts: "InstalledFonts", faces, directory: str | None = None,
                     ) -> tuple[list[str], list[str], list[str]]:
    """The font files a rasteriser needs to draw ``faces`` (``(family, bold, italic)``)
    as Word draws them: ``(files, missing, unaddressable)``.

    **Every installed face is handed over in place**, as the file Word draws it with
    (:meth:`InstalledFonts.drawing_face`); nothing is copied, relabelled or written.
    Where the document's name does not reach the face in the rasteriser (a superfamily
    member: Calibri Light), the SVG names it by a fallback list instead
    (:class:`DrawingName`), which finds the original file.

    A face the document **embeds** exists nowhere but inside the document, and resvg
    reads fonts only from files: it is written, de-obfuscated and otherwise byte for
    byte as embedded, into ``directory`` for as long as the caller keeps it -- and not
    at all without a ``directory``, or when it is a Microsoft face (then it counts as
    missing).  ``missing`` lists the faces this machine cannot supply; ``unaddressable``
    those the SVG's names would not resolve to their own file among ``files``, because
    another face there answers the same family, stretch, style and weight (Avenir's Book
    and Roman are both 400) -- the rasteriser draws one of the two.
    """

    files: list[str] = []
    missing: list[str] = []
    written: dict[str, Face] = {}
    drawn: list[tuple[str, int, str, bool, bool]] = []
    for index, (family, bold, italic) in enumerate(sorted(faces)):
        face = fonts.drawing_face(family, bold, italic)
        if face is None:
            missing.append(family)
            continue
        if face.source.startswith("embedded:"):
            if directory is None or _is_microsoft(face):
                missing.append(family)
                continue
            path = os.path.join(directory, f"embedded-{index}.ttf")
            with open(path, "wb") as handle:
                handle.write(face.data)
            written[path] = Face(face.data, source=path)
        else:
            path = face.source
        if path not in files:
            files.append(path)
        drawn.append((path, face.offset, family, bold, italic))
    pool = [face for path in files
            for face in ([written[path]] if path in written else [f for _, f in _faces_in(Path(path))])]
    unaddressable = []
    for path, offset, family, bold, italic in drawn:
        own = next((other for other in pool if other.source == path and other.offset == offset), None)
        name = fonts.drawing_name(family, bold, italic)
        if own is None or not (resolves_to(own, pool, family, name, bold, italic)
                               or _drawn_apart(own, pool, family, name, bold, italic)):
            unaddressable.append(family)
    return files, missing, unaddressable


def _drawn_apart(face: Face, pool: list, face_name: str, name: "DrawingName | None", bold: bool,
                 italic: bool) -> bool:
    """Whether :func:`drawing_plan` draws ``face`` for what the SVG asks although another
    face of ``pool`` ties with it: ``face`` is the first of the tie whose family (name ID 1)
    is the one the document names, and the faces of other families are other files --
    Aptos Display and Aptos, both filed as "Aptos" at 400 -- so a pass without those draws
    it."""
    families, stretch, style, weight = _query(face_name, name, bold, italic)
    for family in families:
        members = [other for other in pool if family.lower() in (f.lower() for f in other.rasteriser_families)]
        if members:
            found = css_match(members, stretch, style, weight)
            named = [other for other in found if face_name.lower() in _legacy_families(other)]
            return bool(named) and named[0] is face and all(
                other.source != face.source for other in found if other not in named)
    return False


def _legacy_families(face) -> frozenset:
    """A face's family names (name ID 1), lowercased: "Aptos Display" for the face resvg
    files as "Aptos"."""
    try:
        return frozenset(text.lower() for name_id, text in face.names if name_id == 1)
    except (struct.error, TypeError, FontError):
        return frozenset()


# --------------------------------------------------------------------------------------
# Faces filed under one name: drawn in passes of their own
# --------------------------------------------------------------------------------------


@dataclass(frozen=True)
class DrawingPlan:
    """How :func:`docx2svg.png.svg_to_png` draws an SVG whose text asks for two faces the
    rasteriser files under one name (:func:`drawing_plan`): in passes, each over the last.

    ``stacks`` gives each ``font-family`` value drawn after the first pass its pass (the
    rest are drawn in the first); ``withheld`` per pass, the files of those handed over
    that are left out of it -- the faces that would otherwise answer for that pass's own."""

    stacks: dict
    withheld: list

    @property
    def passes(self) -> int:
        return len(self.withheld)


@dataclass(frozen=True)
class _PoolFace:
    """What a plan reads of a face handed to the rasteriser."""

    source: str
    css: tuple
    rasteriser_families: tuple
    legacy: frozenset


@functools.lru_cache(maxsize=4096)
def _pool_faces(path: str, mtime: int, size: int) -> tuple:
    out = []
    for _, face in _scanned_faces(Path(path)):
        try:
            out.append(_PoolFace(path, face.css, tuple(name.lower() for name in face.rasteriser_families),
                                 _legacy_families(face)))
        except (struct.error, TypeError, FontError):
            continue
    return tuple(out)


@functools.lru_cache(maxsize=64)
def _pool(stamps: tuple) -> tuple:
    """The faces of the files ``stamps`` names, in load order, by each name the rasteriser
    files them under; ``None`` where no two faces of different files tie (one name, one
    weight, width and style): nothing to plan, the common case."""
    faces = [face for stamp in dict.fromkeys(stamps) for face in _pool_faces(*stamp)]
    seen: dict = {}
    tie = False
    for face in faces:
        for name in face.rasteriser_families:
            other = seen.setdefault((name, face.css), face.source)
            tie = tie or other != face.source
    if not tie:
        return None
    by_name: dict = {}
    for face in faces:
        for name in dict.fromkeys(face.rasteriser_families):
            by_name.setdefault(name, []).append(face)
    return tuple((name, tuple(members)) for name, members in by_name.items())


_TEXT_TAG = re.compile(r"<text\b[^>]*>")
_SVG_ATTRIBUTE = re.compile(r'\s(font-family|font-weight|font-stretch|font-style)="([^"]*)"')
_GENERIC = frozenset({"serif", "sans-serif", "monospace", "cursive", "fantasy"})


def _text_queries(svg: str) -> dict:
    """``font-family`` value -> the ``(families, usWidthClass, style, weight)`` its texts
    ask for, as :mod:`docx2svg.svg` writes them (on the ``<text>`` element)."""
    out: dict = {}
    for tag in _TEXT_TAG.findall(svg):
        attributes = dict(_SVG_ATTRIBUTE.findall(tag))
        value = attributes.get("font-family")
        if value is None:
            continue
        families = [name.strip().strip("'\"").replace("\\'", "'").replace("&apos;", "'").replace("&amp;", "&")
                    for name in value.split(",")]
        families = tuple(name for name in families if name and name.lower() not in _GENERIC)
        if not families:
            continue
        weight = attributes.get("font-weight", "normal")
        weight = 700 if weight == "bold" else 400 if weight == "normal" else int(weight) if weight.isdigit() else 400
        query = (families, _STRETCH_CLASS.get(attributes.get("font-stretch", "normal"), 5),
                 attributes.get("font-style", "normal"), weight)
        out.setdefault(value, set()).add(query)
    return out


def drawing_plan(svg: str, font_files=(), font_dirs=()) -> DrawingPlan | None:
    """The passes ``svg`` is drawn in with ``font_files`` and ``font_dirs`` (the files a
    rasteriser is handed, :func:`rasteriser_files` and the application's folders), or
    ``None``: one pass, as given, draws every text in its own face.

    resvg files a face under its typographic family (name ID 16) alone, so Aptos Display
    is "Aptos" at 400 as Aptos is, and of two faces that tie the first loaded draws both
    (:func:`css_match`).  A text whose face loses such a tie to a face of another family
    -- its face is the first of the tie whose family (name ID 1) is the name the document
    gives first, and the faces of other families are other files -- is drawn in a pass of its own without those files, over the passes
    before it (:func:`docx2svg.png.pass_svg`); a pass may also only leave files out (a
    folder holding both faces, the document using one).  A tie no file left out undoes
    (two faces of one file) stays as it is.
    """
    from ooxml_common.fonts.rasterdb import file_stamps

    if not font_files and not font_dirs or "<text" not in svg:
        return None
    pool = _pool(tuple(file_stamps([str(path) for path in font_files or ()],
                                   [str(path) for path in font_dirs or ()])))
    if pool is None:
        return None
    queries = _text_queries(svg)
    if not queries:
        return None
    by_name = dict(pool)

    def drawn(query, withheld) -> "_PoolFace | None":
        families, stretch, style, weight = query
        for family in families:
            members = [face for face in by_name.get(family.lower(), ()) if face.source not in withheld]
            if members:
                return css_match(members, stretch, style, weight)[0]      # a tie: the first loaded
        return None

    def wanted(query) -> "tuple[_PoolFace | None, frozenset]":
        """The face the text is to be drawn in, and the files to leave out for it."""
        families, stretch, style, weight = query
        for family in families:
            members = by_name.get(family.lower())
            if members:
                found = css_match(list(members), stretch, style, weight)
                named = [face for face in found if families[0].lower() in face.legacy]
                rivals = frozenset(face.source for face in found if face not in named)
                if named and found[0] not in named and named[0].source not in rivals:
                    return named[0], rivals
                return found[0], frozenset()
        return None, frozenset()

    needs = {value: [(query, *wanted(query)) for query in sorted(found, key=repr)]
             for value, found in queries.items()}
    if not any(rivals for wants in needs.values() for _, _, rivals in wants):
        return None
    passes: list[tuple[list, set]] = [([], set())]      # (values, files left out)

    def holds(values, withheld) -> bool:
        return all(drawn(query, withheld) is face for value in values for query, face, _ in needs[value])

    for value, wants in needs.items():
        rivals = set().union(*(rivals for _, _, rivals in wants))
        for values, withheld in passes:
            if holds(values + [value], withheld | rivals):
                values.append(value)
                withheld.update(rivals)
                break
        else:
            if holds([value], rivals):
                passes.append(([value], set(rivals)))
            else:
                passes[0][0].append(value)              # as it was: no pass draws it apart
    if len(passes) == 1 and not passes[0][1]:
        return None
    return DrawingPlan(stacks={value: number for number, (values, _) in enumerate(passes) if number
                               for value in values},
                       withheld=[frozenset(withheld) for _, withheld in passes])


def default_font_dirs() -> tuple[Path, ...]:
    """:data:`FONT_DIRS` on macOS; elsewhere the usual places, searched in the same spirit
    (system first: :func:`ooxml_common.fonts.office.system_font_dirs`).  Only macOS has
    been measured against Word."""
    if sys.platform == "darwin":
        return FONT_DIRS
    return _office.system_font_dirs()

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


_read = _office.read_file


def _faces_in(path: Path) -> list[tuple[int, Face]]:
    try:
        data = _read(str(path))
        return [(number, Face(data, offset, source=str(path))) for number, offset in enumerate(_offsets(data))]
    except (OSError, struct.error, FontError):
        return []


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
            for number, face in _faces_in(path):
                try:
                    bold, italic = face.style
                    families = face.families()
                except (struct.error, TypeError):
                    continue
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
        try:
            names = _face_at(path, number).rasteriser_families
        except (struct.error, TypeError, FontError):
            continue
        for name in {name.lower() for name in names}:
            out.setdefault(name, []).append((path, number))
    return out


def _rasteriser_family(dirs: tuple[Path, ...], family: str) -> list[Face]:
    return [_face_at(path, number) for path, number in _rasteriser_families_index(dirs).get(family, [])]


def _lookup(table: dict, family: str, bold: bool, italic: bool):
    key = family.lower()
    return table.get((key, bold, italic)) or table.get((key, bold, False)) or table.get((key, False, False))


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
    italic)`` as the vertical model asks.  ``None`` for a face it cannot find: the layout
    then reports the paragraph unmeasurable rather than guess.
    """

    package: bytes | None = None
    dirs: tuple[Path, ...] = FONT_DIRS
    _embedded: dict = field(default_factory=dict, repr=False)

    def __post_init__(self) -> None:
        self._embedded = embedded_faces(self.package) if self.package else {}

    def face(self, family: str, bold: bool = False, italic: bool = False) -> Face | None:
        if not family:
            return None
        found = _lookup(installed_index(self.dirs), family, bold, italic)
        if found is not None:
            return _face_at(*found)
        return _lookup(self._embedded, family, bold, italic)

    def drawing_face(self, family: str, bold: bool = False, italic: bool = False) -> Face | None:
        """The copy of ``family`` Word draws with (:func:`drawing_dirs`): its bundle's where
        it has one, then the installed ones, then the document's embedded face."""
        if not family:
            return None
        found = _lookup(installed_index(drawing_dirs(self.dirs), bundle_first=True), family, bold, italic)
        if found is not None:
            return _face_at(*found)
        return _lookup(self._embedded, family, bold, italic)

    def drawing_name(self, family: str, bold: bool = False, italic: bool = False) -> DrawingName | None:
        """How the SVG names ``family`` so that a rasteriser finds the file Word draws it
        with (:class:`DrawingName`); ``None`` where the document's name and ``bold`` /
        ``italic`` already do -- the face answers to that name and has that style --
        and for a face this machine lacks."""
        face = self.drawing_face(family, bold, italic)
        if face is None:
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
        return None if width is None else (width, face.units_per_em)

    def kern(self, family: str, bold: bool, italic: bool, left: str, right: str) -> int | None:
        face = self.face(family, bold, italic)
        return None if face is None else face.kern(left, right)


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
    import os

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
        if own is None or not resolves_to(own, pool, family, fonts.drawing_name(family, bold, italic), bold, italic):
            unaddressable.append(family)
    return files, missing, unaddressable


def default_font_dirs() -> tuple[Path, ...]:
    """:data:`FONT_DIRS` on macOS; elsewhere the usual places, searched in the same spirit
    (system first: :func:`ooxml_common.fonts.office.system_font_dirs`).  Only macOS has
    been measured against Word."""
    if sys.platform == "darwin":
        return FONT_DIRS
    return _office.system_font_dirs()

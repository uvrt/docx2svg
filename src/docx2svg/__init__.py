"""docx2svg -- render Word documents to SVG, in pure Python.

Typical use::

    from docx2svg import convert_docx_to_svg, convert_docx_to_png

    svgs = convert_docx_to_svg("report.docx")          # one SVG string per page
    pngs = convert_docx_to_png("report.docx")
    both = convert_docx("report.docx")                  # the layout and the SVGs, laid out once

Every drawn element's ``data-docx-path`` is its layout object's ``path``, and
:func:`docx2svg.paths.resolve_path` finds the element it names in the part's XML.

The pipeline is ``.docx -> OPC package -> source model -> resolved properties -> lines ->
pages -> placed glyphs -> SVG -> PNG``: :mod:`docx2svg.opc` and :mod:`docx2svg.parse` read
the file; :mod:`docx2svg.resolve` applies the style cascade; :mod:`docx2svg.linebreak`,
:mod:`docx2svg.vertical` and :mod:`docx2svg.paginate` lay it out, each measured against
Word's own export (ROADMAP.md, Phases 2-4); :mod:`docx2svg.layout` places every glyph
and :mod:`docx2svg.svg` transcribes it; :mod:`docx2svg.png` rasterises.

**What is drawn, and what is not.**  Text in its resolved face, size, weight, slant and
colour, underline and strikethrough, highlight and shading, list labels, paragraph
borders, inline pictures, superscripts and subscripts.  Tables, headers and
footers, and page-number fields (``PAGE``, ``NUMPAGES``, ``SECTIONPAGES``, computed from
the layout's own pages, and in the body ``PAGEREF``, ``REF`` and ``SEQ`` from its pages and
bookmarks; ``STYLEREF`` per page in a header or footer, and in the body; every other field is drawn as Word cached it).  Floating
drawings that text does not wrap around (behind or in front of the text): pictures,
shapes, groups and text boxes, positioned and stacked as Word does, and those text wraps
above and below (``wrapTopAndBottom``), and those text wraps beside.  A drop cap is drawn
in its frame, the text beside it.  Endnotes are numbered and drawn after the text, under
their separators, and footnotes at the foot of their reference's page
(:mod:`docx2svg.notes`), a table cell's too; one in a cell of a section of several text
columns is not drawn (``footnotes-not-drawn``).  VML is read (:mod:`docx2svg.parse.vml`):
a watermark's text path, an embedded object's picture, rectangles, ovals and text boxes.

**The layout carries on** past what its measured rules do not reach (round 16,
:data:`docx2svg.paginate.CARRY_ON`): it lays it out by an approximate rule
(``layout-approximate:<reason>``) or draws it as a marked placeholder of its stated or
estimated height (``layout-placeholder:<reason>``), and every page after it is laid out;
``ConvertOptions.warnings`` and ``ConvertOptions.coverage`` say what, where and why.  With
``CARRY_ON`` off the layout **stops** at the first such item, as it did before: the page it
is on shows a marked band where it starts, no page after it is invented
(``layout-stopped:<reason>``), and a header or footer holding one is drawn up to it
(``story-stopped:<reason>``).

The runtime is standard-library only.  Faces are read in place from where Word finds
them (:mod:`docx2svg.fonts`); PNG output needs the ``png`` extra (resvg-py).
"""

from __future__ import annotations

import os
import tempfile
from contextlib import contextmanager, nullcontext
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Callable, Sequence

from .model import Document, PageMargins, PageSize, Paragraph, Run, Section
from .opc import Package
from .paginate import PageInfo
from .parse import parse_document, parse_package
from .parse.document import KeptBlocks
from .paths import resolve_path
from .png import RasterizerNotAvailable, available_backends, svg_to_png
from ooxml_common.locales import EN_US, Locale
from ooxml_common.opc import OpcPackage

if TYPE_CHECKING:
    from .coverage import Coverage
    from .fonts import Decorations, DrawingName
    from .layout import Layout
    from .measure import Advances
    from .vertical import FaceMetrics

__version__ = "0.11.10"

__all__ = [
    "Conversion",
    "ConvertOptions",
    "Document",
    "PageInfo",
    "PageMargins",
    "PageSize",
    "Package",
    "Paragraph",
    "RasterizerNotAvailable",
    "Run",
    "Section",
    "Warning",
    "available_backends",
    "convert_docx",
    "convert_docx_to_layout",
    "convert_docx_to_png",
    "convert_docx_to_svg",
    "embed_glyphs",
    "glyph_sheet",
    "lay_out_docx",
    "page_svgs",
    "parse_document",
    "parsed_document",
    "parse_package",
    "resolve_path",
    "scan_fonts",
    "svg_to_png",
    "user_font_dirs",
    "warm",
    "__version__",
]


@dataclass(frozen=True)
class Warning:  # noqa: A001 -- the sibling project's name for the same thing
    """Something the output does not show faithfully, with a stable ``code`` a caller can
    fail a build on (``layout-placeholder:drawing``, ``layout-approximate:table``,
    ``footnotes-not-drawn``...)."""

    code: str
    message: str
    page: int | None = None

    def __str__(self) -> str:
        where = f" (page {self.page})" if self.page else ""
        return f"[{self.code}]{where} {self.message}"


@dataclass
class ConvertOptions:
    """Knobs shared by the SVG and PNG entry points."""

    #: 1-based page numbers; ``None`` renders every page the layout reaches.
    pages: Sequence[int] | None = None
    #: Output size in pixels.  For PNG the page is drawn at 300 dpi when both are
    #: ``None``; for SVG they replace the ``width``/``height`` attributes (in points by
    #: default).  Give one and the aspect ratio is kept.
    width: int | None = None
    height: int | None = None
    #: What glyph outlines are scaled to: ``device`` (the size rounded to whole 1/300-inch
    #: pixels, as Word draws its ink) or ``exact``.  Pen positions are the model's either
    #: way.  See :data:`docx2svg.svg.GLYPH_SIZES` and ROADMAP.md, "Phase 5 -- measured".
    glyph_size: str = "device"
    #: The application's own font folders (and the folders under them), searched after
    #: the ones Word uses, for layout and drawing alike.  ``None`` (the default) reads the
    #: ``OOXML_FONT_DIRS`` environment variable (``os.pathsep``-separated, shared with
    #: pptx2svg); an empty list means none, the variable not read.
    font_dirs: Sequence[str] | None = None
    #: Override measurement: an :class:`~docx2svg.measure.Advances`, ``metrics(face,
    #: bold, italic)`` and ``decorations(face, bold, italic)``.  By default all three come
    #: from :class:`docx2svg.fonts.InstalledFonts` over the document.
    advances: Advances | None = None
    metrics: Callable[[str, bool, bool], FaceMetrics | None] | None = None
    decorations: Callable[[str, bool, bool], Decorations | None] | None = None
    #: ``names(face, bold, italic) -> DrawingName | None``: how the SVG names a face its
    #: document name does not reach in a rasteriser.  By default
    #: :meth:`docx2svg.fonts.InstalledFonts.drawing_name`.
    names: Callable[[str, bool, bool], DrawingName | None] | None = None
    #: Lay out and draw an absent face with its open metric compatible substitute
    #: (:data:`docx2svg.fonts.SUBSTITUTES`: Carlito for Calibri, Liberation Sans, Serif
    #: and Mono for Arial, Times New Roman and Courier New) where that is installed.
    #: Never in place of a face that is present; every use is a ``font-substituted``
    #: warning and is listed in :attr:`coverage`.  An absent face with no measured
    #: substitute is laid out with an approximate one (:func:`docx2svg.fonts.default_substitutes`),
    #: and a character a face lacks is measured in another face (``font-glyph-fallback``).
    #: ``False``: none of these; an absent face's paragraphs are placeholders.
    substitute_fonts: bool = True
    #: More substitutions, or other ones: Word's face -> the face to use when it is absent
    #: (``{"Cambria": "Caladea"}``).  One that is not in :data:`docx2svg.fonts.SUBSTITUTES`
    #: was not measured compatible, and is reported as approximate.
    font_substitutes: dict | None = None
    #: The regional settings Word would take from its machine, which a document does not
    #: state (:class:`ooxml_common.locales.Locale`): the character a decimal tab stop
    #: aligns on (Word's is the system's decimal separator, measured, not the document's
    #: language) and the decimal and thousands separators of a chart's numbers.  Always
    #: :data:`~ooxml_common.locales.EN_US` unless given -- never read from the operating
    #: system, so a document draws the same on every server.  Frozen and picklable: it
    #: travels with the options into a worker process.
    locale: Locale = EN_US
    #: How text is written: ``"text"`` (the default) as ``<text>``, drawn by the viewer
    #: with its own fonts; ``"outlines"`` as the paths of the glyphs the PNG is drawn with
    #: -- the same files, passes, fallbacks and shaping resvg uses for
    #: :func:`convert_docx_to_png` (:func:`docx2svg.png.outline_svg`) -- so the SVG needs no
    #: font to be viewed.  Glyph shapes go into the SVG, never font files.  A text that
    #: cannot be drawn so (right-to-left or complex-script text, a colour face) stays text,
    #: with a ``text-not-outlined`` warning.  PNG output is the same either way.
    text: str = "text"
    #: With ``text="outlines"``: each text also kept, transparent, over its outlines, so a
    #: viewer can select, search and copy it (and a screen reader read it).
    selectable_text: bool = False
    #: With ``text="outlines"``: one glyph set for the whole document.  A dict, and every
    #: glyph definition goes into it, by id, instead of into each page's ``<defs>``; the
    #: pages only reference them (``<use href="#g…">``).  An id is made of the glyph's
    #: outline, so it is the same on every page and in every render, and the dict, kept
    #: across renders, gains only the glyphs an edit added.  Put :func:`glyph_sheet` of it
    #: once in the HTML page that shows the pages inline; :func:`embed_glyphs` makes one
    #: page standalone again (to save it, or to hand it to :func:`svg_to_png`).  ``None``
    #: (the default): each page holds its own definitions.
    glyphs: dict | None = None
    #: Collects :class:`Warning` for everything not drawn faithfully.
    warnings: list[Warning] = field(default_factory=list)
    #: Set by every conversion: how much of the document was laid out
    #: (:class:`docx2svg.coverage.Coverage`), also the layout's ``coverage``.
    coverage: Coverage | None = None


def _read(source) -> bytes:
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if isinstance(source, (str, os.PathLike)):
        with open(os.fspath(source), "rb") as handle:
            return handle.read()
    return source.read()


def user_font_dirs(options: ConvertOptions | None = None) -> list[str]:
    """The application's font folders a conversion with ``options`` uses, as strings:
    ``options.font_dirs``, else ``OOXML_FONT_DIRS``
    (:func:`ooxml_common.fonts.office.user_font_dirs`) -- what :func:`convert_docx_to_png`
    hands the rasteriser after the faces it names, and what a caller rasterising
    :func:`page_svgs` itself hands it."""
    return [str(path) for path in _user_dirs(options or ConvertOptions())]


def _user_dirs(options: ConvertOptions, *, subfolders: bool = False) -> tuple:
    """The application's font folders: ``options.font_dirs``, else ``OOXML_FONT_DIRS``
    (:func:`ooxml_common.fonts.office.user_font_dirs`); with ``subfolders``, the folders
    under them too, for the one-level index layout reads."""
    from ooxml_common.fonts.office import user_font_dirs, with_subfolders

    dirs = user_font_dirs(options.font_dirs)
    return with_subfolders(dirs) if subfolders else dirs


def _font_dirs(options: ConvertOptions) -> tuple:
    """The folders a layout with ``options`` finds faces in, in Word's order."""
    from .fonts import default_font_dirs

    return tuple(default_font_dirs()) + _user_dirs(options, subfolders=True)


def _fonts(data: bytes, options: ConvertOptions):
    from .fonts import SUBSTITUTES, InstalledFonts

    dirs = _font_dirs(options)
    substitutes = dict(SUBSTITUTES) if options.substitute_fonts else {}
    substitutes.update({name.lower(): value for name, value in (options.font_substitutes or {}).items()})
    return InstalledFonts(data, dirs=dirs, substitutes=substitutes, recorded_faces=options.substitute_fonts,
                          fallbacks=options.substitute_fonts)


_PARSER_WARNINGS = {
    "styles-part-missing": "the package has no styles part; Word's application defaults apply",
    "document-no-body": "the document has no body",
}


def convert_docx_to_layout(source, options: ConvertOptions | None = None):
    """Parse and lay out a document without writing SVG: a :class:`docx2svg.layout.Layout`
    with every glyph, rule and picture of every page placed."""
    return lay_out_docx(source, options, keep=False)[0]


def lay_out_docx(source, options: ConvertOptions | None = None, *, previous=None, changed_blocks=None,
                 keep: bool = True, measures=None, changed_pages: bool = True) -> tuple:
    """``(layout, data, fonts)``: the document laid out -- every page, whatever
    ``options.pages`` selects -- its bytes, and the faces it was measured with
    (:class:`docx2svg.fonts.InstalledFonts`, ``None`` where ``options`` measures with
    its own ``advances``, ``metrics`` and ``decorations``).  Warnings and coverage go to
    ``options`` as by every conversion.

    The first half of :func:`convert_docx_to_svg`, for a caller that keeps the layout
    and writes pages from it later (:func:`page_svgs`): a render tool laying a document
    out once and drawing a few pages per call.  All three are picklable, so the layout
    can be handed to another process.

    **From a previous layout** (round 19, :mod:`docx2svg.relayout`): ``previous`` is a
    layout this function made of an earlier version of the same document, with the same
    options; the document is laid out again from the page the first changed block is on,
    until its pages line up with ``previous``'s again, and the rest is ``previous``'s.  The
    layout is the one a layout of the new bytes alone makes, page for page.  It says
    which pages to draw again -- ``layout.changed_pages``, the 1-based pages whose drawing
    differs from ``previous``'s page of the same number, and those ``previous`` did not
    have -- and ``layout.relaid`` why the whole document was laid out again where it was
    (:data:`docx2svg.relayout.FALLBACKS`; ``None`` where only what the change reaches
    was), ``layout.relaid_cause`` what caused it (where, which field, which part).  A body
    with page-number fields (a table of contents' ``PAGEREF``, ``PAGE``) or ``SEQ`` and
    ``REF`` is laid out again from the change too (round 20): each of the passes a full
    layout makes until its fields' values settle, from the same pass of ``previous``.  The changed blocks are found by comparing the bytes; ``changed_blocks`` (the
    blocks a caller knows it changed, as paths or ``w14:paraId``) is a hint only, never
    trusted over that.  ``keep=False`` keeps nothing for a later layout (``layout.state``
    is ``None``: smaller, and a ``previous`` it cannot be).

    ``measures`` (round 37), a :class:`docx2svg.relayout.MeasureStore`: what this process
    measured in earlier layouts of the same document, for a ``previous`` handed over without
    its measurements (from another process: they are most of its pickle) -- the blocks
    measured there are not measured again; the layout puts back what it measured.  The
    layout is the same with it or without.  ``changed_pages=False`` leaves
    ``layout.changed_pages`` ``None``: telling them draws each page laid out or drawn again,
    and the page it replaces, to compare them (after an edit that adds a page, every page
    whose footer shows the page count: most of such a layout's time), which a caller
    that compares pages its own way, or not at all, does not need."""
    from . import relayout
    from .layout import _Placer

    options = options if options is not None else ConvertOptions()

    data = _read(source)
    key = relayout.options_key(options)
    plan = relayout.plan(previous, data, key) if previous is not None else None
    fonts = None
    if options.advances is None or options.metrics is None or options.decorations is None:
        # The faces the previous layout measured with, where nothing they are found from changed
        # (a plan that may reuse the measurements holds a state).
        fonts = plan.state.fonts if plan is not None and plan.measured and plan.state is not None \
            else _fonts(data, options)
    advances = options.advances or fonts
    # (``fonts`` is found wherever one of the three is not given.)
    metrics = options.metrics or (fonts.metrics if fonts is not None else None)
    decorations = options.decorations or (fonts.decorations if fonts is not None else None)
    from .coverage import coverage_of

    parts: dict | None
    known: dict | None
    if plan is not None:
        parts, main, segments = plan.parts, plan.main, plan.segments
        known = plan.known
    elif keep:
        parts, main = relayout.part_digests(data)
        known = {}
        segments = relayout.segments(main, keep=known) if main is not None else None
    else:
        parts = main = segments = known = None
    # Round 41: the body's blocks the previous layout's parse kept, taken where their bytes
    # are the same (:class:`docx2svg.parse.document.KeptBlocks`); a layout that keeps its
    # state keeps them for the next.
    kept = getattr(plan.state, "blocks", None) if plan is not None and plan.state is not None else None
    if kept is None and keep:
        kept = KeptBlocks()
    document = parse_package(data, kept=kept, segments=segments) if kept is not None and segments is not None \
        else parse_package(data)
    # The document as parsed, before the layout sets its locale and keeps what it resolved:
    # what :func:`parsed_document` hands a caller reading the same bytes.
    parsed = replace(document) if keep else None
    # The measurements of the previous layout, where the parts they depend on are the same and
    # the blocks' models are what their XML says ("fields-open": a field's result in one block
    # may hang on another's).
    measured = plan.state.memo if plan is not None and plan.state is not None and plan.measured \
        and plan.reason != "fields-open" else None
    # Measurements reused go on with the previous layout's lineage; measured afresh, a new one.
    lineage = (plan.state.lineage if measured is not None and plan is not None and plan.state is not None
               else b"") or relayout.new_lineage()
    if measured is not None and measures is not None:
        kept_here = measures.get(lineage)
        if kept_here:
            measured = {**kept_here, **measured}
    relaid = plan.reason if plan is not None else None
    cause = plan.cause if plan is not None else None
    again = relayout.Again(plan, previous) if plan is not None and plan.reason is None else None

    def placer_for(memo):
        placer = _Placer(document, data, advances, metrics, decorations)
        if keep or plan is not None:
            placer.memo = relayout.Measures(memo)
            placer.segments = segments
            placer.digests = relayout.body_digests(segments)
        return placer

    from . import layout as layout_module

    placer = placer_for(measured)
    try:
        layout = layout_module.lay_out(document, advances, metrics, package=data, decorations=decorations,
                                       locale=options.locale, placer=placer, again=again)
    except relayout.Relaid as why:
        relaid, again = why.args[0], None
        cause = {"reason": relaid}
        placer = placer_for({**(measured or {}), **placer.memo.kept})
        layout = layout_module.lay_out(document, advances, metrics, package=data, decorations=decorations,
                                       locale=options.locale, placer=placer)
    for code in document.warnings:
        options.warnings.append(Warning(code, _PARSER_WARNINGS.get(code, "a body element is not read")))
    for substitution in getattr(fonts, "substitutions", {}).values():
        options.warnings.append(Warning("font-substituted", substitution.message()))
    glyphs: dict = {}
    for (family, char), used in getattr(fonts, "glyph_fallbacks", {}).items():
        glyphs.setdefault((family, used), []).append(char)
    for (family, used), chars in glyphs.items():
        source = "its missing glyph (.notdef)" if used == ".notdef" else repr(used)
        options.warnings.append(Warning(
            "font-glyph-fallback",
            f"{family!r} has no glyph for {chars[0]!r} (U+{ord(chars[0]):04X})"
            + (f" and {len(chars) - 1} more character(s)" if len(chars) > 1 else "")
            + f": measured with {source}, as an approximation (Word's own font fallback is not measured)"))
    for code, message, page in layout.warnings:
        options.warnings.append(Warning(code, message, page))
    layout.coverage = options.coverage = coverage_of(layout, document, data, fonts)
    if keep:
        source_document, original, _, by_page = placer.prepared
        assert parts is not None  # read above wherever ``keep`` is
        state = relayout.record(placer, layout, data, parts, main, fonts, key, placer.approximated, original,
                                by_page)
        state.laid = again.laid if again is not None else range(len(layout.pages))
        state.redrawn = frozenset(again.redrawn) if again is not None else frozenset()
        state.fresh = again.placed_pages if again is not None else frozenset(range(len(layout.pages)))
        state.lineage = lineage
        state.blocks = kept
        state.parsed = parsed
        state.known_segments = known
        layout.state = state
        if measures is not None:
            measures.put(lineage, state.memo)
    if previous is not None:
        layout.relaid = relaid
        layout.relaid_cause = cause if relaid is not None else None
        layout.changed_pages = (relayout.changed_pages(previous, layout, again, data, fonts, relaid)
                                if changed_pages else None)
    return layout, data, fonts


def parsed_document(layout, data: bytes) -> Document:
    """The model of ``data`` (:func:`parse_package`'s), taken from ``layout`` where it is the
    layout of these very bytes that :func:`lay_out_docx` made in this process (round 41): the
    document as it was parsed for it -- before the layout set its locale -- and not parsed
    again; else parsed.  The same document either way (``tests/test_kept_blocks.py``); the
    one handed over is shared with the layout, and nothing may change it."""
    state = getattr(layout, "state", None)
    parsed = getattr(state, "parsed", None)
    if parsed is not None and state is not None and state.data == data:
        return parsed
    return parse_package(data)


def warm(source=None, options: ConvertOptions | None = None, *, scan=None):
    """Do a process's one-time work now -- at a render worker's start -- rather than in
    its first layout, and return what it read of the font folders, to save
    (:class:`docx2svg.fonts.FontScan`, picklable).

    Always: imports every module a layout and its drawing use, and reads the font folders
    ``options`` names (:attr:`ConvertOptions.font_dirs`, else ``OOXML_FONT_DIRS``, after
    the system's and Word's) for the faces in them, in both of the orders a layout and its
    drawing look faces up in.  ``scan``, a :class:`~docx2svg.fonts.FontScan` an earlier
    process returned, spares reading the files it read that have not changed since
    (:func:`docx2svg.fonts.load_scan`; one of another docx2svg version is ignored).
    :func:`scan_fonts` is the scan alone, for a process that starts several workers.

    With ``source`` -- a document, typically the template a worker's documents are made
    from -- also lays it out and draws its first page with ``options``, which loads the
    faces it uses (their files, character maps and advances), the hyphenation patterns it
    asks for, and everything else a first layout does once; its output is discarded and
    ``options``' warnings and coverage are left as they were.

    Nothing it does changes a later conversion's output: it fills the caches the first
    conversion would fill, with the same values (``tests/test_warm.py``).  The caches are
    the process's: a new font installed in a folder already read is not seen until the
    process restarts, as without it."""
    import dataclasses
    import importlib
    import pkgutil

    from . import fonts

    options = options if options is not None else ConvertOptions()
    for module in pkgutil.walk_packages(__path__, f"{__name__}."):
        if module.name != f"{__name__}.cli":
            importlib.import_module(module.name)
    found_scan = scan_fonts(options, scan=scan)
    dirs = _font_dirs(options)
    fonts.installed_index(dirs)
    fonts.installed_index(fonts.drawing_dirs(dirs), bundle_first=True)
    fonts._rasteriser_families_index(fonts.drawing_dirs(dirs))
    if source is not None:
        quiet = dataclasses.replace(options, pages=[1], warnings=[], coverage=None)
        layout, data, found = lay_out_docx(source, quiet)
        page_svgs(layout, data, quiet, fonts=found)
    return found_scan


def scan_fonts(options: ConvertOptions | None = None, *, scan=None):
    """The scan of the font folders :func:`warm` makes, alone -- no imports, no index, no
    layout: a :class:`docx2svg.fonts.FontScan` of every font file in the folders
    ``options`` names (as :func:`warm` reads them), each file opened once, for a few
    kilobytes of its header (its table directory, ``name``, ``OS/2`` and ``head``).

    ``scan``, an earlier one, is taken first for the files whose modification time and
    size it recorded are still theirs (:func:`docx2svg.fonts.load_scan`; one of another
    docx2svg version is ignored), so only the files changed or added since are read.  A
    process that hands the result to its render workers' :func:`warm` (``scan=``) has the
    folders read once for all of them; what this process read is kept for it, as
    :func:`warm` keeps it."""
    from . import fonts

    options = options if options is not None else ConvertOptions()
    if scan is not None:
        fonts.load_scan(scan, __version__)
    return fonts.font_scan(_font_dirs(options), __version__)


#: The old name of :func:`lay_out_docx`, kept for docx-agent's tests until they move.
_lay_out = lay_out_docx


def _selected(options: ConvertOptions, count: int) -> list[int]:
    if options.pages is None:
        return list(range(count))
    return [number - 1 for number in options.pages if 1 <= number <= count]


def _images(data: bytes):
    """``(relationship id, the part it is the part's own) -> (content type, bytes)`` for
    the pictures of the main document (part ``None``) and of its headers and footers."""
    from .opc import content_type, main_part

    package = OpcPackage.open(data)
    main = main_part(package)

    def images(relationship: str, source: str | None = None):
        part = package.related_part(source or main, relationship)
        if part is None:
            return None
        found = package.read(part)
        if found is None:
            return None
        return content_type(package, part), found

    return images


def glyph_sheet(glyphs: dict) -> str:
    """The glyph set of pages outlined with ``ConvertOptions(glyphs={})`` as one ``<svg>``
    element that draws nothing and takes no room, to put once in the HTML page that shows
    them inline (:func:`ooxml_common.outlines.glyph_sheet`)."""
    from ooxml_common.outlines import glyph_sheet as _glyph_sheet

    return _glyph_sheet(glyphs)


def embed_glyphs(svg: str, glyphs: dict) -> str:
    """A page outlined with ``ConvertOptions(glyphs={})`` with the glyph definitions it
    uses put back into it: standalone again, as it would be outlined without the shared
    set (:func:`ooxml_common.outlines.embed_glyphs`)."""
    from ooxml_common.outlines import embed_glyphs as _embed_glyphs

    return _embed_glyphs(svg, glyphs)


def _names(options: ConvertOptions, fonts):
    """How the SVG names faces: the options' ``names``, else the installed faces'."""
    if options.names is not None:
        return options.names
    return fonts.drawing_name if fonts is not None else None


def page_svgs(layout, data: bytes, options: ConvertOptions | None = None, *, device_size: bool = False,
              fonts=None) -> list[str]:
    """The pages ``options.pages`` selects (every page by default) of a layout of
    :func:`lay_out_docx`, as SVG -- byte for byte what :func:`convert_docx_to_svg` writes
    for them.  ``data`` and ``fonts`` are the layout's own (the document's bytes, for its
    pictures; the faces, for how the SVG names them).  ``device_size`` sizes each page in
    device pixels (300 dpi) when ``options`` asks for no size: what
    :func:`convert_docx_to_png` rasterises."""
    from . import svg

    options = options if options is not None else ConvertOptions()

    images = _images(data)
    names = _names(options, fonts)
    documents = []
    for index in _selected(options, len(layout.pages)):
        page = layout.pages[index]
        width, height = options.width, options.height
        if device_size and width is None and height is None:
            width, height = page.width_px, page.height_px

        def warn(code: str, message: str, number: int = index + 1) -> None:
            if all((w.code, w.message) != (code, message) for w in options.warnings):
                options.warnings.append(Warning(code, message, number))

        documents.append(svg.render_page(page, glyph_size=options.glyph_size, images=images, width=width,
                                         height=height, warn=warn, names=names))
    if options.text not in TEXT_MODES:
        raise ValueError(f"ConvertOptions.text must be one of {', '.join(TEXT_MODES)}, not {options.text!r}")
    if options.text == "outlines":
        documents = _outlined(documents, layout, fonts, options)
    return documents


#: ``ConvertOptions.text``'s values.
TEXT_MODES = ("text", "outlines")


def _outlined(documents: list[str], layout, fonts, options: ConvertOptions) -> list[str]:
    """The pages with their text as outlines, drawn with the files
    :func:`convert_docx_to_png` hands the rasteriser (an embedded face written to a
    temporary folder while it is read); every kind of text kept as text reported once, as
    ``text-not-outlined``."""
    from .png import outline_svg

    numbers = [index + 1 for index in _selected(options, len(layout.pages))]
    out = []
    reported: set = set()
    with _raster_fonts(layout, fonts, options) as (files, dirs, skip, sans):
        for document, number in zip(documents, numbers):
            outlined = outline_svg(document, font_dirs=dirs, font_files=files,
                                   skip_system_fonts=skip, sans_serif_family=sans,
                                   selectable=options.selectable_text, glyphs=options.glyphs)
            out.append(outlined.svg)
            for kept in outlined.kept:
                if (kept.reason, kept.family) not in reported:
                    reported.add((kept.reason, kept.family))
                    options.warnings.append(Warning("text-not-outlined", kept.message(), number))
    return out


#: The name the agents called :func:`page_svgs` by; kept.
_svgs = page_svgs


def _render(source, options: ConvertOptions):
    """The SVG pass, keeping the layout, the bytes and the faces it was measured with."""
    layout, data, fonts = lay_out_docx(source, options, keep=False)
    return page_svgs(layout, data, options, fonts=fonts), layout, data, fonts


def convert_docx_to_svg(source, options: ConvertOptions | None = None) -> list[str]:
    """Render a document to one SVG document per page."""
    return _render(source, options or ConvertOptions())[0]


@dataclass
class Conversion:
    """A document laid out once and transcribed: what :func:`convert_docx` returns."""

    #: Every page laid out (:class:`docx2svg.layout.Layout`), whatever ``pages`` selects:
    #: its pages' lines, spans, rules, pictures and floating drawings, each with the
    #: ``path`` its SVG element carries as ``data-docx-path`` (:mod:`docx2svg.paths`).
    layout: Layout
    #: The selected pages' SVG documents, exactly as :func:`convert_docx_to_svg` writes them.
    svgs: list[str]
    #: The 1-based number of the page each SVG is of.
    page_numbers: list[int]


def convert_docx(source, options: ConvertOptions | None = None) -> Conversion:
    """Lay a document out once and return both the layout and the selected pages' SVGs:
    what :func:`convert_docx_to_layout` and :func:`convert_docx_to_svg` return, for the
    cost of one layout.  Warnings are collected in ``options.warnings`` as by
    :func:`convert_docx_to_svg`."""
    options = options or ConvertOptions()
    svgs, layout, _, _ = _render(source, options)
    return Conversion(layout, svgs, [index + 1 for index in _selected(options, len(layout.pages))])


def drawn_faces(layout) -> list[tuple[str, bool, bool]]:
    """Every ``(face, bold, italic)`` a layout draws text in, sorted: its lines' and its
    floating drawings' (a watermark's text path)."""
    faces = {(span.face, span.bold, span.italic) for page in layout.pages for line in page.text_lines()
             for span in line.spans if span.chars}
    faces.update(primitive.face for page in layout.pages for placed in getattr(page, "floats", ())
                 for primitive in getattr(placed, "primitives", None) or () if getattr(primitive, "face", None))
    return sorted(faces)


def convert_docx_to_png(source, options: ConvertOptions | None = None, *, backend: str = "auto",
                        font_dirs: Sequence[str] | None = None, font_files: Sequence[str] | None = None,
                        skip_system_fonts: bool | None = None) -> list[bytes]:
    """Render a document to one PNG per page (``pip install docx2svg[png]``).

    The rasteriser is given **the files Word draws every face with**
    (:meth:`docx2svg.fonts.InstalledFonts.drawing_face`: its own bundle's copy where it
    has one, whose advances are the ones the layout measured), each read in place --
    a face its document name does not reach is named by a fallback list instead
    (:class:`docx2svg.fonts.DrawingName`), so no font file is copied or rewritten --
    and the faces the document embeds (which exist only inside it: written,
    de-obfuscated, to a temporary directory for as long as the rasteriser needs them),
    plus ``font_dirs`` and ``font_files``.  ``skip_system_fonts`` defaults to on when every face drawn was
    found: named by family alone, a face installed twice (Symbol: macOS's and Word's
    SymbolMT) would otherwise be whichever copy the rasteriser meets first, and resvg
    prefers the system's even over a file it is given.  Without ``width``/``height`` a
    page is drawn at 300 dpi, Word's export grid.
    """
    options = options or ConvertOptions()
    layout, data, fonts = lay_out_docx(source, options, keep=False)
    return _rasterise(layout, data, fonts, options, backend=backend, font_dirs=font_dirs, font_files=font_files,
                      skip_system_fonts=skip_system_fonts)


def _rasterise(layout, data: bytes, fonts, options: ConvertOptions, *, backend: str = "auto",
               font_dirs: Sequence[str] | None = None, font_files: Sequence[str] | None = None,
               skip_system_fonts: bool | None = None) -> list[bytes]:
    if options.text != "text":
        # The PNG is drawn from the text itself: outlines would draw the same pixels.
        options = replace(options, text="text")
    documents = page_svgs(layout, data, options, device_size=True, fonts=fonts)
    with _raster_fonts(layout, fonts, options, font_dirs=font_dirs, font_files=font_files,
                       skip_system_fonts=skip_system_fonts, warn=True) as (files, dirs, skip, sans):
        return [svg_to_png(document, backend=backend, font_dirs=dirs, font_files=files,  # type: ignore[arg-type]
                           skip_system_fonts=skip, sans_serif_family=sans)
                for document in documents]


@contextmanager
def _raster_fonts(layout, fonts, options: ConvertOptions, *, font_dirs: Sequence[str] | None = None,
                  font_files: Sequence[str] | None = None, skip_system_fonts: bool | None = None,
                  warn: bool = False):
    """``(files, folders, skip system fonts, sans-serif family)`` the rasteriser is handed
    for ``layout``, for as long as the block runs (a face the document embeds is written,
    de-obfuscated, to a temporary folder meanwhile); ``warn`` reports a face the
    rasteriser may not tell from another (``face-not-addressable``)."""
    from .fonts import rasteriser_files

    # The files Word draws every face with, each read where it is installed; only a face
    # the document embeds is written (de-obfuscated, as embedded), and only while the
    # rasteriser needs it.
    faces = drawn_faces(layout)
    embeds = fonts is not None and any(
        (face := fonts.drawing_face(*drawn)) is not None and face.source.startswith("embedded:") for drawn in faces)
    with tempfile.TemporaryDirectory(prefix="docx2svg-embedded-") if embeds else nullcontext() as directory:
        files: list[str] = []
        sans = None
        missing = fonts is None
        if fonts is not None:
            files, lacking, unaddressable = rasteriser_files(fonts, faces, directory)
            missing = bool(lacking)
            for family in unaddressable if warn else ():
                options.warnings.append(Warning(
                    "face-not-addressable",
                    f"{family!r}: another face given to the rasteriser answers to the same family, weight, "
                    "stretch and style, so the rasteriser may draw that one"))
            # The placeholders' labels are drawn in the generic sans-serif face.
            for family in ("Arial", "Helvetica", "Calibri"):
                found = fonts.drawing_face(family, record=False)
                if found is not None and not found.source.startswith("embedded:"):
                    sans = family
                    if found.source not in files:
                        files.append(found.source)
                    break
        if skip_system_fonts is None:
            skip_system_fonts = not missing
        files += [str(path) for path in font_files or ()]
        dirs = list(font_dirs or ()) + [str(path) for path in _user_dirs(options)]
        yield files, dirs, skip_system_fonts, sans

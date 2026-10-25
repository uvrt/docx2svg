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
bookmarks; every other field is drawn as Word cached it).  Floating
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
from contextlib import nullcontext
from dataclasses import dataclass, field
from typing import Sequence

from .model import Document, PageMargins, PageSize, Paragraph, Run, Section
from .opc import Package
from .paginate import PageInfo
from .parse import parse_document, parse_package
from .paths import resolve_path
from .png import RasterizerNotAvailable, available_backends, svg_to_png
from ooxml_common.locales import EN_US, Locale

__version__ = "0.2.0"

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
    "lay_out_docx",
    "page_svgs",
    "parse_document",
    "parse_package",
    "resolve_path",
    "svg_to_png",
    "user_font_dirs",
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
    advances: object = None
    metrics: object = None
    decorations: object = None
    #: ``names(face, bold, italic) -> DrawingName | None``: how the SVG names a face its
    #: document name does not reach in a rasteriser.  By default
    #: :meth:`docx2svg.fonts.InstalledFonts.drawing_name`.
    names: object = None
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
    #: Collects :class:`Warning` for everything not drawn faithfully.
    warnings: list[Warning] = field(default_factory=list)
    #: Set by every conversion: how much of the document was laid out
    #: (:class:`docx2svg.coverage.Coverage`), also the layout's ``coverage``.
    coverage: object = None


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


def _fonts(data: bytes, options: ConvertOptions):
    from .fonts import SUBSTITUTES, InstalledFonts, default_font_dirs

    dirs = tuple(default_font_dirs()) + _user_dirs(options, subfolders=True)
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
    return lay_out_docx(source, options)[0]


def lay_out_docx(source, options: ConvertOptions | None = None) -> tuple:
    """``(layout, data, fonts)``: the document laid out -- every page, whatever
    ``options.pages`` selects -- its bytes, and the faces it was measured with
    (:class:`docx2svg.fonts.InstalledFonts`, ``None`` where ``options`` measures with
    its own ``advances``, ``metrics`` and ``decorations``).  Warnings and coverage go to
    ``options`` as by every conversion.

    The first half of :func:`convert_docx_to_svg`, for a caller that keeps the layout
    and writes pages from it later (:func:`page_svgs`): a render tool laying a document
    out once and drawing a few pages per call.  All three are picklable, so the layout
    can be handed to another process."""
    from .layout import lay_out

    options = options if options is not None else ConvertOptions()

    data = _read(source)
    document = parse_package(data)
    fonts = None
    if options.advances is None or options.metrics is None or options.decorations is None:
        fonts = _fonts(data, options)
    advances = options.advances or fonts
    metrics = options.metrics or fonts.metrics
    decorations = options.decorations or fonts.decorations
    from .coverage import coverage_of

    layout = lay_out(document, advances, metrics, package=data, decorations=decorations, locale=options.locale)
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
    return layout, data, fonts


#: The old name of :func:`lay_out_docx`, kept for docx-agent's tests until they move.
_lay_out = lay_out_docx


def _selected(options: ConvertOptions, count: int) -> list[int]:
    if options.pages is None:
        return list(range(count))
    return [number - 1 for number in options.pages if 1 <= number <= count]


def _images(data: bytes):
    """``(relationship id, the part it is the part's own) -> (content type, bytes)`` for
    the pictures of the main document (part ``None``) and of its headers and footers."""
    package = Package.open(data)
    main = package.main_document_part

    def images(relationship: str, source: str | None = None):
        part = package.part_by_id(source or main, relationship)
        if part is None or not package.exists(part):
            return None
        return package.content_type(part), package.read(part)

    return images


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
    return documents


#: The name the agents called :func:`page_svgs` by; kept.
_svgs = page_svgs


def _render(source, options: ConvertOptions):
    """The SVG pass, keeping the layout, the bytes and the faces it was measured with."""
    layout, data, fonts = lay_out_docx(source, options)
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
    layout: object
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
    layout, data, fonts = lay_out_docx(source, options)
    return _rasterise(layout, data, fonts, options, backend=backend, font_dirs=font_dirs, font_files=font_files,
                      skip_system_fonts=skip_system_fonts)


def _rasterise(layout, data: bytes, fonts, options: ConvertOptions, *, backend: str = "auto",
               font_dirs: Sequence[str] | None = None, font_files: Sequence[str] | None = None,
               skip_system_fonts: bool | None = None) -> list[bytes]:
    from .fonts import rasteriser_files

    documents = page_svgs(layout, data, options, device_size=True, fonts=fonts)
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
            for family in unaddressable:
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
        return [svg_to_png(document, backend=backend, font_dirs=dirs, font_files=files,  # type: ignore[arg-type]
                           skip_system_fonts=skip_system_fonts, sans_serif_family=sans)
                for document in documents]

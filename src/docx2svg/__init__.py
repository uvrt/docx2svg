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
above and below (``wrapTopAndBottom``).  Floating drawings text wraps beside and
multi-column sections are not laid out yet: the layout **stops**
at the first one, the page it is on shows a marked band where it starts, no page after
it is invented, and ``ConvertOptions.warnings`` says so with a stable code
(``layout-stopped:drawing``...), and a header or footer holding one is drawn up to it
(``story-stopped:drawing``).  A drop cap is drawn in its frame, the text beside it; any
other frame stops the layout (``layout-stopped:frame``).  Endnotes are numbered and drawn
after the text, under their separators, and footnotes at the foot of their reference's
page (:mod:`docx2svg.notes`), a table cell's too; one in a cell of a document with
sections of several text columns is not drawn (``footnotes-not-drawn``).

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

__version__ = "0.1.0"

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
    "parse_document",
    "parse_package",
    "resolve_path",
    "svg_to_png",
    "__version__",
]


@dataclass(frozen=True)
class Warning:  # noqa: A001 -- the sibling project's name for the same thing
    """Something the output does not show faithfully, with a stable ``code`` a caller can
    fail a build on (``layout-stopped:drawing``, ``footnotes-not-drawn``...)."""

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
    #: Extra directories to find faces in, searched after the ones Word uses.
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
    #: Collects :class:`Warning` for everything not drawn faithfully.
    warnings: list[Warning] = field(default_factory=list)


def _read(source) -> bytes:
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if isinstance(source, (str, os.PathLike)):
        with open(os.fspath(source), "rb") as handle:
            return handle.read()
    return source.read()


def _fonts(data: bytes, options: ConvertOptions):
    from pathlib import Path

    from .fonts import InstalledFonts, default_font_dirs

    dirs = tuple(default_font_dirs()) + tuple(Path(d) for d in (options.font_dirs or ()))
    return InstalledFonts(data, dirs=dirs)


_PARSER_WARNINGS = {
    "styles-part-missing": "the package has no styles part; Word's application defaults apply",
    "document-no-body": "the document has no body",
}


def convert_docx_to_layout(source, options: ConvertOptions | None = None):
    """Parse and lay out a document without writing SVG: a :class:`docx2svg.layout.Layout`
    with every glyph, rule and picture of every page placed."""
    return _lay_out(source, options or ConvertOptions())[0]


def _lay_out(source, options: ConvertOptions):
    from .layout import lay_out

    data = _read(source)
    document = parse_package(data)
    fonts = None
    if options.advances is None or options.metrics is None or options.decorations is None:
        fonts = _fonts(data, options)
    advances = options.advances or fonts
    metrics = options.metrics or fonts.metrics
    decorations = options.decorations or fonts.decorations
    layout = lay_out(document, advances, metrics, package=data, decorations=decorations)
    for code in document.warnings:
        options.warnings.append(Warning(code, _PARSER_WARNINGS.get(code, "a body element is not read")))
    for code, message, page in layout.warnings:
        options.warnings.append(Warning(code, message, page))
    return layout, data, fonts


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


def _svgs(layout, data: bytes, options: ConvertOptions, *, device_size: bool = False, fonts=None) -> list[str]:
    """The selected pages as SVG.  ``device_size`` sizes each in device pixels (300 dpi)
    when no size is asked for: what the PNG path rasterises at."""
    from . import svg

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


def _render(source, options: ConvertOptions):
    """The SVG pass, keeping the layout, the bytes and the faces it was measured with."""
    layout, data, fonts = _lay_out(source, options)
    return _svgs(layout, data, options, fonts=fonts), layout, data, fonts


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
    """Every ``(face, bold, italic)`` a layout draws text in, sorted."""
    return sorted({(span.face, span.bold, span.italic) for page in layout.pages for line in page.text_lines()
                   for span in line.spans if span.chars})


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
    layout, data, fonts = _lay_out(source, options)
    return _rasterise(layout, data, fonts, options, backend=backend, font_dirs=font_dirs, font_files=font_files,
                      skip_system_fonts=skip_system_fonts)


def _rasterise(layout, data: bytes, fonts, options: ConvertOptions, *, backend: str = "auto",
               font_dirs: Sequence[str] | None = None, font_files: Sequence[str] | None = None,
               skip_system_fonts: bool | None = None) -> list[bytes]:
    from .fonts import rasteriser_files

    documents = _svgs(layout, data, options, device_size=True, fonts=fonts)
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
                found = fonts.drawing_face(family)
                if found is not None and not found.source.startswith("embedded:"):
                    sans = family
                    if found.source not in files:
                        files.append(found.source)
                    break
        if skip_system_fonts is None:
            skip_system_fonts = not missing
        files += [str(path) for path in font_files or ()]
        dirs = list(font_dirs or ()) + list(options.font_dirs or ())
        return [svg_to_png(document, backend=backend, font_dirs=dirs, font_files=files,  # type: ignore[arg-type]
                           skip_system_fonts=skip_system_fonts, sans_serif_family=sans)
                for document in documents]

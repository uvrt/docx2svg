#!/usr/bin/env python3
"""Word's PDF pages as SVG (PyMuPDF), so both sides of the fidelity score go through resvg.

**Why.**  ``tools/fidelity.py`` used to rasterise Word's PDF with pdfium and our SVG with
resvg, so part of every score was pdfium against resvg.  pdfium is the odd engine out:
it grid-fits glyph outlines (a stem or a bar moves up to a pixel) and widens an
axis-aligned fill to whole pixels (a 4 px border drawn 5 px wide), where resvg and MuPDF
draw the outline as given.  Measured on ``sample-with-table``'s page: pdfium against
MuPDF on the *same PDF* scores SSIM 0.890 -- as low as our render did against Word --
and MuPDF against resvg of this module's SVG 0.996 (ROADMAP.md, "The instrument: one
rasteriser for both sides").  Converting Word's page to SVG and rasterising both sides
with resvg leaves no engine difference to score.

**The converter is part of the instrument**, so :func:`validate` holds it to Word's PDF
before its output is trusted, page by page:

* **same engine, two routes** -- MuPDF's raster of the PDF against MuPDF's raster of the
  SVG: whatever differs is the conversion's, since the engine is the same;
* **no pixel beyond anti-aliasing** -- :func:`beyond_antialiasing` counts the pixels of
  one raster outside the 3x3 neighbourhood range of the other (plus a tolerance): a
  moved glyph, a missing element or a changed colour lands there, an edge rendered a
  fraction of a pixel differently does not;
* **the glyphs are the embedded outlines** -- every font is embedded, every glyph the SVG
  defines is addressed by the embedded *subset's* glyph id, and is redrawn from that
  program, unhinted (:func:`unhinted_outlines`: MuPDF writes outlines with the font's
  hinting applied, which moves PowerPoint's Aptos 30 units); and the SVG draws with no
  text at all, which :func:`rasterise` proves by giving resvg no font to find.

**Font-derived artefacts.**  A converted page carries the glyph outlines of the fonts
Word embedded -- Microsoft's, subset.  They are treated like font files: never written
into the repository, cached only next to the oracle's PDFs (``ORACLE_DIR/svg/``, outside
the tree), and never published.

**Development tooling only**, behind the ``fidelity`` extra (``pip install -e
.[fidelity]``: PyMuPDF, AGPL-3.0 -- fine for a local tool that is not distributed with
the library; nothing under ``src/`` imports it).  Where PyMuPDF is missing,
:func:`available` is false and the tests that need it skip.

Usage::

    python tools/pdf_svg.py --validate            # the committed documents' pages and PROBES'
    python tools/pdf_svg.py --validate PDF...     # any PDFs, e.g. PowerPoint's
    python tools/pdf_svg.py --validate --jobs 1   # serially (default: a page per core)
    python tools/pdf_svg.py PDF                   # convert (cached); print the SVGs' paths

``--validate`` takes a page per process over ``--jobs`` processes (every logical core by
default, capped so the workers fit in memory: :data:`WORKER_MEMORY`), each PDF converted
first so that no two workers convert one at once, and reassembles each PDF's rows in page
order: the output and the verdicts are the serial run's, line for line.  ``--jobs 1`` is
the serial path, with no pool.  Word is never involved: validation reads its cached
exports, and a worker refuses to export (``tools/fidelity.py``'s ``no_word``).
"""

from __future__ import annotations

import argparse
import hashlib
import io
import os
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

#: Rasterising resolution: Word's export grid (fidelity.DPI).
DPI = 300
#: A pixel differs beyond anti-aliasing when a channel is further than this outside the
#: range of the other raster's 3x3 neighbourhood.  An edge pixel two rasterisers cover
#: differently stays inside its neighbours' range; a glyph moved by more than a pixel, a
#: missing stroke or a changed colour does not.
BEYOND_TOLERANCE = 48
#: The largest difference of flat colour (:func:`flat_colour_difference`) a faithful
#: conversion may show: a level of rounding in a colour conversion either way.
MAX_FLAT_COLOUR = 2
#: Isolated pixels beyond anti-aliasing allowed on a page.  MuPDF draws a PDF's text
#: through a glyph cache that places glyphs on a sub-pixel grid, and the SVG's glyphs as
#: paths at their exact position, so a thin stem can land a third of a pixel from the
#: other route's, even under one engine: measured, at most 18 pixels of a page (small
#: Times New Roman and Arial on ``render-15``, a Hebrew stem on ``style-fonts``).  A moved
#: glyph, a missing element or a changed colour is thousands.
MAX_BEYOND = 20
#: The conversion's own version, part of the cache key: 2 redraws the glyphs unhinted.
FORMAT = 2
#: What one validation worker may hold at once, with headroom (``--jobs`` defaults to no
#: more workers than the machine's memory holds at this size).  Measured, a page at the
#: default ``--supersample 4`` (the 1,200 dpi rasters of the 4x comparison are the largest
#: arrays anything here draws): 1.55 GB peak footprint and 1.83 GB peak resident on an A4
#: page, 1.49 GB and 1.78 GB on a Letter page.
WORKER_MEMORY = 2_500_000_000


def available() -> bool:
    """Whether PyMuPDF (the converter) imports."""
    try:
        import pymupdf  # noqa: F401
    except Exception:
        return False
    return True


def converter_version() -> str:
    import pymupdf

    return f"pymupdf-{pymupdf.VersionBind}-f{FORMAT}"


def cache_dir(pdf: Path) -> Path:
    """Where ``pdf``'s converted pages are cached: under the oracle directory (outside the
    repository), keyed by the PDF's content and the converter's version."""
    import oracle

    digest = hashlib.sha256(pdf.read_bytes()).hexdigest()[:16]
    return oracle.ORACLE_DIR / "svg" / f"{pdf.stem}-{digest}-{converter_version()}"


def _refuse_repository(path: Path) -> None:
    repository = HERE.parent.resolve()
    if path.resolve().is_relative_to(repository):
        raise RuntimeError(f"refusing to write a converted page (font-derived) inside the repository: {path}")


def _write_atomically(path: Path, text: str) -> None:
    """Write ``path`` aside and rename it into place.

    The cache is shared by every process of a parallel run (``--jobs``, and
    ``tools/fidelity.py``'s), and two of them may convert one PDF at once.  A rename is
    atomic, so a reader sees a whole file or none -- never one another process is halfway
    through -- and since both write the same page it does not matter which lands last."""
    partial = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    try:
        partial.write_text(text)
        os.replace(partial, path)
    finally:
        partial.unlink(missing_ok=True)


def convert(pdf: Path, cache: bool = True) -> list[tuple[str, dict]]:
    """Every page of ``pdf`` as SVG -- glyphs as paths, redrawn from the embedded programs
    (:func:`unhinted_outlines`) -- with that function's report.  Cached outside the
    repository when ``cache``."""
    import json

    import pymupdf

    pdf = Path(pdf)
    directory = cache_dir(pdf) if cache else None
    if directory is not None and (directory / "done").exists():
        count = int((directory / "done").read_text())
        return [((directory / f"p{i + 1}.svg").read_text(), json.loads((directory / f"p{i + 1}.json").read_text()))
                for i in range(count)]
    document = pymupdf.open(str(pdf))
    out = [unhinted_outlines(document, index, page.get_svg_image(text_as_path=True))
           for index, page in enumerate(document)]
    if directory is not None:
        _refuse_repository(directory)
        directory.mkdir(parents=True, exist_ok=True)
        for index, (svg, report) in enumerate(out):
            _write_atomically(directory / f"p{index + 1}.svg", svg)
            _write_atomically(directory / f"p{index + 1}.json", json.dumps(report))
        _write_atomically(directory / "done", str(len(out)))  # last: it says the rest are there
    return out


def convert_task(pdf: Path) -> None:
    """:func:`convert` of one PDF into the cache, unless it is there: a pool task."""
    if not (cache_dir(pdf) / "done").exists():
        convert(pdf)


def page_svgs(pdf: Path, cache: bool = True) -> list[str]:
    """:func:`convert`'s pages."""
    return [svg for svg, _report in convert(pdf, cache)]


_ROOT = re.compile(r'<svg\b[^>]*?\swidth="([\d.]+)"\s+height="([\d.]+)"\s+viewBox="0 0 ([\d.]+) ([\d.]+)"')


def device_grid(svg: str, dpi: float = DPI) -> str:
    """``svg`` with its root sized in whole device pixels and its viewBox widened to
    match, so that one point is exactly ``dpi / 72`` px.

    PyMuPDF writes the page's size in points (A4: ``595.2`` by ``841.92``), and resvg
    rounds a root's size to whole units *before* zooming: A4 came out 595 by 842, drawn
    2,479 px wide and stretched 1e-4 vertically -- 0.3 px at the foot of the page,
    which the fidelity score read as every line of ours sitting low.  Sizing the root in
    device pixels (the page's size rounded, as pdfium and our SVG round it) and widening
    the viewBox by the fraction of a pixel that adds leaves the drawing unscaled."""
    found = _ROOT.search(svg)
    if found is None:
        raise ValueError("a converted page's root has no width, height and viewBox from the origin")
    width_pt, height_pt = float(found.group(3)), float(found.group(4))
    width_px, height_px = round(width_pt * dpi / 72), round(height_pt * dpi / 72)
    root = found.group(0)
    sized = (root.replace(f'width="{found.group(1)}"', f'width="{width_px}"', 1)
             .replace(f'height="{found.group(2)}"', f'height="{height_px}"', 1)
             .replace(f'viewBox="0 0 {found.group(3)} {found.group(4)}"',
                      f'viewBox="0 0 {width_px * 72 / dpi!r} {height_px * 72 / dpi!r}"', 1))
    return svg.replace(root, sized, 1)


#: A number in a clip path's data this far out is Quartz's "no clip": Word's export wraps
#: every drawing in ``-257698032 -257697190 515396064 515396064 re W n`` (points), and
#: MuPDF writes it into the page's clip path as it is.  resvg cannot hold a coordinate
#: that far out -- it drops the clip's area, and every drawing inside it with it -- so
#: such a number is brought in to :data:`CLIP_BOUND`, which is still far past any page.
#: Found when a local template's shapes (drawn at last by the shared DrawingML renderers)
#: scored *worse*: the converted page had lost them all, where MuPDF's own raster of the
#: same SVG, which the validation compares, had them.
CLIP_FAR = 10_000_000
CLIP_BOUND = 1_000_000
_CLIP_DATA = re.compile(r'(<clipPath\b[^>]*>\s*<path\b[^>]*?\bd=")([^"]*)')
_FAR_NUMBER = re.compile(r"-?\d{8,}(?:\.\d+)?")


def bounded_clips(svg: str) -> str:
    """``svg`` with every clip path's far-out coordinates brought in (:data:`CLIP_FAR`)."""
    def bound(number: re.Match) -> str:
        value = float(number.group(0))
        return number.group(0) if abs(value) < CLIP_FAR else str(CLIP_BOUND if value > 0 else -CLIP_BOUND)

    return _CLIP_DATA.sub(lambda found: found.group(1) + _FAR_NUMBER.sub(bound, found.group(2)), svg)


def rasterise(svg: str, dpi: float = DPI, scale: int = 1):
    """resvg's raster of a converted page as an RGB array, on white, with **no fonts at all**:
    a converted page draws its glyphs as paths, and a ``<text>`` would draw nothing.  The
    page is drawn on the device grid (:func:`device_grid`), then zoomed by ``scale``."""
    import numpy as np
    import resvg_py
    from PIL import Image

    Image.MAX_IMAGE_PIXELS = None
    png = resvg_py.svg_to_bytes(svg_string=device_grid(bounded_clips(svg), dpi), zoom=float(scale), background="white",
                                skip_system_fonts=True)
    return np.asarray(Image.open(io.BytesIO(bytes(png))).convert("RGB"))


def truth_pages(pdf: Path) -> list:
    """Word's pages at :data:`DPI` through the converter and resvg."""
    return [rasterise(svg) for svg in page_svgs(pdf)]


# -- validation -------------------------------------------------------------------------


def beyond_antialiasing(a, b, tolerance: int = BEYOND_TOLERANCE):
    """The mask of pixels where either raster leaves the 3x3 neighbourhood range of the
    other by more than ``tolerance`` in some channel."""
    import numpy as np

    def neighbourhood(image):
        padded = np.pad(image, ((1, 1), (1, 1), (0, 0)), mode="edge")
        rows, cols = image.shape[:2]
        low, high = image.copy(), image.copy()
        for dy in range(3):
            for dx in range(3):
                window = padded[dy:dy + rows, dx:dx + cols]
                np.minimum(low, window, out=low)
                np.maximum(high, window, out=high)
        return low, high

    bad = np.zeros(a.shape[:2], bool)
    for x, y in ((a, b), (b, a)):
        x, y = x.astype(np.int16), y.astype(np.int16)
        low, high = neighbourhood(x)
        bad |= ((y < low - tolerance) | (y > high + tolerance)).any(axis=2)
    return bad


def _fit(a, b):
    rows, cols = min(a.shape[0], b.shape[0]), min(a.shape[1], b.shape[1])
    return a[:rows, :cols], b[:rows, :cols]


def flat_colour_difference(a, b) -> int:
    """The largest channel difference between two rasters where both are flat -- a pixel
    whose 3x3 neighbourhood is one colour in each: a fill, a line's core, a glyph's
    interior.  A changed colour shows here, whatever anti-aliasing does at edges.  (The
    fidelity score's 64-bin histogram cannot say it: a fill of 215 against 216 crosses a
    bin boundary, and PowerPoint's ``table-test`` scored 0.85 on one level of one
    channel.)"""
    import numpy as np

    def flat(image):
        image = image.astype(np.int16)
        padded = np.pad(image, ((1, 1), (1, 1), (0, 0)), mode="edge")
        rows, cols = image.shape[:2]
        same = np.ones((rows, cols), bool)
        for dy in range(3):
            for dx in range(3):
                same &= (padded[dy:dy + rows, dx:dx + cols] == image).all(axis=2)
        return same

    both = flat(a) & flat(b)
    if not both.any():
        return 0
    return int(np.abs(a.astype(np.int16) - b.astype(np.int16))[both].max())


def compare(a, b, images=()) -> dict:
    """SSIM and histogram (fidelity.score), mean and max |difference|, the pixels beyond
    anti-aliasing and the largest difference of flat colour, of two rasters of one page.
    Inside ``images`` (rectangles in pt, :func:`image_regions`) the pixels beyond
    anti-aliasing are counted apart, and ``flat_colour`` leaves them out
    (``flat_colour_all`` does not): see :func:`faithful`."""
    import numpy as np

    import fidelity

    a, b = _fit(a, b)
    row = fidelity.score(b, a)
    difference = np.abs(a.astype(np.int16) - b.astype(np.int16)).max(axis=2)
    beyond = beyond_antialiasing(a, b)
    inside = np.zeros(beyond.shape, bool)
    for x0, y0, x1, y1, *_size in images:
        k = DPI / 72
        inside[max(int(y0 * k) - 2, 0):int(y1 * k) + 3, max(int(x0 * k) - 2, 0):int(x1 * k) + 3] = True
    outside_a, outside_b = a.copy(), b.copy()
    outside_b[inside] = outside_a[inside]
    return {"ssim": row["ssim"], "histogram": row["histogram"], "mean_abs": round(float(difference.mean()), 3),
            "max_abs": int(difference.max()), "beyond": int((beyond & ~inside).sum()),
            "beyond_in_images": int((beyond & inside).sum()),
            "flat_colour": flat_colour_difference(outside_a, outside_b), "flat_colour_all": flat_colour_difference(a, b)}


def _mupdf(document, index: int, dpi: float = DPI):
    import numpy as np

    pixmap = document[index].get_pixmap(dpi=int(dpi), alpha=False)
    return np.frombuffer(pixmap.samples, np.uint8).reshape(pixmap.height, pixmap.width, 3)


def _pdfium(pdf: Path, index: int, dpi: float = DPI):
    import numpy as np
    import pypdfium2 as pdfium

    return np.asarray(pdfium.PdfDocument(str(pdf))[index].render(scale=dpi / 72).to_pil().convert("RGB"))


def _downsample(image, factor: int):
    import numpy as np

    rows, cols = image.shape[0] // factor * factor, image.shape[1] // factor * factor
    out = np.empty((rows // factor, cols // factor, 3), np.uint8)
    strip = 256  # output rows at a time: the whole page in float would be gigabytes
    for top in range(0, rows // factor, strip):
        bottom = min(top + strip, rows // factor)
        blocks = image[top * factor:bottom * factor, :cols].reshape(bottom - top, factor, cols // factor, factor, 3)
        out[top:bottom] = blocks.astype(np.float32).mean(axis=(1, 3)).round().astype(np.uint8)
    return out


_GLYPH = re.compile(r'<path id="font_(\d+)_(\d+)" d="([^"]*)"')
_USE = re.compile(r'<use data-text="[^"]*" xlink:href="#font_(\d+)_(\d+)" transform="matrix\(([^)]*)\)"')


def _programs(document, index: int) -> tuple[dict, list[str]]:
    """The embedded font programs of page ``index`` by ``/BaseFont`` (``None`` where
    fontTools cannot read one), and the fonts not embedded."""
    from fontTools.ttLib import TTFont

    programs, not_embedded = {}, []
    for font in document[index].get_fonts(full=True):
        buffer = document.extract_font(font[0])[3]
        if not buffer:
            not_embedded.append(font[3])
            continue
        try:
            programs[font[3]] = TTFont(io.BytesIO(buffer), fontNumber=0)
        except Exception:
            programs[font[3]] = None
    return programs, sorted(set(not_embedded))


def _traced_names(page) -> dict:
    """MuPDF's text trace: the font name of every glyph drawn, by glyph id and origin."""
    out = {}
    for span in page.get_texttrace():
        for _unicode, gid, origin, _bbox in span["chars"]:
            out[(gid, round(origin[0], 1), round(origin[1], 1))] = span["font"]
    return out


def _glyph_bounds(font, gid: int):
    from fontTools.pens.boundsPen import BoundsPen

    glyph_set = font.getGlyphSet()
    pen = BoundsPen(glyph_set)
    glyph_set[font.getGlyphOrder()[gid]].draw(pen)
    return pen.bounds


def _glyph_path(font, gid: int) -> str:
    """The glyph's outline as SVG path data, in ems, y up (as MuPDF writes its glyphs)."""
    from fontTools.pens.svgPathPen import SVGPathPen
    from fontTools.pens.transformPen import TransformPen

    glyph_set = font.getGlyphSet()
    scale = 1 / font["head"].unitsPerEm
    pen = SVGPathPen(glyph_set, ntos=lambda v: f"{v:.9g}")
    glyph_set[font.getGlyphOrder()[gid]].draw(TransformPen(pen, (scale, 0, 0, scale, 0, 0)))
    return pen.getCommands()


def unhinted_outlines(document, index: int, svg: str) -> tuple[str, dict]:
    """``svg`` (MuPDF's page ``index``) with every glyph redrawn from the embedded font
    program it was drawn with, unhinted; and the report :func:`validate` holds.

    **Why redraw.**  MuPDF's SVG writes a glyph's outline as FreeType loads it *with the
    font's hinting* (at one pixel a font unit): for Word's Calibri and Cambria subsets
    that is within a unit or two of the outline, but PowerPoint's Aptos subsets come out
    with their x-height pulled up 30 units (1.5% of the em, 2.7 px at 44 pt) -- measured,
    ``table-test.pdf``.  Both PDF rasterisers draw the outline unhinted at that scale,
    so the converter's own reading of a glyph is replaced by the program's.

    **Which program.**  MuPDF names each SVG font ``font_<n>`` and each glyph by its id in
    the *embedded subset* (an id only that subset gives meaning to -- a font lookup would
    carry the installed face's ids).  Font ``n`` is the embedded program named by MuPDF's
    text trace for its glyphs' origins, and among subsets of one face (Word writes a face
    once per run of formatting) the one whose outlines are nearest MuPDF's, in font
    units.  The report says how far MuPDF's hinted outline was from the one drawn
    (``hinting_units``), and fails a font no embedded program accounts for."""
    from fontTools.pens.boundsPen import BoundsPen
    from fontTools.svgLib.path import parse_path

    programs, not_embedded = _programs(document, index)
    traced = _traced_names(document[index])
    names: dict[str, set] = {}
    for n, gid, matrix in _USE.findall(svg):
        values = [float(v) for v in matrix.split(",")]
        name = traced.get((int(gid), round(values[4], 1), round(values[5], 1)))
        if name is not None:
            names.setdefault(n, set()).add(name)
    glyphs: dict[str, list] = {}
    for n, gid, d in _GLYPH.findall(svg):
        pen = BoundsPen(None)
        parse_path(d, pen)
        glyphs.setdefault(n, []).append((int(gid), pen.bounds))

    choice, unmatched, hinting = {}, [], 0.0
    for n, rows in sorted(glyphs.items(), key=lambda item: int(item[0])):
        wanted = names.get(n, set())
        best = None
        for base, font in programs.items():
            if font is None or (wanted and base.split("+", 1)[-1] not in wanted):
                continue
            order_size, units = len(font.getGlyphOrder()), font["head"].unitsPerEm
            if any(gid >= order_size for gid, _ in rows):
                continue
            total, worst, fits = 0.0, 0.0, True
            for gid, bounds in rows:
                theirs = _glyph_bounds(font, gid)
                if (bounds is None) != (theirs is None):
                    fits = False
                    break
                if bounds is not None:
                    off = max(abs(p * units - q) for p, q in zip(bounds, theirs))
                    total, worst = total + off, max(worst, off)
            if fits and (best is None or total < best[1]):
                best = (base, total, worst)
        if best is None:
            unmatched.append(n)
        else:
            choice[n] = best[0]
            hinting = max(hinting, best[2])

    def redraw(match) -> str:
        n, gid = match.group(1), int(match.group(2))
        if n not in choice:
            return match.group(0)
        return f'<path id="font_{n}_{gid}" d="{_glyph_path(programs[choice[n]], gid)}"'

    out = _GLYPH.sub(redraw, svg)
    text_elements = len(re.findall(r"<text\b", out))
    report = {"fonts": len(programs) + len(not_embedded), "not_embedded": not_embedded,
              "unreadable": sorted(b for b, f in programs.items() if f is None), "svg_fonts": len(glyphs),
              "matched": {n: [choice[n], len(glyphs[n])] for n in sorted(choice, key=int)},
              "unmatched": unmatched, "named": all(n in names for n in choice), "text_elements": text_elements,
              "hinting_units": round(hinting, 2)}
    report["ok"] = not not_embedded and not unmatched and not text_elements and report["named"]
    return out, report


def image_regions(pdf: Path, index: int) -> list[tuple]:
    """The rectangles (pt) where page ``index`` draws a raster image, and its pixel size."""
    import pymupdf

    page = pymupdf.open(str(pdf))[index]
    out = []
    for info in page.get_image_info():
        x0, y0, x1, y1 = info["bbox"]
        out.append((round(x0, 1), round(y0, 1), round(x1, 1), round(y1, 1), info["width"], info["height"]))
    return out


def validate(pdf: Path, supersample: int = 4, pages: list[int] | None = None, cache: bool = True) -> list[dict]:
    """Per page: the conversion under one engine (MuPDF's PDF raster against MuPDF's SVG
    raster), resvg's raster of the SVG against MuPDF's and pdfium's rasters of the PDF --
    at 1x and, when ``supersample``, both drawn at that factor and area-averaged down --
    and :func:`unhinted_outlines`' report."""
    import pymupdf

    converted = convert(pdf, cache)
    document = pymupdf.open(str(pdf))
    out = []
    for index, (svg, report) in enumerate(converted):
        if pages is not None and index not in pages:
            continue
        ours = rasterise(svg)
        images = image_regions(pdf, index)
        truth = _mupdf(document, index)
        row = {"page": index + 1, "outlines": report, "images": images,
               "mupdf_pdf_vs_resvg_svg": compare(truth, ours, images),
               "pdfium_vs_resvg_svg": compare(_pdfium(pdf, index), ours, images)}
        # MuPDF's own SVG reader ignores <mask> (an image's soft mask: PowerPoint's chart
        # scenes), drawing the mask's image black; there the same-engine route says
        # nothing about the conversion, and resvg's route stands alone.
        if "<mask" not in svg:
            svg_document = pymupdf.open(stream=svg.encode(), filetype="svg")
            row["mupdf_pdf_vs_mupdf_svg"] = compare(truth, _mupdf(svg_document, 0), images)
        if supersample:
            big = _downsample(_pdfium(pdf, index, DPI * supersample), supersample)
            row[f"pdfium_vs_resvg_svg_{supersample}x"] = compare(big, _downsample(rasterise(svg, scale=supersample),
                                                                                  supersample), images)
        out.append(row)
    return out


def validate_page(task) -> list[dict]:
    """:func:`validate` of one page, ``(pdf, page index, supersample, cache)``: a pool
    task.  Pages are validated independently of one another, so a PDF's rows are its
    pages' rows concatenated in page order."""
    pdf, page, supersample, cache = task
    return validate(pdf, supersample, [page], cache=cache)


def default_jobs() -> int:
    """Every logical core, unless memory holds fewer validation workers of
    :data:`WORKER_MEMORY`."""
    import parallel

    return parallel.default_jobs(WORKER_MEMORY)


#: Word's pages the converter is validated on besides the committed documents' (the
#: newest cached export of each ``tools/oracle.py`` stem): text in several faces, borders,
#: shading and highlights, tables, pictures, super/subscripts and the ink probe.
PROBES = (
    "ink", "render-15", "border-probe", "picture-15", "table-border-15", "table-content-15", "table-merge-15",
    "script-15", "style-fonts", "line-advance-sizes-verdana", "line-advance-sizes-aptos", "kerning-probe",
)


def committed_pdf(name: str) -> Path:
    """Word's cached export of a committed document (``fidelity.COMMITTED``), as the
    fidelity scores find it; never exports."""
    import oracle

    import read_render

    if name == "layout-sweep.docx":
        return oracle.ORACLE_DIR / "layout-sweep.pdf"
    data = (read_render.FIXTURES / name).read_bytes()
    return oracle.ORACLE_DIR / f"{Path(name).stem}-{hashlib.sha256(data).hexdigest()[:16]}.pdf"


def validation_pdfs() -> list[Path]:
    """The committed documents' exports, then :data:`PROBES`' -- those cached here."""
    import fidelity
    import oracle

    out = [pdf for pdf in map(committed_pdf, fidelity.COMMITTED) if pdf.exists()]
    for stem in PROBES:
        found = sorted(oracle.ORACLE_DIR.glob(f"{stem}-" + "[0-9a-f]" * 16 + ".pdf"), key=lambda p: p.stat().st_mtime)
        if found:
            out.append(found[-1])
    return out


def faithful(row: dict) -> bool:
    """Whether :func:`validate`'s row passes: every glyph the embedded program's; no flat
    colour off by more than :data:`MAX_FLAT_COLOUR`; and at most :data:`MAX_BEYOND`
    isolated pixels beyond anti-aliasing -- between MuPDF's raster of the PDF and resvg's
    of the SVG over the whole page, embedded bitmaps included, and between MuPDF's two
    rasters outside the bitmaps.

    MuPDF's own SVG reader is not a faithful reader of everything the conversion writes,
    measured where resvg's and pdfium's rasters agree with MuPDF's PDF raster exactly: it
    ignores ``<mask>`` (an image's soft mask: PowerPoint's chart scenes drawn black), and
    it filters a stretched bitmap its own way (``picture-15``'s 1 x 1 px colour swatch
    stretched over 118 pt, 128 levels off at its edges).  So the same-engine route is
    skipped on a page with a mask and holds only outside the bitmaps."""
    same, cross = row.get("mupdf_pdf_vs_mupdf_svg"), row["mupdf_pdf_vs_resvg_svg"]
    return (row["outlines"]["ok"]
            and (same is None or (same["beyond"] <= MAX_BEYOND and same["flat_colour"] <= MAX_FLAT_COLOUR))
            and cross["beyond"] + cross["beyond_in_images"] <= MAX_BEYOND
            and cross["flat_colour_all"] <= MAX_FLAT_COLOUR)


def _print_row(name: str, row: dict, supersample: int) -> bool:
    same, cross, pdfium = row.get("mupdf_pdf_vs_mupdf_svg"), row["mupdf_pdf_vs_resvg_svg"], row["pdfium_vs_resvg_svg"]
    big = row.get(f"pdfium_vs_resvg_svg_{supersample}x")
    outlines = row["outlines"]
    ok = faithful(row)

    def route(r) -> str:
        return (f"{r['ssim']:.4f} colour {r['flat_colour']} beyond {r['beyond']:>4}"
                + (f" (+{r['beyond_in_images']} in bitmaps)" if r["beyond_in_images"] else ""))

    print(f"{name[:34]:34} p{row['page']:<3} MuPDF pdf/svg " + (route(same) if same else "n/a (<mask>)")
          + " | MuPDF/resvg " + route(cross) + " | pdfium/resvg " + route(pdfium)
          + (f" | {supersample}x " + route(big) if big else "")
          + f" | glyphs {sum(n for _, n in outlines['matched'].values())} in {outlines['svg_fonts']} fonts"
          + f" (MuPDF's hinting moved them <= {outlines['hinting_units']} units)"
          + ("" if outlines["named"] else " UNNAMED")
          + (f" UNMATCHED {outlines['unmatched']}" if outlines["unmatched"] else "")
          + (f" NOT EMBEDDED {outlines['not_embedded']}" if outlines["not_embedded"] else "")
          + (f" images {row['images']}" if row["images"] else "")
          + ("" if ok else "  FAIL"))
    return ok


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pdf", nargs="*", type=Path)
    parser.add_argument("--validate", action="store_true",
                        help="hold the converter to the PDFs (default: the committed documents' and PROBES')")
    parser.add_argument("--supersample", type=int, default=4,
                        help="also compare pdfium and resvg at this factor (0: no)")
    parser.add_argument("--max-pages", type=int, help="validate at most this many pages of each PDF")
    parser.add_argument("--no-cache", action="store_true",
                        help="convert in memory only (e.g. another project's PDFs: nothing is written)")
    parser.add_argument("--jobs", "-j", type=int, default=None,
                        help="processes to validate pages in (default: every logical core, fewer if memory is "
                        "short); 1 is the serial path.  Output and verdicts are identical either way")
    args = parser.parse_args(argv[1:])
    jobs = default_jobs() if args.jobs is None else args.jobs
    if jobs < 1:
        parser.error("--jobs must be at least 1")
    if not available():
        print("PyMuPDF is not installed: pip install -e .[fidelity] (see tools/pdf_svg.py)")
        return 2
    if not args.validate:
        for pdf in args.pdf:
            page_svgs(pdf)
            print(cache_dir(pdf))
        return 0
    pdfs = args.pdf or validation_pdfs()
    cache = not args.no_cache
    if jobs == 1:
        # The serial path: a PDF validated, then printed, then the next.
        results = (validate(pdf, args.supersample, list(range(args.max_pages)) if args.max_pages else None,
                            cache=cache) for pdf in pdfs)
    else:
        # A page per task (a page at the default 4x is the best part of a minute), every
        # PDF converted first so that no two workers convert one at once; each PDF's rows
        # are then its pages' rows in page order, which is what validate() returns for the
        # whole PDF.
        import pymupdf
        from fidelity import no_word
        from parallel import pool_map

        if cache:
            pool_map(convert_task, [pdf for pdf in pdfs if not (cache_dir(pdf) / "done").exists()], jobs, no_word)
        counts = [pymupdf.open(str(pdf)).page_count for pdf in pdfs]
        tasks = [(number, pdf, page) for number, (pdf, count) in enumerate(zip(pdfs, counts))
                 for page in range(min(count, args.max_pages) if args.max_pages else count)]
        found = pool_map(validate_page, [(pdf, page, args.supersample, cache) for _number, pdf, page in tasks],
                         jobs, no_word)
        results = [[] for _pdf in pdfs]
        for (number, _pdf, _page), rows in zip(tasks, found):
            results[number].extend(rows)
    failures = 0
    for pdf, rows in zip(pdfs, results):
        for row in rows:
            failures += not _print_row(pdf.stem, row, args.supersample)
    print(f"{failures} page(s) failed" if failures else "every page converted within anti-aliasing")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

#!/usr/bin/env python3
"""Score our raster against Word's: SSIM and colour-histogram correlation over foreground.

The method is the sibling ``pptx2svg``'s ``tools/fidelity.py``, adapted here rather than
imported (Phase 1 decided the harness is not shared code): two gates that answer
different questions.

* **SSIM** on greyscale compares local means, variances and covariance -- "is the same
  thing in the same place" -- so a missing line or a shifted page moves it and half a
  pixel of anti-aliasing does not.
* **Colour-histogram correlation** (64 bins a channel) is blind to position and catches
  wrong colours, which SSIM on grey cannot see.

Both are taken over **foreground pixels only** -- grey below 245 in either image -- or a
mostly white page scores near-perfect whatever is on it; under 1.5% foreground a page is
too sparse to score and counts 1.0.

**The caveat, kept from the sibling:** SSIM is a *mean over the foreground mask*, so
removing wrong ink shrinks the denominator as well as the error, and a correction can
lower SSIM while the picture improves.  So every row also carries the **unnormalised
structural loss** -- the sum of ``1 - SSIM`` over the mask -- and the mean absolute
difference; when SSIM falls alone while those improve, look at the pictures.

**Both sides draw with the real faces, or the score measures font substitution.**  Word
draws with Microsoft's installed faces, so ours is rasterised with the same files, read
in place and never copied: for each face the copy Word *draws* with
(``InstalledFonts.drawing_face``: its bundle's, whose version its PDF embeds), and the
document's own embedded faces (de-obfuscated into a temporary directory for the
rasteriser, then deleted), with the host's other fonts switched off.  A document that
draws a face this machine does not have is skipped with the reason.

**A low score is a failure to explain, not a number to record.**  ``sample-with-table``
was recorded at SSIM 0.787 against its neighbours' 0.92-0.96 and read as a gain: the
text of every shaded table row was painted over by its shading (ROADMAP.md, "Tables --
measured").  So every run lists the **outliers** -- a page drawn in full whose SSIM is
more than :data:`OUTLIER_DROPS` below the median page of its document or of the corpus --
with the unnormalised loss beside it (:func:`outliers`), and ``tests/test_fidelity.py``
fails on one that is not explained there.

**Every page either side made is scored.**  A layout that makes a page more than Word, or
a page fewer, puts every page after the one that differs somewhere else, and scoring only
the pages both made would never show it (an extra page went unnoticed that way).  So a
page only one side made is scored as wholly wrong -- SSIM and histogram 0
(:func:`fully_wrong`) -- in every mean; each document's line reports our page count and
Word's; and a document whose count is not Word's is reported as a ``PAGE COUNT
MISMATCH`` (:func:`page_count_mismatches`), makes the run exit 1, cannot be recorded, and
fails ``tests/test_fidelity.py``.

**Ink or advances.**  Word's PDF places every glyph at the pen position the layout
computes but scales its outline to the size *rounded to whole device pixels* (11 pt is
45.83 px, drawn at 46: Quartz's ``Tm``).  ``--glyph-size`` renders ours either way:
``device`` by default, ``exact`` or ``both`` when asked for (``--ink`` always renders both).

**One rasteriser for both sides** (``--truth``).  Word's PDF used to be rasterised by
pdfium and our SVG by resvg, so part of every score was pdfium against resvg: pdfium
grid-fits glyph outlines and widens an axis-aligned fill to whole pixels, and on
``sample-with-table``'s page it scores 0.890 against MuPDF on the *same* PDF.  The
default truth is now ``svg``: Word's page converted to SVG by ``tools/pdf_svg.py``
(PyMuPDF; glyphs as the embedded outlines, validated against the PDF page by page) and
rasterised by resvg like ours, so a pixel that differs is a layout or drawing
difference.  ``--truth pdfium`` is the old instrument, kept for comparison (ROADMAP.md,
"The instrument: one rasteriser for both sides").

**Only the default is scored unless asked for**: the ``svg`` truth at the ``device`` glyph
size.  ``--truth pdfium`` / ``both`` and ``--glyph-size exact`` / ``both`` score the rest,
and so do their baselines: ``tests/fidelity-baselines.json`` keeps the ``exact`` entries
beside the ``device`` ones and the ``pdfium`` truth under its name, a default ``--record``
carries every one of them over exactly as recorded (:func:`baselines_payload`, which
refuses rather than keep one that no longer describes its document), and ``--record
--truth both --glyph-size both`` re-records them all.

Dev-only: numpy, pillow, pypdfium2 and resvg-py, and PyMuPDF for the ``svg`` truth (the
``fidelity`` extra); Word's exports come from ``tools/oracle.py`` (cached by content).

Usage::

    python tools/fidelity.py                      # the committed documents (svg truth, device glyphs)
    python tools/fidelity.py DOCX...              # any documents (their scores are printed only)
    python tools/fidelity.py --truth pdfium       # the old instrument: Word's PDF through pdfium
    python tools/fidelity.py --truth both --glyph-size both   # everything that is recorded
    python tools/fidelity.py --record             # rewrite the svg/device baselines, keep the rest
    python tools/fidelity.py --record --truth both --glyph-size both   # re-record everything
    python tools/fidelity.py --jobs 1             # serially (default: every core)
    python tools/fidelity.py --verify-cache       # re-draw everything, hold the raster cache to it
    python tools/fidelity.py --no-cache           # draw everything, cache nothing

**Parallel.**  Pages are scored over ``--jobs`` processes (every logical core by default,
capped so the workers fit in memory: :data:`WORKER_MEMORY`), a task per document and page:
the task lays the document out and renders its SVGs as the serial run does, rasterises
that one page of ours at each glyph size and of Word's under each truth, and scores them.
Each page is computed whole in one process and the rows are reassembled in document and
page order, so every score, every baseline field and every printed line is the serial
run's bit for bit; ``--jobs 1`` is the serial path itself, with no pool.  Word is never
launched by a worker: the exports are looked up (``tools/oracle.py``, cached by content)
in this process, one at a time, before any pool starts, and a worker refuses to export
(:func:`no_word`).

**Rasters are cached, and the cache is checked on every run** (``tools/raster_cache.py``,
the sibling's, unchanged).  Our raster and Word's are kept in the oracle directory's
``svg/rasters/`` (:func:`raster_cache_root`) -- outside the repository, since a raster of
Word's page holds Microsoft's glyph shapes -- under a key of every input that moves their
pixels (:func:`our_components`, :func:`truth_components`: the SVG's and the PDF's bytes,
the *contents* of every font file resvg is handed, the converter, the rasteriser and
decoder versions, the resolution and options, :data:`HARNESS_VERSION`), stored beside
each raster and compared on every read.  A key is only as complete as what its author
knew moves pixels, so every run also re-draws :data:`VERIFY_SAMPLE` pages, spread across
the corpus and rotated by date, and holds them to the cache byte for byte: **any
difference discards the whole cache and stops the run** (exit 3) without scoring.
``--verify-cache`` re-draws every page; ``--no-cache`` bypasses the cache.  (pdfium's
rasters are not cached: that truth runs only when asked for.)

**SSIM is computed over the content only** (:func:`content_box`): the bounding box of the
union of both images' foreground, grown by the SSIM window's radius.  That is exact, not
an approximation -- the same floats in the same order -- and the same sampled pages are
scored both ways on every run and held equal, bit for bit (:func:`score_checked`).
"""

from __future__ import annotations

import argparse
import io
import json
import os
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import parallel  # noqa: E402  (tools/parallel.py: standard library only)
from parallel import pool_map  # noqa: E402

REPO = HERE.parent
BASELINES = REPO / "tests" / "fidelity-baselines.json"

#: Pixels at least this bright in both images are background.
FOREGROUND_MAX = 245
#: Below this much foreground a page is too sparse to score.
MIN_FOREGROUND = 0.015
#: How far below the typical page (the median of its document's pages, or of the
#: corpus's) a page's SSIM -- or a document's mean below the median document's -- may
#: fall before it is an outlier to explain.  The banded table hidden under its shading
#: was 0.11 below the corpus's median page and 0.14 below the median document.  Per truth
#: (:data:`TRUTHS`): 0.05 was set against pdfium, whose pages spread 0.84-0.96; with one
#: rasteriser the typical page is 0.99 and the committed pages spread 0.96-1.0, so 0.05 would
#: let a page lose five times the whole residual unremarked.  0.02 is the old drop scaled
#: to the new spread (ROADMAP.md, 5.13).
OUTLIER_DROPS = {"svg": 0.02, "pdfium": 0.05}
OUTLIER_DROP = OUTLIER_DROPS["svg"]
#: Rendering resolution: Word's export grid, 300 dpi, so a device pixel is a pixel.
DPI = 300
#: How Word's page is rasterised: ``svg`` -- converted by ``tools/pdf_svg.py`` and drawn
#: by resvg, like ours (the default); ``pdfium`` -- the PDF drawn by pdfium (the old
#: instrument, kept for comparison).
TRUTHS = ("svg", "pdfium")
#: The default truth: the only one a run scores unless asked for another (``--truth``).
DEFAULT_TRUTH = "svg"
#: How our glyphs are sized (``ConvertOptions.glyph_size``): ``device`` -- the outline
#: scaled to the size Word's PDF draws it at, whole device pixels -- and ``exact``.
GLYPH_SIZES = ("device", "exact")
#: The glyph size a run renders unless asked for both (``--glyph-size``).
DEFAULT_GLYPH_SIZE = "device"


# -- metrics ---------------------------------------------------------------------------


#: The SSIM window's radius: an 11-tap Gaussian (sigma 1.5, truncated at 5 px).  A pixel's
#: SSIM depends on the pixels within this many rows and columns of it and on nothing else,
#: which is what makes cropping to the content exact (:func:`content_box`).
SSIM_RADIUS = 5


def _gaussian_kernel(sigma: float = 1.5, radius: int = SSIM_RADIUS):
    import numpy as np

    x = np.arange(-radius, radius + 1, dtype=np.float64)
    k = np.exp(-(x ** 2) / (2 * sigma ** 2))
    return k / k.sum()


def _blur(image, kernel):
    """Separable Gaussian blur with edge padding: scipy's when it is installed (the same
    kernel: sigma 1.5, radius 5, normalised; edge mode), plain numpy otherwise."""
    import numpy as np

    try:
        from scipy.ndimage import gaussian_filter1d
    except ImportError:
        gaussian_filter1d = None
    if gaussian_filter1d is not None and len(kernel) == 11:
        out = gaussian_filter1d(image.astype(np.float64), 1.5, axis=0, mode="nearest", truncate=5 / 1.5)
        return gaussian_filter1d(out, 1.5, axis=1, mode="nearest", truncate=5 / 1.5)

    radius = len(kernel) // 2
    padded = np.pad(image, ((radius, radius), (0, 0)), mode="edge")
    out = np.zeros_like(image, dtype=np.float64)
    for i, weight in enumerate(kernel):
        out += weight * padded[i:i + image.shape[0], :]
    padded = np.pad(out, ((0, 0), (radius, radius)), mode="edge")
    out = np.zeros_like(image, dtype=np.float64)
    for i, weight in enumerate(kernel):
        out += weight * padded[:, i:i + image.shape[1]]
    return out


def ssim_map(a, b, data_range: float = 255.0):
    import numpy as np

    a = a.astype(np.float64)
    b = b.astype(np.float64)
    kernel = _gaussian_kernel()
    c1 = (0.01 * data_range) ** 2
    c2 = (0.03 * data_range) ** 2
    mu_a, mu_b = _blur(a, kernel), _blur(b, kernel)
    var_a = np.maximum(_blur(a * a, kernel) - mu_a * mu_a, 0.0)
    var_b = np.maximum(_blur(b * b, kernel) - mu_b * mu_b, 0.0)
    cov = _blur(a * b, kernel) - mu_a * mu_b
    return ((2 * mu_a * mu_b + c1) * (2 * cov + c2)) / ((mu_a ** 2 + mu_b ** 2 + c1) * (var_a + var_b + c2))


def masked_ssim(a, b, mask, data_range: float = 255.0):
    """``ssim_map(a, b)[mask]``, bit for bit, with the per-pixel arithmetic after the blurs
    done for the masked pixels only.  Every step of :func:`ssim_map` after its blurs is
    elementwise -- each output element the same IEEE operations on the same input
    elements, in the same order -- so gathering the blurred fields first changes no value
    and none of their order."""
    import numpy as np

    a = a.astype(np.float64)
    b = b.astype(np.float64)
    kernel = _gaussian_kernel()
    c1 = (0.01 * data_range) ** 2
    c2 = (0.03 * data_range) ** 2
    mu_a, mu_b = _blur(a, kernel)[mask], _blur(b, kernel)[mask]
    var_a = np.maximum(_blur(a * a, kernel)[mask] - mu_a * mu_a, 0.0)
    var_b = np.maximum(_blur(b * b, kernel)[mask] - mu_b * mu_b, 0.0)
    cov = _blur(a * b, kernel)[mask] - mu_a * mu_b
    return ((2 * mu_a * mu_b + c1) * (2 * cov + c2)) / ((mu_a ** 2 + mu_b ** 2 + c1) * (var_a + var_b + c2))


def _gray(rgb):
    """``rgb.mean(axis=2)``, bit for bit: the channels' sum is a whole number either way
    (exact in float64), and the mean divides it by 3 once, correctly rounded, as this
    does -- without numpy's reduction along a three-element axis."""
    import numpy as np

    total = rgb[:, :, 0].astype(np.uint16)
    total += rgb[:, :, 1]
    total += rgb[:, :, 2]
    return total / 3.0


def _mean_abs_difference(a_rgb, b_rgb) -> float:
    """``np.abs(a - b).mean()`` in int16, bit for bit: the sum of whole numbers is exact
    either way (below 2**53), and the mean is that sum divided by the count, correctly
    rounded, as this is."""
    import numpy as np

    difference = np.maximum(a_rgb, b_rgb)
    difference -= np.minimum(a_rgb, b_rgb)
    return int(difference.sum(dtype=np.uint64)) / difference.size


def _histogram_vectors(a_rgb, b_rgb, mask) -> list:
    import numpy as np

    vectors = []
    for image in (a_rgb, b_rgb):
        parts = []
        for channel in range(3):
            counts, _ = np.histogram(image[:, :, channel][mask], bins=64, range=(0, 256))
            parts.append(counts.astype(np.float64))
        vectors.append(np.concatenate(parts))
    return vectors


def _correlation(vectors) -> float:
    import numpy as np

    first, second = vectors[0] - vectors[0].mean(), vectors[1] - vectors[1].mean()
    denominator = np.sqrt((first ** 2).sum() * (second ** 2).sum())
    return 1.0 if denominator == 0 else float((first * second).sum() / denominator)


def histogram_correlation(a_rgb, b_rgb, mask) -> float:
    return _correlation(_histogram_vectors(a_rgb, b_rgb, mask))


def content_box(mask, margin: int = SSIM_RADIUS) -> tuple[slice, slice]:
    """The rows and columns SSIM and the histograms have to be computed over for the pixels
    of ``mask``: the bounding box of ``mask`` grown by ``margin`` on every side, clamped to
    the page.

    ``mask`` is the foreground of :func:`score` -- the **union** of both images' ink, never
    one side's.  Exact, not approximate: :func:`_blur` is separable, so a pixel's blurred
    value is a fixed-order sum over the pixels within :data:`SSIM_RADIUS` of it, and with
    ``margin`` at least that, every one of them lies inside the box -- or, at a clamped
    edge, is the page's own edge row, which the blur's edge padding replicates identically
    either way.  Every masked pixel's SSIM is therefore the same float cropped or not, and
    the masked values come out in the same (row-major) order, so the mean and the loss are
    the same sums; the histograms count the same pixels.  ``tests/test_fidelity.py`` holds
    that on every scored page, and every run re-checks a sample (:func:`score_corpus`).

    A page whose foreground reaches every edge -- a page colour, a full-bleed picture --
    gets the whole page back: no speed-up, and nothing cut.  What needs the whole page --
    the foreground's coverage and the mean absolute difference -- is never cropped."""
    import numpy as np

    rows = np.flatnonzero(mask.any(axis=1))
    cols = np.flatnonzero(mask.any(axis=0))
    if not len(rows):
        return slice(0, mask.shape[0]), slice(0, mask.shape[1])
    return (slice(max(int(rows[0]) - margin, 0), min(int(rows[-1]) + margin + 1, mask.shape[0])),
            slice(max(int(cols[0]) - margin, 0), min(int(cols[-1]) + margin + 1, mask.shape[1])))


class CropMismatch(Exception):
    """A page whose SSIM or histograms cropped to the content are not the full page's, bit
    for bit."""


def _score(ours_rgb, truth_rgb, crop: bool):
    """:func:`score`'s row, with the masked SSIM values and the histogram vectors it was
    computed from (``None`` for a page too sparse to score)."""
    import numpy as np

    if crop:
        ours_gray, truth_gray = _gray(ours_rgb), _gray(truth_rgb)
        mask = (ours_gray < FOREGROUND_MAX) | (truth_gray < FOREGROUND_MAX)
        coverage = int(np.count_nonzero(mask)) / mask.size
        mean_abs = _mean_abs_difference(ours_rgb, truth_rgb)
    else:  # the literal full-page computation, which the cropped one is held to
        ours_gray, truth_gray = ours_rgb.mean(axis=2), truth_rgb.mean(axis=2)
        mask = (ours_gray < FOREGROUND_MAX) | (truth_gray < FOREGROUND_MAX)
        coverage = float(mask.mean())
        mean_abs = float(np.abs(ours_rgb.astype(np.int16) - truth_rgb.astype(np.int16)).mean())
    result = {"coverage": round(coverage, 4), "mean_abs": round(mean_abs, 3)}
    if coverage < MIN_FOREGROUND:
        result.update(ssim=1.0, histogram=1.0, loss=0.0, sparse=True)
        return result, None, None
    if crop:
        box = content_box(mask)
        ssim = masked_ssim(ours_gray[box], truth_gray[box], mask[box])
    else:
        box = (slice(None), slice(None))
        ssim = ssim_map(ours_gray, truth_gray)[mask]
    vectors = _histogram_vectors(ours_rgb[box], truth_rgb[box], mask[box])
    result.update(ssim=round(float(ssim.mean()), 4), histogram=round(_correlation(vectors), 4),
                  loss=round(float((1 - ssim).sum()), 1), sparse=False)
    return result, ssim, vectors


def score(ours_rgb, truth_rgb, crop: bool = True) -> dict:
    """SSIM, histogram correlation, the unnormalised structural loss and the mean
    absolute difference over the foreground, for one page.

    ``crop`` computes SSIM and the histograms over :func:`content_box` only, which gives
    the full page's numbers bit for bit (see there), and the rest by the equivalent
    shortcuts :func:`_gray`, :func:`_mean_abs_difference` and :func:`masked_ssim`;
    ``crop=False`` is the literal full-page computation, kept for the checks that hold the
    two equal (:func:`score_checked`)."""
    return _score(ours_rgb, truth_rgb, crop)[0]


def score_checked(ours_rgb, truth_rgb, label: str = "") -> dict:
    """:func:`score`, computed cropped *and* over the full page, and held equal: the masked
    SSIM values and the histogram vectors bit for bit, and the rows.  Raises
    :class:`CropMismatch` on any difference."""
    row, ssim, vectors = _score(ours_rgb, truth_rgb, True)
    full_row, full_ssim, full_vectors = _score(ours_rgb, truth_rgb, False)
    same = row == full_row and (ssim is None) == (full_ssim is None)
    if same and ssim is not None:
        same = (ssim.shape == full_ssim.shape and ssim.tobytes() == full_ssim.tobytes()
                and all(a.tobytes() == b.tobytes() for a, b in zip(vectors, full_vectors)))
    if not same:
        raise CropMismatch(f"{label}: cropped {row} but full page {full_row}")
    return row


#: What a page only one side made is, in its row's ``unmatched``.
OURS_ONLY, WORD_ONLY = "ours only", "Word only"


def fully_wrong(image, unmatched: str) -> dict:
    """The row of a page only one side made (``unmatched``: :data:`OURS_ONLY`, a page our
    layout makes past Word's last, or :data:`WORD_ONLY`, one of Word's past ours): scored
    as wholly wrong -- SSIM and histogram correlation 0, the structural loss every
    foreground pixel of the page that exists (``1 - 0`` over the mask), whatever it holds,
    an empty page included -- so that a document with a page too many or too few scores
    below one with Word's pages, and a mismatch cannot pass unseen.  ``image`` is the
    existing page's raster; the mean absolute difference is taken against white paper."""
    import numpy as np

    gray = _gray(image)
    mask = gray < FOREGROUND_MAX
    coverage = int(np.count_nonzero(mask)) / mask.size
    return {"coverage": round(coverage, 4), "mean_abs": round(float((255 - image.astype(np.int16)).mean()), 3),
            "ssim": 0.0, "histogram": 0.0, "loss": float(np.count_nonzero(mask)),
            "sparse": coverage < MIN_FOREGROUND, "unmatched": unmatched}


# -- the two rasters -------------------------------------------------------------------------


def word_pages(pdf: Path, truth: str = "svg") -> list:
    """Word's pages at :data:`DPI`, as RGB arrays: through ``tools/pdf_svg.py`` and resvg
    (``svg``) or through pdfium (``pdfium``)."""
    import numpy as np
    import pypdfium2 as pdfium

    if truth == "svg":
        import pdf_svg

        return pdf_svg.truth_pages(pdf)
    if truth != "pdfium":
        raise ValueError(f"unknown truth {truth!r}: one of {TRUTHS}")

    document = pdfium.PdfDocument(str(pdf))
    out = []
    for page in document:
        image = page.render(scale=DPI / 72).to_pil().convert("RGB")
        out.append(np.asarray(image))
    return out


def drawn_face_files(layout, data: bytes, directory: str) -> tuple[list[str], str | None]:
    """The files the faces of ``layout`` are drawn with (``docx2svg.fonts.rasteriser_files``:
    Word's copies, read in place, and the document's embedded faces written, as embedded,
    into ``directory``); or, for a face this machine lacks or the rasteriser cannot tell
    from another, the reason."""
    from docx2svg import drawn_faces
    from docx2svg.fonts import InstalledFonts, rasteriser_files

    fonts = InstalledFonts(data)
    files, missing, unaddressable = rasteriser_files(fonts, drawn_faces(layout), directory)
    if missing:
        return [], f"no face {missing[0]!r} on this machine"
    if unaddressable:
        return [], f"face {unaddressable[0]!r} is not addressable among the files drawn with"
    arial = fonts.drawing_face("Arial")
    if arial is not None and arial.source not in files:
        files.append(arial.source)
    return files, None


def word_page(pdf: Path, index: int, truth: str = "svg"):
    """Page ``index`` of :func:`word_pages`, rasterised alone: the same pixels."""
    import numpy as np
    import pypdfium2 as pdfium

    if truth == "svg":
        import pdf_svg

        return pdf_svg.rasterise(pdf_svg.page_svgs(pdf)[index])
    if truth != "pdfium":
        raise ValueError(f"unknown truth {truth!r}: one of {TRUTHS}")
    return np.asarray(pdfium.PdfDocument(str(pdf))[index].render(scale=DPI / 72).to_pil().convert("RGB"))


def word_page_count(pdf: Path, truth: str = "svg") -> int:
    """``len(word_pages(pdf, truth))``, rasterising nothing."""
    import pypdfium2 as pdfium

    if truth == "svg":
        import pdf_svg

        return len(pdf_svg.page_svgs(pdf))
    if truth != "pdfium":
        raise ValueError(f"unknown truth {truth!r}: one of {TRUTHS}")
    return len(pdfium.PdfDocument(str(pdf)))


def _our_svgs(data: bytes, glyph_size: str, directory: str):
    """Our layout of ``data``, its pages as SVG at :data:`DPI` (the SVG's own device-pixel
    size), and the files its faces are drawn with (:func:`drawn_face_files`, into
    ``directory``) or the reason there are none."""
    from docx2svg import ConvertOptions, _lay_out, _svgs

    options = ConvertOptions(glyph_size=glyph_size)
    layout, data, fonts = _lay_out(data, options)
    documents = _svgs(layout, data, options, device_size=True, fonts=fonts)
    files, missing = drawn_face_files(layout, data, directory)
    return layout, documents, files, missing


#: How our SVG is rasterised, besides its font files: shared by :func:`_our_raster` and
#: the raster cache's key (:func:`our_components`), so the two cannot drift apart.
OUR_RASTER_OPTIONS = {"backend": "resvg", "skip_system_fonts": True, "sans_serif_family": "Arial"}


def _our_raster(document: str, files: list[str]):
    """One page of ours (:func:`_our_svgs`), rasterised with its faces, as an RGB array."""
    import numpy as np
    from PIL import Image

    from docx2svg.png import svg_to_png

    png = svg_to_png(document, font_files=files, **OUR_RASTER_OPTIONS)
    return np.asarray(Image.open(io.BytesIO(png)).convert("RGB"))


def our_pages(data: bytes, glyph_size: str):
    """Our pages at :data:`DPI` (the SVG's own device-pixel size), and the layout."""
    with tempfile.TemporaryDirectory(prefix="docx2svg-fidelity-") as directory:
        layout, documents, files, missing = _our_svgs(data, glyph_size, directory)
        if missing:
            return None, layout, missing
        out = [_our_raster(document, files) for document in documents]
    return out, layout, None


def _same_shape(image, shape):
    """``image`` cropped or padded with white to ``shape``: pdfium renders a page one pixel
    taller or wider than its box when the box times the scale is not whole (792 pt at
    300 dpi comes out 3,301 rows).  Resampling instead would blur every glyph."""
    import numpy as np

    out = np.full(shape, 255, dtype=image.dtype)
    rows, cols = min(shape[0], image.shape[0]), min(shape[1], image.shape[1])
    out[:rows, :cols] = image[:rows, :cols]
    return out


def score_document(name: str, data: bytes, pdf: Path, glyph_sizes=GLYPH_SIZES,
                   max_pages: int | None = None, truth: str = "svg", cache_mode: str = "off",
                   cache_root: Path | None = None) -> dict:
    """Per glyph size: every page's score, and the means over the pages the layout draws
    in full (no stop band) and over every page it made; Word's page rasterised as
    ``truth`` says (:data:`TRUTHS`)."""
    return score_truths(name, data, pdf, glyph_sizes, max_pages, (truth,), cache_mode, cache_root)[truth]


def score_truths(name: str, data: bytes, pdf: Path, glyph_sizes=GLYPH_SIZES,
                 max_pages: int | None = None, truths=TRUTHS, cache_mode: str = "off",
                 cache_root: Path | None = None) -> dict:
    """:func:`score_document` for each of ``truths``, rendering ours once: one document
    through :func:`score_corpus`, in this process."""
    return score_corpus([(name, data, pdf)], glyph_sizes, max_pages, truths, 1, cache_mode, cache_root)[0]


def _entry(pages: list[dict], word_pages: int, pages_made: int) -> dict:
    """One truth's and glyph size's result from its page rows: ``pages_made`` is our
    layout's page count, ``word_pages`` Word's.  The rows include a page only one side
    made (:func:`fully_wrong`), so a page too many or too few lowers every mean."""
    full = [p for p in pages if not p["stopped"]]

    def mean(rows, key):
        return round(sum(r[key] for r in rows) / len(rows), 4) if rows else None

    return {
        "pages": pages,
        "ssim": mean(full, "ssim"), "histogram": mean(full, "histogram"), "loss": mean(full, "loss"),
        "ssim_all": mean(pages, "ssim"), "pages_scored": len(full), "pages_made": pages_made,
        "word_pages": word_pages,
    }


# -- the raster cache's keys (tools/raster_cache.py) -------------------------------------

#: Bump when anything this file does to a raster changes in a way no other component of
#: the key records.  Part of every cache key.
HARNESS_VERSION = 1

#: Pages a run re-draws to hold the raster cache to what drawing gives today, and to hold
#: the cropped SSIM to the full page's (:func:`score_corpus`): spread through the corpus
#: and rotated by date, so that every page is re-drawn within ``ceil(pages /
#: VERIFY_SAMPLE)`` days -- nine, at the 52 pages scored today.
VERIFY_SAMPLE = 6

#: How a run uses the raster cache (``--no-cache``, ``--verify-cache``): ``use`` reads and
#: fills it, and re-draws :data:`VERIFY_SAMPLE` pages to hold it to what drawing gives;
#: ``verify`` re-draws every page and holds every cached raster so; ``read`` reads it and
#: writes nothing (the tests); ``off`` neither reads nor writes it.
CACHE_MODES = ("use", "verify", "read", "off")


def raster_cache_root() -> Path:
    """Where rasters are cached: beside the converted pages in the oracle directory,
    outside any repository -- Word's rasters hold Microsoft's glyph shapes, and ours are
    drawn with its fonts."""
    import oracle

    return oracle.ORACLE_DIR / "svg" / "rasters"


#: Content digests read in this run, by path and ``stat`` (cleared by :func:`_init_worker`,
#: so a run never trusts a digest from before it started).
_DIGESTS: dict = {}


def content_digest(path) -> str:
    """The sha256 of a file's *content* -- never its path or date: an Office update can
    change a face in place -- read once per run."""
    import raster_cache

    stat = os.stat(path)
    memo = (str(path), stat.st_size, stat.st_mtime_ns, stat.st_ino)
    if memo not in _DIGESTS:
        _DIGESTS[memo] = raster_cache.digest_file(path)
    return _DIGESTS[memo]


_VERSIONS: dict = {}


def _rasteriser_versions() -> dict:
    """resvg-py's version and the digest of its compiled module; Pillow's version (it
    decodes resvg's PNG into the array that is cached)."""
    if not _VERSIONS:
        import importlib.metadata

        import PIL
        import resvg_py

        binaries = sorted(path for path in Path(resvg_py.__file__).parent.iterdir()
                          if path.suffix in (".so", ".pyd", ".dylib"))
        _VERSIONS.update(resvg=importlib.metadata.version("resvg-py"),
                         resvg_binary=[content_digest(path) for path in binaries], pillow=PIL.__version__)
    return dict(_VERSIONS)


def _font_name(path: str, directory: str | None) -> str:
    """A font file as the key names it: in place, by its path; written for the page (a
    document's embedded face, :func:`drawn_face_files`), by its name in the temporary
    directory, which is another directory on every run."""
    if directory is not None and Path(path).parent == Path(directory):
        return "<embedded>/" + Path(path).name
    return str(path)


def our_components(svg: str, files: list[str], directory: str | None = None) -> dict:
    """Every input that decides our raster's pixels: the SVG's bytes; the font files resvg
    is handed, in order, each by name and *content*; the options; the code that turns them
    into resvg's call (``docx2svg/png.py``); resvg and Pillow; the resolution and
    :data:`HARNESS_VERSION`."""
    import raster_cache

    import docx2svg.png

    return {
        "side": "ours", "harness": HARNESS_VERSION, "dpi": DPI,
        "svg": raster_cache.digest_bytes(svg.encode("utf-8")),
        "options": dict(sorted(OUR_RASTER_OPTIONS.items())),
        "font_files": [[_font_name(path, directory), content_digest(path)] for path in files],
        "png_py": content_digest(docx2svg.png.__file__),
        **_rasteriser_versions(),
    }


def truth_components(pdf: Path, index: int, svg: str) -> dict:
    """Every input that decides Word's raster under the ``svg`` truth: the PDF's bytes and
    the page, the converted page's bytes, the converter's version and source
    (``tools/pdf_svg.py``, which also rasterises it), PyMuPDF's version, the resolution and
    options, resvg and Pillow, and :data:`HARNESS_VERSION`."""
    import pdf_svg
    import pymupdf
    import raster_cache

    return {
        "side": "truth", "harness": HARNESS_VERSION, "dpi": pdf_svg.DPI,
        "pdf": content_digest(pdf), "page": index,
        "svg": raster_cache.digest_bytes(svg.encode("utf-8")),
        "converter": pdf_svg.converter_version(),
        "pdf_svg_py": content_digest(pdf_svg.__file__),
        "pymupdf": pymupdf.VersionBind,
        "options": {"device_grid": True, "zoom": 1.0, "background": "white", "skip_system_fonts": True},
        **_rasteriser_versions(),
    }


# -- in parallel -------------------------------------------------------------------------

#: What one scoring worker may hold at once, with headroom (``--jobs`` defaults to no more
#: workers than the machine's memory holds at this size).  Measured, a page of the
#: committed documents (A4 and Letter at 300 dpi) scored at both glyph sizes against both
#: truths: 1.88 GB peak footprint, 2.59 GB peak resident -- the float64 SSIM maps of
#: 8.7 million pixels.  The converter's validation has its own (``pdf_svg.WORKER_MEMORY``).
WORKER_MEMORY = 3_000_000_000


def default_jobs() -> int:
    """Every logical core, unless memory holds fewer scoring workers of :data:`WORKER_MEMORY`."""
    return parallel.default_jobs(WORKER_MEMORY)


def _refuse_export(*_args, **_kwargs):
    raise RuntimeError("a worker process may not export through Word: exports are looked up before the pool starts")


def no_word() -> None:
    """A pool worker's initializer: :func:`oracle.export` raises there.  Word is one
    instance per machine, and an export opens a document and then exports the *active*
    one, so two at once would each export the other's pages -- and each begins by quitting
    Word under the other.  (In the calling process, where ``--jobs 1`` runs, it does
    nothing: the serial path is the serial path.)"""
    import multiprocessing

    if multiprocessing.parent_process() is None:
        return
    import oracle

    oracle.export = _refuse_export


#: A scoring process's raster cache and how it uses it (:data:`CACHE_MODES`), set by
#: :func:`_init_worker`.
_WORKER: dict = {}


def _init_worker(cache_root: str | None = None, cache_mode: str = "off") -> None:
    """A pool's initializer (and, at ``--jobs 1``, this process's): :func:`no_word`, and
    the raster cache as the run uses it."""
    no_word()
    _DIGESTS.clear()
    _WORKER["cache_mode"] = cache_mode
    if cache_root is not None and cache_mode != "off":
        import raster_cache

        _WORKER["cache"] = raster_cache.RasterCache(Path(cache_root), REPO)
    else:
        _WORKER.pop("cache", None)


#: A worker's render of the document it is scoring, per glyph size (:func:`_our_svgs`),
#: kept while its pages come in: the tasks are handed out in page order, so a worker
#: mostly scores consecutive pages of one document, and laying a 36-page document out
#: again for every page would cost more than rasterising it.  The key is the document's
#: SHA-256; the directory holding its embedded faces lives as long as the entry.
_RENDERED: dict = {}


def _rendered(data: bytes, glyph_size: str):
    """:func:`_our_svgs` of ``data`` at ``glyph_size``, from this worker's cache, and the
    directory its embedded faces are written to."""
    import hashlib

    key = (hashlib.sha256(data).hexdigest(), glyph_size)
    if key not in _RENDERED:
        for digest, _size in list(_RENDERED):
            if digest != key[0]:
                _RENDERED.pop((digest, _size))[0].cleanup()
        directory = tempfile.TemporaryDirectory(prefix="docx2svg-fidelity-")
        _RENDERED[key] = (directory, *_our_svgs(data, glyph_size, directory.name))
    return (*_RENDERED[key][1:], _RENDERED[key][0].name)


#: A worker's converted pages of the export it is scoring (``pdf_svg.page_svgs``), kept
#: while its pages come in rather than read from the converter's cache for every page.
_TRUTH_SVGS: dict = {}


def _truth_svgs(pdf: Path) -> list[str]:
    import pdf_svg

    key = (str(Path(pdf).resolve()), Path(pdf).stat().st_mtime_ns)
    if key not in _TRUTH_SVGS:
        _TRUTH_SVGS.clear()
        _TRUTH_SVGS[key] = pdf_svg.page_svgs(pdf)
    return _TRUTH_SVGS[key]


def _plan(task) -> dict:
    """What :func:`score_truths` knows of a document before it scores a page: whether it is
    skipped, and how many pages each side has -- ours per glyph size, with the stop flags;
    Word's per truth.  Rasterises nothing."""
    data, pdf, glyph_sizes, truths = task
    word = {truth: word_page_count(pdf, truth) for truth in truths}
    ours = {}
    for glyph_size in glyph_sizes:
        layout, _documents, _files, missing, _directory = _rendered(data, glyph_size)
        if missing:
            return {"skipped": missing}
        ours[glyph_size] = len(layout.pages)
    return {"ours": ours, "word": word}


def _raster(kind: str, components_of, draw, verify: bool, checks: list, label: str):
    """A raster from the worker's cache (:data:`_WORKER`), or drawn -- and, when
    ``verify``, drawn *and* compared byte for byte with what the cache holds."""
    cache, mode = _WORKER.get("cache"), _WORKER.get("cache_mode", "off")
    if cache is None or mode == "off":
        return draw()
    components = components_of()
    if verify:
        fresh = draw()
        status = cache.verify(kind, components, fresh, store=mode != "read", label=label)
        checks.append({"page": label, "kind": kind, "cache": status})
        return fresh
    cached = cache.get(kind, components)
    if cached is not None:
        checks.append({"page": label, "kind": kind, "cache": "hit"})
        return cached
    fresh = draw()
    if mode != "read":
        cache.put(kind, components, fresh, label)
    checks.append({"page": label, "kind": kind, "cache": "drawn"})
    return fresh


def _page_rows(data: bytes, pdf: Path, index: int, limits: dict, label: str, recheck: bool, crop_check: bool,
               checks: list) -> dict:
    """Page ``index`` of one document: ``{truth: {glyph size: row}}`` for each pair whose
    page :func:`score_truths` scores (``limits``: per glyph size and truth, ``(both, ours,
    word)`` -- the pages both sides made, and each side's count, all at most ``max_pages``).
    Ours is rasterised once per glyph size, Word's once per truth, each from the cache where
    it holds them (:func:`_raster`).  A page only one side made is scored
    :func:`fully_wrong` from the side that made it.

    The spot-checks, each recorded in ``checks`` and raised on a mismatch: ``recheck``
    re-draws the page's rasters and holds them to the cache byte for byte; ``crop_check``
    scores each pair cropped and over the full page and holds the two equal, bit for bit
    (:func:`score_checked`)."""
    import pdf_svg

    references: dict = {}
    out: dict = {}

    def reference_of(truth):
        if truth not in references:
            if truth == "svg":
                page = _truth_svgs(pdf)[index]
                references[truth] = _raster("truth", lambda: truth_components(pdf, index, page),
                                            lambda: pdf_svg.rasterise(page), recheck, checks, label)
            else:
                references[truth] = word_page(pdf, index, truth)
        return references[truth]

    for glyph_size, by_truth in limits.items():
        wanted = [truth for truth, (both, ours, word) in by_truth.items() if index < max(ours, word)]
        if not wanted:
            continue
        image = None
        if any(index < ours for _both, ours, _word in by_truth.values()):
            layout, documents, files, _missing, directory = _rendered(data, glyph_size)
            svg = documents[index]
            image = _raster("ours", lambda: our_components(svg, files, directory), lambda: _our_raster(svg, files),
                            recheck, checks, f"{label} ({glyph_size})")
        for truth in wanted:
            both, ours, word = by_truth[truth]
            if index >= both:
                if index < ours:
                    row = dict(fully_wrong(image, OURS_ONLY), stopped=layout.pages[index].stop is not None)
                else:
                    row = dict(fully_wrong(reference_of(truth), WORD_ONLY), stopped=False)
                out.setdefault(truth, {})[glyph_size] = row
                continue
            reference = _same_shape(reference_of(truth), image.shape)
            if crop_check:
                row = score_checked(image, reference, f"{label} ({glyph_size}, {truth} truth)")
                checks.append({"page": label, "kind": f"crop-{glyph_size}-{truth}", "crop": "equal"})
            else:
                row = score(image, reference)
            row["stopped"] = layout.pages[index].stop is not None
            out.setdefault(truth, {})[glyph_size] = row
    return out


def _score_page(task) -> dict:
    """One page, as a pool task: its rows (:func:`_page_rows`), and what its checks found.
    A mismatch is returned rather than raised, so that :func:`score_corpus` can name every
    one of them."""
    import raster_cache

    data, pdf, index, limits, label, recheck, crop_check = task
    checks: list = []
    try:
        rows = _page_rows(data, pdf, index, limits, label, recheck, crop_check, checks)
    except raster_cache.CacheMismatch as mismatch:
        return {"rows": None, "checks": checks, "cache_mismatch": f"{label}: {mismatch}"}
    except CropMismatch as mismatch:
        return {"rows": None, "checks": checks, "crop_mismatch": str(mismatch)}
    return {"rows": rows, "checks": checks}


def score_corpus(documents, glyph_sizes=GLYPH_SIZES, max_pages: int | None = None, truths=TRUTHS,
                 jobs: int = 1, cache_mode: str = "off", cache_root: Path | None = None, stats: dict | None = None,
                 day: int | None = None, crop_all: bool = False) -> list[dict]:
    """:func:`score_truths` of each ``(name, data, pdf)`` of ``documents``, in order, over
    ``jobs`` processes, a page per task: the serial results bit for bit.

    Every export is converted first (the ``svg`` truth: one task per PDF, into the cache),
    so that no two workers convert one PDF at once; then each document is planned
    (:func:`_plan`, a task each) and its pages scored (:func:`_score_page`).

    Rasters come from the raster cache (``cache_root``, by default :func:`raster_cache_root`)
    as ``cache_mode`` says (:data:`CACHE_MODES`, ``tools/raster_cache.py``).  Every run
    spot-checks :data:`VERIFY_SAMPLE` pages, chosen by ``raster_cache.rotating_sample`` for
    ``day`` (today): each is re-drawn and held to the cache byte for byte, and scored
    cropped and full-page and held equal.  ``verify`` re-draws every page, and ``crop_all``
    holds every page's crop.  **Any cache mismatch discards the whole cache and raises**
    ``raster_cache.CacheMismatch``; a crop mismatch raises :class:`CropMismatch`.  Neither
    is a warning: nothing is scored.  ``stats``, when given, receives what the checks
    found."""
    import raster_cache

    if cache_mode not in CACHE_MODES:
        raise ValueError(f"unknown cache mode {cache_mode!r}: one of {CACHE_MODES}")
    documents = list(documents)
    cache = None
    if cache_mode != "off":
        cache = raster_cache.RasterCache(Path(cache_root) if cache_root is not None else raster_cache_root(), REPO)
    init = (str(cache.root) if cache is not None else None, cache_mode)
    if "svg" in truths:
        import pdf_svg

        pending = [pdf for pdf in sorted({pdf for _name, _data, pdf in documents})
                   if not (pdf_svg.cache_dir(pdf) / "done").exists()]
        pool_map(pdf_svg.convert_task, pending, jobs, _init_worker, init)
    plans = pool_map(_plan, [(data, pdf, tuple(glyph_sizes), tuple(truths)) for _name, data, pdf in documents],
                     jobs, _init_worker, init)
    tasks, owners = [], []
    for number, ((name, data, pdf), plan) in enumerate(zip(documents, plans)):
        if "skipped" in plan:
            continue
        # Exactly the pages score_truths scores: every page either side made, up to
        # max_pages -- the pages both made compared, the rest scored as wholly wrong.
        cap = max_pages if max_pages is not None else max(plan["word"].values()) + max(plan["ours"].values())
        limits = {glyph_size: {truth: (min(plan["ours"][glyph_size], plan["word"][truth], cap),
                                       min(plan["ours"][glyph_size], cap), min(plan["word"][truth], cap))
                               for truth in truths} for glyph_size in glyph_sizes}
        for index in range(max(max(ours, word) for by_truth in limits.values()
                               for _both, ours, word in by_truth.values())):
            tasks.append((data, pdf, index, limits, f"{name} page {index + 1}"))
            owners.append(number)
    everything = set(range(len(tasks)))
    sampled = set(raster_cache.rotating_sample(list(range(len(tasks))), VERIFY_SAMPLE, day))
    recheck = everything if cache_mode == "verify" else sampled
    cropped = everything if crop_all or cache_mode == "verify" else sampled
    outcomes = pool_map(_score_page, [task + (number in recheck, number in cropped)
                                      for number, task in enumerate(tasks)], jobs, _init_worker, init)

    checks = [check for outcome in outcomes for check in outcome["checks"]]
    if stats is not None:
        stats.update(pages=len(tasks), sampled=[task[4] for number, task in enumerate(tasks)
                                                if number in recheck | cropped], checks=checks)
    cache_mismatches = [outcome["cache_mismatch"] for outcome in outcomes if "cache_mismatch" in outcome]
    if cache_mismatches:
        cache.discard()
        raise raster_cache.CacheMismatch(
            f"the raster cache under {cache.root} does not hold what drawing gives today:\n  "
            + "\n  ".join(cache_mismatches)
            + "\nSome input that moves pixels is missing from the cache key (tools/fidelity.py, "
            "our_components / truth_components), or the cache was corrupted.  The whole cache has been "
            "discarded and nothing was scored.  Find the input before trusting a cached run again; "
            "--verify-cache re-draws every page and compares.")
    crop_mismatches = [outcome["crop_mismatch"] for outcome in outcomes if "crop_mismatch" in outcome]
    if crop_mismatches:
        raise CropMismatch("SSIM cropped to the content is not the full page's:\n  " + "\n  ".join(crop_mismatches)
                           + "\nNothing was scored.  content_box() must never cut a pixel the score depends on.")
    rows = [outcome["rows"] for outcome in outcomes]

    out = []
    for number, plan in enumerate(plans):
        if "skipped" in plan:
            out.append({truth: {"skipped": plan["skipped"]} for truth in truths})
            continue
        pages = [row for owner, row in zip(owners, rows) if owner == number]
        out.append({truth: {glyph_size: _entry([page[truth][glyph_size] for page in pages
                                                if glyph_size in page.get(truth, {})], plan["word"][truth],
                                               plan["ours"][glyph_size])
                            for glyph_size in glyph_sizes} for truth in truths})
    return out


def _median(values: list[float]) -> float:
    ordered = sorted(values)
    middle = len(ordered) // 2
    return ordered[middle] if len(ordered) % 2 else (ordered[middle - 1] + ordered[middle]) / 2


def _scored_pages(results: dict, glyph_size: str) -> dict:
    """The pages both sides made, drawn in full and dense enough to score, per document:
    what the medians and the outliers are taken over.  A page only one side made is
    reported by :func:`page_count_mismatches` instead."""
    return {name: [(index, page) for index, page in enumerate(result[glyph_size]["pages"])
                   if not page["stopped"] and not page["sparse"] and "unmatched" not in page]
            for name, result in results.items() if "skipped" not in result}


def page_count_mismatches(results: dict, glyph_size: str = "device") -> list[dict]:
    """The documents of ``results`` whose layout makes another number of pages than Word:
    a regression whatever the scores say, since every page after the first that differs is
    somewhere else.  The pages only one side made are in the scores as wholly wrong
    (:func:`fully_wrong`)."""
    out = []
    for name, result in results.items():
        if "skipped" in result:
            continue
        entry = result[glyph_size]
        if entry["pages_made"] != entry["word_pages"]:
            unmatched = [index + 1 for index, page in enumerate(entry["pages"]) if "unmatched" in page]
            out.append({"document": name, "glyph_size": glyph_size, "ours": entry["pages_made"],
                        "word": entry["word_pages"], "unmatched": unmatched})
    return out


def describe_page_count_mismatch(row: dict) -> str:
    pages = row["unmatched"]
    span = (f"page {pages[0]}" if len(pages) == 1 else f"pages {pages[0]}-{pages[-1]}") if pages else "no page"
    side = "ours only" if row["ours"] > row["word"] else "Word's only"
    return (f"PAGE COUNT MISMATCH {row['document']} ({row['glyph_size']}): ours {row['ours']} pages, Word "
            f"{row['word']} -- {span} ({side}) scored as wholly wrong (SSIM 0); a regression")


def page_median(results: dict, glyph_size: str = "device") -> float | None:
    """The median SSIM of every page of ``results`` drawn in full and dense enough to
    score: the corpus's typical page (recorded with ``--record``)."""
    everything = [page["ssim"] for rows in _scored_pages(results, glyph_size).values() for _, page in rows]
    return round(_median(everything), 4) if everything else None


def outliers(results: dict, glyph_size: str = "device", corpus: float | None = None,
             drop: float = OUTLIER_DROP) -> list[dict]:
    """The pages of ``results`` (``{name: score_document(...)}``) drawn in full and not
    too sparse to score whose SSIM is more than ``drop`` (:data:`OUTLIER_DROPS` of the
    truth scored) below the median page
    of their document *or* of the corpus (:func:`page_median` of ``results``, or
    ``corpus``), with their loss and mean absolute difference: each one is to be
    explained, not recorded."""
    typical = page_median(results, glyph_size) if corpus is None else corpus
    out = []
    for name, rows in _scored_pages(results, glyph_size).items():
        if not rows:
            continue
        own = _median([page["ssim"] for _, page in rows])
        for index, page in rows:
            if page["ssim"] < max(own, typical) - drop:
                out.append({"document": name, "page": index + 1, "glyph_size": glyph_size, "ssim": page["ssim"],
                            "document_median": round(own, 4), "corpus_median": round(typical, 4),
                            "loss": page["loss"], "mean_abs": page["mean_abs"]})
    return out


def describe_outlier(row: dict) -> str:
    return (f"OUTLIER {row['document']} page {row['page']} ({row['glyph_size']}): SSIM {row['ssim']} against a "
            f"document median {row['document_median']} and a corpus median {row['corpus_median']}; loss "
            f"{row['loss']}, mean |diff| {row['mean_abs']} -- explain it (look at the pictures) before recording")


def corpus_median(baselines: dict, glyph_size: str = "device") -> float:
    """The median of the recorded documents' SSIM means (but a document too sparse to
    score, which counts 1.0): the typical score a document's own is held against."""
    return _median([r[glyph_size]["ssim"] for r in baselines["documents"].values()
                    if "skipped" not in r and r[glyph_size]["ssim"] is not None and r[glyph_size]["loss"]])


COMMITTED = (
    "layout-sweep.docx", "style-document.docx", "samplelib/sample-long.docx", "samplelib/sample-resume.docx",
    "samplelib/sample-simple.docx", "wordto/sample-1page.docx", "wordto/sample-5pages.docx",
    "wordto/sample-10pages.docx", "wordto/sample-with-images.docx", "wordto/sample-with-table.docx",
)


def ink(truth: str = "svg") -> int:
    """``make_ink_probe.py``: per page (one size each), both glyph sizes against Word's page,
    and the size Word's PDF scales the outlines by (Quartz's ``Tm``)."""
    import make_ink_probe
    import oracle
    import quartz_pdf

    data = make_ink_probe.build()
    pdf = oracle.export(data, name="ink")
    drawn = [sorted({run.size_px for run in page.runs}) for page in quartz_pdf.read(pdf)]
    result = score_document("ink", data, pdf, truth=truth)
    print(f"{'size':>8} {'exact px':>9} {'Word draws':>10}  {'SSIM device':>11} {'SSIM exact':>10}"
          f"  {'loss device':>11} {'loss exact':>10}")
    for index, half_points in enumerate(make_ink_probe.SIZES):
        device, exact = result["device"]["pages"][index], result["exact"]["pages"][index]
        print(f"{half_points / 2:>6} pt {half_points / 2 * 300 / 72:>9.3f} {str(drawn[index]):>10}  "
              f"{device['ssim']:>11} {exact['ssim']:>10}  {device['loss']:>11} {exact['loss']:>10}")
    return 0


def _summary(results: dict, glyph_sizes=GLYPH_SIZES) -> dict:
    """What ``--record`` keeps of one truth's results at ``glyph_sizes``: the committed
    documents' means (no pages) and the median page."""
    committed = {n: r for n, r in results.items() if n in COMMITTED}
    return {
        "documents": {name: {size: {k: v for k, v in r[size].items() if k != "pages"}
                             for size in glyph_sizes} if "skipped" not in r else r
                      for name, r in committed.items()},
        "page_median": {size: page_median(committed, size) for size in glyph_sizes},
    }


#: The recording's description (``--record``).
ABOUT = ("Raster fidelity of the committed documents against Word 16.106's PDF at 300 dpi: SSIM "
         "and colour-histogram correlation over foreground pixels, the unnormalised structural "
         "loss, per glyph size; pages the layout draws in full. Both sides draw with Word's "
         "installed faces, read in place. 'documents' and 'page_median' are the default truth, "
         "'svg': Word's page converted by tools/pdf_svg.py (PyMuPDF) and rasterised by resvg, like "
         "ours. 'pdfium' holds the same for Word's PDF rasterised by pdfium, the instrument before "
         "it. Regenerate with tools/fidelity.py --record.")


class KeptBaselineMismatch(Exception):
    """A recorded entry a run that did not score it cannot keep honestly."""


def _keep(what: str, name: str, kept: dict | None, fresh: dict, recorded: dict | None) -> None:
    """Refuse to carry ``kept`` -- a document's recorded entry for a truth or glyph size this
    run did not score -- over beside ``fresh``, the document's summary at the default truth
    as this run scored it (``recorded``: as it was recorded), where keeping it would
    misdescribe it."""
    why = None
    if kept is None:
        why = "nothing is recorded for it"
    elif ("skipped" in fresh) != ("skipped" in kept) or fresh.get("skipped") != kept.get("skipped"):
        why = f"it is {'skipped' if 'skipped' in fresh else 'scored'} now, and was {'skipped' if 'skipped' in kept else 'scored'} (or skipped for another reason) when recorded"
    elif "skipped" not in fresh:
        device = fresh[DEFAULT_GLYPH_SIZE]
        if any(entry["word_pages"] != device["word_pages"] for entry in kept.values()):
            why = "Word's export of it has another number of pages"
        elif recorded is None or recorded.get(DEFAULT_GLYPH_SIZE, {}).get("pages_made") != device["pages_made"]:
            why = "our layout of it makes another number of pages"
    if why:
        raise KeptBaselineMismatch(f"{name}: its recorded {what} baseline cannot be kept as it is, because {why}.  "
                                   "Re-record everything: --record --truth both --glyph-size both")


def baselines_payload(fresh: dict, previous: dict | None, converter: str) -> dict:
    """What ``--record`` writes, from ``fresh`` -- ``{truth: _summary(...)}`` of the truths
    scored, the default truth among them -- and ``previous``, the recording it replaces.

    **Only what was scored is recorded.**  The default run scores the ``svg`` truth at the
    ``device`` glyph size; every other recorded entry -- the ``exact`` glyph size's, and
    the ``pdfium`` truth's -- is carried over from ``previous`` *as it is*, never deleted
    or changed, and so is its ``page_median``.  It refuses (:class:`KeptBaselineMismatch`)
    where keeping one would misdescribe it (:func:`_keep`): the document is skipped now, or
    was; Word's export or our layout of it has another number of pages; the ``svg``
    truth's kept entries were converted by another converter; a document is gone from the
    run.  ``--record --truth both --glyph-size both`` re-records everything."""
    previous = previous or {}
    svg = fresh[DEFAULT_TRUTH]
    sizes = [size for size in GLYPH_SIZES if size in svg["page_median"]]
    if DEFAULT_GLYPH_SIZE not in sizes:
        raise ValueError(f"a recording scores the {DEFAULT_GLYPH_SIZE} glyph size")
    kept_sizes = [size for size in GLYPH_SIZES if size not in sizes]
    recorded = {TRUTHS[0]: previous, **({"pdfium": previous["pdfium"]} if "pdfium" in previous else {})}

    def merged(truth: str) -> dict:
        prior = recorded.get(truth) or {}
        documents = prior.get("documents") or {}
        if truth in fresh:
            if not kept_sizes:
                return fresh[truth]
            what, keep = f"{truth} {'/'.join(kept_sizes)}", kept_sizes
        else:
            what, keep = truth, GLYPH_SIZES
        if truth == DEFAULT_TRUTH and previous.get("converter") != converter and keep:
            raise KeptBaselineMismatch(f"the recorded {what} baselines were converted by {previous.get('converter')}, "
                                       f"not {converter}.  Re-record everything: --record --truth both --glyph-size "
                                       "both")
        gone = sorted(set(documents) - set(svg["documents"]))
        if gone:
            raise KeptBaselineMismatch(f"{', '.join(gone)}: recorded, but not in this run; a run that does not score "
                                       f"the {what} baselines does not delete one.  Re-record everything: --record "
                                       "--truth both --glyph-size both")
        out_documents = {}
        for name, summary in svg["documents"].items():
            prior_entry = documents.get(name)
            kept = None if prior_entry is None else (
                prior_entry if "skipped" in prior_entry else {size: prior_entry[size] for size in keep
                                                               if size in prior_entry} or None)
            _keep(what, name, kept, summary, previous.get("documents", {}).get(name))
            if "skipped" in summary:
                out_documents[name] = summary
            elif truth in fresh:
                out_documents[name] = {**fresh[truth]["documents"][name], **kept}
            else:
                out_documents[name] = kept
        page_median = {**(fresh[truth]["page_median"] if truth in fresh else {}),
                       **{size: prior["page_median"][size] for size in keep}}
        return {"documents": out_documents, "page_median": page_median}

    payload = {"_about": ABOUT, "truth": DEFAULT_TRUTH, "converter": converter, **merged(DEFAULT_TRUTH)}
    payload["pdfium"] = merged("pdfium")
    return payload


def _choice(value: str, everything: tuple) -> tuple:
    """``--truth`` / ``--glyph-size``: one, or ``both`` (the default first)."""
    return everything if value == "both" else (value,)


def _describe_checks(stats: dict, cache_mode: str, root: Path | None) -> str:
    """One line on what the run's spot-checks found, for stderr."""
    import raster_cache

    checks = stats.get("checks", [])
    count: dict = {}
    for check in checks:
        if "cache" in check:
            count[check["cache"]] = count.get(check["cache"], 0) + 1
    crops = sum(1 for check in checks if "crop" in check)
    parts = [f"{len(stats.get('sampled', []))} of {stats.get('pages', 0)} pages spot-checked"
             + (" (" + ", ".join(stats["sampled"]) + ")" if cache_mode != "verify" else ""),
             f"rasters re-drawn and identical to the cache: {count.get('match', 0)}",
             f"crops equal to the full page: {crops}"]
    if cache_mode != "off" and root is not None:
        entries, size = raster_cache.RasterCache(root, REPO).size()
        parts.append(f"cache: {count.get('hit', 0)} hits, {count.get('drawn', 0) + count.get('stored', 0)} drawn; "
                     f"{entries} rasters, {size / 1e6:.1f} MB in {root}")
    else:
        parts.append("cache: off")
    return "; ".join(parts)


def main(argv: list[str]) -> int:
    import raster_cache
    import read_render

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("docx", nargs="*", type=Path)
    parser.add_argument("--record", action="store_true",
                        help=f"rewrite {BASELINES.name} (committed documents) for what was scored: by default the svg "
                        "truth at the device glyph size, every other recorded entry kept exactly as it is; "
                        "--truth both --glyph-size both re-records everything")
    parser.add_argument("--truth", choices=TRUTHS + ("both",), default=DEFAULT_TRUTH,
                        help="how Word's page is rasterised (default svg: converted, then resvg like ours; pdfium: "
                        "the old instrument, scored only when asked for; both)")
    parser.add_argument("--glyph-size", choices=GLYPH_SIZES + ("both",), default=DEFAULT_GLYPH_SIZE,
                        help="how our glyphs are sized (default device: as Word's PDF draws them; exact, scored "
                        "only when asked for; both)")
    parser.add_argument("--probe", action="store_true", help="also score the render probes")
    parser.add_argument("--ink", action="store_true", help="score make_ink_probe.py page by page (size by size)")
    parser.add_argument("--pages", action="store_true", help="print every page's score")
    parser.add_argument("--max-pages", type=int, help="score at most this many pages of each document")
    parser.add_argument("--jobs", "-j", type=int, default=None,
                        help="processes to score pages in (default: every logical core, fewer if memory is short); "
                        "1 is the serial path.  Scores and output are identical either way")
    caching = parser.add_mutually_exclusive_group()
    caching.add_argument("--no-cache", action="store_true",
                         help="draw every raster; neither read nor write the raster cache")
    caching.add_argument("--verify-cache", action="store_true",
                         help="re-draw every raster and compare it with the cache byte for byte (and every page's "
                         "cropped SSIM with the full page's); any difference discards the cache and fails")
    args = parser.parse_args(argv[1:])
    jobs = default_jobs() if args.jobs is None else args.jobs
    if jobs < 1:
        parser.error("--jobs must be at least 1")
    if args.ink:
        return ink(args.truth if args.truth != "both" else DEFAULT_TRUTH)
    truths = _choice(args.truth, TRUTHS)
    glyph_sizes = _choice(args.glyph_size, GLYPH_SIZES)
    if args.record and (DEFAULT_TRUTH not in truths or DEFAULT_GLYPH_SIZE not in glyph_sizes):
        parser.error(f"--record scores the {DEFAULT_TRUTH} truth at the {DEFAULT_GLYPH_SIZE} glyph size, which every "
                     "other entry is kept beside; re-record the others with --truth both / --glyph-size both")
    if args.docx:
        documents = [(path.name, path.read_bytes(), path.stem) for path in args.docx]
    else:
        documents = [(name, (read_render.FIXTURES / name).read_bytes(), Path(name).stem) for name in COMMITTED]
    if args.probe:
        import make_render_probe

        documents += [(f"render-{s}", make_render_probe.build(s), f"render-{s}") for s in make_render_probe.SETTINGS]
    cache_mode = "off" if args.no_cache else "verify" if args.verify_cache else "use"
    root = raster_cache_root() if cache_mode != "off" else None
    # Every export looked up here, one at a time, before any worker starts (or, at
    # --jobs 1, before any page is scored).
    exported = [(name, data, read_render.export(name if name in COMMITTED else stem, data))
                for name, data, stem in documents]
    stats: dict = {}
    try:
        scored_documents = zip([name for name, _data, _pdf in exported],
                               score_corpus(exported, glyph_sizes, args.max_pages, truths, jobs, cache_mode, root,
                                            stats))
    except (raster_cache.CacheMismatch, CropMismatch) as mismatch:
        print(f"FIDELITY RUN ABORTED: {mismatch}", file=sys.stderr)
        return 3
    print(_describe_checks(stats, cache_mode, root), file=sys.stderr)
    results: dict = {truth: {} for truth in truths}
    for name, scored in scored_documents:
        for truth in truths:
            result = results[truth][name] = scored[truth]
            if "skipped" in result:
                print(f"{name:34} skipped: {result['skipped']}")
                continue
            for size in glyph_sizes:
                r = result[size]
                count = ("" if r["pages_made"] == r["word_pages"] else "  PAGE COUNT MISMATCH")
                print(f"{name:34} {truth:6} {size:6} SSIM {r['ssim']}  histogram {r['histogram']}  loss {r['loss']}  "
                      f"(pages ours {r['pages_made']}, Word {r['word_pages']}{count}; drawn in full "
                      f"{r['pages_scored']}; SSIM over every page {r['ssim_all']})")
                if args.pages:
                    for index, page in enumerate(r["pages"]):
                        print(f"{'':34} {'':6} {'':6}   page {index + 1:>3}: SSIM {page['ssim']}  histogram "
                              f"{page['histogram']}  loss {page['loss']}  mean |diff| {page['mean_abs']}"
                              + ("  (stopped)" if page["stopped"] else "") + ("  (sparse)" if page["sparse"] else "")
                              + (f"  ({page['unmatched']}: wholly wrong)" if "unmatched" in page else ""))
    mismatched = [row for truth in truths for size in glyph_sizes
                  for row in page_count_mismatches(results[truth], size)]
    for row in mismatched:
        print(describe_page_count_mismatch(row))
    for truth in truths:
        flagged = [row for size in glyph_sizes
                   for row in outliers(results[truth], size, drop=OUTLIER_DROPS[truth])]
        for row in flagged:
            print(f"[{truth}] " + describe_outlier(row))
        if not flagged:
            print(f"[{truth}] no outliers: every page scored is within {OUTLIER_DROPS[truth]} of its document's and the "
                  "corpus's median")
    if args.record:
        refused = sorted({row["document"] for row in mismatched if row["document"] in COMMITTED})
        if refused:
            print(f"not written: {', '.join(refused)} made another number of pages than Word; a page count that is "
                  "not Word's is a regression, not a baseline", file=sys.stderr)
            return 2
        previous = json.loads(BASELINES.read_text()) if BASELINES.exists() else None
        try:
            payload = baselines_payload({truth: _summary(results[truth], glyph_sizes) for truth in truths}, previous,
                                        __import__("pdf_svg").converter_version())
        except KeptBaselineMismatch as refused:
            print(f"not written: {refused}", file=sys.stderr)
            return 2
        BASELINES.write_text(json.dumps(payload, indent=1, sort_keys=True) + "\n")
        kept = [what for what, scored in (("the exact glyph size", "exact" in glyph_sizes),
                                          ("the pdfium truth", "pdfium" in truths)) if not scored]
        print(f"wrote {BASELINES}" + (f"; {' and '.join(kept)} kept as recorded (--truth both --glyph-size both "
                                      "re-records them)" if kept else ""))
    return 1 if mismatched else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

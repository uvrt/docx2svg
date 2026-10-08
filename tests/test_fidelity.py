"""Raster fidelity of the committed documents, against Word's PDF (``tools/fidelity.py``).

**Slow, and needs this machine**: Word's exports (``tools/oracle.py``'s cache in the
Office group container), Word's installed faces, pypdfium2, numpy, pillow and resvg-py,
and PyMuPDF (the ``fidelity`` extra) for the default instrument.  Marked ``slow``: run
with ``--run-slow``.  Skipped wherever any of it is missing, which is everywhere but the
machine with Word -- and, for the ``svg`` truth, wherever PyMuPDF is not installed (CI).

**Two instruments, both recorded** (``tools/fidelity.py``'s ``TRUTHS``): ``svg``, the
default -- Word's page converted by ``tools/pdf_svg.py`` and rasterised by resvg, like ours,
so no rasteriser difference is scored -- and ``pdfium``, the one before it, kept for
comparison.  The scores are pinned in ``tests/fidelity-baselines.json`` (``tools/fidelity.py
--record``: the ``svg`` truth at the top level, ``pdfium`` under its name) and may not drop
by more than :data:`MAX_SSIM_DROP`, as the sibling's harness allows.  A rise is not a
failure; record it.  **Scored only when asked for**, as the harness does: the tests that
render default to the ``svg`` truth at the ``device`` glyph size; ``pytest --pdfium`` runs
the ``pdfium`` variants (marked ``pdfium``) and ``pytest --exact`` adds the ``exact`` glyph
size.  The recorded scores of both are checked offline either way.  Read ``tools/fidelity.py``'s docstring on why SSIM can fall while the
picture improves before rebaselining a drop.

**A low score is to be explained, not recorded.**  A document or page far below the
typical score (``fidelity.OUTLIER_DROPS`` of the truth below the median of its document's pages or of
the corpus) fails here unless the truth's explanations (:data:`EXPLAINED`) say why:
``sample-with-table``'s 0.787, recorded as a gain, was its shaded rows' text painted over
by the shading.  The recorded means are checked in the default suite, offline; every page,
with ``--run-slow``.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

BASELINES = REPO / "tests" / "fidelity-baselines.json"
MAX_SSIM_DROP = 0.02
#: The pdfium truth's outliers, as they were explained under it (ROADMAP.md, 5.12 and
#: Tables), by ``(document, page)`` (``None``: the document's mean): the SSIM (the lower
#: glyph size's) whose shortfall was looked at, and what it is.  An explanation covers that
#: number, not the document: a score more than :data:`MAX_SSIM_DROP` under it is an
#: outlier again.  Under the ``svg`` truth none of them is an outlier (5.13): what they
#: named was mostly pdfium's.
TABLE_LINES = ("the table's border lines: Word's rectangles are the model's to a few hundred px "
               "(render-observations.json), but pdfium draws a 4 px line whose edge falls off the pixel "
               "grid 5 px wide, and the left border a pixel left (Tables stage 5); SSIM, a mean over thin "
               "foreground, punishes a one-pixel line along its whole length")
DENSE_PROSE = ("a page of dense Cambria 11 pt prose: every glyph at Word's pen position, but Word's PDF "
               "draws the glyphs inside a long text object up to ~0.9 px off it (Phase 5.3), and the "
               "glyph-edge residual of two rasterisers is per glyph; identical before and after the "
               "paint-order fix")
EXPLAINED_PDFIUM: dict[tuple[str, int | None], tuple[float, str]] = {
    ("wordto/sample-with-table.docx", None): (0.8875, TABLE_LINES + ": 0.9443 away from them"),
    ("wordto/sample-5pages.docx", None): (0.8609, "page 1 below"),
    ("wordto/sample-5pages.docx", 1): (0.8446, DENSE_PROSE + "; and " + TABLE_LINES + ": 0.8787 away from them"),
    ("wordto/sample-10pages.docx", None): (0.8766, "pages 1 and 2 are dense prose; page 3, " + TABLE_LINES
                                           + ": 0.9512 away from them"),
    ("wordto/sample-10pages.docx", 1): (0.8467, DENSE_PROSE),
    ("samplelib/sample-long.docx", 1): (0.8571, DENSE_PROSE + "; its document's median is its picture pages'"),
    ("samplelib/sample-simple.docx", 1): (0.8857, TABLE_LINES + ": 0.9293 away from them; its document's "
                                          "median is a near-empty page 2's"),
}
#: The ``svg`` truth's outliers (ROADMAP.md, 5.13), at its drop of 0.02.
QUARTZ_DRIFT = ("Word's PDF draws the glyphs inside a long text object where Quartz's encoded advances put "
                "them, which drifts from Word's own pen position (Phase 5.3, tools/glyphs.py): on the line at "
                "y = 1208 px our words run 0.05, 0.25, 0.43, 0.70 and 1.52 px left of Word's ink, left to "
                "right, every text object's start exact; the page is sparse, so the four lines of that "
                "paragraph weigh on its mean")
EXPLAINED_SVG: dict[tuple[str, int | None], tuple[float, str]] = {
    ("layout-sweep.docx", 1): (0.9548, QUARTZ_DRIFT),
}
EXPLAINED = {"svg": EXPLAINED_SVG, "pdfium": EXPLAINED_PDFIUM}
TRUTHS = ("svg", "pdfium")


def _explained(truth: str, name: str, page: int | None, ssim: float) -> bool:
    found = EXPLAINED[truth].get((name, page))
    return found is not None and ssim >= found[0] - MAX_SSIM_DROP


def _baselines() -> dict:
    return json.loads(BASELINES.read_text(encoding="utf-8"))


def _recorded(truth: str) -> dict:
    """One truth's recording: ``documents`` and ``page_median``."""
    baselines = _baselines()
    return baselines if truth == baselines.get("truth", "svg") else baselines[truth]


@pytest.mark.parametrize("truth", TRUTHS)
def test_no_recorded_document_is_an_unexplained_outlier(truth):
    import fidelity

    baselines = _recorded(truth)
    for size in ("device", "exact"):
        typical = fidelity.corpus_median(baselines, size)
        for name, result in baselines["documents"].items():
            if "skipped" in result or result[size]["ssim"] is None:
                continue
            ssim = result[size]["ssim"]
            assert ssim >= typical - fidelity.OUTLIER_DROPS[truth] or _explained(truth, name, None, ssim), (
                f"{name} ({truth}, {size}) is recorded at SSIM {ssim}, loss {result[size]['loss']}, against a "
                f"median document's {typical}: explain it (look at the pictures) before recording it")


@pytest.mark.parametrize("truth", TRUTHS)
def test_every_explanation_is_still_needed(truth):
    """An explanation of a document that is no longer an outlier is stale: remove it."""
    import fidelity

    baselines = _recorded(truth)
    for (name, page), (ssim, _reason) in EXPLAINED[truth].items():
        if page is None:
            typical = fidelity.corpus_median(baselines, "device")
            assert baselines["documents"][name]["device"]["ssim"] < typical - fidelity.OUTLIER_DROPS[truth], name


def test_both_truths_are_recorded_for_the_same_documents():
    baselines = _baselines()
    assert baselines["truth"] == "svg"
    assert sorted(baselines["documents"]) == sorted(baselines["pdfium"]["documents"])


#: The truths a scoring test runs under: ``svg``, and ``pdfium`` -- the old instrument,
#: which the harness scores only when asked for -- under ``pytest --pdfium``.
SCORED_TRUTHS = ["svg", pytest.param("pdfium", marks=pytest.mark.pdfium)]


def _glyph_sizes(config) -> tuple[str, ...]:
    """The glyph sizes a scoring test renders: ``device``, as the harness does by default,
    and ``exact`` too under ``pytest --exact``."""
    return ("device", "exact") if config.getoption("--exact") else ("device",)


def _needs_the_harness(truth: str) -> None:
    for module in ("numpy", "PIL", "pypdfium2", "resvg_py"):
        pytest.importorskip(module)
    import pdf_svg

    from docx2svg.fonts import InstalledFonts

    if truth == "svg" and not pdf_svg.available():
        pytest.skip("PyMuPDF (the fidelity extra, tools/pdf_svg.py) is not installed here")
    if InstalledFonts().face("Calibri") is None:
        pytest.skip("Word's faces are not installed here")


def _committed(name: str) -> tuple[bytes, Path]:
    """A committed document and Word's cached export of it, or skip: never exports."""
    import pdf_svg
    import read_render

    pdf = pdf_svg.committed_pdf(name)
    if not pdf.exists():
        pytest.skip("Word's export of this document is not cached here")
    return (read_render.FIXTURES / name).read_bytes(), pdf


@pytest.mark.slow
@pytest.mark.parametrize("truth", SCORED_TRUTHS)
@pytest.mark.parametrize("name", sorted(json.loads(BASELINES.read_text(encoding="utf-8"))["documents"]) if BASELINES.exists() else [])
def test_fidelity_does_not_regress(name, truth, request):
    """Rasters come from the harness's cache where it has them -- read-only, every entry's
    key components and pixel digest checked on the way in -- and are drawn otherwise."""
    _needs_the_harness(truth)
    import fidelity

    data, pdf = _committed(name)
    sizes = _glyph_sizes(request.config)
    expected_all = _recorded(truth)
    expected = expected_all["documents"][name]
    result = fidelity.score_document(name, data, pdf, sizes, truth=truth, cache_mode="read")
    mismatched = [row for size in sizes for row in fidelity.page_count_mismatches({name: result}, size)]
    assert not mismatched, "\n".join(fidelity.describe_page_count_mismatch(row) for row in mismatched)
    for size in sizes:
        flagged = [row for row in fidelity.outliers({name: result}, size, corpus=expected_all["page_median"][size],
                                                    drop=fidelity.OUTLIER_DROPS[truth])
                   if not _explained(truth, name, row["page"], row["ssim"])]
        assert not flagged, "\n".join(fidelity.describe_outlier(row) for row in flagged)
    for size in sizes:
        if expected[size]["ssim"] is None:
            continue
        assert result[size]["ssim"] >= expected[size]["ssim"] - MAX_SSIM_DROP, (name, size)
        assert result[size]["histogram"] >= expected[size]["histogram"] - MAX_SSIM_DROP, (name, size)


# --------------------------------------------------------------------------------------
# A page too many or too few: scored, reported, failing
# --------------------------------------------------------------------------------------

def test_a_page_only_one_side_made_is_wholly_wrong():
    """Its SSIM and histogram are 0 whatever it holds -- an empty page included -- and its
    loss is every foreground pixel of the page that exists."""
    import fidelity

    np = _numpy()
    page = np.full((90, 160, 3), 255, dtype=np.uint8)
    empty = fidelity.fully_wrong(page, fidelity.OURS_ONLY)
    assert (empty["ssim"], empty["histogram"], empty["loss"], empty["sparse"]) == (0.0, 0.0, 0.0, True)
    page[10:40, 20:120] = 0
    inked = fidelity.fully_wrong(page, fidelity.WORD_ONLY)
    assert (inked["ssim"], inked["histogram"], inked["loss"]) == (0.0, 0.0, 3000.0)
    assert inked["unmatched"] == fidelity.WORD_ONLY and inked["mean_abs"] == round(255 * 3000 / (90 * 160), 3)


def _synthetic(ours: int, word: int) -> dict:
    """A result as :func:`fidelity.score_corpus` builds one: ``ours`` pages made, ``word``
    Word's, every page both made scored 0.99."""
    import fidelity

    row = {"ssim": 0.99, "histogram": 1.0, "loss": 10.0, "mean_abs": 0.1, "coverage": 0.05, "sparse": False,
           "stopped": False}
    pages = [dict(row) for _ in range(min(ours, word))]
    side = fidelity.OURS_ONLY if ours > word else fidelity.WORD_ONLY
    pages += [dict(row, ssim=0.0, histogram=0.0, loss=500.0, unmatched=side) for _ in range(abs(ours - word))]
    return {"device": fidelity._entry(pages, word, ours)}


@pytest.mark.parametrize("ours, word", [(3, 2), (2, 3), (2, 2)])
def test_a_page_count_that_is_not_words_is_reported_and_lowers_the_score(ours, word):
    import fidelity

    result = _synthetic(ours, word)
    entry = result["device"]
    assert (entry["pages_made"], entry["word_pages"]) == (ours, word)
    mismatched = fidelity.page_count_mismatches({"doc": result})
    if ours == word:
        assert not mismatched and entry["ssim"] == 0.99
        return
    assert entry["ssim"] < 0.99 and len(entry["pages"]) == max(ours, word)
    (row,) = mismatched
    assert (row["ours"], row["word"], row["unmatched"]) == (ours, word, [3])
    assert "PAGE COUNT MISMATCH doc" in fidelity.describe_page_count_mismatch(row)
    # Never an outlier of its own, and never in the medians the outliers are held to.
    assert fidelity.outliers({"doc": result}) == []
    assert fidelity.page_median({"doc": result}) == 0.99


# --------------------------------------------------------------------------------------
# SSIM cropped to the content: exact, or it is not done
# --------------------------------------------------------------------------------------

def _numpy():
    return pytest.importorskip("numpy")


def _page(np, rows=90, cols=160):
    return np.full((rows, cols, 3), 255, dtype=np.uint8)


def _both_ways(a, b):
    import fidelity

    cropped, full = fidelity.score(a, b), fidelity.score(a, b, crop=False)
    assert cropped == full, (cropped, full)
    assert fidelity.score_checked(a, b) == cropped    # the masked values and histograms, bit for bit
    return cropped


def test_the_crop_box_is_the_union_of_both_images_ink_plus_the_window():
    """Never one side's ink: what only Word drew is inside the box as surely as what only
    we drew, or the score would not see what we left out."""
    import fidelity

    np = _numpy()
    ours, theirs = _page(np), _page(np)
    ours[30:40, 50:60] = 0           # ink only we drew
    theirs[60:70, 100:120] = 0       # ink only Word drew
    mask = (ours.mean(axis=2) < fidelity.FOREGROUND_MAX) | (theirs.mean(axis=2) < fidelity.FOREGROUND_MAX)
    rows, cols = fidelity.content_box(mask)
    radius = fidelity.SSIM_RADIUS
    assert radius == len(fidelity._gaussian_kernel()) // 2 == 5
    assert (rows.start, rows.stop) == (30 - radius, 70 + radius)
    assert (cols.start, cols.stop) == (50 - radius, 120 + radius)
    _both_ways(ours, theirs)


def test_cropped_ssim_is_the_full_pages_bit_for_bit():
    """Random content in random places, margins from none to wide: every score field is
    the same float either way, not merely close."""
    np = _numpy()
    rng = np.random.default_rng(7)
    for _trial in range(40):
        ours, theirs = _page(np), _page(np)
        top, left = int(rng.integers(0, 80)), int(rng.integers(0, 150))
        bottom = top + int(rng.integers(1, 90 - top + 1))
        right = left + int(rng.integers(1, 160 - left + 1))
        ours[top:bottom, left:right] = rng.integers(0, 256, size=(bottom - top, right - left, 3))
        shift = int(rng.integers(-4, 5))
        theirs[max(top + shift, 0):max(bottom + shift, 1), left:right] = 40
        _both_ways(ours, theirs)


def test_the_shortcuts_are_the_literal_computations_bit_for_bit():
    """The grey level, the mean absolute difference and the SSIM gathered before its
    per-pixel arithmetic are the original expressions' floats, not merely close: random
    pages, and the extremes of every channel."""
    import fidelity

    np = _numpy()
    rng = np.random.default_rng(11)
    pages = [rng.integers(0, 256, size=(70, 90, 3), dtype=np.uint8) for _ in range(6)]
    pages += [np.zeros((70, 90, 3), np.uint8), np.full((70, 90, 3), 255, np.uint8)]
    for a, b in zip(pages, pages[1:] + pages[:1]):
        assert fidelity._gray(a).tobytes() == a.mean(axis=2).tobytes()
        assert fidelity._mean_abs_difference(a, b) == float(np.abs(a.astype(np.int16) - b.astype(np.int16)).mean())
        mask = rng.random((70, 90)) < 0.3
        gray_a, gray_b = a.mean(axis=2), b.mean(axis=2)
        assert fidelity.masked_ssim(gray_a, gray_b, mask).tobytes() == fidelity.ssim_map(gray_a, gray_b)[mask].tobytes()


def test_a_page_with_ink_at_its_edges_is_not_cropped_at_all():
    """Content to the edges, a page colour, a full-bleed picture: the box is the page, so
    there is nothing to gain and nothing is cut."""
    import fidelity

    np = _numpy()
    ours, theirs = _page(np), _page(np)
    ours[0, 0] = 0
    theirs[-1, -1] = 0
    ours[40:50, 40:50] = 0
    mask = (ours.mean(axis=2) < fidelity.FOREGROUND_MAX) | (theirs.mean(axis=2) < fidelity.FOREGROUND_MAX)
    assert fidelity.content_box(mask) == (slice(0, 90), slice(0, 160))
    _both_ways(ours, theirs)

    tinted = np.full((90, 160, 3), 200, dtype=np.uint8)      # a page colour under 245
    tinted[20:30, 20:60] = 30
    mask = tinted.mean(axis=2) < fidelity.FOREGROUND_MAX
    assert fidelity.content_box(mask) == (slice(0, 90), slice(0, 160))
    assert _both_ways(tinted, tinted)["ssim"] == 1.0


def test_ink_within_the_window_of_an_edge_is_clamped_not_cut():
    """Content none to six pixels from the edges: the box is clamped to the page, and the
    page's own edge rows are what the blur's padding replicates either way."""
    import fidelity

    np = _numpy()
    for gap in range(0, fidelity.SSIM_RADIUS + 2):
        ours, theirs = _page(np), _page(np)
        ours[gap:gap + 20, gap:gap + 30] = 0
        theirs[gap + 1:gap + 21, gap:gap + 30] = 0
        theirs[90 - gap - 12:90 - gap, 160 - gap - 25:160 - gap] = 90
        _both_ways(ours, theirs)


def test_a_crop_that_cut_a_pixel_would_be_caught(monkeypatch):
    """The check is not vacuous: a box one pixel short of the window changes the masked
    SSIM values, and :func:`fidelity.score_checked` raises."""
    import fidelity

    np = _numpy()
    ours, theirs = _page(np), _page(np)
    ours[30:60, 40:100] = 0
    theirs[31:61, 40:100] = 0
    ours[25, :] = 250    # background (not ink), exactly the window's radius above the ink
    real = fidelity.content_box
    monkeypatch.setattr(fidelity, "content_box", lambda mask: real(mask, fidelity.SSIM_RADIUS - 1))
    with pytest.raises(fidelity.CropMismatch):
        fidelity.score_checked(ours, theirs, "short box")


@pytest.mark.slow
@pytest.mark.parametrize("truth", SCORED_TRUTHS)
def test_cropped_ssim_equals_the_full_page_on_every_scored_page(truth, request):
    """The committed documents themselves, every page the harness scores: the run holds
    each page's cropped SSIM and histograms equal to the full page's, bit for bit, and
    raises on the first that is not.

    Rasters come from the harness's cache where it has them -- read-only -- and are drawn
    otherwise; this test writes nothing."""
    _needs_the_harness(truth)
    import os

    import fidelity

    documents = [(name, *_committed(name)) for name in fidelity.COMMITTED]
    stats: dict = {}
    sizes = _glyph_sizes(request.config)
    jobs = 1 if os.environ.get("PYTEST_XDIST_WORKER") else fidelity.default_jobs()
    results = fidelity.score_corpus(documents, sizes, None, (truth,), jobs, cache_mode="read", stats=stats,
                                    crop_all=True)
    scored = sum(len(result[truth][size]["pages"]) for result in results if "skipped" not in result[truth]
                 for size in sizes)
    assert scored and stats["pages"] == scored // len(sizes)
    crops = [check for check in stats["checks"] if check.get("crop") == "equal"]
    assert len(crops) == scored, (len(crops), scored)


# --------------------------------------------------------------------------------------
# The raster cache: a stale entry can never be used quietly
# --------------------------------------------------------------------------------------

def _cache(tmp_path):
    _numpy()
    pytest.importorskip("PIL")
    import raster_cache

    return raster_cache, raster_cache.RasterCache(tmp_path / "rasters")


def test_the_cache_returns_what_it_stored_and_only_for_the_same_components(tmp_path):
    np = _numpy()
    raster_cache, cache = _cache(tmp_path)
    image = np.arange(4 * 5 * 3, dtype=np.uint8).reshape(4, 5, 3)
    components = {"svg": "abc", "fonts": "def", "dpi": 300}
    assert cache.get("ours", components) is None
    cache.put("ours", components, image)
    assert np.array_equal(cache.get("ours", components), image)
    assert cache.get("ours", dict(components, fonts="xyz")) is None     # another key: a miss
    assert cache.get("truth", components) is None                        # another side
    assert cache.verify("ours", components, image) == "match"


def test_a_cached_raster_that_is_not_what_was_stored_is_refused(tmp_path):
    """Components that do not match the ones asked for, or pixels that are not the ones
    written, are a corrupted cache -- raised, not treated as a miss."""
    np = _numpy()
    raster_cache, cache = _cache(tmp_path)
    image = np.zeros((4, 5, 3), dtype=np.uint8)
    components = {"svg": "abc"}
    cache.put("ours", components, image)
    meta = cache.root / "ours" / f"{raster_cache.key(components)}.json"

    stored = json.loads(meta.read_text(encoding="utf-8"))
    meta.write_text(json.dumps(dict(stored, components={"svg": "other"})))
    with pytest.raises(raster_cache.CacheMismatch, match="components"):
        cache.get("ours", components)

    meta.write_text(json.dumps(dict(stored, pixels="0" * 64)))
    with pytest.raises(raster_cache.CacheMismatch, match="pixels"):
        cache.get("ours", components)


def test_a_fresh_raster_that_differs_from_the_cache_by_one_value_fails(tmp_path):
    np = _numpy()
    raster_cache, cache = _cache(tmp_path)
    image = np.full((4, 5, 3), 255, dtype=np.uint8)
    cache.put("truth", {"pdf": "p"}, image)
    fresh = image.copy()
    fresh[2, 3, 1] = 254
    with pytest.raises(raster_cache.CacheMismatch, match="1 pixels"):
        cache.verify("truth", {"pdf": "p"}, fresh)


def test_the_cache_refuses_to_live_inside_a_repository(tmp_path):
    """Rasters of Word's pages hold Microsoft's glyph shapes."""
    raster_cache, _ = _cache(tmp_path)
    with pytest.raises(RuntimeError, match="repository"):
        raster_cache.RasterCache(REPO / "scratch" / "rasters", REPO)
    with pytest.raises(RuntimeError, match="git checkout"):
        raster_cache.RasterCache(REPO / "tests" / "rasters")


def test_the_harness_caches_outside_the_repository():
    import fidelity
    import oracle

    root = fidelity.raster_cache_root()
    assert root.is_relative_to(oracle.ORACLE_DIR) and not root.resolve().is_relative_to(REPO.resolve())


def test_the_spot_check_sample_is_spread_and_visits_every_page():
    import raster_cache

    pages = list(range(52))
    day = 739_000
    first = raster_cache.rotating_sample(pages, 6, day)
    assert len(first) == 6
    assert max(first) - min(first) >= 52 * 0.75       # spread through the corpus
    assert first == raster_cache.rotating_sample(pages, 6, day)
    stride = -(-52 // 6)
    seen: set = set()
    for offset in range(stride):
        seen |= set(raster_cache.rotating_sample(pages, 6, day + offset))
    assert seen == set(pages)


def test_a_font_changed_in_place_changes_the_key(tmp_path):
    """An Office update can change a face without changing its path, its name, its size
    or its date.  The key reads the file's contents -- and names a document's embedded
    face by its place in the page's temporary directory, not by that directory."""
    import os

    for module in ("numpy", "PIL", "resvg_py"):
        pytest.importorskip(module)
    import fidelity

    face = tmp_path / "Face.ttf"
    face.write_bytes(b"one face")
    embedded = tmp_path / "docx2svg-fidelity-abc" / "embedded-0.ttf"
    embedded.parent.mkdir()
    embedded.write_bytes(b"an embedded face")
    files = [str(face), str(embedded)]

    before = fidelity.our_components("<svg/>", files, str(embedded.parent))
    assert before["font_files"][1][0] == "<embedded>/embedded-0.ttf"
    fidelity._DIGESTS.clear()                              # a new run
    stamp = face.stat()
    face.write_bytes(b"two face")                          # same path, same size, same date
    os.utime(face, ns=(stamp.st_atime_ns, stamp.st_mtime_ns))
    fidelity._DIGESTS.clear()
    after = fidelity.our_components("<svg/>", files, str(embedded.parent))
    assert after != before
    moved = tmp_path / "docx2svg-fidelity-xyz"
    moved.mkdir()
    (moved / "embedded-0.ttf").write_bytes(b"an embedded face")
    assert fidelity.our_components("<svg/>", [str(face), str(moved / "embedded-0.ttf")], str(moved)) == after
    assert fidelity.our_components("<svg />", files, str(embedded.parent)) != after
    assert fidelity.our_components("<svg/>", files[::-1], str(embedded.parent)) != after   # the order resvg reads


def test_a_stale_cached_raster_discards_the_cache_and_stops_the_run(tmp_path):
    """The detection the key cannot provide for itself.  An entry whose key is right and
    whose pixels are not -- exactly what a missing key input leaves behind -- is found by
    the spot-check, the whole cache is discarded, and the run raises instead of scoring.
    The cache here is the test's own, under ``tmp_path``."""
    _needs_the_harness("svg")
    import fidelity
    import raster_cache

    name = "wordto/sample-with-table.docx"
    documents = [(name, *_committed(name))]
    root = tmp_path / "rasters"
    stats: dict = {}
    first = fidelity.score_corpus(documents, ("device",), None, ("svg",), 1, "use", root, stats)
    entries = sorted((root / "ours").glob("*.json"))
    assert len(entries) == 1 and len(list((root / "truth").glob("*.json"))) == 1
    assert stats["sampled"] == [f"{name} page 1"]              # one page: always sampled
    assert fidelity.score_corpus(documents, ("device",), None, ("svg",), 1, "use", root) == first
    assert fidelity.score_corpus(documents, ("device",), None, ("svg",), 1, "off") == first

    # Stale but self-consistent: other pixels, recorded under the right key.
    stored = json.loads(entries[0].read_text(encoding="utf-8"))
    image = raster_cache._array(entries[0].with_suffix(".png").read_bytes()).copy()
    image[10, 10] = 255 - image[10, 10]
    entries[0].with_suffix(".png").write_bytes(raster_cache._png(image))
    entries[0].write_text(json.dumps(dict(stored, pixels=raster_cache.pixel_digest(image))))

    with pytest.raises(raster_cache.CacheMismatch, match="discarded"):
        fidelity.score_corpus(documents, ("device",), None, ("svg",), 1, "use", root)
    assert not root.exists()
    assert fidelity.score_corpus(documents, ("device",), None, ("svg",), 1, "verify", root) == first


# --------------------------------------------------------------------------------------
# The recording: what a run did not score is kept, never deleted or changed
# --------------------------------------------------------------------------------------

def _fresh(baselines: dict) -> dict:
    """The default run's summary that would reproduce ``baselines``' own: the svg truth at
    the device glyph size."""
    return {"svg": {"documents": {name: entry if "skipped" in entry else {"device": entry["device"]}
                                  for name, entry in baselines["documents"].items()},
                    "page_median": {"device": baselines["page_median"]["device"]}}}


def _payload(fresh: dict, baselines: dict) -> dict:
    import copy

    import fidelity

    return fidelity.baselines_payload(fresh, copy.deepcopy(baselines), baselines["converter"])


def test_a_default_record_keeps_every_entry_it_did_not_score_exactly():
    """The default ``--record`` scores the svg truth at the device glyph size; the exact
    and pdfium entries it writes are the ones already recorded, and the file it writes is
    the same bytes."""
    baselines = _baselines()
    payload = _payload(_fresh(baselines), baselines)
    assert json.dumps(payload, indent=1, sort_keys=True) + "\n" == BASELINES.read_text(encoding="utf-8")


def test_a_full_record_writes_what_it_scored():
    import copy

    baselines = _baselines()
    fresh = {"svg": {k: baselines[k] for k in ("documents", "page_median")}, "pdfium": baselines["pdfium"]}
    assert json.dumps(_payload(fresh, baselines), indent=1, sort_keys=True) + "\n" == BASELINES.read_text(encoding="utf-8")
    changed = copy.deepcopy(fresh)
    name = next(iter(changed["svg"]["documents"]))
    changed["pdfium"]["documents"][name]["exact"]["ssim"] = 0.5
    payload = _payload(changed, baselines)
    assert payload["pdfium"]["documents"][name]["exact"]["ssim"] == 0.5
    assert payload["documents"][name] == baselines["documents"][name]


def test_a_default_record_refuses_what_it_cannot_keep_honestly():
    """A kept entry shares its document's export and layout; where those moved, keeping it
    would misdescribe it, and deleting it is not a default run's to do."""
    import fidelity

    baselines = _baselines()
    name = next(name for name, entry in baselines["documents"].items() if "skipped" not in entry)

    fresh = _fresh(baselines)
    fresh["svg"]["documents"][name]["device"] = dict(fresh["svg"]["documents"][name]["device"], word_pages=99)
    with pytest.raises(fidelity.KeptBaselineMismatch, match="Word's export"):
        _payload(fresh, baselines)

    fresh = _fresh(baselines)
    fresh["svg"]["documents"][name]["device"] = dict(fresh["svg"]["documents"][name]["device"], pages_made=99)
    with pytest.raises(fidelity.KeptBaselineMismatch, match="our layout"):
        _payload(fresh, baselines)

    fresh = _fresh(baselines)
    fresh["svg"]["documents"][name] = {"skipped": "no face 'X' on this machine"}
    with pytest.raises(fidelity.KeptBaselineMismatch, match="skipped"):
        _payload(fresh, baselines)

    fresh = _fresh(baselines)
    del fresh["svg"]["documents"][name]
    with pytest.raises(fidelity.KeptBaselineMismatch, match="not in this run"):
        _payload(fresh, baselines)

    with pytest.raises(fidelity.KeptBaselineMismatch, match="converted by"):
        fidelity.baselines_payload(_fresh(baselines), baselines, "pymupdf-0.0-f0")

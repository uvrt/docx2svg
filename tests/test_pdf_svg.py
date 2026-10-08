"""The fidelity instrument's converter (``tools/pdf_svg.py``): Word's PDF page to SVG.

The instrument's own arithmetic runs everywhere (numpy where it is installed).  The
converter's faithfulness to Word's PDF -- same engine, two routes; no pixel beyond
anti-aliasing; glyphs as the embedded outlines -- needs PyMuPDF (the ``fidelity`` extra)
and Word's cached exports, so it is ``slow`` and skips where either is missing, CI
among them.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import pdf_svg  # noqa: E402  (imports nothing outside the standard library at module level)


A4 = ('<svg xmlns="http://www.w3.org/2000/svg" version="1.1" width="595.2" height="841.92" '
      'viewBox="0 0 595.2 841.92">\n<path d="M0 0H10V10Z"/>\n</svg>\n')


def test_the_page_is_drawn_on_the_device_grid():
    """resvg rounds a root's size in points before zooming (A4 came out 2,479 px wide and
    stretched 1e-4 tall); the root is sized in device pixels and the viewBox widened to
    match, so one point is exactly 300/72 px."""
    sized = pdf_svg.device_grid(A4)
    assert 'width="2480" height="3508"' in sized
    width, height = (float(v) for v in sized.split('viewBox="0 0 ')[1].split('"')[0].split())
    assert abs(2480 / width - 300 / 72) < 1e-9 and abs(3508 / height - 300 / 72) < 1e-9
    assert sized.endswith('<path d="M0 0H10V10Z"/>\n</svg>\n')


def test_quartzs_no_clip_is_brought_within_what_resvg_can_hold():
    # Word wraps every drawing in a clip rectangle half a billion points across; resvg
    # drops an area that far out, and the drawing with it.
    svg = ('<svg><defs><clipPath id="clip_1"><path transform="matrix(1,0,0,-1,0,612)" '
           'd="M-257698030-257697430H257698030V257698640H-257698030Z"/></clipPath></defs>'
           '<path d="M12345678 0"/></svg>')
    bounded = pdf_svg.bounded_clips(svg)
    assert 'd="M-1000000-1000000H1000000V1000000H-1000000Z"' in bounded
    assert '<path d="M12345678 0"/>' in bounded  # only clip paths are touched


def test_an_edge_is_anti_aliasing_and_a_moved_stroke_is_not():
    np = pytest.importorskip("numpy")
    page = np.full((40, 40, 3), 255, np.uint8)
    page[10:30, 10:14] = 0
    softer = page.copy()
    softer[10:30, 14] = 128  # the same stroke's edge covered differently
    assert pdf_svg.beyond_antialiasing(page, softer).sum() == 0
    moved = np.full_like(page, 255)
    moved[10:30, 13:17] = 0  # three pixels right
    assert pdf_svg.beyond_antialiasing(page, moved).sum() > 0
    recoloured = page.copy()
    recoloured[10:30, 10:14] = (200, 0, 0)
    assert pdf_svg.beyond_antialiasing(page, recoloured).sum() > 0


def test_a_changed_colour_is_seen_where_the_histogram_would_not_say():
    np = pytest.importorskip("numpy")
    page = np.full((20, 20, 3), 255, np.uint8)
    page[5:15, 5:15] = (204, 210, 215)
    rounded = page.copy()
    rounded[5:15, 5:15] = (204, 210, 216)  # a level of rounding: faithful
    assert pdf_svg.flat_colour_difference(page, rounded) == 1
    shifted = page.copy()
    shifted[5:15, 5:15] = (204, 200, 215)
    assert pdf_svg.flat_colour_difference(page, shifted) == 10


def test_a_converted_page_is_never_written_into_the_repository(tmp_path):
    with pytest.raises(RuntimeError):
        pdf_svg._refuse_repository(REPO / "tests" / "anything")
    pdf_svg._refuse_repository(tmp_path)


#: A representative sample of Word's pages (every page of every committed export took 27
#: minutes; ``python tools/pdf_svg.py --validate`` runs them all): dense prose, a table
#: with ruled lines and shading, a picture, an A4 page in several faces, and a resume.
SAMPLE = (("wordto/sample-with-table.docx", 0), ("wordto/sample-10pages.docx", 0),
          ("wordto/sample-with-images.docx", 0), ("layout-sweep.docx", 0), ("style-document.docx", 0),
          ("samplelib/sample-resume.docx", 0))


@pytest.mark.slow
def test_the_converter_is_faithful_to_words_pdf():
    """A sample of the committed documents' exports: MuPDF's raster of the PDF and of the
    SVG, and resvg's raster of the SVG, differ only at anti-aliasing scale, in no colour;
    every glyph is the embedded program's, and the SVG has no text."""
    if not pdf_svg.available():
        pytest.skip("PyMuPDF (the fidelity extra) is not installed here")
    for module in ("numpy", "PIL", "resvg_py", "fontTools", "pypdfium2"):
        pytest.importorskip(module)
    checked = 0
    for name, page in SAMPLE:
        pdf = pdf_svg.committed_pdf(name)
        if not pdf.exists():
            continue
        for row in pdf_svg.validate(pdf, supersample=0, pages=[page]):
            where = f"{pdf.stem} page {row['page']}"
            assert pdf_svg.faithful(row), (where, row)
            if "mupdf_pdf_vs_mupdf_svg" in row:
                assert row["mupdf_pdf_vs_mupdf_svg"]["ssim"] >= 0.98, (where, row["mupdf_pdf_vs_mupdf_svg"])
            checked += 1
    if not checked:
        pytest.skip("Word's exports are not cached here")

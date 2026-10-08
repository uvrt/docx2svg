"""The loop, closed.

Every test here compares something this project *read* out of its own fixture against
the same thing measured out of the PDF **Word** produced from that fixture.  They are the
only tests in the suite that can say this project is wrong, as opposed to merely
inconsistent with itself.  They skip where Word is absent, which is everywhere but the
reference Mac.

The measurements are all *differences* between glyph ink boxes.  That is what lets them
be exact without a font file: the left side bearing, the face's vertical metrics and
Word's page-box rounding are identical on both sides of every subtraction and cancel.
No Microsoft font file is read, shipped or committed by anything here.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))

from docx2svg import parse_package  # noqa: E402
from docx2svg.units import device_page_extent_pt, twips_to_pt  # noqa: E402

import make_layout_sweep as sweep  # noqa: E402

#: pdfium returns coordinates as C floats, so a value that is exact in the PDF comes back
#: with about 1e-5 pt of float32 noise on a 600 pt page.  Anything this project could get
#: *wrong* is orders of magnitude larger -- a twip is 0.05 pt -- so this tolerance
#: separates "exact" from "not exact" rather than granting slack.
FLOAT32_NOISE_PT = 1e-3

#: Every test here reads Word's export of the layout sweep (``oracle_pdf``), which exports
#: it through Word when it is not cached: they run serially, or on one xdist worker
#: (tests/conftest.py).
pytestmark = pytest.mark.word


@pytest.fixture(scope="module")
def document(layout_sweep_path: Path):
    return parse_package(layout_sweep_path)


def test_word_opens_and_exports_our_generated_document(oracle_pdf: Path):
    """The acceptance test for hand-written OOXML, and the only one that counts.

    Well-formed is not the bar.  Word silently *repairs* a document with wrong child
    order or a malformed content-type override, and a repaired document exports under a
    different name or not at all -- which reads as a broken oracle rather than a broken
    fixture.  If this passes, the fixture is real.
    """
    assert oracle_pdf.exists() and oracle_pdf.stat().st_size > 1000


def test_word_drew_the_face_the_document_asked_for(oracle_pdf: Path):
    """No substitution crept in.

    ``hello.docx``, which named Calibri but carried no ``w:docDefaults``, came back with
    an ``Aptos`` subset alongside its Calibri ones -- Word resolved the theme's minor
    font for the runs that named no face.  A fixture that substitutes is measuring font
    resolution, not layout, and every number taken from it is worth less than it looks.
    """
    names = {
        match.group(1).decode("latin-1").split("+")[-1]
        for match in re.finditer(rb"/BaseFont\s*/([A-Za-z0-9+#-]+)", oracle_pdf.read_bytes())
    }
    assert names == {sweep.FACE, f"{sweep.FACE}-Bold"}, names


def test_the_exported_page_box_is_the_authored_size_snapped_to_device_pixels(
    document, oracle_lines
):
    """Word's ``/MediaBox`` is **not** ``w:pgSz``.

    A4's 11906 twips is 595.3 pt authored and 595.2 pt exported.  Both exported extents
    are whole 1/300-inch device pixels, and section 2 of the fixture exists to falsify
    that: it asks for 10000 x 13000 twips, which is *exactly* 500 x 650 pt, so a page box
    that were simply the authored size would read 500 x 650.  It reads 499.92 x 649.92 --
    2083 and 2708 device pixels.  The law survived the test designed to kill it.
    """
    page_sizes, _ = oracle_lines
    assert len(page_sizes) == 3  # two A4 pages, then the odd-size section

    for section, page in ((document.sections[0], 0), (document.sections[-1], -1)):
        expected = (
            device_page_extent_pt(section.page_size.width_twips),
            device_page_extent_pt(section.page_size.height_twips),
        )
        assert page_sizes[page][0] == pytest.approx(expected[0], abs=FLOAT32_NOISE_PT)
        assert page_sizes[page][1] == pytest.approx(expected[1], abs=FLOAT32_NOISE_PT)

    # And the falsification, spelled out rather than implied.
    assert page_sizes[-1][0] != pytest.approx(twips_to_pt(sweep.ODD_WIDTH_TWIPS), abs=0.05)


def test_the_indent_ladder_lands_exactly_where_it_was_authored(document, oracle_lines):
    """**The measurement that closes the loop.**

    Five paragraphs beginning with the same glyph at five authored indents.  The glyph's
    left side bearing is identical in all five, so the differences between their ink-box
    left edges are the authored indent differences and nothing else -- no font metric
    enters the comparison at any point.
    """
    _, lines = oracle_lines
    ladder_lines = [line for line in lines if line.text.startswith("H indent")]
    ladder_paragraphs = [
        paragraph for paragraph in document.paragraphs if paragraph.text.startswith("H indent ")
    ]
    assert len(ladder_lines) == len(ladder_paragraphs) == len(sweep.INDENT_LADDER_TWIPS)

    base_left = ladder_lines[0].left
    for line, twips in zip(ladder_lines, sweep.INDENT_LADDER_TWIPS):
        assert line.left - base_left == pytest.approx(
            twips_to_pt(twips), abs=FLOAT32_NOISE_PT
        ), f"indent {twips} twips"


def test_tab_stops_are_measured_from_the_text_column_not_the_page(document, oracle_lines):
    _, lines = oracle_lines
    tab_line = next(line for line in lines if line.text.count("H") >= 4 and len(line.text) <= 10)
    probes = [glyph for glyph in tab_line.glyphs if glyph.char == "H"]
    assert len(probes) == len(sweep.TAB_STOPS_TWIPS) + 1

    for glyph, twips in zip(probes[1:], sweep.TAB_STOPS_TWIPS):
        assert glyph.left - probes[0].left == pytest.approx(
            twips_to_pt(twips), abs=FLOAT32_NOISE_PT
        ), f"tab stop {twips} twips"


def test_every_baseline_sits_on_a_whole_device_pixel(oracle_lines):
    """The same 1/300-inch grid the page box snaps to.

    Seventy numbered single-line paragraphs, and every one of their baselines is a whole
    device pixel to within float32 noise.  Measured from **one glyph** -- the ``L`` every
    label starts with -- because a baseline estimated from a whole line moves with which
    of its glyphs overshoot, and that artefact read as a four-valued line advance the
    first time this was measured.

    The accumulator that puts them there is the next test.
    """
    page_sizes, lines = oracle_lines
    numbered = [line for line in lines if re.fullmatch(r"L\d\d", line.text)]
    assert len(numbered) == sweep.PAGINATION_LINES

    for line in numbered:
        probe = next(glyph for glyph in line.glyphs if glyph.char == "L")
        from_top_px = (page_sizes[line.page][1] - probe.bottom) * 300 / 72
        assert from_top_px == pytest.approx(round(from_top_px), abs=0.01), line.text


def test_pagination_is_emergent_and_this_is_where_word_put_the_break(oracle_lines):
    """An **observation**, pinned so a change to it is noticed.

    Nothing in the fixture says "page 2".  Word decided, and where it decided is the
    thing Phase 3 has to reproduce.  This test does not prove the renderer right -- there
    is no renderer -- it freezes the answer so that the day a model predicts 33 instead
    of 32, the disagreement is visible rather than discovered later.
    """
    _, lines = oracle_lines
    numbered = [line for line in lines if re.fullmatch(r"L\d\d", line.text)]
    by_page: dict[int, list[str]] = {}
    for line in numbered:
        by_page.setdefault(line.page, []).append(line.text)

    assert sorted(by_page) == [0, 1]
    assert by_page[0][0] == "L01" and by_page[0][-1] == "L32"
    assert by_page[1][0] == "L33" and by_page[1][-1] == "L70"


def test_the_vertical_model_places_every_baseline_of_the_fixture(oracle_pdf: Path):
    """All 91 baselines, to 0 device px: the Phase 0.6 accumulator, reproduced.

    Phase 0 could reproduce only 61 of the 70 numbered baselines, with ``round(origin +
    n * a)``.  ``docx2svg.vertical`` gets every line of every page: 14 pt headings mixed
    with 11 pt body lines, the page-2 restart, and section 2's different top margin.
    Line tops accumulate exactly; the baseline rounds against the line's top or bottom
    edge depending on the pitch's fractional pixel (see that module).  Read from the
    content stream, where Quartz writes baselines as integers, rather than from ink.
    """
    import quartz_pdf

    from docx2svg.vertical import FaceMetrics, baseline_px, line_pitch_px, twips_to_px

    calibri = FaceMetrics(2048, 1950, 550, 0)
    tops = [twips_to_px(sweep.MARGIN_TWIPS)] * 2 + [twips_to_px(sweep.ODD_MARGIN_TWIPS)]
    pages = quartz_pdf.read(oracle_pdf)
    assert len(pages) == 3
    checked = 0
    for page, top in zip(pages, tops):
        for line in quartz_pdf.lines([page]):
            # The heading runs are drawn at 58 device px (14 pt), body text at 46 (11 pt).
            half_points = sweep.HEADING_HALF_POINTS if line.runs[0].size_px == 58 else sweep.BODY_HALF_POINTS
            assert baseline_px(top, calibri, half_points) == round(line.y), line.text
            top += line_pitch_px(calibri, half_points)
            checked += 1
    assert checked == 91

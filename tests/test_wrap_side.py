"""Text beside a floating drawing against what Word drew, offline (ROADMAP.md, "Floating
drawings -- measured", F.8).

``tests/fixtures/wrap-side-observations.json`` holds, for ``tools/make_wrap_side_probe.py``,
Word's text objects and pictures and every face number the renderer asked for
(``tools/read_wrap_side_probe.py``).  Each document is laid out here from those numbers
alone, and every score is held to the recording.
"""

from __future__ import annotations

import json
import sys
from fractions import Fraction
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_wrap_side_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))

#: Lines of Word's not within half a pixel of the model's, per document and family, and why.
KNOWN = {
    # An edge on exactly half a pixel -- the text column's in mode 15, the page's below it --
    # goes to the pixel either way, by a rule not found: the tie family's pages, and the
    # polygon past its extent with its distance (tight and through, 13 lines each).
    "wrap-side-15": {"polygon": 26},
    "wrap-edge-15-1442": {"tie": 20},
    "wrap-edge-15-1440": {},
    "wrap-edge-none-1442": {"tie": 15},
    "wrap-edge-none-1440": {"tie": 15},
    # No settings: a tie as above (``min``, and a tab's segment in ``more``), and what is
    # recorded but not modelled below mode 15 -- a line no segment takes text at goes further
    # down than the drawing's foot; wrapText left with no room left of the drawing puts the
    # text right of it (``square``, and ``more``), and beside a 400-twip gap breaks words
    # into it; a hanging first line beside a drawing inside the indent.
    "wrap-side-none": {"min": 5, "more": 133, "square": 13},
}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", reader.documents(), ids=[name for name, _ in reader.documents()])
def test_text_beside_drawings_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["glyphs"] == recorded["glyphs"]
    assert result["lines"] == recorded["lines"]
    assert result["families"] == recorded["families"]
    assert result["images"] == recorded["images_score"]
    assert result["images"][0] == result["images"][1]
    assert result["warnings"] == recorded["warnings"]
    misses = {family: total - agree for family, (agree, total) in result["families"].items() if agree != total}
    assert misses == KNOWN[name]
    # Every glyph Word drew is drawn, and none it did not.
    score = result["result"]
    assert score.extra == 0 and score.not_drawn == 0


def test_a_segment_narrower_than_360_twips_takes_no_text():
    from docx2svg.wrap import Column, Wrap, segments

    column = Column(Fraction(300), 300, Fraction(1909 * 983), True, Fraction(1909))
    px = Fraction(300, 1440)

    def left_gap(twips):
        wrap = Wrap(Fraction(0), Fraction(100), 300 + twips * px, Fraction(1000), Fraction(0), Fraction(0),
                    "bothSides", Fraction(0), ())
        found, beside = segments([wrap], Fraction(10), Fraction(50), column)
        return [s for s in found if not s.edge], beside

    assert left_gap(359)[0] == [] and len(left_gap(360)[0]) == 1
    assert left_gap(360)[1]


def test_a_line_that_only_touches_a_drawing_is_not_beside_it():
    from docx2svg.wrap import Column, Wrap, segments

    column = Column(Fraction(300), 300, Fraction(1909 * 983), True, Fraction(1909))
    wrap = Wrap(Fraction(100), Fraction(200), Fraction(800), Fraction(1000), Fraction(0), Fraction(0), "bothSides",
                Fraction(100), ())
    assert segments([wrap], Fraction(200), Fraction(260), column) == (None, [])
    assert segments([wrap], Fraction(40), Fraction(100), column) == (None, [])
    assert segments([wrap], Fraction(199), Fraction(260), column)[0] is not None


def test_largest_takes_the_left_on_a_tie_and_the_wider_side_otherwise():
    from docx2svg.wrap import Column, Wrap, segments

    column = Column(Fraction(300), 300, Fraction(1000 * 983), True, Fraction(1000))

    def sides(left):
        wrap = Wrap(Fraction(0), Fraction(100), Fraction(300 + left), Fraction(300 + left + 200), Fraction(0),
                    Fraction(0), "largest", Fraction(0), ())
        found, _ = segments([wrap], Fraction(10), Fraction(50), column)
        return ["right" if s.edge else "left" for s in found]

    assert sides(400) == ["left"]
    assert sides(399) == ["right"]
    assert sides(401) == ["left"]


def test_a_polygon_span_is_read_over_the_band():
    from docx2svg.wrap import polygon_span

    triangle = ((Fraction(0), Fraction(100)), (Fraction(100), Fraction(100)), (Fraction(0), Fraction(0)))
    assert polygon_span(triangle, Fraction(10), Fraction(30)) == (0, 30)
    assert polygon_span(triangle, Fraction(150), Fraction(160)) is None

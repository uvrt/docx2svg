"""Drop caps against what Word drew, offline (ROADMAP.md, "Floating drawings -- measured",
F.19).

``tests/fixtures/drop-cap-observations.json`` holds, for ``tools/make_drop_cap_probe.py``,
Word's text objects and filled rectangles and every face number the renderer asked for
(``tools/read_drop_cap_probe.py``).  Each document is laid out here from those numbers
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

import read_drop_cap_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = [(name, data) for name, data in reader.documents() if name in DATA["documents"]]


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[name for name, _ in DOCUMENTS])
def test_drop_caps_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["glyphs"] == recorded["glyphs"]
    assert result["lines"] == recorded["lines"]
    assert result["families"] == recorded["families"]
    assert result["warnings"] == recorded["warnings"]
    # Every letter and every line beside it and after it where Word put it.
    assert result["lines"][0] == result["lines"][1]
    assert result["result"].extra == 0


def _drop(width_px: Fraction, hspace: int = 0):
    from docx2svg.paginate import DropCap
    from docx2svg.vertical import LAYOUT_UNIT_PX

    return DropCap(None, None, [], [], [], 0, width_px / LAYOUT_UNIT_PX, "drop", hspace, "text")


def test_the_text_beside_a_drop_cap_starts_after_its_width_and_half_a_twip():
    from docx2svg.vertical import LAYOUT_UNIT_PX

    extra = 103 * LAYOUT_UNIT_PX
    # Mode 15, the margin on the pixel grid: the width and 103 units, no rounding.
    start, shift = _drop(Fraction(1066, 10)).text_start(300, Fraction(300), True)
    assert (start, shift) == (300 + Fraction(1066, 10) + extra, 0)
    # The margin 5/12 into its pixel, the width 0.17 into its own: a pixel more.
    start, shift = _drop(Fraction(12817, 100)).text_start(300, Fraction(3605, 12), True)
    assert shift == 1
    # Below mode 15: on the page's nearest twip, the text on that pixel.
    start, shift = _drop(Fraction(12817, 100)).text_start(300, Fraction(3605, 12), False)
    assert (start * 1440 / 300).denominator == 1
    assert start + shift == 429


def test_a_frame_that_is_not_a_drop_cap_still_stops_the_layout():
    from docx2svg.paginate import is_drop_cap

    assert is_drop_cap({"framePr": {"dropCap": "drop", "lines": "3"}})
    assert is_drop_cap({"framePr": {"dropCap": "margin"}})
    assert not is_drop_cap({"framePr": {"dropCap": "none", "x": "100"}})
    assert not is_drop_cap({"framePr": {"x": "100", "y": "200"}})
    assert not is_drop_cap({})

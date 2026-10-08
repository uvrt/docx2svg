"""A line break that ends a paragraph puts the paragraph's mark on a line of its own
(ROADMAP.md, "Floating drawings -- measured", F.9): ``tools/make_break_end_probe.py``,
recorded by ``tools/read_break_end_probe.py`` in
``tests/fixtures/break-end-observations.json``.  Every glyph -- the lines after one, two
and three trailing breaks under every line rule included -- at Word's position and
baseline (820 of 1,034 baselines before the rule)."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_break_end_probe  # noqa: E402
import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(read_break_end_probe.OBSERVATIONS.read_text(encoding="utf-8"))


@pytest.mark.parametrize("name,data", read_break_end_probe.documents(),
                         ids=[n for n, _ in read_break_end_probe.documents()])
def test_the_mark_after_a_trailing_break_has_its_line(name, data):
    fonts = render_record.RecordedFonts(DATA["faces"])
    recorded = DATA["documents"][name]
    result, _rects, warnings = read_render.score(name, data, fonts, recorded["objects"], [])
    assert read_story_probe.row(result) == recorded["glyphs"]
    assert warnings == recorded["warnings"]
    assert result.extra == 0 and result.not_drawn == 0 and result.matched == result.word
    assert result.x_agree == result.y_agree == result.face_agree == result.size_agree == result.matched


def test_a_trailing_break_ends_a_line_of_its_own():
    from docx2svg.linebreak import BREAK, GLYPH, Geometry, Piece, break_pieces

    items = [Piece(GLYPH, "a", 100, None, 0), Piece(BREAK, "\n", 0, None, 1)]
    lines = break_pieces(items, Geometry(0, 0, 10 ** 6), None)
    assert [(line.start, line.end) for line in lines] == [(0, 2), (2, 2)]
    items[1] = Piece(BREAK, "\f", 0, None, 1)
    assert len(break_pieces(items, Geometry(0, 0, 10 ** 6), None)) == 1

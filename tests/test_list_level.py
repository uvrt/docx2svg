"""List levels a paragraph's numbering does not give as named (round 24): a direct
``w:numPr`` of only ``w:ilvl`` takes the level's indent over the style's; any level of a
single-level definition is its level 0, counted on; a level a multilevel definition does
not have draws no number, its first line at its left indent; a style's own list level tied
to another style draws no number.  ``tools/make_list_level_probe.py``, recorded by
``tools/read_list_level_probe.py`` in ``tests/fixtures/list-level-observations.json``."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import read_list_level_probe as reader  # noqa: E402
import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = reader.documents()


def _score(name, data):
    fonts = render_record.RecordedFonts(DATA["faces"])
    return read_render.score(name, data, fonts, DATA["documents"][name]["objects"], [])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[n for n, _ in DOCUMENTS])
def test_every_label_and_line_where_word_draws_it(name, data):
    pytest.importorskip("numpy")
    result, _rects, warnings = _score(name, data)
    assert read_story_probe.row(result) == DATA["documents"][name]["glyphs"]
    assert warnings == DATA["documents"][name]["warnings"]
    assert result.matched == result.word == result.model
    assert result.y_agree == result.face_agree == result.matched
    # Three labels "i." a thousandth of a pixel from Word's.
    assert result.x_agree == result.matched - 3


@pytest.mark.parametrize("index,expected", [
    (0, (312, 305, 312, 1, 5)),  # a direct w:ilvl alone: the style's indent over the level's
    (1, (296, 283, 275, 25, 21)),  # a level past the definition's, or tied elsewhere, taken as named
])
@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[n for n, _ in DOCUMENTS])
def test_refuted_rules_keep_their_scores(name, data, index, expected):
    pytest.importorskip("numpy")
    _what, variant = reader.refuted()[index]
    with variant():
        result, _rects, _warnings = _score(name, data)
    assert (result.matched, result.x_agree, result.y_agree, result.extra, result.not_drawn) == expected

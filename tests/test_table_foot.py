"""Table rows at a page's foot (ROADMAP.md, "Tables -- measured", stage 6a): a row holding a
cell merged down splits as any row does, its merged cell's lines going on at the next
page's top; a header row is never split, and in mode 15 is not left at the foot with no
row after it.  ``tools/make_table_foot_probe.py``, recorded by
``tools/read_table_foot_probe.py`` in ``tests/fixtures/table-foot-observations.json``."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import read_table_foot_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = reader.documents()


def _score(name, data, fills=True):
    fonts = render_record.RecordedFonts(DATA["faces"])
    recorded = DATA["documents"][name]
    return read_render.score(name, data, fonts, recorded["objects"], recorded["fills"] if fills else [])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[n for n, _ in DOCUMENTS])
def test_every_glyph_and_border_where_word_draws_it(name, data):
    pytest.importorskip("numpy")
    result, rects, warnings = _score(name, data)
    assert read_story_probe.row(result) == DATA["documents"][name]["glyphs"]
    assert warnings == [] and rects["all"] == DATA["documents"][name]["rects"]
    assert result.matched == result.word == result.model
    assert result.x_agree == result.y_agree == result.face_agree == result.matched
    assert rects["all"][2:] == [0, 0]


@pytest.mark.parametrize("index,setting,expected", [
    (0, "none", (7447, 7318, 84, 84)), (0, "15", (7405, 7368, 6, 6)),  # a header row split like any row
    (1, "none", (7531, 7531, 0, 0)), (1, "15", (7411, 7411, 120, 0)),  # mode 15: a header left at the foot
])
def test_refuted_rules_keep_their_scores(index, setting, expected):
    name, data = dict(enumerate(DOCUMENTS))[0 if setting == "none" else 1]
    _what, variant = reader.refuted()[index]
    with variant():
        result, _rects, _warnings = _score(name, data, fills=False)
    assert (result.matched, result.y_agree, result.extra, result.not_drawn) == expected

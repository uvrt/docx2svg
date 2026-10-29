"""Paragraphs whose indents leave their text little or no width (round 24): in a
fixed-layout table and in the body Word breaks inside the words, a character to a line, as
the renderer does; an autofit table's column is widened to hold a cell's indents and its
widest word -- in mode 15 also in a table as wide as the column, which is wider than its
room by its borders' halves.  ``tools/make_narrow_text_probe.py``, recorded by
``tools/read_narrow_text_probe.py`` in ``tests/fixtures/narrow-text-observations.json``.

Not yet as Word: a word exactly as wide as a column at its narrowest is broken inside it
(its last letter on the next line), and in mode 15 an autofit table whose cells' indents
leave less than their widest word goes past the room, which the renderer keeps to."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import read_narrow_text_probe as reader  # noqa: E402
import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = reader.documents()


def _score(name, data):
    fonts = render_record.RecordedFonts(DATA["faces"])
    return read_render.score(name, data, fonts, DATA["documents"][name]["objects"], [])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[n for n, _ in DOCUMENTS])
def test_every_glyph_scores_as_recorded(name, data):
    pytest.importorskip("numpy")
    result, _rects, warnings = _score(name, data)
    assert read_story_probe.row(result) == DATA["documents"][name]["glyphs"]
    assert warnings == DATA["documents"][name]["warnings"]


def test_mode_15_column_wide_table_refused_scores_lower():
    pytest.importorskip("numpy")
    name, data = DOCUMENTS[1]
    _what, variant = reader.refuted()[0]
    with variant():
        result, _rects, _warnings = _score(name, data)
    assert (result.matched, result.x_agree, result.y_agree, result.extra, result.not_drawn) == (1059, 834, 714, 331, 331)
    assert result.matched < DATA["documents"][name]["glyphs"][0]

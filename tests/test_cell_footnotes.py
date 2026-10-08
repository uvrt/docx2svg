"""Footnotes referenced in a table cell, laid out and drawn (ROADMAP.md, "Phase 4 --
measured", 4.19).  ``tools/make_cell_footnote_probe.py``, recorded by
``tools/read_cell_footnote_probe.py`` in ``tests/fixtures/cell-footnote-observations.json``."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest import mock

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import read_cell_footnote_probe as reader  # noqa: E402
import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = reader.documents()
#: With no ``settings.xml``, in the split case filled with 44 lines Word moves the line
#: above the one with the reference too (not found): its 24 glyphs a page lower, and the
#: reference's number with them.
KNOWN = {"cellfn-none": (24, 2), "cellfn-15": (0, 0)}


def _score(name, data):
    fonts = render_record.RecordedFonts(DATA["faces"])
    return read_render.score(name, data, fonts, DATA["documents"][name]["objects"], [])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[n for n, _ in DOCUMENTS])
def test_a_cell_s_footnotes_where_word_draws_them(name, data):
    result, _rects, warnings = _score(name, data)
    assert read_story_probe.row(result) == DATA["documents"][name]["glyphs"]
    assert warnings == []
    off, extra = KNOWN[name]
    assert result.word == 22363 and result.matched - result.y_agree == off
    assert result.extra == extra and result.face_agree == result.size_agree == result.matched
    # Every glyph off Word's x is an object start a layout unit or less off (0.0004 px).
    assert all(abs(float(problem.rsplit("dx=", 1)[1])) < 0.001 for problem in result.problems
               if "dx=" in problem and problem.split("y=")[1].split()[0] == problem.split("y=")[2].split()[0])


def test_before_a_cell_s_footnotes_took_no_room():
    name, data = DOCUMENTS[1]
    _what, patch = reader.refuted()[0]
    with mock.patch.object(*patch):
        result, _rects, _warnings = _score(name, data)
    assert (result.matched, result.y_agree) == (4391, 3201)

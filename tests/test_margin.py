"""A left margin off the device-pixel grid (ROADMAP.md, "Headers, footers and fields --
measured", H.5): Word starts the text column at the margin rounded to a whole pixel, held
in whole layout units.  ``tools/make_margin_probe.py``, recorded by
``tools/read_margin_probe.py`` in ``tests/fixtures/margin-observations.json``."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_margin_probe  # noqa: E402
import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(read_margin_probe.OBSERVATIONS.read_text(encoding="utf-8"))

#: Glyphs off Word's pen x: in mode 15, three right-aligned footer lines a layout unit
#: off (0.001 px) -- not found.
KNOWN = {"margin-none": 0, "margin-15": 3}


@pytest.mark.parametrize("name,data", read_margin_probe.documents(), ids=[n for n, _ in read_margin_probe.documents()])
def test_every_line_starts_where_word_starts_it(name, data):
    fonts = render_record.RecordedFonts(DATA["faces"])
    recorded = DATA["documents"][name]
    result, _rects, warnings = read_render.score(name, data, fonts, recorded["objects"], [])
    assert read_story_probe.row(result) == recorded["glyphs"]
    assert warnings == recorded["warnings"]
    assert result.extra == 0 and result.not_drawn == 0 and result.matched == result.word
    assert result.matched - result.x_agree == KNOWN[name]
    assert result.y_agree == result.face_agree == result.size_agree == result.matched

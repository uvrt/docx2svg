"""A square-wrapped text box beside several tables (ROADMAP.md, "Floating drawings --
measured", F.22): docx-agent saw Word fit three lines more on such a page; isolated here,
every page break is Word's.  ``tools/make_wrap_tables_probe.py``, recorded by
``tools/read_wrap_tables_probe.py`` in ``tests/fixtures/wrap-tables-observations.json``."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import read_wrap_tables_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = reader.documents()


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[n for n, _ in DOCUMENTS])
def test_every_glyph_on_word_s_page(name, data):
    fonts = render_record.RecordedFonts(DATA["faces"])
    recorded = DATA["documents"][name]
    result, _rects, warnings = read_render.score(name, data, fonts, recorded["objects"], [])
    assert read_story_probe.row(result) == recorded["glyphs"]
    assert warnings == []
    assert result.matched == result.word == result.model == 2333 and result.x_agree == result.matched
    # Every page's lines Word's, on Word's baselines: two tables with nothing between them
    # (``adjacent``) stand on one edge, the second the narrower border higher (243 glyphs
    # 2 px low before; ``paginate.table_gap_px``).
    assert result.matched == result.y_agree

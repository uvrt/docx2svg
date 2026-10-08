"""The renderer against what Word drew, offline (ROADMAP.md, "Phase 5 -- measured").

``tests/fixtures/render-observations.json`` holds, for every committed document and both
render probes (``tools/make_render_probe.py``), every text object and filled rectangle of
Word's PDF, and every face number the renderer asked for.  Each document is rendered
here from those numbers alone, and:

* **the SVG is a transcription of the layout**: every glyph's ``x`` and ``y`` in the
  emitted SVG are the layout's (``read_render.score`` asserts it);
* **the layout is Word's**: on every committed document every glyph the renderer draws
  is matched to one Word drew, at Word's pen position (a text object's start exactly; an
  advance inside one within Quartz's encoding), on Word's baseline, in Word's face and
  size -- and the scores of the probes are pinned, their known disagreements included;
* the rectangles cover the pixels they covered when recorded, per document.

Needs numpy for the rectangle comparison (the ``dev`` extra); skipped without it.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_render  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(read_render.OBSERVATIONS.read_text(encoding="utf-8"))

#: What each probe is known to get wrong, with the reason (ROADMAP.md, "Phase 5 --
#: measured"): the glyphs off Word's position, and what they are.
KNOWN = {
    # The last lines of the distributed paragraphs, and the lines of the narrow columns
    # where Word spreads slack over characters as well as spaces (871 glyphs).
    "render-none": 871,
    # The distributed paragraphs, which in mode 15 also break otherwise than the model
    # (not measured: Phase 3.8 measured justified text); and five glyphs of compressed
    # justified lines a layout unit off (0.001 px: how Word rounds each space's share).
    "render-15": 1093,
}


def _documents():
    probes = dict(read_render.documents())
    return [(name, probes[name]) for name in DATA["documents"]]


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", _documents(), ids=[name for name, _ in _documents()])
def test_render_against_word(name, data, fonts):
    pytest.importorskip("numpy")
    recorded = DATA["documents"][name]
    result, rects, warnings = read_render.score(name, data, fonts, recorded["objects"], recorded["fills"])
    assert read_render.row(result) == recorded["glyphs"]
    assert rects == recorded["rects"]
    assert warnings == recorded["warnings"]
    disagreeing = result.matched - min(result.x_agree, result.y_agree, result.face_agree, result.size_agree)
    if name in KNOWN:
        assert result.matched - result.x_agree == KNOWN[name]
    else:
        # Every glyph drawn is one Word drew, and every one is where Word drew it.
        assert result.extra == 0 and result.matched == result.model
        assert disagreeing == 0, result.problems[:5]


def test_the_committed_documents_are_exact_to_the_pixel_where_they_draw_rules(fonts):
    """``sample-long`` and ``sample-resume`` draw paragraph borders; ``sample-simple`` a
    border and an underline; ``sample-simple``, ``sample-5pages``, ``sample-10pages`` and
    ``sample-with-table`` Light Grid tables: every rectangle exact.  The tables' two
    residuals of Tables stage 5 -- the left border a pixel right of Word's, the shading
    painted up to the lines where Word leaves a pixel column white -- are closed
    (``make_table_line_probe.py``, ROADMAP.md, Phase 5.14), and so is stage 6's: the
    border under a row at a page's foot (Phase 5.15)."""
    pytest.importorskip("numpy")
    for name in ("samplelib/sample-long.docx", "samplelib/sample-resume.docx", "samplelib/sample-simple.docx",
                 "wordto/sample-5pages.docx", "wordto/sample-10pages.docx", "wordto/sample-with-table.docx"):
        assert DATA["documents"][name]["rects"]["all"][2:] == [0, 0], name

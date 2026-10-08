"""Headers, footers and fields against what Word drew, offline (ROADMAP.md, "Headers,
footers and fields -- measured").

``tests/fixtures/story-observations.json`` holds, for every document of
``tools/make_story_probe.py``, ``make_story_select_probe.py`` and ``make_field_probe.py``,
Word's text objects and filled rectangles and every face number the renderer asked for.
Each document is rendered here from those numbers alone, and every glyph -- body, header
and footer -- is matched to Word's (``tools/read_render.score``), as the committed
documents' are in ``test_render.py``.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(read_story_probe.OBSERVATIONS.read_text(encoding="utf-8"))

#: What each document is known to get wrong, and why: glyphs matched but off Word's
#: position (x, baseline), and header and footer glyphs matched but not exact.
KNOWN = {
    # The two ``w:vAlign`` cases: Word centres or bottom-aligns the body, which is drawn
    # from the top here (102 glyphs' baselines, warned); their stories are exact.
    **{f"story-{setting}": {"x": 0, "y": 102, "story": 0} for setting in ("none", "12", "14", "15")},
    # Fields Word computes again on export that are not computed here (DATE, TIME, IF) and
    # number forms in the application's language: drawn as cached, and warned of.  REF,
    # PAGEREF and SEQ are computed (H.8): 115 off before.
    "fields-word": {"x": 110, "y": 0, "story": 102},
}


def _documents():
    probes = dict(read_story_probe.documents())
    return [(name, probes[name]) for name in DATA["documents"]]


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", _documents(), ids=[name for name, _ in _documents()])
def test_stories_against_word(name, data, fonts):
    pytest.importorskip("numpy")
    recorded = DATA["documents"][name]
    result, rects, warnings = read_render.score(name, data, fonts, recorded["objects"], recorded["fills"])
    assert read_story_probe.row(result) == recorded["glyphs"]
    assert rects == recorded["rects"]
    assert warnings == recorded["warnings"]
    known = KNOWN.get(name, {"x": 0, "y": 0, "story": 0})
    assert result.matched - result.x_agree == known["x"]
    assert result.matched - result.y_agree == known["y"]
    assert result.story_matched - result.story_exact == known["story"]
    assert result.face_agree == result.size_agree == result.matched
    if name != "fields-word":
        # Every glyph drawn is one Word drew, and every glyph Word drew is drawn.
        assert result.extra == 0 and result.not_drawn == 0
        assert result.story_matched == result.story_model > 0
        assert rects["all"][2:] == [0, 0]


def test_the_breaker_measures_the_computed_page_number():
    """In the section numbered from 1,000,000 of ``fields-computed``, the line breaks
    Word made (every glyph on its baseline, above) need the computed value: with the
    cached ``1`` measured instead, some ``PAGE`` and ``NUMPAGES`` paragraphs break
    elsewhere."""
    import make_field_probe

    from docx2svg import convert_docx_to_layout, layout as layout_module

    fonts = render_record.RecordedFonts(DATA["faces"])
    data = make_field_probe.build("computed")

    def lines(options):
        laid = convert_docx_to_layout(data, options)
        return [(page.number, line.baseline, "".join("".join(s.chars) for s in line.spans))
                for page in laid.pages for line in page.lines if line.story is None]

    computed = lines(render_record.options(fonts))
    original = layout_module.substitute_document
    try:
        layout_module.substitute_document = lambda document, values: document
        cached = lines(render_record.options(fonts))
    finally:
        layout_module.substitute_document = original
    wide = [line for line in computed if line[2].startswith("W")]
    assert wide and computed != cached
    assert sum(1 for line in cached if line not in computed) >= 2

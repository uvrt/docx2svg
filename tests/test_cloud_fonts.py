"""Office's cloud-font cache as a font source (ROADMAP.md, "Phase 5 -- measured", 5.11):
the faces Office for Mac downloads -- Aptos Display, Word 365's heading face, among them --
read in place after every installed face.  ``tools/make_cloud_font_probe.py``, recorded by
``tools/read_cloud_font_probe.py`` in ``tests/fixtures/cloud-font-observations.json``."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))
sys.path.insert(0, str(REPO / "src"))

import read_cloud_font_probe  # noqa: E402
import read_render  # noqa: E402
import read_story_probe  # noqa: E402
import render_record  # noqa: E402

from docx2svg import fonts  # noqa: E402

DATA = json.loads(read_cloud_font_probe.OBSERVATIONS.read_text(encoding="utf-8"))


@pytest.mark.parametrize("name,data", read_cloud_font_probe.documents(),
                         ids=[n for n, _ in read_cloud_font_probe.documents()])
def test_every_glyph_word_draws_in_a_cloud_face_is_drawn(name, data):
    """Every glyph drawn and matched -- before, the layout stopped at the first paragraph,
    0 of 4,336 drawn.  The residuals are faces' own vertical metrics (Raleway's baselines
    1 px, Anton's and Merriweather Sans Light's) and, with no compatibility mode, Segoe
    UI's synthesised bold italic: recorded, not fitted."""
    faces = render_record.RecordedFonts(DATA["faces"])
    recorded = DATA["documents"][name]
    result, _rects, warnings = read_render.score(name, data, faces, recorded["objects"], [])
    assert read_story_probe.row(result) == recorded["glyphs"]
    assert warnings == recorded["warnings"] == []
    assert result.extra == 0 and result.not_drawn == 0 and result.matched == result.word == 4336


def test_aptos_display_is_exact():
    """The face every new Word 365 document's headings are in: every glyph at Word's pen
    position, baseline, face and size, in both settings."""
    for name, data in read_cloud_font_probe.documents():
        faces = render_record.RecordedFonts(DATA["faces"])
        result, _rects, _warnings = read_render.score(name, data, faces, DATA["documents"][name]["objects"], [])
        assert not [problem for problem in result.problems if problem.startswith("p1 ")], name


def test_the_cache_is_searched_after_every_installed_folder(tmp_path):
    for family in ("Zeta", "Alpha"):
        (tmp_path / family).mkdir()
    (tmp_path / "loose.ttf").write_bytes(b"")
    assert fonts.cloud_font_dirs(tmp_path) == (tmp_path / "Alpha", tmp_path / "Zeta")
    assert fonts.cloud_font_dirs(tmp_path / "missing") == ()
    cloud = fonts.cloud_font_dirs()
    assert fonts.FONT_DIRS[len(fonts.FONT_DIRS) - len(cloud):] == cloud
    assert fonts.WORD_FONTS in fonts.FONT_DIRS[:len(fonts.FONT_DIRS) - len(cloud)]

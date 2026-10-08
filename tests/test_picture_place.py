"""Where inline pictures are drawn, and at what size, against Word.

``tests/fixtures/picture-place-observations.json`` holds Word's image boxes for
``tools/make_picture_place_probe.py`` and ``tools/make_picture_probe.py``
(``tools/read_picture_place_probe.py``).  Offline, from the recorded face numbers:

* a picture is drawn at its extent scaled, both ways alike, to fit the extent truncated
  to whole twips (``layout.picture_size``);
* its bottom, with its effect extent, stands on the text's baseline -- under a taller
  mark, at the bottom of the mark's height; an ``auto`` multiple's extra below it, an
  ``atLeast`` line's above it; in an ``exact`` line four fifths of the way down
  (``lines.LineHeight.object_drop``).

The one box 5.14 left off was a line-height question, not a drawing one, since
settled: in mode 15 a picture's line is at least its run's text height (ROADMAP.md,
Phase 5.16).
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_picture_place_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
KNOWN: dict[str, int] = {}


@pytest.mark.parametrize("name,data", reader.documents(), ids=[name for name, _ in reader.documents()])
def test_pictures_where_word_drew_them(name, data):
    from docx2svg import _render

    fonts = render_record.RecordedFonts(DATA["faces"])
    _, layout, _, _ = _render(data, render_record.options(fonts))
    agree, scored, problems = reader.compare(DATA["documents"][name], reader.model_boxes(layout))
    assert scored == len(DATA["documents"][name])
    assert scored - agree == KNOWN.get(name, 0), problems

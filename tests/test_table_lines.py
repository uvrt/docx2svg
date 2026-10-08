"""Table border lines and cell shading, to the device pixel, against Word.

``tools/make_table_line_probe.py`` sweeps a table's grid lines over every fraction of a
pixel a twip can reach, on both sides of the text column's edge, with the page's left
margin on and off the pixel grid; ``tests/fixtures/table-line-observations.json`` holds
what Word drew at every place the model draws a line or a shading
(``tools/read_table_line_probe.py``).  Offline, from the recorded face numbers.

* every **vertical line** where Word drew it: its centre is the text column's left edge
  rounded, plus the grid line's distance from it rounded half away from zero
  (``layout.grid_line_px``);
* every **horizontal line** where Word drew it: the band below the last row on a page is
  drawn up from its bottom rounded (Tables stage 6's residual, Phase 5.15);
* every **shading** edge where Word drew it but those at exactly half a pixel, which Word
  rounds either way by a rule not found (ROADMAP.md, Phase 5.14).
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_table_line_probe as probe  # noqa: E402
import read_table_line_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))

#: Shading edges at exactly half a pixel that Word rounds down (cells whose shading
#: differs from Word's, of 864 a document); every other edge is exact.
SHADING_TIES = {"none": 2, "15": 98, "none-m1442": 7, "15-m1442": 0, "none-m1438": 76, "15-m1438": 32}


@pytest.mark.parametrize("setting", list(probe.SETTINGS))
def test_table_lines_and_shading_where_word_drew_them(setting):
    from docx2svg import _render

    recorded = DATA["documents"][f"table-line-{setting}"]
    fonts = render_record.RecordedFonts(DATA["faces"])
    _, layout, _, _ = _render(probe.build(setting), render_record.options(fonts))
    ours = reader.observe(layout, reader.model_fills(layout))
    keys = sorted(ours)
    assert reader.keys_digest(keys) == recorded["keys"], "the model draws other rules than when recorded"
    word = {key: value for key, value in zip(keys, recorded["word"])}
    score = reader.compare(word, ours)
    assert score["v agree"] == score["v scored"] == 1920
    assert score["h agree"] == score["h scored"] == 2376
    assert score["s scored"] == 864
    assert score["s scored"] - score["s agree"] == SHADING_TIES[setting]

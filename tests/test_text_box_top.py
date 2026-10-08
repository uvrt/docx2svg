"""Where a text box's text starts, swept by fractions of a pixel (ROADMAP.md, "Floating
drawings -- measured", F.14): ``tools/make_text_box_top_probe.py``, recorded by
``tools/read_text_box_top_probe.py`` in ``tests/fixtures/text-box-top-observations.json``.

Word holds a text box in whole twips -- its corner the nearest twip, each side's inset
plus its distance in from the shape's edge rounded to the nearest twip (the outline's half
whole EMU, truncated) -- and starts the text at that point in the nearest layout unit: its
lines on that point's device pixel, a half up.  Below mode 15 the room between the insets
is a twip shorter (204 layout units).  The box's last line keeps the rest of the room
below its text, and its baseline rounds as a line box with space below it does.  Every
one of 1,710 swept lines, in all three settings, on Word's pen x and baseline (1,178,
1,178 and 1,211 baselines and 1,706 pen x before); the four turned boxes are recorded,
not laid out as Word turns them."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_text_box_top_probe as probe  # noqa: E402
import read_text_box_top_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))


@pytest.mark.parametrize("setting", list(probe.SETTINGS))
def test_swept_text_boxes_against_word(setting):
    fonts = render_record.RecordedFonts(DATA["faces"])
    recorded = DATA["documents"][setting]
    model = reader.model_positions(probe.build(setting), fonts)
    assert reader.misses(recorded["word"], model, 1) == recorded["x_misses"] == []
    assert reader.misses(recorded["word"], model, 2) == recorded["y_misses"] == []


def test_turned_boxes_are_recorded():
    """A box flipped horizontally keeps its text where the unturned box would (the
    model's); flipped vertically or turned, Word turns the text with the box, which the
    model does not: recorded, with the model's miss, for a later stage."""
    word = DATA["documents"]["15"]["word"]
    fonts = render_record.RecordedFonts(DATA["faces"])
    model = reader.model_positions(probe.build("15"), fonts)
    for box in probe.BOXES:
        if box.sweep.family != "turn":
            continue
        w, m = word[box.label], model[box.label]
        if box.sweep.xfrm == "flipH":
            assert w == [m[0], m[1], m[2]]
        else:
            assert abs(w[1] - m[1]) > 10 and abs(w[2] - m[2]) > 10

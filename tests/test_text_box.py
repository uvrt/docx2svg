"""Where a text box's text starts (ROADMAP.md, "Floating drawings -- measured", F.13 and
F.14): ``tools/make_text_box_probe.py``, recorded by ``tools/read_text_box_probe.py`` in
``tests/fixtures/text-box-observations.json``.  The insets count from the geometry's text
rectangle, itself inside half the outline's width -- drawn or ``a:noFill``, stated or the
style's (F.13: 34 / 115 first glyphs within half a pixel of Word's before, 100 and 101
after); each side in whole twips, the text at its point's layout unit and pixel, a half
up, and the box's last line keeping its room below it (F.14): 113 / 115 in both settings.
What is left is pinned: centred text under ``wrap="none"``, which Word does not centre on
the box's inner width."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import make_text_box_probe  # noqa: E402
import read_text_box_probe  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(read_text_box_probe.OBSERVATIONS.read_text(encoding="utf-8"))


@pytest.mark.parametrize("setting", list(make_text_box_probe.SETTINGS))
def test_text_box_first_glyphs_against_word(setting):
    fonts = render_record.RecordedFonts(DATA["faces"])
    recorded = DATA["documents"][setting]
    model = read_text_box_probe.model_positions(make_text_box_probe.build(setting), fonts)
    assert read_text_box_probe.misses(recorded["word"], model) == recorded["misses"]
    assert recorded["misses"] == ["B87", "B88"]


def test_the_outline_and_the_geometry_move_the_text_in():
    """Every left-aligned, wrapped box of the line, inset and geometry families -- an
    outline of 0.5 to 8 pt, drawn or ``a:noFill``, stated or the style's; a rounded
    rectangle, an ellipse, a triangle, an octagon, a custom text rectangle -- puts its first
    glyph's pen x on Word's, in both settings."""
    fonts = render_record.RecordedFonts(DATA["faces"])
    for setting in make_text_box_probe.SETTINGS:
        model = read_text_box_probe.model_positions(make_text_box_probe.build(setting), fonts)
        word = DATA["documents"][setting]["word"]
        for number, box in enumerate(make_text_box_probe.CASES):
            if box.family in ("line", "inset", "geometry") and box.jc is None and box.wrap == "square":
                key = f"B{number}"
                assert abs(word[key][1] - model[key][1]) <= 0.5, (setting, key, box.note)

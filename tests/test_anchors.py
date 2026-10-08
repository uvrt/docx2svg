"""Floating drawings against what Word drew, offline (ROADMAP.md, "Floating drawings --
measured").

``tests/fixtures/anchor-observations.json`` holds, for ``tools/make_anchor_probe.py`` and
``tools/make_drawing_probe.py``, Word's text objects, filled rectangles and pictures (box,
paint order and layer) and every face number the renderer asked for
(``tools/read_anchor_probe.py``).  Each document is laid out here from those numbers
alone.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_anchor_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))

#: What each document is known to get wrong, and why.
KNOWN = {
    # Pictures: relativeHeight 1-3, below the range Word writes, are not stacked by value
    # (two pages, five pictures); allowOverlap="0" moves the second picture clear of the
    # first (warned, two).  The last five pages' anchors in a table cell are where Word
    # drew them since F.16 (they stopped the layout before).
    **{f"anchor-{s}": {"pictures": 7, "baselines": 0} for s in ("none", "15")},
    # A text box in a group Word draws with its own spacing (one picture, and 11 glyphs'
    # advances).  Every text box line on Word's baseline since F.14 (a text box's last
    # line keeps its room below it; the box in whole twips): 59 and 519 missed before.
    "drawing-none": {"pictures": 1, "baselines": 0},
    "drawing-15": {"pictures": 1, "baselines": 0},
    # A header's and a footer's drawings: every picture where Word drew it, in its order.
    **{f"story-anchor-{s}": {"pictures": 0, "baselines": 0} for s in ("none", "15")},
    # Text above and below a wrapTopAndBottom drawing: every line and picture Word's.
    **{f"wrap-anchor-{s}": {"pictures": 0, "baselines": 0} for s in ("none", "15")},
    # Such drawings in consecutive paragraphs, each positioned below the text the one
    # before pushed down, the last with its paragraph on the next page (F.24).
    **{f"wrap-stack-{s}": {"pictures": 0, "baselines": 0} for s in ("none", "15")},
}


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", reader.documents(), ids=[name for name, _ in reader.documents()])
def test_floating_drawings_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["glyphs"] == recorded["glyphs"]
    assert result["images"] == recorded["images_score"]
    assert result["fills"] == recorded["fills_score"]
    assert result["warnings"] == recorded["warnings"]
    score = result["result"]
    known = KNOWN[name]
    assert score.extra == 0
    assert score.face_agree == score.size_agree == score.matched
    assert score.matched - score.y_agree == known["baselines"]
    assert result["images"][1] - result["images"][0] == known["pictures"]
    # Every rectangle a shape fills is Word's, to the precision Quartz writes.
    assert result["fills"][0] == result["fills"][1]


def test_the_body_does_not_move():
    """No glyph of the anchor probe's body moves: a drawing text does not wrap around takes
    no room in its line (every body glyph matched is at Word's position and baseline)."""
    for name in ("anchor-none", "anchor-15"):
        row = DATA["documents"][name]["glyphs"]
        matched, _model, _word, exact, exact_scored, step, step_scored, y = row[:8]
        assert exact == exact_scored and step == step_scored and y == matched


def test_offsets_and_extents_truncate_to_twips():
    from fractions import Fraction

    from docx2svg.floating import Frames, emu_twips, horizontal, vertical
    from docx2svg.model import Anchor, AnchorPosition

    frames = Frames(11906, 16838, 1442, 1300, 1500, 1600, 700, 650, 1, True, Fraction(1768), Fraction(1768),
                    Fraction(2037), Fraction(1442))
    assert emu_twips(400003) == 629 and emu_twips(-250001) == -393
    anchor = Anchor((1248113, 1181039), h=AnchorPosition("margin", align="right"),
                    v=AnchorPosition("margin", align="bottom"))
    # The right margin at 10606 twips less the extent truncated (1965): 8641.
    assert horizontal(anchor, frames) == 8641
    assert vertical(anchor, frames) == 16838 - 1600 - 1859
    centred = Anchor((1244412, 1144036), h=AnchorPosition("margin", align="center"))
    assert horizontal(centred, frames) == Fraction(10089, 2)


def test_theme_colour_transforms():
    from ooxml_common.drawingml.model import ColorTransform, SchemeColor, SrgbColor

    from docx2svg.drawing import Colours

    colours = Colours({"accent1": "4472C4", "accent2": "ED7D31", "dk1": "000000"}, {})
    assert colours.resolve(SchemeColor("accent1", [ColorTransform("lumMod", 75000)])).hex == "#2F5597"
    assert colours.resolve(SchemeColor("accent2", [ColorTransform("lumMod", 60000),
                                                   ColorTransform("lumOff", 40000)])).hex == "#F4B183"
    # 127.5 levels: a half rounds down.
    assert colours.resolve(SchemeColor("tx1", [ColorTransform("lumMod", 50000),
                                               ColorTransform("lumOff", 50000)])).hex == "#7F7F7F"
    assert colours.resolve(SrgbColor("4472C4", [ColorTransform("lumOff", 40000),
                                                ColorTransform("lumMod", 60000)])).hex == "#517CC8"
    assert not colours.warnings


def test_word_colour_map_is_spelled_as_drawingml_slots():
    from ooxml_common.drawingml.model import SchemeColor

    from docx2svg.drawing import Colours, color_map_slots

    assert color_map_slots({"bg1": "light1", "t1": "dark1", "hyperlink": "hyperlink"}) == {
        "bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2", "hlink": "hlink"}
    swapped = Colours({"lt1": "FFFFFF", "dk1": "000000"}, {"bg1": "dark1", "t1": "light1"})
    assert swapped.resolve(SchemeColor("tx1")).hex == "#FFFFFF"
    assert swapped.resolve(SchemeColor("bg1")).hex == "#000000"
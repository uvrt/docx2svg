"""Picture bullets against what Word drew, offline (ROADMAP.md, "List labels -- measured",
"Picture bullets").

``tests/fixtures/picture-bullet-observations.json`` holds, for
``tools/make_picture_bullet_probe.py``, Word's text objects, filled rectangles and image boxes
and every face number the renderer asked for (``tools/read_picture_bullet_probe.py``).  Each
document is laid out here from those numbers alone, and every score is held to the
recording.
"""

from __future__ import annotations

import json
import sys
from fractions import Fraction
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_picture_bullet_probe as reader  # noqa: E402
import render_record  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
DOCUMENTS = [(name, data) for name, data in reader.documents() if name in DATA["documents"]]


@pytest.fixture(scope="module")
def fonts():
    return render_record.RecordedFonts(DATA["faces"])


@pytest.mark.parametrize("name,data", DOCUMENTS, ids=[name for name, _ in DOCUMENTS])
def test_picture_bullets_against_word(name, data, fonts):
    recorded = DATA["documents"][name]
    result = reader.score(name, data, fonts, recorded)
    assert result["glyphs"] == recorded["glyphs"]
    assert result["lines"] == recorded["lines"]
    assert result["families"] == recorded["families"]
    assert result["pictures"] == recorded["pictures"]
    assert result["warnings"] == recorded["warnings"]
    # Every line and every bullet where Word put it, in every setting.
    assert result["lines"][0] == result["lines"][1]
    assert result["pictures"][0] == result["pictures"][1]
    assert result["result"].extra == 0


def test_a_picture_bullet_is_sized_by_the_label_and_the_pictures_pixels():
    from docx2svg.linebreak import picture_bullet_units

    # 11 pt, a 16 px picture: 10.125 pt; 8 px twice that; 40 pt 42.75 pt.
    assert picture_bullet_units(22, (16, 16)) == Fraction(10125, 1000) * 4096
    assert picture_bullet_units(22, (8, 8)) == 2 * picture_bullet_units(22, (16, 16))
    assert picture_bullet_units(80, (16, 16)) == Fraction(4275, 100) * 4096
    # The height counts, not the width.
    assert picture_bullet_units(22, (16, 32)) == picture_bullet_units(22, (32, 32))


def test_the_picture_bullets_picture_is_read_from_the_numbering_part():
    import make_picture_bullet_probe as probe

    from docx2svg import parse_package

    document = parse_package(probe.build("none"))
    level = document.numbering[1][0]
    assert level.picture_bullet == 0
    bullet = document.picture_bullets[0]
    assert (bullet.part, bullet.pixels) == ("word/numbering.xml", (16, 16))
    drawing = next(k for k, case in enumerate(probe.CASES) if case.kind == "drawing")
    assert document.picture_bullets[drawing].relationship == "rIdImg1"

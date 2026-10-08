"""Where a superscript and a subscript are drawn: ``resolve.script_raise_half_points``
against Word, offline (ROADMAP.md, "The script offset -- measured").

``tests/fixtures/script-offset-observations.json`` holds what Word drew for every line of
``tools/make_script_offset_probe.py`` -- faces of our own, each moving one field -- and
the numbers of the real faces of ``make_script_probe.py``'s size sweep, whose drawn
offsets ``script-observations.json`` already holds.  The rule's scores are pinned, and
so are the refuted alternative's.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import read_script_offset_probe as reader  # noqa: E402

from docx2svg.resolve import script_raise_half_points  # noqa: E402
from docx2svg.vertical import FaceMetrics  # noqa: E402

DATA = json.loads(reader.OBSERVATIONS.read_text(encoding="utf-8"))
LINES = [row for row in DATA["lines"] if row[0] == "size"]


def test_every_offset_is_whole_half_points():
    import math
    from fractions import Fraction

    possible = {math.floor(Fraction(k * 25, 12) + Fraction(1, 2)) for k in range(200)}
    assert all(row[3] in possible and row[4] in possible for row in DATA["lines"])


def test_the_probe():
    score = reader.score_probe(LINES)
    # The drawn size, the fallback included (a face whose offsets Word refuses is drawn at
    # 3/5 however valid its size): every line.
    assert score["size"] == [690, 690]
    assert score["superscript"] == [545, 661, 690]
    assert score["subscript"] == [445, 670, 690]
    # Refuted: the face's own superscript offset alone.
    assert score["offset alone"][0] == 309


def test_the_real_faces():
    score = reader.score_real(DATA["real_faces"])
    assert score["superscript"] == [786, 1068, 1140]
    assert score["subscript"] == [965, 1135, 1140]


def test_a_script_is_not_bounded_by_its_line():
    """Beside a 36 pt run, and at 24 pt among 11 pt text, a script moves as it does alone
    at its own size."""
    alone = {(row[1], row[2]): row for row in DATA["lines"] if row[0] == "size"}
    for kind, face, _hp, sup, sub, *_ in DATA["lines"]:
        if kind == "large":
            assert (sup, sub) == (alone[(face, 48)][3], alone[(face, 48)][4])
        elif kind == "tall" and face == "Dx Sup 600":
            # The face's own offset, 22 x 600 / 2048 = 6.4 half points: 13 px, as alone.
            assert (sup, sub) == (13, 4)


def test_the_bound_is_what_holds_calibri_down():
    calibri = FaceMetrics(*DATA["real_faces"]["Calibri"])
    # Its offset (0.477 em) would raise a 48 half-point superscript by 23 half points; the
    # ascents of 48 and of the drawn 31 differ by 16, which is what Word draws.
    assert script_raise_half_points(48, "superscript", calibri) == 16

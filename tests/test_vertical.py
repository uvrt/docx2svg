"""The vertical model against every baseline Word drew in the Phase 2 probes.

These run without Word: ``tests/fixtures/line-advance-observations.json`` is the probes'
measured baselines (integers, device px), recorded by
``tools/read_line_advance_probe.py --record``.  The assertion is equality, not a
tolerance -- a baseline is an integer and the model produces integers.
"""

from __future__ import annotations

import json
from fractions import Fraction
from pathlib import Path

import pytest

from docx2svg.vertical import (
    LAYOUT_UNIT_PT,
    PX_PER_PT,
    FaceMetrics,
    baseline_px,
    line_pitch_px,
    natural_height_px,
    quantise,
    round_half_up,
    twips_to_px,
)

OBSERVATIONS = Path(__file__).parent / "fixtures" / "line-advance-observations.json"


def _load():
    data = json.loads(OBSERVATIONS.read_text(encoding="utf-8"))
    faces = {name: FaceMetrics(**fields) for name, fields in data["faces"].items()}
    groups = []
    for g in data["groups"]:
        ys = [g["first"]]
        for digit in g["extra"]:
            ys.append(ys[-1] + g["gap"] + int(digit))
        groups.append((g, ys))
    return faces, groups


FACES, GROUPS = _load()
TOP = twips_to_px(1440)


def _predict(g, count):
    face = FACES[g["face"]]
    pitch = line_pitch_px(face, g["half_points"], g["rule"], g["line"])
    return [
        baseline_px(TOP + k * pitch, face, g["half_points"], g["rule"], g["line"])
        for k in range(count)
    ]


@pytest.mark.parametrize(
    "g, ys",
    GROUPS,
    ids=lambda v: f"{v['face']}-{v['half_points']}-{v['rule']}{v['line']}" if isinstance(v, dict) else "",
)
def test_every_baseline_to_the_pixel(g, ys):
    assert _predict(g, len(ys)) == ys


def test_the_probes_cover_what_phase_2_asked_for():
    """Done-condition: >= 5 sizes and all three line rules, to 0 px -- now including the
    135 ``auto`` multiple groups Phase 2 left at 115 (ROADMAP.md, "The line box")."""
    rules = {g["rule"] for g, _ in GROUPS}
    sizes = {g["half_points"] for g, _ in GROUPS}
    assert rules == {"auto", "exact", "atLeast"}
    assert len(sizes) >= 5
    assert sum(len(ys) for _, ys in GROUPS) == 23859
    multiples = [g for g, _ in GROUPS if g["rule"] == "auto" and g["line"] > 240]
    assert len(multiples) == 135


def test_an_auto_multiple_is_a_line_with_space_below_its_text():
    """The extra height of ``auto`` > 240 goes below the text, so the line rounds in three
    parts (``baseline_in_box``) and never against its bottom edge: every one of the 135
    groups keeps one offset from ``round(top)`` whatever the top's fractional pixel."""
    for g, ys in GROUPS:
        if g["rule"] == "auto" and g["line"] > 240:
            face = FACES[g["face"]]
            pitch = line_pitch_px(face, g["half_points"], g["rule"], g["line"])
            offsets = {y - round_half_up(TOP + k * pitch) for k, y in enumerate(ys)}
            assert len(offsets) == 1, g


def test_the_layout_unit_is_a_4096th_of_a_point():
    unit_px = LAYOUT_UNIT_PT * PX_PER_PT
    # 7 twips is 1433.6 units; Word tracks by 1434 of them (measured: 1.45874 px/glyph).
    assert twips_to_px(7) == 1434 * unit_px
    assert float(twips_to_px(7)) == pytest.approx(1.458740, abs=1e-6)
    # A whole twip count that is a multiple of 5 is exact.
    assert twips_to_px(1440) == 300
    assert quantise(Fraction(1, 3)) == round(Fraction(1, 3) / unit_px) * unit_px


def test_single_spacing_is_the_hhea_sum_and_is_exact_for_half_point_sizes():
    calibri = FACES["Calibri"]
    # 11 pt: 2500/2048 em = 55.948893 px, the Phase 0 ratio.
    assert natural_height_px(calibri, 22) == Fraction(2500 * 11, 2048) * PX_PER_PT
    # Arial and Times New Roman share 2355 units, which is why their gaps are identical.
    assert natural_height_px(FACES["Arial"], 22) == natural_height_px(FACES["Times New Roman"], 22)


def test_the_refuted_single_rounding_model_really_is_refuted():
    """``round(top + ascent + k*h)`` -- the model Phase 0 left at 61 of 70."""
    calibri = FACES["Calibri"]
    g, ys = next((g, ys) for g, ys in GROUPS if g["face"] == "Calibri" and g["half_points"] == 22
                 and g["rule"] == "auto" and g["line"] == 240)
    h = natural_height_px(calibri, 22)
    ascent = Fraction(1950 * 11, 2048) * PX_PER_PT
    naive = [int((TOP + ascent + k * h + Fraction(1, 2)) // 1) for k in range(len(ys))]
    assert naive != ys
    assert _predict(g, len(ys)) == ys

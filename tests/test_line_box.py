"""The line box against every baseline of the line-box probe, without Word.

``tests/fixtures/line-box-observations.json`` holds the baselines Word drew for every
probe in ``tools/make_line_box_probe.py`` (integers, device px, one list per page) and
the metric integers of each face, recorded by ``tools/read_line_box_probe.py --record``.
The paragraphs are regenerated from the generator; the page split is checked so that a
change to the generator cannot silently re-pair baselines with other paragraphs.

The assertion is equality: 12,069 baselines, space before / after / both, bottom borders,
``exact`` and ``auto`` multiples in steps of one unit, and the other line rules combined
with spacing.  ROADMAP.md ("The line box") has the derivation and the refutations.
"""

from __future__ import annotations

import json
from fractions import Fraction
from pathlib import Path

import pytest

import make_line_box_probe as probe
import read_line_box_probe as reader

from docx2svg.vertical import (
    FaceMetrics, LineBox, baseline_in_box, border_px, paragraph_gap_px, round_half_up, twips_to_px,
)

DATA = json.loads((Path(__file__).parent / "fixtures" / "line-box-observations.json").read_text(encoding="utf-8"))
FACES = {name: FaceMetrics(*values) for name, values in DATA["faces"].items()}
PROBES = {p.name: p for p in probe.PROBES}


@pytest.mark.parametrize("name", sorted(DATA["probes"]))
def test_every_baseline_of_the_line_box_probe(name):
    p = PROBES[name]
    observed = DATA["probes"][name]
    assert [len(page) for page in observed] == [len(page) for page in probe.pages(p)]
    assert reader.predict(p, FACES[p.face]) == observed


def test_the_probe_is_what_the_roadmap_counts():
    assert set(DATA["probes"]) == set(PROBES)
    assert sum(len(page) for pages in DATA["probes"].values() for page in pages) == 12069


def _score(**kwargs):
    hits = total = 0
    for name, observed in DATA["probes"].items():
        p = PROBES[name]
        predicted = reader.predict(p, FACES[p.face], **kwargs)
        hits += sum(a == b for pa, pb in zip(predicted, observed) for a, b in zip(pa, pb))
        total += sum(len(page) for page in observed)
    return hits, total


def test_rounding_each_line_in_its_own_pitch_is_refuted():
    """Phase 2's arrangement -- spacing and borders stacked between lines but outside the
    box the baseline rounds in -- misses about one line in six."""
    hits, total = _score(line_box=False)
    assert total == 12069
    assert hits < 10400


def test_border_width_is_truncated_to_whole_twips():
    # 27/8 pt is 67.5 twips; Word takes 67.  Even eighths are whole twips already.
    assert border_px(27, 0) == twips_to_px(67)
    assert border_px(8, 4) == twips_to_px(20) + twips_to_px(80)
    assert border_px(27, 0) != twips_to_px(Fraction(135, 2))


def test_paragraph_spacing_collapses_in_twips():
    gap, box = paragraph_gap_px(7, 8)
    # The previous paragraph keeps its 7 twips; this box gets the 1 twip left over,
    # converted as 1 twip (205 units), not as the difference of 8 and 7 converted (204).
    assert box == twips_to_px(1)
    assert gap == twips_to_px(7) + twips_to_px(1)
    assert paragraph_gap_px(10, 4) == (twips_to_px(10), 0)


def test_space_below_the_text_anchors_the_line_at_its_top():
    """With anything below the text the offset from ``round(top)`` does not depend on the
    top's fractional pixel; without it, a pitch with f >= 1/2 rounds against its bottom."""
    em = Fraction(11) * Fraction(300, 72) / 2048
    plain = LineBox(pitch=Fraction(2500) * em, text_above=1950 * em, text_below=550 * em,
                    natural=Fraction(2500) * em)
    spaced = LineBox(**{**plain.__dict__, "space_after": twips_to_px(200)})
    tops = [Fraction(300) + Fraction(k, 10) for k in range(10)]
    assert len({baseline_in_box(t, spaced) - round_half_up(t) for t in tops}) == 1
    assert len({baseline_in_box(t, plain) - round_half_up(t) for t in tops}) == 2


MIXED = json.loads((Path(__file__).parent / "fixtures" / "mixed-line-observations.json").read_text(encoding="utf-8"))
MIXED_FACES = {name: FaceMetrics(*values) for name, values in MIXED["faces"].items()}


def test_every_baseline_of_the_mixed_line_probe():
    """440 baselines of lines mixing faces or sizes, under every line rule: the line box
    on ``baselines.pitch``, whose ``auto`` extra is over the tallest text item alone."""
    import make_mixed_line_probe as mixed
    import read_mixed_line_probe as mixed_reader

    assert sorted(g.name for g in mixed.GROUPS) == sorted(MIXED["groups"])
    for group in mixed.GROUPS:
        ys = MIXED["groups"][group.name]
        assert mixed_reader.predict(group, MIXED_FACES, len(ys)) == ys, group.name


def test_the_multiple_extra_over_the_combined_extent_is_refuted():
    """An inline Symbol run under ``auto`` 276 / 360: scaling the combined extent (Cambria's
    descent under SymbolMT's ascent) instead of the tallest item leaves 19 of 44 off."""
    from dataclasses import replace

    import make_mixed_line_probe as mixed
    import read_mixed_line_probe as mixed_reader

    misses = 0
    for group in mixed.GROUPS:
        if group.extra != "symbol" or group.rule != "auto" or group.line == 240:
            continue
        box, candidates = mixed_reader.models(group, MIXED_FACES)
        wrong = replace(box, pitch=candidates["combined"])
        ys = MIXED["groups"][group.name]
        top = twips_to_px(1440)
        misses += sum(baseline_in_box(top + k * wrong.pitch, wrong) != y for k, y in enumerate(ys))
    assert misses == 19

"""Consecutive bordered paragraphs, top borders and between borders, without Word.

``tests/fixtures/border-group-observations.json`` holds the lines Word drew for every
document of ``tools/make_border_group_probe.py``; every baseline goes through
``baselines.predict`` (ROADMAP.md, "Phase 4 -- measured": borders that are one box).
"""

from __future__ import annotations

import read_border_group_probe as reader


def test_every_line_of_every_setting():
    assert reader.exact(reader.offline()) == (408, 408)


def test_each_border_counted_on_its_own_is_refuted(monkeypatch):
    """The model before: each paragraph's own bottom border, one line wide whatever its
    style, and no top or between border -- 280 of the 408 lines."""
    from fractions import Fraction

    from docx2svg import vertical

    def alone(pp, previous=None, following=None):
        bottom = pp.get("pBdr.bottom")
        drawn = bottom and bottom[0] not in (None, "nil", "none")
        return Fraction(0), vertical.border_px(bottom[1], bottom[2]) if drawn else Fraction(0)

    monkeypatch.setattr(vertical, "paragraph_borders_px", alone)
    assert reader.exact(reader.offline()) == (280, 408)

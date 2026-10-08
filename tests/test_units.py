"""Units, and the two conversions that are easy to get wrong."""

from __future__ import annotations

import pytest

from docx2svg.units import (
    device_page_extent_pt,
    emu_to_pt,
    half_points_to_pt,
    pt_to_twips,
    twips_to_pt,
)


def test_twips_round_trip():
    for twips in (0, 1, 360, 1440, 11906):
        assert pt_to_twips(twips_to_pt(twips)) == pytest.approx(twips)


def test_a_twip_is_a_twentieth_of_a_point():
    assert twips_to_pt(1440) == 72.0
    assert twips_to_pt(360) == 18.0


def test_a_font_size_is_in_half_points():
    """``w:sz w:val="22"`` is 11 pt.

    The one quantity in the format measured this way, and the one place a factor of two
    can hide for a long time -- a document rendered at twice its size still looks like a
    document.
    """
    assert half_points_to_pt(22) == 11.0
    assert half_points_to_pt(28) == 14.0


def test_emu_is_the_drawingml_unit():
    assert emu_to_pt(914400) == 72.0


@pytest.mark.parametrize(
    "twips, expected",
    [
        # A4, measured out of Word's own export.  Authored 595.3 / 841.9.
        (11906, 595.2),
        (16838, 841.92),
        # The fixture's second section, which exists to falsify the rule: 10000 twips is
        # exactly 500 pt, so a page box that were the authored size would read 500.
        (10000, 499.92),
        (13000, 649.92),
        # US Letter is a whole number of device pixels either way and is unchanged --
        # which is why a corpus of Letter documents would never have revealed this.
        (12240, 612.0),
        (15840, 792.0),
    ],
)
def test_the_page_box_is_snapped_to_whole_device_pixels(twips, expected):
    assert device_page_extent_pt(twips) == pytest.approx(expected, abs=1e-9)

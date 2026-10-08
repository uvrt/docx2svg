"""Unit conversions.

WordprocessingML measures almost everything in **twips** -- a twentieth of a point, so
1440 to the inch -- and that is the unit this project keeps internally wherever an
authored value is being carried around unchanged.  It is exact in integers, which the
point is not: a 360-twip indent is 18 pt exactly, but a 1-twip one is 0.05 pt and a
sequence of them accumulates float error that a sequence of integers does not.

DrawingML, which arrives inside ``w:drawing``, uses **EMU** instead -- 914400 to the
inch, 635 to the twip.  Both appear in the same document, and confusing them is a factor
of 635, which is large enough to be obvious and therefore not the dangerous mistake.  The
dangerous one is twips against **half-points**: ``w:sz w:val="22"`` is 11 pt, not 22, and
not 1.1.  Font sizes are the only quantity in the format measured that way, and every
site that reads one goes through :func:`half_points_to_pt` so the factor appears once.
"""

from __future__ import annotations

TWIPS_PER_POINT = 20
TWIPS_PER_INCH = 1440
POINTS_PER_INCH = 72
EMU_PER_POINT = 12700
EMU_PER_TWIP = 635

#: Word's PDF export snaps the page box to whole device pixels at this resolution.
#: Measured, not assumed -- see ROADMAP.md, "The page box is not the page size".
DEVICE_DPI = 300


def twips_to_pt(twips: float) -> float:
    """Twentieths of a point to points."""
    return twips / TWIPS_PER_POINT


def pt_to_twips(points: float) -> float:
    return points * TWIPS_PER_POINT


def half_points_to_pt(half_points: float) -> float:
    """``w:sz``/``w:szCs`` to points.  The one place this factor is written down."""
    return half_points / 2


def emu_to_pt(emu: float) -> float:
    return emu / EMU_PER_POINT


def emu_to_twips(emu: float) -> float:
    return emu / EMU_PER_TWIP


def device_page_extent_pt(twips: float, dpi: int = DEVICE_DPI) -> float:
    """The page extent Word's PDF export *actually* writes, for an authored ``w:pgSz``.

    Word does not put the authored size in the ``/MediaBox``.  It rounds it to a whole
    device pixel at :data:`DEVICE_DPI` first, and the residual is visible at the first
    decimal place -- A4's 11906 twips is 595.3 pt authored and **595.2 pt** exported.
    Any comparison against the oracle that skips this step is wrong by up to 0.12 pt on
    every page, in a direction that depends on the page size, which is exactly the kind
    of error that gets attributed to something else.

    Established on two page sizes with opposite rounding directions; the derivation and
    its falsification test are in ROADMAP.md.
    """
    inches = twips / TWIPS_PER_INCH
    return round(inches * dpi) / dpi * POINTS_PER_INCH

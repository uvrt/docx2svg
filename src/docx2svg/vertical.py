"""Where Word puts a baseline: the line pitch, the line box, and the device-pixel rounding.

Every number here was measured out of Word 16.106's PDF export by the probes in
``tools/`` and is written up, with its observations and residuals, in ROADMAP.md
("Phase 2", and "The line box").  Nothing in this module reads a font file: a face
enters as four integers from its ``hhea`` table (or ``OS/2`` typo metrics when
``USE_TYPO_METRICS`` is set), which are facts about the face and not the face.

The model
---------

Word lays out in **1/4096 pt** (:data:`LAYOUT_UNIT_PT`) and only rounds to the device
grid -- 1/300 inch, :data:`~docx2svg.units.DEVICE_DPI` -- when it draws.  A line's
*pitch* ``h`` is exact in that unit and lines accumulate it exactly, together with the
paragraph spacing and borders between them (:func:`paragraph_gap_px`, :func:`border_px`):
no rounding is carried from one line to the next.  What rounds is the **baseline inside
the line**, and it rounds inside a **line box** that is more than the text:

* the space before a paragraph -- what is left of it after the previous paragraph's
  space after, :func:`paragraph_gap_px` -- sits in the box of the paragraph's first
  line, above the text;
* the space after, a bottom border, and the extra height of an ``auto`` multiple above
  one sit in the box *below* the text (the space after and the border in the last
  line's box only).

A box with **nothing below its text** (``E = 0``) rounds as Phase 2 measured, against one
of its two edges chosen by ``f = frac(H)``, the fractional pixel of the box height:
bottom-anchored ``round(top + H) - round(below)`` when ``f >= 1/2``; otherwise
top-anchored, ``round(top) + round(above - f/2)`` for a plain line, ``round(top) +
floor(H) - round(below)`` for a line with space above its text (``atLeast`` taller than
the text, or space before), and ``round(top) + round(above - f/4)`` for ``exact``.

A box with **space below its text** (``E > 0``) is always top-anchored, and it rounds
in three parts: the space above the text to whole pixels, the space below to whole
pixels, and the text in what is left, bottom-first but never higher than its own
ascent below the top -- and never lower than its own descent above the box's real
bottom edge.  :func:`baseline_in_box` is the whole rule; ROADMAP.md has its derivation.

``exact`` lines put the baseline at 4/5 of the pitch whatever the face, with the part
below it (``h/5``) truncated to the layout unit.

A line's extent is the largest ascent and the largest descent over what is on it
(:func:`line_extent`) -- its characters but spaces, or its paragraph mark alone when it
holds nothing else (:func:`line_items`) -- each character's from :func:`item_extent`: its
face at its
``w:sz`` -- a superscript or subscript too, unreduced and on the baseline -- moved by
``w:position`` and grown by a run border (ROADMAP.md, "Superscripts, subscripts,
position and run borders").

What this reproduces, to 0 device pixels: every baseline of every probe -- 19,389 bare
lines in Phase 2 and the 135 ``auto`` multiple groups it left open, 12,000 lines of the
line-box probe (space after, space before, both, borders, ``exact`` and ``auto``
multiples swept in steps of one unit), the bordered-line probe -- and every in-scope
line of ``layout-sweep.docx``, ``style-document.docx`` and the three samplelib documents.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from fractions import Fraction

from .units import DEVICE_DPI, POINTS_PER_INCH, TWIPS_PER_POINT

#: Word's internal layout unit.  Measured three ways, all agreeing (ROADMAP.md, Phase 2):
#: off-grid indents land on multiples of it to 5e-5 px; tracking of 7 twips (1433.6
#: units) comes back as exactly 1434 units per glyph; and ``auto`` multiples reproduce
#: only when the pitch is rounded to it (134 of 135 groups consistent, against 123 for
#: the unrounded pitch and at most 127 for any other unit tried).
LAYOUT_UNIT_PT = Fraction(1, 4096)

PX_PER_PT = Fraction(DEVICE_DPI, POINTS_PER_INCH)

#: The layout unit in device px.
LAYOUT_UNIT_PX = LAYOUT_UNIT_PT * PX_PER_PT

#: ``exact`` lines put the baseline this fraction of the pitch above the line's bottom,
#: independently of the face and size: measured on 11 pitches x 5 sizes x 3 faces, then
#: on every pitch from 240 to 480 twips in four faces (``make_line_box_probe.py``).
EXACT_BELOW_FRACTION = Fraction(1, 5)

#: How far the rounding slack ``f = frac(H)`` moves a top-anchored baseline up, per kind
#: of line with nothing below its text: ``round(above - LAMBDA * f)``.  Plain lines split
#: it evenly (feasible [0.386, 0.592) over Phase 2's 81 groups).  ``exact`` lines move by
#: less: anything in (0, 1/2) reproduces all 7,672 ``exact`` baselines and nothing
#: outside it does (0 fails at ties, 1/2 fails 36), so 1/4 is the middle of an interval,
#: not a measured value.  Lines with space above their text use ``floor(H) - round(below)``,
#: which is ``LAMBDA = 1`` up to ties.
LAMBDA_PLAIN = Fraction(1, 2)
LAMBDA_EXACT = Fraction(1, 4)

_HALF = Fraction(1, 2)


@dataclass(frozen=True)
class FaceMetrics:
    """The four integers the vertical model needs from a face, in font units.

    ``descent`` is positive (the magnitude of ``hhea.descender``).  Take them from
    ``hhea`` unless ``OS/2.fsSelection`` bit 7 (``USE_TYPO_METRICS``) is set, in which case
    take ``sTypoAscender``/``-sTypoDescender``/``sTypoLineGap``.

    **Which copy of a face** matters and is not obvious: for Times New Roman, Word 16.106
    embeds the copy in its own bundle (v7.00, ``hhea.lineGap`` 0) but lays out with the
    macOS system copy's metrics (v5.01, ``lineGap`` 87) -- only 87 reproduces the pitch,
    and only the system copy's kern pairs reproduce kerned Greek.  See ROADMAP.md.
    """

    units_per_em: int
    ascent: int
    descent: int
    line_gap: int = 0
    #: ``OS/2.ySuperscriptYSize`` and ``ySubscriptYSize``: what a ``w:vertAlign`` run is
    #: drawn at, in font units of an em (:func:`docx2svg.resolve.script_half_points`).
    #: They do not enter the line height.  ``None`` when not recorded.
    superscript_size: int | None = None
    subscript_size: int | None = None
    #: ``OS/2.ySuperscriptYOffset`` and ``ySubscriptYOffset`` (the latter positive down):
    #: whether Word takes the face's script fields at all, and how far it moves a script
    #: (:func:`docx2svg.resolve.script_raise_half_points`).  ``None`` when not recorded.
    superscript_offset: int | None = None
    subscript_offset: int | None = None


def quantise(px: Fraction) -> Fraction:
    """Round a length in device px to Word's layout unit (half up)."""
    return math.floor(px / LAYOUT_UNIT_PX + _HALF) * LAYOUT_UNIT_PX


def quantise_down(px: Fraction) -> Fraction:
    """Truncate a length in device px to Word's layout unit."""
    return math.floor(px / LAYOUT_UNIT_PX) * LAYOUT_UNIT_PX


def twips_to_px(twips: int | Fraction) -> Fraction:
    """An authored twip length, as Word holds it, in device px."""
    return quantise(Fraction(twips) / TWIPS_PER_POINT * PX_PER_PT)


def round_half_up(x: Fraction) -> int:
    return math.floor(x + _HALF)


def _em_px(face: FaceMetrics, half_points: int) -> Fraction:
    return Fraction(half_points, 2) * PX_PER_PT / face.units_per_em


def natural_height_px(face: FaceMetrics, half_points: int) -> Fraction:
    """Single-spaced pitch: ``(ascent + descent + lineGap) / upm x size``, exact."""
    return quantise((face.ascent + face.descent + face.line_gap) * _em_px(face, half_points))


def line_pitch_px(face: FaceMetrics, half_points: int, rule: str = "auto", line: int = 240) -> Fraction:
    """The distance from one line's top to the next, for a ``w:spacing`` rule.

    ``auto``: ``line`` is in 240ths of the natural height.  ``exact``/``atLeast``: ``line``
    is in twips; ``atLeast`` never goes below the natural height.
    """
    natural = natural_height_px(face, half_points)
    if rule == "auto":
        return quantise(natural * line / 240)
    if rule == "exact":
        return twips_to_px(line)
    if rule == "atLeast":
        return max(natural, twips_to_px(line))
    raise ValueError(f"unknown line rule {rule!r}")


# -- paragraph spacing and borders ----------------------------------------------------


def paragraph_gap_px(previous_after: int, before: int) -> tuple[Fraction, Fraction]:
    """The space between two paragraphs, and the part of it in the second one's box.

    Both arguments in twips, after contextual spacing has zeroed whichever applies.  The
    two collapse to the larger (ROADMAP.md, "Style inheritance"), and the collapse is
    taken **in twips**: the previous paragraph keeps its whole space after (in its last
    line's box, below the text) and this paragraph's box gets ``before - after`` twips
    above its text when that is positive.  Returns ``(gap, space above the first line's
    text)`` in px.

    Measured by ``make_line_box_probe.py`` (``combo-both``: before = after = n on every
    paragraph, so each box gets 1 twip): taking ``max`` of the two lengths already in px
    instead leaves the page one layout unit off per paragraph, which moved two baselines
    at a rounding tie; giving the whole space to the second paragraph's box instead
    fails ``style-document.docx`` and every heading of the samplelib documents.
    """
    box_before = twips_to_px(before - previous_after) if before > previous_after else Fraction(0)
    return twips_to_px(previous_after) + box_before, box_before


def keeps_space_before_at_page_top(*, section_start: bool, page_break_before: bool = False,
                                   compatibility_mode: int | None = None) -> bool:
    """Whether a paragraph that starts a page keeps its space before there.

    * **The first paragraph of a section keeps it** -- the document's first paragraph, and
      the first of every ``nextPage`` section -- in every compatibility setting.
    * **A ``w:pageBreakBefore`` paragraph keeps it below compatibility mode 15**, and
      when no mode is stated (no ``settings.xml``, no ``w:compat``, an empty one); in
      mode 15 it loses it.  ``w:suppressSpBfAfterPgBrk`` changes neither.
    * Any other paragraph that starts a page loses it: after a manual page break (in
      every setting), after a natural break.

    Measured by ``make_page_top_probe.py`` (space before swept 1..120 twips, seven
    compatibility settings), every line exact; ROADMAP.md, "The page top -- measured".
    Under the earlier rule, which dropped it at every page top, every section start and
    every ``wordto`` document's first page was low by exactly its space before, and so
    were ``filesamples/sample1``'s ``pageBreakBefore`` headings (mode 12).
    """
    if section_start:
        return True
    return page_break_before and (compatibility_mode is None or compatibility_mode < 15)


def page_top_gap_px(before: int, *, keeps: bool, previous_after: int = 0) -> tuple[Fraction, Fraction]:
    """The space before of a paragraph that starts a page: ``(how far it moves the text
    down from the top margin, the space above the text in the first line's box)``.

    Kept, it is what the space between two paragraphs puts in the second one's box
    (:func:`paragraph_gap_px`), collapsed with ``previous_after`` -- the space after of
    the paragraph before the break, on the page before: after 200 twips, 240 before
    keeps 40.  Measured by ``make_page_top_collapse_probe.py`` (``pageBreakBefore``
    below mode 15 and section starts in every mode, after 0-360 twips, before 1-480):
    724 of 724 lines, against 308 keeping it whole.  **Dropped, it moves nothing but is still in the box**:
    the box hangs above the top margin by the space before, and the baseline rounds in
    it (:func:`baseline_in_box` on ``top - space``).  So a dropped space still decides
    which edge the line rounds against.  Measured by ``make_page_top_probe.py``: with 12,
    17, 36, 41, 65, 89 and 113 twips dropped (after a manual break, a natural break, or
    ``pageBreakBefore`` in mode 15) Calibri 11 pt is drawn at 343, not at 344 like every
    other page top -- 22 of 22 such lines exact against 0, and nothing else moves.
    """
    if keeps and previous_after:
        # Kept, it still collapses with the space after of the paragraph before the
        # break, at the foot of the page before: what is left of it is kept, and is what
        # the box holds (holding the whole space instead is not discriminated: 724 of
        # 724 lines either way).
        _, box = paragraph_gap_px(previous_after, before)
        return box, box
    kept, box = paragraph_gap_px(0, before)
    return (kept if keeps else Fraction(0)), box


def autospace_kept(*, numbered: bool, neighbour_numbered: bool | None) -> bool:
    """Whether an autospaced space before or after (14 pt, :mod:`docx2svg.resolve`) is
    there at all, given the paragraph on that side.

    It is not **between two list items** -- of the same list or of two lists -- and not
    **before the document's first paragraph** (``neighbour_numbered=None``: nothing on
    that side).  Everywhere else it is, including before a section's first paragraph,
    a ``pageBreakBefore`` paragraph that keeps its space (below mode 15), and on either
    side of a list next to a paragraph that is not a list item.  Measured by
    ``make_autospacing_probe.py`` (ROADMAP.md, "Autospacing -- measured"): three
    autospaced items are drawn with no space between them, and the first item's space
    after is out of its box too (it moves the item by 1 px).
    """
    if neighbour_numbered is None:
        return False
    return not (numbered and neighbour_numbered)


def border_px(width_eighths: int, space_points: int) -> Fraction:
    """The height a ``w:pBdr/w:bottom`` adds below its paragraph: space plus width.

    ``w:sz`` (eighths of a point) is converted to **whole twips, truncated**, before the
    layout unit: a 27/8 pt border takes 67 twips, not 67.5.  Measured on 128 bordered
    paragraphs per face and size (``make_line_box_probe.py``, ``border``): keeping the
    half twip leaves the lines after an odd width 0.1 px high and misses 190 of 1,618
    baselines; truncating reproduces all of them.
    """
    return twips_to_px(space_points * 20) + twips_to_px(width_eighths * 20 // 8)


#: Border styles drawn as more than one line, and how many line widths tall they are
#: (``make_border_group_probe.py``: ``double`` at ``w:sz`` 4 and 12 takes three widths).
LINES_TALL = {"double": 3}
#: The sides of ``w:pBdr``.
BORDER_SIDES = ("top", "left", "bottom", "right", "between", "bar")


def _drawn(border) -> bool:
    return bool(border) and border[0] not in (None, "nil", "none")


def side_px(border, *, between: bool = False) -> Fraction:
    """The height one ``w:pBdr`` side takes: its space and its width (:func:`border_px`),
    a ``double`` line three widths; a ``w:between`` border its space on both sides."""
    if not _drawn(border):
        return Fraction(0)
    width = border[1] * LINES_TALL.get(border[0], 1)
    height = border_px(width, border[2])
    return height + twips_to_px(border[2] * 20) if between else height


def border_box(pp) -> tuple:
    """What decides whether two consecutive paragraphs share one border box: every side
    as stated (style, width, space, colour) and the indents."""
    return (tuple(pp.get(f"pBdr.{side}") for side in BORDER_SIDES),
            pp.get("ind.left", 0) or 0, pp.get("ind.right", 0) or 0)


def paragraph_borders_px(pp, previous=None, following=None) -> tuple[Fraction, Fraction]:
    """``(above, below)``: the room a paragraph's borders take above its first line and
    below its last, given the paragraphs next to it (resolved, ``None`` for none).

    **Consecutive paragraphs whose borders and indents are the same are one box**: the
    top border only over the first, the bottom border only under the last, and the
    ``w:between`` border (its space above and below the line) between each two.  A
    different width, space, colour, style or left indent makes two boxes; different
    paragraph spacing does not.  A top border takes its space and width above the
    text, as the bottom one does below.  Measured by ``make_border_group_probe.py``,
    in four settings that agree.
    """
    box = border_box(pp)
    if not any(_drawn(side) for side in box[0]):
        return Fraction(0), Fraction(0)
    after_same = following is not None and border_box(following) == box
    before_same = previous is not None and border_box(previous) == box
    between = pp.get("pBdr.between")
    # A between border's space above its line is in the upper paragraph's box, its line
    # and the space below in the lower one's: 408 of 408 lines; all of it in either box,
    # 404.  (The space below in the lower box and the rest in the upper also gives 408:
    # which side of the line its width is on is not discriminated.)
    upper = twips_to_px(between[2] * 20) if _drawn(between) else Fraction(0)
    above = (side_px(between, between=True) - upper) if before_same else side_px(pp.get("pBdr.top"))
    below = upper if after_same else side_px(pp.get("pBdr.bottom"))
    return above, below


def run_border_px(width_eighths: int, space_points: int) -> Fraction:
    """What a ``w:bdr`` adds to its run above the text, and again below: space plus width.

    The width is truncated to whole twips, as a paragraph border's is
    (:func:`border_px`), and then *truncated* to the layout unit, where a paragraph
    border's is rounded: ``w:sz`` 27 is 67 twips, 13,721.6 units, and only 13,721
    reproduces the three lines (of 432 with ``w:sz`` 27) that sit at a rounding tie
    (``make_script_probe.py``, ``bdr27x0``); rounding misses them, and keeping the half
    twip misses 171 of 5,184 border lines.
    """
    return twips_to_px(space_points * 20) + quantise_down(
        Fraction(width_eighths * 20 // 8, 20) * PX_PER_PT)


# -- the line box -----------------------------------------------------------------------


@dataclass(frozen=True)
class LineBox:
    """One line's box, in exact device px, from the top of its first part.

    ``pitch`` is the line rule's pitch ``h``; ``space_before`` the paragraph space in
    this box above the text (first line only); ``space_after`` what sits below the
    line's pitch -- the paragraph's space after and bottom border, last line only.
    ``text_above``/``text_below`` are the text's extent about the baseline (the largest
    ascent with lineGap, and the largest descent, over everything on the line), and
    ``natural`` the natural pitch of the text the line rule scales or compares with.
    """

    pitch: Fraction
    text_above: Fraction
    text_below: Fraction
    natural: Fraction
    rule: str = "auto"
    line: int = 240
    space_before: Fraction = Fraction(0)
    space_after: Fraction = Fraction(0)

    @property
    def multiple_extra(self) -> Fraction:
        """The part of an ``auto`` multiple above one that is not text: it goes below."""
        if self.rule == "auto" and self.line > 240:
            return self.pitch - self.text_above - self.text_below
        return Fraction(0)

    @property
    def height(self) -> Fraction:
        return self.space_before + self.pitch + self.space_after

    @property
    def below_text(self) -> Fraction:
        """``E``: everything in the box below the text line."""
        return self.multiple_extra + self.space_after


def exact_below_px(pitch: Fraction) -> Fraction:
    """How far below the baseline an ``exact`` line ends: ``h/5``, truncated to the unit.

    Truncated rather than rounded: at ``w:line`` 372, 396 and 420 (``h/5`` an odd
    multiple of a half pixel to within a unit) only truncation matches, and it changes
    nothing elsewhere (``make_line_box_probe.py``, ``exact``).
    """
    return quantise_down(pitch * EXACT_BELOW_FRACTION)


def _line_extent(box: LineBox) -> tuple[Fraction, Fraction, bool]:
    """(above, below) of the baseline within the line's pitch, and whether the line puts
    its slack above the text -- the Phase 2 model of a line on its own."""
    h, above, below, natural = box.pitch, box.text_above, box.text_below, box.natural
    if box.rule == "exact":
        below = exact_below_px(h)
        return h - below, below, True
    if box.rule == "atLeast" and h > natural:
        return h - below, below, True
    if box.rule == "auto" and box.line < 240:
        return above * h / natural, below * h / natural, False
    if box.rule == "auto" and box.line > 240:
        return above, below, False
    return above, h - above, False


def _text_line(box: LineBox) -> tuple[Fraction, Fraction, Fraction, Fraction]:
    """(space above the text inside the pitch, the text's own pitch, above, below)."""
    h = box.pitch
    if box.rule == "atLeast" and h > box.natural:
        return h - box.natural, box.natural, box.text_above, box.natural - box.text_above
    if box.rule == "auto" and box.line > 240:
        return Fraction(0), box.text_above + box.text_below, box.text_above, box.text_below
    above, below, _ = _line_extent(box)
    return Fraction(0), h, above, below


def baseline_in_box(top: Fraction, box: LineBox) -> int:
    """The device-pixel row of a line's baseline.  ``top`` is the exact top of the box.

    **Nothing below the text** (``E = 0``): Phase 2's rule, applied to the box.  With
    ``H`` the box height and ``f = frac(H)``, the baseline rounds against the bottom edge
    when ``f >= 1/2`` and against the top edge otherwise, by how the line puts its slack.

    **Space below the text** (``E > 0``): always against the top edge, in three parts.
    The space above the text (paragraph space before, and an ``atLeast`` line's extra)
    rounds to whole pixels, ``Xa``; so does the space below, ``E``; the text keeps what
    is left, ``P = t + (Xa - round(Xa)) + (E - round(E))`` for a text pitch ``t``, and sits
    on its bottom -- ``round(P) - round(below)`` -- but no higher than ``round(above)``
    under its top.  Separately the baseline is never lower than ``round(below)`` above
    the box's real bottom edge ``round(top + H)``; that only binds when ``E`` is a
    fraction of a pixel.
    """
    H = box.height
    E = box.below_text
    if E == 0:
        above, below, extra_above = _line_extent(box)
        above += box.space_before
        f = H - math.floor(H)
        if f >= _HALF:
            return round_half_up(top + H) - round_half_up(below)
        if box.rule == "exact" and box.space_before == 0:
            return round_half_up(top) + round_half_up(above - LAMBDA_EXACT * f)
        if extra_above or box.space_before:
            return round_half_up(top) + math.floor(H) - round_half_up(below)
        return round_half_up(top) + round_half_up(above - LAMBDA_PLAIN * f)

    inner, t, above, below = _text_line(box)
    xa = box.space_before + inner
    rounded_xa = round_half_up(xa)
    rest = t + (xa - rounded_xa) + (E - round_half_up(E))
    from_top = round_half_up(top) + rounded_xa + max(round_half_up(above),
                                                     round_half_up(rest) - round_half_up(below))
    return min(from_top, round_half_up(top + H) - round_half_up(below))


# -- a line of several faces and sizes, and its list label --------------------------------


def line_items(chars, mark) -> list:
    """What on one line takes part in its height: ``chars`` are ``(item, character)``
    for every character Word put on the line, ``mark`` the paragraph mark's item.
    ROADMAP.md, "The paragraph mark in the line height", has the probe and the scores.

    **A space takes no part**, whatever its size: a 30 pt space after 24 pt text, or
    between two words of it, leaves the line as tall as the text alone.  (A tab takes
    none either; the parser keeps it out of the run's text.)

    **The paragraph mark takes part only in a line that holds nothing but spaces** -- an
    empty paragraph, or one of spaces alone -- where it alone is the line: a 30 pt space
    under a 24 pt mark is a 24 pt line, a 24 pt space under a 30 pt mark a 30 pt one.  On
    a line that holds text it takes no part: not by its size, not by its face's ascent
    or descent, not on a paragraph's last line (where it sits) nor on any other.

    Measured by ``make_mark_probe.py``: marks 30 pt over 24 pt text in six faces, a
    same-size mark in a face reaching 0.27 em higher or 0.19 em lower than the text's, a
    smaller mark reaching higher or lower, three-line paragraphs, numbered paragraphs,
    spaces of 16, 26 and 30 pt alone, after and inside text, under ``auto`` 240 and 276,
    ``exact``, ``atLeast`` above and below the text, in four compatibility settings that
    draw alike.  Counting the mark on every line (the model before) misses every such
    case, and so does counting it on the last line only; counting spaces misses every
    line after a space larger than its text.
    """
    if any(char != " " for _, char in chars):
        return [item for item, char in chars if char != " "]
    return [mark]


def label_items(label, *, first_line: bool) -> list:
    """What of a list label takes part in a line's height: all of it on its paragraph's
    first line, where Word draws it, and nothing on the others.

    Measured by ``make_mark_probe.py`` (ROADMAP.md, "The paragraph mark in the line
    height"): three-line numbered paragraphs of Calibri and Courier New text whose label
    is 30 pt over 24 pt text -- from the level's ``w:sz`` or from the mark's -- under
    ``auto`` 240 and 276, ``exact``, ``atLeast`` above and below the text, in four
    compatibility settings: every line, against 248 of 600 with the label on every line
    (the model before), whose second and third lines were each a label's extra ascent low.
    """
    return list(label) if first_line else []


def line_extent(text_items, metrics, label_items=()) -> tuple[FaceMetrics, int, bool] | None:
    """One line's extent: the largest ascent (with lineGap, which Word puts above) and the
    largest descent over what is on it, maxed *separately*; **a list label takes part by
    its ascent only**.

    ``text_items`` and ``label_items`` are resolved character formats (anything with
    ``face``, ``bold``, ``italic`` and ``half_points``: what of the line's characters and
    its paragraph mark takes part, :func:`line_items`; the label's characters), ``metrics(face, bold, italic)`` a
    :class:`FaceMetrics` or ``None``.  Returns ``(face, half_points, mixed)``: for a line
    of one face and size that face; otherwise a synthetic :class:`FaceMetrics` at the
    size of the item with the largest ascent, with ``mixed`` true.

    Separate maxima rather than the tallest item: single-spaced Cambria 11 pt with a
    SymbolMT glyph (2059 above / 450 below against Cambria's 1946 / 455) has a pitch of
    56.24-56.33 px (``make_mixed_line_probe.py``): 56.26 for separate maxima, 56.15 for
    the tallest item.

    The label's descent: ``make_label_probe.py`` put labels in SymbolMT, Wingdings,
    Courier New and Arial Black, at 11 and 20 pt, before Tahoma, Arial, Cambria and
    Calibri text (descents 0.207-0.269 em), single and ``auto`` 276, in three
    compatibility settings.  Wherever the label's descent exceeds the text's, the line
    is as tall as the text's descent under the label's ascent (Courier New 20 pt before
    Tahoma: 78.857 px drawn, 78.84 so, 94.38 with the label's descent).  ROADMAP.md, "A
    list label's descent -- measured", has the scores.  The same glyphs as inline
    *runs* count with their descent, as text does.
    """
    extents = []
    for item, label in [(item, False) for item in text_items] + [(item, True) for item in label_items]:
        picture = getattr(item, "picture_px", None)
        if picture is not None:
            # A picture bullet (``lines.PictureLabel``): on the baseline, moved by its
            # ``w:position``, nothing below it (``make_picture_bullet_probe.py``).
            extents.append((picture + twips_to_px(item.position * 10), Fraction(0), item.half_points, None, True,
                            True))
            continue
        face = metrics(item.face, item.bold, item.italic) if item.face else None
        if face is None:
            continue
        above, below = item_extent(item, face)
        moved = (above, below) != _face_extent(face, item.half_points)
        extents.append((above, Fraction(0) if label else below, item.half_points, face, label, moved))
    if not extents:
        return None
    above = max(e[0] for e in extents)
    below = max(e[1] for e in extents)
    half_points = max(extents, key=lambda e: e[0])[2]
    distinct = {(e[3], e[2]) for e in extents}
    if len(distinct) == 1 and not all(e[4] for e in extents) and not any(e[5] for e in extents):
        # One face and size: a label in the text's own face and size adds nothing.
        face, half_points = next(iter(distinct))
        return face, half_points, False
    em = Fraction(half_points, 2) * PX_PER_PT
    return FaceMetrics(1, above / em, below / em, 0), half_points, True


def _face_extent(face: FaceMetrics, half_points: int) -> tuple[Fraction, Fraction]:
    em = _em_px(face, half_points)
    return (face.ascent + face.line_gap) * em, face.descent * em


def item_extent(item, face: FaceMetrics) -> tuple[Fraction, Fraction]:
    """How far one character reaches above and below the baseline, in px.

    * Its face at its ``w:sz`` -- **a ``w:vertAlign`` run too**: it takes part at the
      size it would have without it, and on the baseline, although it is drawn smaller
      (:func:`docx2svg.resolve.script_half_points`) and raised or lowered.  Measured by
      ``make_script_probe.py``: superscripts, subscripts and both, at the text's size,
      twice it and half it, six faces, four line rules, three compatibility settings --
      13,824 of 13,824 baselines; counting the run at its drawn size, 11,232.
    * Moved by **``w:position``**: raised by ``p`` half points, it reaches ``p`` higher
      and ``p`` less far down; lowered, the other way (``p`` in twips, then the layout
      unit, as any authored length).  The same probe: +-1..48 half points, six faces,
      four rules, and on runs half and twice the text's size, 13,824 of 13,824.  It is
      not in the natural height an ``auto`` multiple scales (:func:`tallest_natural`):
      counting it there, 11,130; a raised run keeping its descent, 13,404.
    * Grown by its **run border** (``w:bdr``), its space and width, above and below
      (:func:`run_border_px`).  The border is the run's: a bordered run smaller than the
      text grows the line only where its own extent and border pass the text's (a
      border around the whole line instead: 0 of the 288 lines per setting with a
      half-size bordered run).  The same probe: ``w:sz`` 2..96 and ``w:space`` 0..31,
      six faces, four rules, runs half and twice the text's size -- 15,552 of 15,552.
    """
    above, below = _face_extent(face, item.half_points)
    shift = twips_to_px(getattr(item, "position", 0) * 10)
    above, below = above + shift, below - shift
    border = getattr(item, "border", None)
    if border is not None:
        pad = run_border_px(*border)
        above, below = above + pad, below + pad
    return above, below


def tallest_natural(items, metrics) -> Fraction | None:
    """The largest natural height of any one item (:func:`item_natural_px`)."""
    heights = [item_natural_px(item, face) for item in items if item.face
               for face in [metrics(item.face, item.bold, item.italic)] if face is not None]
    return max(heights) if heights else None


def item_natural_px(item, face: FaceMetrics) -> Fraction:
    """One item's natural height, the one an ``auto`` multiple scales: its face's at its
    size (ascent + descent + lineGap), **and its run border twice** -- a bordered run's
    line under 1.15 is 1.15 times its bordered height (``make_script_probe.py``, 1,152 of
    1,152 bordered lines under ``auto`` 276; with the border outside the natural height,
    101).  ``w:position`` does not enter it (:func:`item_extent`)."""
    natural = natural_height_px(face, item.half_points)
    border = getattr(item, "border", None)
    return natural + 2 * run_border_px(*border) if border is not None else natural


def mixed_line_pitch(extent, text_extent, rule: str, line: int,
                     text_natural: Fraction | None = None) -> Fraction:
    """The pitch of a line of several items: under an ``auto`` multiple, the natural
    height of everything on the line (:func:`line_extent`, label included) plus the
    multiple's extra over the **tallest text item alone**; otherwise
    :func:`line_pitch_px` of the combined extent.

    ``text_extent`` is the text without the label (``None`` when there is none),
    ``text_natural`` the largest natural height of any single text item
    (:func:`tallest_natural`).  Measured by ``make_mixed_line_probe.py`` (22 lines per
    group): under ``auto`` 276 and 360 a Symbol bullet *label* adds its extra height
    once (Cambria 61.79 -> 64.33, 80.60 -> 83.13 px) -- the extra is over the text
    alone.  An inline Symbol *run* is text, and the extra is over its own natural height
    (SymbolMT 56.150 px at 11 pt), not over the line's combined extent (Cambria's descent
    under SymbolMT's ascent, 56.262): 64.685 and 84.337 px, which put every one of the 44
    baselines of both groups on the pixel.  The combined extent gave 64.701 and 84.393 --
    19 of the 44 baselines one pixel off.
    """
    face, half_points, _ = extent
    if rule != "auto" or line == 240 or (text_extent is None and text_natural is None):
        return line_pitch_px(face, half_points, rule, line)
    if text_natural is None:
        text_face, text_half_points, _ = text_extent
        text_natural = natural_height_px(text_face, text_half_points)
    everything = natural_height_px(face, half_points)
    return quantise(everything + text_natural * (Fraction(line, 240) - 1))


# -- a line on its own ------------------------------------------------------------------


def face_line_box(face: FaceMetrics, half_points: int, rule: str = "auto", line: int = 240, *,
                  space_before: Fraction = Fraction(0), space_after: Fraction = Fraction(0)) -> LineBox:
    """The box of a line of one face and size."""
    em = _em_px(face, half_points)
    return LineBox(
        pitch=line_pitch_px(face, half_points, rule, line),
        text_above=(face.ascent + face.line_gap) * em,
        text_below=face.descent * em,
        natural=natural_height_px(face, half_points),
        rule=rule,
        line=line,
        space_before=space_before,
        space_after=space_after,
    )


def baseline_px(top: Fraction, face: FaceMetrics, half_points: int, rule: str = "auto",
                line: int = 240) -> int:
    """The device-pixel row of a line's baseline, given the line's exact top in px.

    A line with no paragraph spacing or border in its box -- what every Phase 2 probe
    line was.  :func:`baseline_in_box` takes the general case.
    """
    return baseline_in_box(top, face_line_box(face, half_points, rule, line))


def baseline_offset(face: FaceMetrics, half_points: int, rule: str = "auto", line: int = 240) -> Fraction:
    """The exact (unrounded) baseline offset below the line top, for reference.

    The *drawn* baseline is :func:`baseline_px`, which is not ``round(top + this)``: that
    simpler model -- one rounding of the exact position -- is refuted (61 of 70 baselines
    of the Phase 0 fixture, 1 px off on the rest).
    """
    above, _, _ = _line_extent(face_line_box(face, half_points, rule, line))
    return above

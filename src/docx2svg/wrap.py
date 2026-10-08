"""Text beside a floating drawing: ``wrapSquare``, ``wrapTight`` and ``wrapThrough``.

Every rule here was read off Word's PDF by ``tools/make_wrap_side_probe.py`` and its
readers (ROADMAP.md, "Floating drawings -- measured", F.8), not taken from ECMA-376
20.4.2.17-20, which says what the elements mean but not which lines are beside a
drawing, where their text starts and stops, or what Word does with a narrow gap:

* **Which lines.**  A line is beside a drawing when it reaches into the drawing's box
  (its extent in whole twips and its effect extent) widened by ``distT`` above and
  ``distB`` below: from its top down its pitch -- **in mode 15 less an ``auto``
  multiple's extra below the text**, below mode 15 the whole pitch and, on a paragraph's
  last line, its space after; a line that only touches the box's edge is not beside it.
  The drawing is positioned first, from the page laid out without it, and stays there
  (:mod:`docx2svg.paginate`).
* **What it keeps text off.**  The box widened by ``distL`` and ``distR``.  For
  ``wrapTight`` and ``wrapThrough`` (alike) the wrap polygon instead -- each point in
  whole twips from the drawing's corner, truncated, scaled from 21,600 to the extent --
  across its least and greatest x within the line's *text* (its pitch less a multiple's
  extra, in every mode), widened by ``distL`` / ``distR``; its height is the polygon's,
  widened by ``distT`` / ``distB`` in mode 15 and not below it.  ``wrapText``:
  ``bothSides`` keeps text on both sides, ``left`` only left of it, ``right`` only right,
  ``largest`` the side with more room in the column, indents aside (the left on a tie).
* **Where the text goes.**  What is left of the column, in pieces (*segments*) from left
  to right; a segment narrower than **360 twips** takes no text, whatever would fit in it.
  Text fills the segments greedily, as a line (:func:`docx2svg.linebreak.break_line`),
  each up to its right edge exactly -- the drawing's left edge less ``distL`` measured from
  the left margin's exact edge, to the layout unit -- and a segment whose first word does
  not fit is passed over (no word is broken in one).  A line no segment takes text on
  goes down to the foot of the drawings it is beside, and is tried again there.
* **Where a segment's text starts.**  At the drawing's right edge plus ``distR``,
  measured from the left margin's exact edge (``w``), or at the paragraph's left indent
  where that is further right; the first line's indent moves it from there, a hanging one
  no further left than the edge.  Then the whole segment moves right by a rounding Word
  makes on the device-pixel grid: to the edge's pixel on the page, ``round(E)``, from
  ``C + w`` (``C`` the text column's start, a whole pixel) below mode 15, from
  ``round(C + w)`` in mode 15 -- so below mode 15 the text starts on a whole pixel, and in
  mode 15 on the edge where the margin is on the pixel grid and up to a pixel right of it
  where it is not.  An edge on exactly half a pixel goes either way, by a rule not found
  (:func:`segments` takes the page's half up and the column's half down in mode 15, the
  page's half down below it: the most measured cases).
* **Tabs** in a segment right of a drawing: stops from the column's left edge in mode 15,
  from the segment's start below it.
* **Alignment** is each segment's own: centred, right-aligned and justified text is set
  in its segment as a line is in the column.

Standard library only.
"""

from __future__ import annotations

import dataclasses
import math
from dataclasses import dataclass
from fractions import Fraction

from .linebreak import Geometry, Line, Piece, break_line
from .vertical import LAYOUT_UNIT_PX

#: The narrowest segment that takes text: 360 twips (``make_wrap_side_probe.py``, gaps of
#: 351 to 359 twips left and right of a drawing take none, 360 and wider take text, at 8,
#: 11 and 20 pt alike).
MIN_SEGMENT_TWIPS = 360
#: Device px per twip.
TWIP_PX = Fraction(300, 1440)
#: The layout unit (1/4096 pt) per device px.
UNITS_PER_PX = 1 / LAYOUT_UNIT_PX


@dataclass(frozen=True)
class Wrap:
    """A drawing text wraps beside, on one page: device px from the page's top left.

    ``top`` and ``bottom`` are its box widened by ``distT`` and ``distB``; ``left`` and
    ``right`` its box with its effect extent, and ``dist_left`` / ``dist_right`` what text
    keeps further off; ``polygon`` its wrap polygon in page px (``wrapTight``,
    ``wrapThrough``), which then decides both how far down it reaches and how far across
    at each line; ``side`` its ``wrapText``.  ``y`` is the drawing's top as positioned,
    ``key`` its anchor (as :class:`docx2svg.paginate.Band`'s)."""

    top: Fraction
    bottom: Fraction
    left: Fraction
    right: Fraction
    dist_left: Fraction
    dist_right: Fraction
    side: str
    y: Fraction
    key: tuple
    polygon: tuple | None = None
    dist_top: Fraction = Fraction(0)
    dist_bottom: Fraction = Fraction(0)
    #: The drawing's left edge as positioned.
    x: Fraction | None = None
    #: Whether the pixel rounding of the text right of it takes a half up where a
    #: drawing's takes it down: a floating table's (``make_float_table_probe.py``: below
    #: mode 15 edges at 642.5 and 1,112.5 px put the text at 643 and 1,113; in mode 15 an
    #: edge at 937.917 px, whose text is at 937.5 from the column's start, leaves it there).
    half_up: bool = False
    #: How far the text right of it is moved from its right edge, where that is not the
    #: drawings' pixel rounding: a drop cap's frame (:meth:`docx2svg.paginate.DropCap.
    #: text_start`); ``None`` for a drawing.
    text_shift: Fraction | None = None


def _half_up(x: Fraction) -> int:
    return math.floor(x + Fraction(1, 2))


def _half_down(x: Fraction) -> int:
    return math.ceil(x - Fraction(1, 2))


def polygon_span(points, y0, y1) -> tuple[Fraction, Fraction] | None:
    """The least and greatest x of the region ``points`` encloses between heights ``y0``
    and ``y1`` (every edge's part in the band, and every vertex in it), or ``None`` when
    the band misses it."""
    xs: list[Fraction] = []
    closed = list(points)
    if closed and closed[0] != closed[-1]:
        closed.append(closed[0])
    for (ax, ay), (bx, by) in zip(closed, closed[1:]):
        lo, hi = (ay, by) if ay <= by else (by, ay)
        if hi < y0 or lo > y1:
            continue
        if ay == by:
            xs += [ax, bx]
            continue
        for y in (max(lo, y0), min(hi, y1)):
            xs.append(ax + (bx - ax) * (y - ay) / (by - ay))
    if not xs:
        return None
    return min(xs), max(xs)


def blocked(wrap: Wrap, y0: Fraction, y1: Fraction, text: Fraction | None = None) -> tuple[Fraction, Fraction] | None:
    """What of the band from ``y0`` to ``y1`` ``wrap`` keeps text off, as ``(lo, hi)`` page
    px (distances included), or ``None`` when the line is not beside it.  A wrap polygon
    is read over the line's text only, down to ``text`` (its pitch less an ``auto``
    multiple's extra, in every mode)."""
    if not (y0 < wrap.bottom and y1 > wrap.top):
        return None
    if wrap.polygon is None:
        return wrap.left - wrap.dist_left, wrap.right + wrap.dist_right
    span = polygon_span(wrap.polygon, y0 - wrap.dist_bottom, (y1 if text is None else text) + wrap.dist_top)
    if span is None:
        return None
    return span[0] - wrap.dist_left, span[1] + wrap.dist_right


@dataclass(frozen=True)
class Segment:
    """A piece of a line beside drawings, in the layout unit from the text column's left
    edge: from ``left`` (the column's start, or a drawing's right edge plus ``distR``,
    measured from the margin's exact edge) to ``right``; its text is set out from
    ``left`` (or the indent) and then moved right by ``shift`` (the pixel rounding)."""

    left: Fraction
    right: Fraction
    shift: Fraction = Fraction(0)
    #: Whether ``left`` is a drawing's edge rather than the column's start.
    edge: bool = False
    #: Whether its tab stops are measured from ``left`` (below mode 15) rather than
    #: from the column's left edge.
    tabs_here: bool = False


@dataclass(frozen=True)
class Column:
    """Where the page's text column is: its left margin's exact edge and its start as
    drawn (a whole device pixel), device px; its width in the layout unit; the mode."""

    exact: Fraction
    start: int
    width: Fraction
    mode15: bool
    #: The width in device px, exactly (the column's twips).
    width_px: Fraction = Fraction(0)


def free_spans(wraps, y0: Fraction, y1: Fraction, column: Column,
               text: Fraction | None = None) -> tuple[list[tuple[Fraction, Fraction, bool]], list[Wrap]]:
    """What ``wraps`` leave of the column from ``y0`` down to ``y1``: ``(left, right,
    edge)`` spans, device px from the left margin's exact edge, left to right -- ``edge``
    the drawing (a :class:`Wrap`) whose side ``left`` is, ``False`` where it is the
    column's start -- and the drawings the band is beside (none: the whole column is
    free)."""
    beside: list[Wrap] = []
    spans: list[tuple[Fraction, Fraction]] = []
    width_px = column.width_px
    for wrap in wraps:
        span = blocked(wrap, y0, y1, text)
        if span is None:
            continue
        beside.append(wrap)
        lo, hi = span[0] - column.exact, span[1] - column.exact
        side = wrap.side or "bothSides"
        if side == "largest":
            side = "left" if lo >= width_px - hi else "right"
        if side == "left":
            hi = Fraction(10 ** 9)
        elif side == "right":
            lo = Fraction(-10 ** 9)
        spans.append((lo, hi, wrap))
    free = [(Fraction(0), width_px, False)]
    for lo, hi, wrap in sorted(spans, key=lambda span: span[:2]):
        out = []
        for a, b, edge in free:
            if hi <= a or lo >= b:
                out.append((a, b, edge))
                continue
            if lo > a:
                out.append((a, lo, edge))
            if hi < b:
                out.append((hi, b, wrap))
        free = out
    return free, beside


def segments(wraps, y0: Fraction, y1: Fraction, column: Column,
             text: Fraction | None = None) -> tuple[list[Segment] | None, list[Wrap]]:
    """The segments a line from ``y0`` down to ``y1`` has beside ``wraps``, left to right,
    and the drawings it is beside; ``(None, [])`` when it is beside none (the whole
    column is its)."""
    free, beside = free_spans(wraps, y0, y1, column, text)
    if not beside:
        return None, []
    result = []
    for a, b, edge in free:
        if (b - a) / TWIP_PX < MIN_SEGMENT_TWIPS:
            continue
        shift = Fraction(0)
        if edge:
            page = column.exact + a
            text = column.start + a
            if edge.text_shift is not None:
                shift = edge.text_shift
            elif column.mode15:
                shift = 0 if page == text else _half_up(page) - (_half_up if edge.half_up else _half_down)(text)
            elif edge.half_up:
                shift = _half_up(page) - text
            else:
                shift = _half_down(page) - text
        result.append(Segment(a * UNITS_PER_PX, b * UNITS_PER_PX, shift * UNITS_PER_PX, bool(edge),
                              bool(edge) and not column.mode15))
    return result, beside


@dataclass
class WrappedLine(Line):
    """A line beside drawings: its pieces from ``start`` to ``end`` in one or more
    segments, each ``(line, segment, first)`` -- ``first`` where the paragraph's first
    line's indent applies to it -- and ``drop``, the top the line was moved down to where
    no segment at its first place took text (``None`` when it was not moved)."""

    segments: tuple = ()
    drop: Fraction | None = None


def segment_geometry(geometry: Geometry, segment: Segment, first: bool) -> Geometry:
    """``geometry`` for the text of one segment: it starts at the segment's left edge or
    the indent, whichever is further right, and the first line's indent moves it from
    there -- a hanging indent no further left than the segment's edge; it ends at the
    segment's right edge or the indent's."""
    right = min(_unit(segment.right), geometry.right)
    if not segment.edge:
        return dataclasses.replace(geometry, right=right)
    edge = Fraction(segment.left)
    start = max(edge, Fraction(geometry.start))
    first_start = max(edge, start + geometry.first_start - geometry.start) if first else start
    return dataclasses.replace(geometry, first_start=_unit(first_start), start=_unit(start), right=right,
                               hanging_stop=None, tab_origin=_unit(edge) if segment.tabs_here else 0)


def _unit(value) -> int | Fraction:
    value = Fraction(value)
    return int(value) if value.denominator == 1 else value


def fits_in(items: list[Piece], line: Line, geometry: Geometry, first: bool) -> bool:
    """Whether a line broken in a segment keeps within it: not a word broken for want of
    room (Word passes such a segment over), and its first piece not past its edge."""
    if line.emergency:
        return False
    if line.end == line.start:
        return True
    x = geometry.first_start if first else geometry.start
    piece = items[line.start]
    if piece.kind in ("space", "break", "tab", "soft hyphen"):
        return True
    return x + piece.width <= geometry.right


def break_across(items: list[Piece], start: int, segs: list[Segment], geometry: Geometry, advances, *,
                 first: bool) -> WrappedLine | None:
    """One line from piece ``start`` across ``segs``: each takes as much as fits in it, in
    order; a segment its next word does not fit in is passed over.  ``None`` where no
    segment takes anything (the line must go down)."""
    n = len(items)
    placed = []
    i = start
    if i >= n:
        if not segs:
            return None
        return WrappedLine(start, start, segments=((Line(start, start), segs[0], first),))
    for segment in segs:
        if i >= n:
            break
        here = first and not placed
        g = segment_geometry(geometry, segment, here)
        if g.right <= (g.first_start if here else g.start):
            continue
        line = break_line(items, i, g, advances, first=here)
        if not fits_in(items, line, g, here):
            continue
        placed.append((line, segment, here))
        i = line.end
        if line.forced:
            break
    if not placed:
        return None
    last = placed[-1][0]
    return WrappedLine(placed[0][0].start, last.end, forced=last.forced, hyphenated=last.hyphenated,
                       segments=tuple(placed))

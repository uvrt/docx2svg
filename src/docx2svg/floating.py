"""Where Word puts a floating drawing (``wp:anchor``): the positioning rules, measured.

Every rule here was read off Word's PDF by ``tools/make_anchor_probe.py`` /
``tools/read_anchor_probe.py`` (ROADMAP.md, "Floating drawings -- measured"), not taken
from ECMA-376 20.4.2.10-11 and 20.4.3.1-2, which say what each ``relativeFrom`` names but
not how Word rounds, where a paragraph or a line starts, or what an alignment does with
the effect extent.  The model works in **twips**, as Word does:

* an offset (``wp:posOffset``) is truncated to whole twips, towards zero (400,003 EMU is
  629 twips, not 630);
* the extent an alignment uses is truncated to whole twips on each axis (a picture is
  then drawn at :func:`docx2svg.layout.picture_size`, a shape at that extent);
* a centred drawing may sit on a half twip;
* an alignment places the box of the extent **and** its effect extent (``left`` puts the
  effect extent's left edge on the frame's; ``right`` its right edge; ``center`` its
  middle), and the drawing is inside it; an offset places the drawing itself;
* the horizontal frames: ``page`` the page's edges; ``margin`` and ``column`` the left
  margin's exact edge (not the text column's, which is rounded to a device pixel) and
  the right margin; ``leftMargin`` / ``rightMargin`` the page's edge to the margin;
  ``insideMargin`` / ``outsideMargin`` one of those by the page's parity -- inside is the
  left margin on an odd page in mode 15, the right one below it; ``character`` the
  anchor's place in its line (:func:`docx2svg.layout.anchor_character_x`), a frame of no
  width;
* the vertical frames: ``page``; ``margin``; ``topMargin`` / ``bottomMargin``;
  ``insideMargin`` / ``outsideMargin`` (inside is the top margin on an odd page, in every
  mode; below mode 15 outside is too); ``paragraph`` -- the top of the paragraph, which is where the paragraph before
  it ends **after its own space after** (the rest of this paragraph's space before is
  inside it), or the top of the page's text for the page's first paragraph -- and any
  alignment against it is its top; ``line`` -- the anchor's line's box, whose top is the
  paragraph's top on its first line and the line's top on the others: in mode 15 the
  box is as tall as it reaches down to the line's pitch, and below mode 15 it has no
  height (``center`` and ``bottom`` put the drawing's middle and bottom on its top);
* ``inside`` / ``outside`` against the page or the margin: inside is left and top on an
  odd page; **against the page vertically, the top is half the header distance below the
  page's top edge and the bottom half the footer distance above its foot**;
* ``simplePos``: the point, in whole twips, from the page's top left corner --
  vertically through the paragraph: its distance from the paragraph's top rounded to
  whole twips.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from fractions import Fraction

from .model import Anchor, AnchorPosition

#: Device px per twip.
TWIP_PX = Fraction(300, 1440)
EMU_PER_TWIP = 635


def emu_twips(value: int) -> int:
    """EMU to whole twips, truncated towards zero, as Word takes an offset or an extent."""
    return int(Fraction(value, EMU_PER_TWIP))


@dataclass(frozen=True)
class Frames:
    """What an anchor on one page may be positioned against, in twips from the page's top
    left (``paragraph_top``, ``line_top``, ``line_bottom`` and ``character`` exact)."""

    page_width: int
    page_height: int
    left: int
    right: int
    top: int
    bottom: int
    header: int
    footer: int
    #: The page's place, 1-based: its parity picks inside and outside.
    page: int
    mode15: bool
    paragraph_top: Fraction
    line_top: Fraction
    line_bottom: Fraction
    character: Fraction
    #: Where ``page`` (and ``leftMargin`` / ``topMargin``) start: the page's corner, or in
    #: a table cell the cell's (``layout._Placer._cell_anchors``).
    page_x: Fraction = Fraction(0)
    page_y: Fraction = Fraction(0)
    #: The frames are a table cell's: ``margin`` has no height there, and every
    #: alignment against it is its top (``make_cell_anchor_probe.py``).
    cell: bool = False
    #: In a section of several text columns, the anchor's column: its left edge and width,
    #: twips (``column`` is the column's; ``margin`` stays the page's text area).
    column: tuple | None = None


def _box(anchor: Anchor) -> tuple[int, int, int, int, int, int]:
    """The extent, truncated to twips, and the effect extent's four sides, in twips."""
    cx, cy = (emu_twips(v) for v in anchor.extent)
    el, et, er, eb = (emu_twips(v) for v in anchor.effect)
    return cx, cy, el, et, er, eb


def _aligned(start, width, size: int, before: int, after: int, align: str | None, first: bool):
    """``size`` (with ``before`` and ``after`` of effect extent) aligned in a frame from
    ``start``, ``width`` wide: ``first`` says whether inside/outside is the frame's start."""
    if align in ("inside", "outside"):
        align = "left" if (align == "inside") == first else "right"
    if align in ("left", "top"):
        return start + before
    if align in ("right", "bottom"):
        return start + width - size - after
    if align == "center":
        return start + Fraction(width - (before + size + after), 2) + before
    return None


def horizontal(anchor: Anchor, frames: Frames) -> Fraction | None:
    """The drawing's left edge, twips from the page's left edge; ``None`` when the anchor
    names a frame this model does not know."""
    cx, _cy, el, _et, er, _eb = _box(anchor)
    if anchor.simple is not None:
        return Fraction(emu_twips(anchor.simple[0]))
    position: AnchorPosition = anchor.h
    odd = frames.page % 2 == 1
    inside_left = odd if frames.mode15 else not odd
    relative = position.relative or "column"
    frame = {
        "page": (frames.page_x, frames.page_width),
        "margin": (frames.left, frames.page_width - frames.left - frames.right),
        "column": frames.column or (frames.left, frames.page_width - frames.left - frames.right),
        "leftMargin": (frames.page_x, frames.left),
        "rightMargin": (frames.page_width - frames.right, frames.right),
        "insideMargin": (0, frames.left) if inside_left else (frames.page_width - frames.right, frames.right),
        "outsideMargin": (frames.page_width - frames.right, frames.right) if inside_left else (0, frames.left),
        "character": (frames.character, 0),
    }.get(relative)
    if frame is None:
        return None
    start, width = frame
    if position.offset is not None:
        return start + emu_twips(position.offset)
    if relative == "character":
        # A frame of no width: left puts the drawing's left edge on the character, right
        # its right edge, centre its middle.
        return _aligned(start, 0, cx, el, er, position.align, odd)
    return _aligned(start, width, cx, el, er, position.align or "left", odd)


def vertical(anchor: Anchor, frames: Frames) -> Fraction | None:
    """The drawing's top edge, twips from the page's top edge."""
    _cx, cy, _el, et, _er, eb = _box(anchor)
    if anchor.simple is not None:
        above = Fraction(emu_twips(anchor.simple[1])) - frames.paragraph_top
        return frames.paragraph_top + math.floor(above + Fraction(1, 2))
    position: AnchorPosition = anchor.v
    odd = frames.page % 2 == 1
    relative = position.relative or "paragraph"
    height = frames.page_height
    margins = {
        "page": (frames.page_y, height),
        "margin": (frames.top, height - frames.top - frames.bottom),
        "topMargin": (frames.page_y, frames.top),
        "bottomMargin": (height - frames.bottom, frames.bottom),
        "insideMargin": (0, frames.top) if odd else (height - frames.bottom, frames.bottom),
        # Below mode 15 the outside margin is the inside one (measured: both are the top
        # margin on an odd page).
        "outsideMargin": ((height - frames.bottom, frames.bottom) if odd else (0, frames.top)) if frames.mode15
        else ((0, frames.top) if odd else (height - frames.bottom, frames.bottom)),
    }
    if relative in margins:
        start, extent = margins[relative]
        if position.offset is not None:
            return start + emu_twips(position.offset)
        align = position.align or "top"
        if frames.cell and relative == "margin":
            align = "top"
        if relative == "page" and align in ("inside", "outside"):
            if (align == "inside") == odd:
                return Fraction(frames.header, 2) + et
            return height - Fraction(frames.footer, 2) - cy - eb
        return _aligned(start, extent, cy, et, eb, align, odd)
    if relative == "paragraph":
        if position.offset is not None:
            return frames.paragraph_top + emu_twips(position.offset)
        return frames.paragraph_top + et
    if relative == "line":
        top = frames.line_top
        if position.offset is not None:
            return top + emu_twips(position.offset)
        align = position.align or "top"
        bottom = frames.line_bottom if frames.mode15 else top
        if align == "top":
            return top + et
        if align == "bottom":
            return bottom - cy - eb
        if align == "center":
            return Fraction(top + bottom, 2) - Fraction(et + cy + eb, 2) + et
        return None
    return None


#: What a drawing anchored in a table cell may be positioned by, as measured
#: (``make_cell_anchor_probe.py``): ``(axis, relativeFrom, offset or the alignment)``.
#: In the cell (``layoutInCell``, or mode 15 whatever it says) and against the page (below
#: mode 15 without ``layoutInCell``).
CELL_POSITIONS = frozenset({
    ("h", "column", "offset"), ("h", "column", "left"), ("h", "column", "center"), ("h", "column", "right"),
    ("h", "margin", "offset"), ("h", "margin", "left"), ("h", "margin", "center"), ("h", "margin", "right"),
    ("h", "page", "offset"), ("h", "leftMargin", "offset"), ("h", "character", "offset"),
    ("v", "paragraph", "offset"), ("v", "paragraph", "top"), ("v", "paragraph", "center"),
    ("v", "paragraph", "bottom"), ("v", "line", "offset"), ("v", "margin", "offset"), ("v", "margin", "top"),
    ("v", "margin", "center"), ("v", "margin", "bottom"), ("v", "page", "offset"), ("v", "topMargin", "offset"),
})
PAGE_POSITIONS = CELL_POSITIONS - {("h", "character", "offset")}


def in_cell(anchor: Anchor, mode15: bool) -> bool:
    """Whether a drawing anchored in a table cell is positioned against the cell: with
    ``layoutInCell``, in mode 15 whatever it says, and below mode 15 without it where
    either axis is against the anchor's ``character`` or ``line`` -- both axes then
    (``make_cell_anchor_probe.py``: a ``character`` offset takes ``paragraph`` from the
    cell, a ``line`` offset ``column``)."""
    return (mode15 or anchor.layout_in_cell or (anchor.h.relative or "column") == "character"
            or (anchor.v.relative or "paragraph") == "line")


def cell_position_modelled(anchor: Anchor, mode15: bool) -> str | None:
    """Why a drawing anchored in a table cell cannot be positioned by the model (``None``
    where it can): a position :data:`CELL_POSITIONS` / :data:`PAGE_POSITIONS` does not hold."""
    if anchor.simple is not None:
        return "simplePos"
    known = CELL_POSITIONS if in_cell(anchor, mode15) else PAGE_POSITIONS
    for axis, position in (("h", anchor.h), ("v", anchor.v)):
        relative = position.relative or ("column" if axis == "h" else "paragraph")
        how = "offset" if position.offset is not None else (position.align or ("left" if axis == "h" else "top"))
        if (axis, relative, how) not in known:
            return f"{relative} {how}"
    return None

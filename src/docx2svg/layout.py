"""Where every glyph, rule and picture of a page goes: the model's layout, placed.

Phase 5's premise is that **emission is a transcription of verified positions**.  Nothing
here decides where a line breaks, where a page ends or where a baseline sits: those are
:mod:`docx2svg.linebreak`, :mod:`docx2svg.paginate` and :mod:`docx2svg.vertical`, each
measured against Word (ROADMAP.md, Phases 2-4).  This module walks their result and
writes down, per page, the numbers an SVG needs:

* **baselines** -- the page stack of :func:`paginate._scan` (the same gaps, borders and
  page-top rules), with each line's baseline rounded in its line box by
  :func:`vertical.baseline_in_box`, as ``tools/baselines.predict`` does;
* **pen positions** -- every glyph's x from the line's start, its advance and its kern
  pair with the next glyph, in Word's layout unit (1/4096 pt) exactly, then tab stops
  and the paragraph's alignment (:func:`line_positions`);
* **the drawn size and offset** of superscripts, subscripts and ``w:position`` runs;
* **decorations** -- underline, strikethrough, highlight, shading, paragraph borders --
  and inline pictures.

And, as honestly as the rest: where the paginator stops (a table, a floating drawing, a
paragraph it cannot measure, a multi-column section), the page records *that* it
stopped and where (:class:`Stop`), and no page after it is invented.

Standard library only.
"""

from __future__ import annotations

import dataclasses
import math
from dataclasses import dataclass, field
from fractions import Fraction
from types import SimpleNamespace

from . import linebreak, paginate, vertical
from .lines import label_format
from .model import Document, Paragraph, Table
from .resolve import (
    character_format, resolve_mark, resolve_paragraph, resolve_run, script_half_points, script_raise_half_points,
)
from .resolve.cascade import Resolved
from .units import device_page_extent_pt
from .vertical import LAYOUT_UNIT_PX, PX_PER_PT, LineBox, baseline_in_box, round_half_up, twips_to_px

#: Device px per half point: a raised or lowered run moves by whole half points.
PX_PER_HALF_POINT = PX_PER_PT / 2


def units_px(units) -> Fraction:
    """A length in Word's layout unit (1/4096 pt) in device px."""
    return Fraction(units) * LAYOUT_UNIT_PX


def half_points_offset_px(half_points: int) -> int:
    """How many device px a run raised by ``half_points`` is drawn above its line's
    baseline (negative: below).

    Measured on ``w:position`` by ``make_script_probe.py`` (``position-*``: ±1, 2, 3, 6,
    12, 24 and 48 half points, 384 groups): every drawn offset is the distance rounded
    half up *in magnitude* -- +6 and -6 half points (12.5 px) are both drawn 13 px from
    the baseline -- so the raise is taken whole and the device rounds it, symmetrically.
    """
    magnitude = round_half_up(abs(half_points) * PX_PER_HALF_POINT)
    return magnitude if half_points >= 0 else -magnitude


# -- what a page holds -----------------------------------------------------------------


#: The glyph Word draws each tab leader with (``make_tab_leader_probe.py``, ``kinds``:
#: ``heavy`` is the underscore too).
LEADER_GLYPHS = {"dot": ".", "hyphen": "-", "underscore": "_", "heavy": "_", "middleDot": "\u00b7"}


@dataclass
class Span:
    """Consecutive glyphs of one run, in one format, on one line: one SVG ``<text>``."""

    path: str
    chars: list[str]
    #: Pen x of every glyph, device px from the page's left edge, exact.
    xs: list[Fraction]
    #: The baseline the glyphs are drawn on, device px from the page's top (the line's
    #: baseline, moved by a script offset or ``w:position``).
    y: int
    #: The pen x after the last glyph.
    end: Fraction
    face: str
    bold: bool
    italic: bool
    #: The size the glyphs are drawn at, in half points (a script's is its OS/2 size).
    half_points: int
    color: str | None = None
    #: ``w:u``'s value, ``w:strike``/``w:dstrike``, ``w:highlight``'s colour name, the
    #: run's shading fill.
    underline: str | None = None
    strike: bool = False
    double_strike: bool = False
    highlight: str | None = None
    shading: str | None = None
    #: ``label`` for a list label, ``hyphen`` for the hyphen a soft hyphen draws at a
    #: line's end, ``leader`` for a tab's leader, ``tab`` for the space an underlined tab
    #: is underlined under, ``text`` otherwise.
    kind: str = "text"
    #: The run's text effects, ``(fill, outline)`` as ``ooxml_common.drawingml.model``
    #: values (either ``None``: the colour's fill, no outline), or ``None`` for plain text.
    paint: object = None
    #: The line's baseline (before any script or position offset).
    line_baseline: int = 0
    #: The line's text extent about its baseline, device px, for highlight and shading.
    line_above: Fraction = Fraction(0)
    line_below: Fraction = Fraction(0)
    #: ``superscript`` / ``subscript`` / ``None``.
    vertical_align: str | None = None


@dataclass
class Picture:
    """An inline picture (``wp:inline`` with an ``a:blip``), or a drawing it cannot show
    (``relationship`` ``None``: drawn as a placeholder of its extent)."""

    path: str
    relationship: str | None
    x: Fraction
    y: Fraction
    width: Fraction
    height: Fraction
    #: The part the relationship is the part's own (a header's picture); ``None``: the
    #: main document's.
    part: str | None = None


@dataclass
class Rule:
    """A filled rectangle the text needs: a paragraph border, shading, a highlight, an
    underline or strike (device px, exact)."""

    kind: str
    path: str
    x: Fraction
    y: Fraction
    width: Fraction
    height: Fraction
    color: str | None = None
    #: A border's or underline's style (``single``, ``double``, ``dotted``...).
    style: str | None = None


@dataclass
class Line:
    path: str
    number: int
    baseline: int
    #: The exact top of the line's pitch, and its pitch, device px.
    top: Fraction
    pitch: Fraction
    spans: list[Span] = field(default_factory=list)
    #: The paragraph's ``w14:paraId``, when the file has one.
    paragraph_id: str | None = None
    #: The device-px rows the line's text occupies (:func:`text_band`).
    band: tuple[int, int] | None = None
    #: ``header`` or ``footer`` for a line of a story, and the story's part
    #: (``word/header1.xml``); ``None`` for the body.
    story: str | None = None
    part: str | None = None
    #: The 0-based text column the line is in, in a section of several columns (a
    #: footnote's line in the column it is drawn under); ``None`` in a section of one
    #: column, and for a header's, a footer's or a text box's line.  Metadata only:
    #: nothing is drawn from it.
    column: int | None = None


@dataclass
class Stop:
    """Where the model stopped: the obstacle, why, and how far down the page it starts."""

    path: str
    reason: str
    y: Fraction
    #: Top-level paragraphs and tables from the obstacle to the end that are not drawn.
    remaining: int
    #: What is known of the obstacle's extent, device px from the page's top left: a
    #: floating drawing's box when its anchor states it against the page or margin.
    extent: tuple[Fraction, Fraction, Fraction, Fraction] | None = None


@dataclass
class Float:
    """A floating drawing (``wp:anchor``) placed on its page: the drawing's box (device px
    from the page's top left), how it stacks, and what it paints -- the primitives of
    :mod:`docx2svg.drawing` and a text box's lines, in page coordinates."""

    path: str
    x: Fraction
    y: Fraction
    width: Fraction
    height: Fraction
    #: Painted before everything on the page (``behindDoc``), or after everything.
    behind: bool
    #: ``(relativeHeight, the order it was met in)``: the order floats of one layer paint in.
    order: tuple
    primitives: list = field(default_factory=list)
    lines: list = field(default_factory=list)
    rules: list = field(default_factory=list)
    #: ``header`` or ``footer`` for a story's drawing, with the story's part.
    story: str | None = None
    part: str | None = None


@dataclass
class Page:
    #: The page's 0-based place in the document.
    number: int
    #: The page Word exports: whole 1/300-inch device pixels (Phase 0.5).
    width_px: int
    height_px: int
    width_pt: float
    height_pt: float
    #: The text column's left and right edges and the text area's top and bottom.
    column: tuple[Fraction, Fraction, Fraction, Fraction]
    lines: list[Line] = field(default_factory=list)
    rules: list[Rule] = field(default_factory=list)
    pictures: list[Picture] = field(default_factory=list)
    stop: Stop | None = None
    #: What the paginator knows of the page (its number as ``PAGE`` shows it, its
    #: section, whether the section starts on it, its story kind, whether it is a blank
    #: page Word inserts): :class:`docx2svg.paginate.PageInfo`, public; ``None`` on a
    #: page made after the layout stopped.
    info: object = None
    #: Floating drawings (:class:`Float`), in the order they were placed.
    floats: list = field(default_factory=list)
    #: What the drawings' paint refers to -- gradients, patterns, filters -- as SVG
    #: definitions, with the ids they are referred to by (``ooxml_common.drawingml.svg.Defs``);
    #: ``None`` until a drawing needs one.
    defs: object = None
    #: Where the body's stack ends on the page (``None`` for a blank page): the top of
    #: its footnotes under ``w:pos`` ``beneathText``.
    body_end: Fraction | None = None

    def layers(self) -> tuple[list, list]:
        """The floats behind the text and those in front of it, each in paint order:
        ascending ``relativeHeight``, then the order met (``make_anchor_probe.py``, ``z``)."""
        ordered = sorted(self.floats, key=lambda f: f.order)
        return [f for f in ordered if f.behind], [f for f in ordered if not f.behind]

    def paint(self) -> list[tuple]:
        """What the page paints, in Word's order (``make_anchor_probe.py``,
        ``make_story_anchor_probe.py``): ``("floats", [...])``, ``("under",)`` (the rules
        text sits on), ``("lines", [...])``, ``("over",)`` (the rules drawn over text),
        ``("pictures",)``.  The body's drawings behind the text first, then its rules,
        text and pictures, then its drawings in front -- and where a header or footer has
        drawings, the story before all of that: its drawings behind, its text, its
        drawings in front (Word paints a header's picture in front of the header's text
        and under the body's, and the body's drawings behind its text over the header's
        text)."""
        behind, front = self.layers()
        story_behind = [f for f in behind if f.story]
        story_front = [f for f in front if f.story]
        body_behind = [f for f in behind if not f.story]
        body_front = [f for f in front if not f.story]
        # A page with no drawing a story's text could be above or below is painted as
        # before drawings were drawn (the body's lines first, then its stories').
        if not (story_behind or story_front or (body_behind and any(line.story for line in self.lines))):
            return [("floats", body_behind), ("under",), ("lines", list(self.lines)), ("over",), ("pictures",),
                    ("floats", body_front)]
        return [("floats", story_behind), ("lines", [line for line in self.lines if line.story]),
                ("floats", story_front), ("floats", body_behind), ("under",),
                ("lines", [line for line in self.lines if not line.story]), ("over",), ("pictures",),
                ("floats", body_front)]

    def text_lines(self) -> list:
        """Every line of text the page draws, in paint order (:meth:`paint`), text boxes
        included."""
        out: list = []
        for step in self.paint():
            if step[0] == "lines":
                out.extend(step[1])
            elif step[0] == "floats":
                out.extend(line for placed in step[1] for line in placed.lines)
        return out


@dataclass
class Layout:
    pages: list[Page]
    #: ``(code, message, 1-based page or None)``, deduplicated, in order.
    warnings: list[tuple[str, str, int | None]] = field(default_factory=list)
    #: How much of the document this covers (:class:`docx2svg.coverage.Coverage`): set by
    #: :func:`docx2svg.convert_docx_to_layout` and its siblings, ``None`` from :func:`lay_out`.
    coverage: object = None

    def warn(self, code: str, message: str, page: int | None = None) -> None:
        entry = (code, message, page)
        if entry not in self.warnings:
            self.warnings.append(entry)


# -- horizontal ------------------------------------------------------------------------

#: The character a decimal tab stop aligns on: the *system's* decimal separator, not the
#: document's language.  Measured by ``make_render_probe.py`` (``tabs``): on this machine
#: (a metric macOS, whose ``Normal.dotm`` also gives the 708-twip default tab stop)
#: ``12.345`` and ``1234.5`` after a decimal stop in an ``en-GB`` document are both drawn
#: ending at the stop, as after a right stop -- Word found no separator in them.  Like
#: :data:`linebreak.APPLICATION_DEFAULT_TAB`, a fact about the application.
APPLICATION_DECIMAL_SEPARATOR = ","


#: Dotted and dashed underlines: device px on, and the period (ROADMAP.md, 5.15).
UNDERLINE_DASHES = {"dotted": (6, 12), "dash": (16, 24)}

#: One twip in device px.
TWIP_PX = PX_PER_PT / 20


def _in_column(lines, frame) -> None:
    """Say which text column ``lines`` were placed in: ``frame``'s, when its section has
    several (:attr:`Line.column`)."""
    if getattr(frame, "multi", True):
        for line in lines:
            line.column = frame.column


def picture_size(cx: int, cy: int) -> tuple[Fraction, Fraction]:
    """The size Word draws an inline picture of extent ``cx`` x ``cy`` EMU, device px: the
    extent scaled, both ways alike, to fit the extent truncated to whole twips
    (``make_picture_probe.py``, ``make_picture_place_probe.py``: every picture's width and
    height to 0.001 px; the exact extent is up to 0.8 px wider and 0.2 px taller)."""
    emu = Fraction(PX_PER_PT, 12700)
    if cx <= 0 or cy <= 0:
        return cx * emu, cy * emu
    scale = min(Fraction(cx // 635 * 635, cx), Fraction(cy // 635 * 635, cy))
    return cx * scale * emu, cy * scale * emu


def round_half_away(x: Fraction) -> int:
    """``x`` rounded to the nearest integer, a half away from zero."""
    return -round_half_up(-x) if x < 0 else round_half_up(x)


def grid_line_px(page_left_twips, grid_twips) -> int:
    """The device px a table's grid line is drawn centred on: the text column's left edge
    rounded to a whole px, and the grid line's distance from it, rounded half away from
    zero (``make_table_line_probe.py``: 7,680 / 7,680 vertical lines, every width, both
    sides of the edge, a margin on and off the pixel grid; rounded half up from the page's
    edge, 5,968: the left border of a table whose first grid line is *n*.5 px left of the
    text column's edge went a pixel right)."""
    return round_half_up(page_left_twips * TWIP_PX) + round_half_away(grid_twips * TWIP_PX)


def _segment_width(items, start: int, end: int, advances) -> int | Fraction:
    total = 0
    for j in range(start, end):
        piece = items[j]
        if piece.kind in (linebreak.TAB, linebreak.BREAK, linebreak.SOFT_HYPHEN):
            continue
        total += piece.width + (linebreak._kern(piece, items[j + 1], advances) if j + 1 < len(items) else 0)
    return total


def line_positions(items, line: linebreak.Line, geometry: linebreak.Geometry, advances, *, first_line: bool,
                   alignment: str | None, last_line: bool, leaders: dict | None = None) -> dict[int, int | Fraction]:
    """Piece index -> pen x from the column's left edge, in 1/4096 pt, for one line.

    The pen advances by each piece's width (tracking included) and its kern pair with
    the next piece -- the breaker's own arithmetic (:func:`linebreak.break_pieces`), so a
    line is drawn exactly as wide as it was measured.  A left tab jumps to its stop; the
    text after a right, centre or decimal stop is moved so it ends at, centres on or
    puts its decimal point at the stop, but never left of the tab.  Then the line moves
    by its paragraph's alignment: centred or right-aligned between the line's start and
    the right edge, its trailing spaces left out; justified (``both``, ``distribute``)
    lines but the last and those a break ends get the slack spread over the spaces after
    their last tab; a justified line past the edge (mode 15) the overrun taken from them
    in proportion to their widths.

    ``leaders``, when given, is filled with every tab: piece index -> ``(its stop's leader
    or None, the pen at the tab, the pen after it)``, before the line is aligned.
    """
    x: int | Fraction = geometry.first_start if first_line else geometry.start
    start_x = x
    out: dict[int, int | Fraction] = {}
    j = line.start
    end = line.end
    while j < end:
        piece = items[j]
        if piece.kind == linebreak.TAB:
            out[j] = x
            stop, kind, _default = linebreak.next_tab(x, geometry, first_line)
            nxt = next((k for k in range(j + 1, end) if items[k].kind == linebreak.TAB), end)
            if kind in ("right", "center", "decimal"):
                content_end = nxt
                while content_end > j + 1 and items[content_end - 1].kind in (linebreak.SPACE, linebreak.BREAK):
                    content_end -= 1
                if kind == "decimal":
                    point = next((k for k in range(j + 1, content_end)
                                  if items[k].char == APPLICATION_DECIMAL_SEPARATOR), content_end)
                    width = _segment_width(items, j + 1, point, advances)
                    begin = stop - width
                else:
                    width = _segment_width(items, j + 1, content_end, advances)
                    begin = stop - (width if kind == "right" else Fraction(width, 2))
                x = max(x, begin)
            else:
                x = stop
            if leaders is not None:
                leaders[j] = (linebreak.tab_leader(out[j], geometry, first_line), out[j], x)
            j += 1
            continue
        out[j] = x
        if piece.kind not in (linebreak.BREAK, linebreak.SOFT_HYPHEN):
            x = x + piece.width + (linebreak._kern(piece, items[j + 1], advances) if j + 1 < len(items) else 0)
        j += 1
    # The line's content ends at its last piece that is not a space, a break or an
    # undrawn soft hyphen.
    last = end - 1
    while last >= line.start and items[last].kind in (linebreak.SPACE, linebreak.BREAK) or (
            last >= line.start and items[last].kind == linebreak.SOFT_HYPHEN and not line.hyphenated):
        last -= 1
    if last < line.start:
        return out
    content_end = out[last] + (items[last].hyphen if items[last].kind == linebreak.SOFT_HYPHEN else
                               items[last].width if items[last].kind != linebreak.TAB else 0)
    slack = geometry.right - content_end
    alignment = alignment or "left"
    if alignment in ("center",) and slack > 0:
        # Half the slack, truncated to the layout unit: an odd slack puts the line half a
        # unit left of the middle (``style-document.docx``: two centred quotes, drawn at
        # 754006 and 757515 units where the middle is 754006.5 and 757515.5).
        shift = math.floor(Fraction(slack, 2))
        return {k: v + shift for k, v in out.items()}
    if alignment in ("right", "end") and slack > 0:
        return {k: v + slack for k, v in out.items()}
    if alignment == "distribute" and slack > 0 and (last_line or line.forced):
        # A distributed paragraph's last line is spread too, over every gap between its
        # glyphs; Word gives the gaps after a space more than the others, by a rule not
        # measured (make_render_probe.py: 51.02 and 64.53 px on one line, 5.43 and 15.34
        # on another), so the gaps are equal here and those lines are known to differ.
        drawn = [k for k in sorted(out) if k <= last]
        if len(drawn) > 1:
            share = Fraction(slack, len(drawn) - 1)
            return {k: out[k] + share * min(drawn.index(k) if k in drawn else len(drawn) - 1, len(drawn) - 1)
                    for k in out}
    # A justified line past the edge (mode 15, :func:`linebreak.squeezes`) has its spaces
    # compressed by the overrun -- the last line of its paragraph too.
    squeezed = alignment == "both" and slack < 0 and geometry.squeeze
    if squeezed or (alignment in ("both", "distribute") and slack > 0 and not last_line and not line.forced):
        tab = max((k for k in range(line.start, last + 1) if items[k].kind == linebreak.TAB), default=line.start - 1)
        spaces = [k for k in range(tab + 1, last) if items[k].kind == linebreak.SPACE]
        if spaces:
            share = Fraction(slack, len(spaces))
            moved: dict[int, int | Fraction] = {}
            extra = Fraction(0)
            total_width = sum(items[k].width for k in spaces)
            before = 0
            for k in sorted(out):
                if slack < 0:
                    # Compressed: the overrun taken from the spaces in proportion to their
                    # widths, the shift after each rounded to the unit (make_justify_probe.py:
                    # spaces of two sizes; every glyph of the first round's compressed lines
                    # exact, the second round's to one unit but for 35 of 35,404).
                    moved[k] = out[k] + math.floor(Fraction(slack * before, total_width) + Fraction(1, 2))
                else:
                    moved[k] = out[k] + extra
                if k in spaces:
                    extra += share
                    before += items[k].width
            return moved
    del start_x
    return out


def wrapped_positions(items, line, geometry: linebreak.Geometry, advances, *, first_line: bool,
                      alignment: str | None, last_line: bool, leaders: dict | None = None) -> dict[int, int | Fraction]:
    """:func:`line_positions` of a line, or of each segment of a line beside a drawing
    (:class:`docx2svg.wrap.WrappedLine`): set in its segment as a line is in the column --
    aligned, centred or justified there, the paragraph's last segment as its last line --
    and moved by the segment's shift (:mod:`docx2svg.wrap`)."""
    from .wrap import segment_geometry

    segments = getattr(line, "segments", ())
    # A line broken for another of unequal columns is set in that column
    # (:class:`paginate.ColumnLine`).
    geometry = getattr(line, "geometry", None) or geometry
    if not segments:
        return line_positions(items, line, geometry, advances, first_line=first_line, alignment=alignment,
                              last_line=last_line, leaders=leaders)
    out: dict[int, int | Fraction] = {}
    for k, (part, segment, first) in enumerate(segments):
        here = segment_geometry(geometry, segment, first)
        found: dict = {}
        positions = line_positions(items, part, here, advances, first_line=first, alignment=alignment,
                                   last_line=last_line and k == len(segments) - 1, leaders=found)
        shift = segment.shift
        if leaders is not None:
            leaders.update({j: (leader, x0 + shift, x1 + shift) for j, (leader, x0, x1) in found.items()})
        out.update({j: x + shift if shift else x for j, x in positions.items()})
    return out


# -- the page walk ---------------------------------------------------------------------


@dataclass
class _Paragraph:
    """What the page walk needs of a flow paragraph beyond :class:`paginate.Para`."""

    paragraph: Paragraph
    resolved: Resolved
    pieces: list
    lines: list
    geometry: linebreak.Geometry
    label: str | None


def _color(value) -> str | None:
    if not value or value == "auto":
        return None
    return value if len(value) == 6 and all(c in "0123456789abcdefABCDEF" for c in value) else None


#: ``w:highlight``'s names (ECMA-376 17.18.40) as colours.
HIGHLIGHT = {
    "black": "000000", "blue": "0000FF", "cyan": "00FFFF", "green": "00FF00", "magenta": "FF00FF",
    "red": "FF0000", "yellow": "FFFF00", "white": "FFFFFF", "darkBlue": "000080", "darkCyan": "008080",
    "darkGreen": "008000", "darkMagenta": "800080", "darkRed": "800000", "darkYellow": "808000",
    "darkGray": "808080", "lightGray": "C0C0C0",
}


def _shading(value) -> str | None:
    if not value:
        return None
    pattern, _color_value, fill = value
    if pattern in ("nil",):
        return None
    return _color(fill)


class _Placer:
    def __init__(self, document: Document, package: bytes | None, advances, metrics, decorations) -> None:
        self.document = document
        self.advances = advances
        self.metrics = metrics
        self.decorations = decorations
        self.layout = Layout(pages=[])
        self.run_cache: dict[int, Resolved] = {}
        from .scope import paragraph_features

        self.features = paragraph_features(package) if package is not None else None
        from .chart import Frames

        #: Draws charts and SmartArt from the package's parts (:mod:`docx2svg.chart`).
        self.frames = Frames(package, document)

    # -- the flow ------------------------------------------------------------------------

    def features_with_columns(self) -> list[str | None]:
        """:func:`scope.paragraph_features` per paragraph of the body, a note's paragraph
        placed in it left as laid out.  (A section of several text columns is laid out
        by column, :func:`paginate._lay_columns`: it no longer stops the layout.)"""
        document = self.document
        features = list(self.features) if self.features is not None else None
        top_level = 0
        index = 0
        out: list[str | None] = []

        placed = getattr(self, "placed_notes", set())

        def walk(item, section, in_table, note=False):
            nonlocal index
            note = note or id(item) in placed
            if isinstance(item, Table):
                for row in item.rows:
                    for cell in row:
                        for inner in cell:
                            walk(inner, section, True, note)
                return
            if note:
                # A note's paragraph: not in the file where it now stands.
                out.append(None)
                return
            reason = features[index] if features is not None and index < len(features) else None
            out.append(reason)
            index += 1

        for item in document.body:
            section = paginate._section_of(document, top_level)
            walk(item, section, False)
            if not isinstance(item, Table):
                top_level += 1
        return out

    def run(self) -> Layout:
        from .notes import footnote_references, with_endnotes, with_footnotes

        source = self.document
        # The endnotes, numbered and placed in the flow as paragraphs of it.
        self.document, self.placed_notes = with_endnotes(self.document)
        # SEQ counted and REF replaced by its bookmark's text, once: neither depends on the
        # layout (``make_xref_field_probe.py``); PAGEREF is computed with PAGE, below.
        self.document = self._cross_references(self.document)
        original = self.document
        # A field is laid out now (its instruction is never text, and a computed field's
        # value is this layout's own): the renderer no longer stops at one.
        # A picture bullet is laid out too (``linebreak.picture_bullet_units``); one whose
        # picture is not read stops the layout as unmeasurable.
        features = [None if reason in ("field", "drawing (picture bullet)") else reason
                    for reason in self.features_with_columns()]
        # A drop cap's frame is laid out too (``paginate.DropCap``): one before a paragraph
        # of the body; any other frame still stops the layout.
        body = list(original.body)
        drops = {id(item) for k, item in enumerate(body)
                 if isinstance(item, Paragraph) and k + 1 < len(body) and isinstance(body[k + 1], Paragraph)
                 and paginate.is_drop_cap(resolve_paragraph(original, item))}
        features = [None if reason == "frame" and id(paragraph) in drops else reason
                    for reason, paragraph in zip(features, _all_paragraphs(original.body))]
        # The footnotes' references and the notes' own marks, numbered (:mod:`docx2svg.notes`);
        # numbers that restart on each page are numbered again from the pages found.
        by_page = any(section.footnote_restart == "eachPage" for section in original.sections)
        if any(in_table for *_, in_table in footnote_references(original)) and any(
                paginate.is_multi_column(section) for section in original.sections):
            # In one column a cell's footnotes are laid out (``paginate._table_notes``); in a
            # section of several, not measured.
            self.layout.warn("footnotes-not-drawn", "a footnote referenced in a table cell in a document with "
                             "sections of several text columns takes no room at the page foot and is not drawn "
                             "(its reference's number is)")
        footnote_pages: dict | None = None
        document = with_footnotes(original)
        values: dict = {}
        for _attempt in range(4):
            items = paginate.flow(document, self.advances, self.metrics, features=features)
            result = paginate.paginate(document, items)
            self.document, self.items = document, items
            found = self._body_field_values(items, result)
            pages = _footnote_pages(items, result) if by_page else None
            if found == values and pages == footnote_pages:
                break
            # A computed field's value moves what it is measured in (a wider page number
            # can break its line elsewhere): lay the body out again with the values found,
            # until they are the ones it was laid out with.  So does a note's number.
            values, footnote_pages = found, pages
            document = substitute_document(with_footnotes(original, footnote_pages), values)
        else:
            self.layout.warn("field-layout-unsettled", "the body's page-number fields did not settle on one "
                             "layout in four passes; the last is drawn")
        self.references = paginate.story_references(document)
        self._notes(items)
        details = self._details(items)
        ends = result.starts[1:] + [paginate.Position(result.stopped, result.stopped_line)
                                    if result.stopped is not None else paginate.Position(len(items))]
        self.result = result
        page_number = 0
        drawn_pages = [page for page in sorted(result.info) if not result.info[page].blank]
        for index, (start, end) in enumerate(zip(result.starts, ends)):
            number = drawn_pages[index] if index < len(drawn_pages) else paginate._page_number(result, index)
            # Blank pages Word inserts (an ``oddPage`` / ``evenPage`` section on the wrong
            # parity, a footnote continuation page) have no lines, and no header or footer
            # (``make_story_select_probe.py``): emit them empty.
            while page_number < number:
                blank = self._blank(start, page_number, result.info.get(page_number))
                self._footnotes(blank, page_number)
                self.layout.pages.append(blank)
                page_number += 1
            stopped = result.stopped if (index == len(result.starts) - 1 and result.stopped is not None) else None
            info = result.info.get(number)
            # The bands of the page's wrapTopAndBottom drawings, as the paginator found them.
            self._bands = result.bands.get(number, [])
            self._wraps = result.wraps.get(number, [])
            page = self._page(items, details, start, end, number, stopped, result.stopped_line, info)
            self._bands = []
            self._wraps = []
            self._draw_stories(page, info)
            self._footnotes(page, number)
            self.layout.pages.append(page)
            page_number = number + 1
        if not result.starts and result.stopped is not None:
            self.layout.pages.append(self._page(items, details, paginate.Position(0), paginate.Position(0), 0,
                                                result.stopped))
        if not self.layout.pages and document.sections:
            # A body with nothing in it is still a page (Word exports one), and shows its
            # section's first page's stories.
            start = document.sections[0].page_number_start
            number = 1 if start is None else start
            info = paginate.PageInfo(number, 0, True, paginate.story_kind(document, 0, section_first=True,
                                                                          number=number))
            page = self._new_page(0, 0, info.story)
            page.info = info
            self._draw_stories(page, info)
            self.layout.pages.append(page)
        self._line_numbers(source)
        return self.layout

    def _line_numbers(self, source: Document) -> None:
        """Line numbers (``w:lnNumType``) beside the body's lines, as Word draws them
        (``make_line_number_probe.py``, every case, with no ``settings.xml`` and in mode 15):

        * **which lines**: every line of the body's own paragraphs, an empty one's too --
          not a table's, a header's, a footer's, a note's or a text box's, nor a paragraph
          under ``w:suppressLineNumbers``, which are not counted either;
        * **the count** starts at ``w:start`` + 1 and restarts at each page (``newPage``,
          unstated) or section (``newSection``); ``continuous`` goes on from the section
          before.  A number is drawn where the count is a multiple of ``w:countBy``;
        * **where**: on the line's baseline, ending ``w:distance`` (auto: 360 twips) left of
          the text column's edge -- the paragraph's indent does not move it -- that edge
          rounded to a whole device px, each digit advancing by whole device px as a note
          number's does (:func:`linebreak.note_number_units`);
        * **in** the ``LineNumber`` character style over the document's defaults.

        A section of several text columns is not numbered (not measured): warned
        ``line-numbers-not-drawn``."""
        sections = source.sections
        if not any(section.line_numbering for section in sections):
            return
        from .model import Run, RunProperties
        from .resolve import character_format, resolve_paragraph, resolve_run

        section_of: dict[str, int] = {}
        suppressed: set[str] = set()
        for k, section in enumerate(sections):
            for paragraph in source.paragraphs[section.first_paragraph:section.last_paragraph]:
                section_of[paragraph.path] = k
                if section.line_numbering and resolve_paragraph(source, paragraph).get("suppressLineNumbers"):
                    suppressed.add(paragraph.path)
        for k, section in enumerate(sections):
            if section.line_numbering and section.column_count > 1:
                self.layout.warn("line-numbers-not-drawn", "a section of several text columns numbers its lines "
                                 "(w:lnNumType), which is not drawn there (not measured)")
        style = "LineNumber" if source.styles is not None and "LineNumber" in source.styles.styles else None
        run = Run(text="0", properties=RunProperties(style_id=style))
        resolved = resolve_run(source, Paragraph(runs=(run,)), run)
        count = None
        last_section = None
        for page in self.layout.pages:
            if page.info is None or page.info.blank:
                continue
            first_on_page = True
            lines = sorted((line for line in page.lines if line.story is None and line.path in section_of),
                           key=lambda line: line.baseline)
            for line in lines:
                k = section_of[line.path]
                numbering = sections[k].line_numbering
                if numbering is None or sections[k].column_count > 1 or line.path in suppressed:
                    continue
                count_by, start, distance, restart = numbering
                if count is None or (restart == "newPage" and first_on_page) or (
                        restart == "newSection" and k != last_section):
                    count = start
                count += 1
                first_on_page, last_section = False, k
                if count % count_by:
                    continue
                span = self._line_number(page, line, k, str(count), 360 if distance is None else distance,
                                         resolved, source)
                if span is not None:
                    line.spans.insert(0, span)

    def _line_number(self, page: Page, line: Line, section: int, text: str, distance: int, resolved,
                     source: Document) -> Span | None:
        from .resolve import character_format

        fmt = character_format(resolved, text[0], source)
        units = []
        for char in text:
            advance = self.advances.advance(fmt.face, fmt.bold, fmt.italic, char) if fmt.face else None
            if advance is None:
                self.layout.warn("line-numbers-not-drawn", "a line number whose face has no advances is not drawn")
                return None
            units.append(linebreak.note_number_units(advance, fmt.half_points) * PX_PER_PT / 4096)
        end = round_half_up(page.column[0] - twips_to_px(distance))
        x = Fraction(end) - sum(units, Fraction(0))
        xs = []
        for width in units:
            xs.append(x)
            x += width
        return Span(f"w:sectPr[{section + 1}]/w:lnNumType", list(text), xs, line.baseline, x, fmt.face, fmt.bold,
                    fmt.italic, fmt.half_points, _color(resolved.get("color")), kind="line-number")

    # -- fields ------------------------------------------------------------------------------

    def _cross_references(self, document: Document) -> Document:
        """The body with its ``SEQ`` fields counted and its ``REF`` fields' runs replaced
        (:func:`docx2svg.fields.cross_references`); a warning for each drawn as cached."""
        from .fields import cross_references

        if not any(run.field is not None and run.field.keyword in ("SEQ", "REF")
                   for paragraph in _all_paragraphs(document.body) for run in paragraph.runs):
            return document
        body, problems = cross_references(document.body)
        for keyword, why in problems:
            self.layout.warn(f"field-not-computed:{keyword}",
                             f"a {keyword} field in the body is drawn as Word cached it ({why})")
        return dataclasses.replace(document, body=body, paragraphs=[b for b in body if isinstance(b, Paragraph)])

    def _bookmark_pages(self, items, result, details, drawn) -> dict:
        """Bookmark name -> the :class:`paginate.PageInfo` of the page its start is on: the
        page of the line the start stands in (``make_xref_field_probe.py``, ``late``)."""
        from .fields import bookmarks

        paragraphs = _all_paragraphs(self.document.body)
        out = {}
        for name, ((block, run), _end) in bookmarks(paragraphs).items():
            page = self._page_of(items, result, details, drawn, block,
                                 sum(len(r.text) for r in paragraphs[block].runs[:run]))
            if page is not None and result.info.get(page) is not None:
                out[name] = result.info[page]
        return out

    def _page_of(self, items, result, details, drawn, block: int, offset: int) -> int | None:
        """The page the body paragraph ``block``'s character ``offset`` is drawn on."""
        index = self._flow_of.get(block)
        if index is not None and index in result.pages and index in details:
            info = details[index]
            line_of = paginate._line_of_sources(info.pieces, info.lines)
            line = line_of.get(offset, line_of.get(offset - 1, 0))
            found = result.pages[index]
            return found[min(line, len(found) - 1)] if found else None
        if drawn.get(block):
            return drawn[block][0]
        return None

    def _field_text(self, run, info: paginate.PageInfo | None, where: str) -> str | None:
        """A computed field's text on the page ``info`` describes, from this layout's own
        pagination; ``None`` (the cached result is drawn) where it cannot be computed,
        with a warning."""
        from .fields import Uncomputable, field_text

        result = self.result_for_fields
        if info is None:
            return None
        try:
            return field_text(run.field.instruction, run.field.keyword, page=info.number,
                              pages=self._page_count(result), section_pages=self._section_pages(result, info.section),
                              section_format=self.document.sections[info.section].page_number_format)
        except Uncomputable as why:
            self.layout.warn(f"field-not-computed:{run.field.keyword}",
                             f"a {run.field.keyword} field {where} is drawn as Word cached it ({why.args[0]})")
            return None

    def _reference_text(self, run, targets: dict) -> str | None:
        """A ``PAGEREF``'s text: its bookmark's page as ``PAGE`` shows it there -- its
        number in its own section's format (``make_xref_field_probe.py``); ``None``, with a
        warning, where the bookmark is not laid out or the field is not computable."""
        from .fields import Uncomputable, _argument, field_text

        name = _argument(run.field.instruction)
        target = targets.get(name) if name else None
        if target is None:
            self.layout.warn("field-not-computed:PAGEREF",
                             "a PAGEREF field in the body is drawn as Word cached it (its bookmark is not in the "
                             "body, or not laid out)")
            return None
        try:
            return field_text(run.field.instruction, "PAGEREF", page=target.number, pages=None, section_pages=None,
                              section_format=self.document.sections[target.section].page_number_format)
        except Uncomputable as why:
            self.layout.warn("field-not-computed:PAGEREF",
                             f"a PAGEREF field in the body is drawn as Word cached it ({why.args[0]})")
            return None

    @staticmethod
    def _page_count(result) -> int | None:
        """``NUMPAGES``: every page, blank ones included; unknown past a stop."""
        if result.stopped is not None:
            return None
        return max(result.info) + 1 if result.info else 1

    def _section_pages(self, result, section: int) -> int | None:
        """``SECTIONPAGES``: the pages a section has content on (``make_story_select_probe.py``,
        ``blank``: a section sharing its one page with two continuous ones counts 1, one
        going on to the next page 2; a blank page before it is not its)."""
        items = self.items
        if result.stopped is not None and section >= items[result.stopped].section:
            return None
        pages = {page for index, found in result.pages.items() if items[index].section == section for page in found}
        return len(pages)

    def _body_field_values(self, items, result) -> dict:
        """``(block, run) -> text`` for every computed field of the body, on the page its
        line is on (``baselines.blocks`` order, table cells included)."""
        self.result_for_fields = result
        keys = [(block, k) for block, paragraph in enumerate(_all_paragraphs(self.document.body))
                for k, run in enumerate(paragraph.runs) if run.field is not None
                and run.field.keyword in PAGE_FIELDS]
        if not keys:
            return {}
        drawn = paginate.drawn_pages(items, result)
        details = self._details(items)
        self._flow_of = {item.block: index for index, item in enumerate(items) if isinstance(item, paginate.Para)}
        paragraphs = _all_paragraphs(self.document.body)
        targets = None
        out = {}
        for block, k in keys:
            run = paragraphs[block].runs[k]
            if run.field.keyword == "PAGEREF":
                if targets is None:
                    targets = self._bookmark_pages(items, result, details, drawn)
                text = self._reference_text(run, targets)
                if text is not None:
                    out[(block, k)] = text
                continue
            page = self._page_of(items, result, details, drawn, block,
                                 sum(len(r.text) for r in paragraphs[block].runs[:k]))
            if page is None:
                continue
            text = self._field_text(paragraphs[block].runs[k], result.info.get(page), "in the body")
            if text is not None:
                out[(block, k)] = text
        return out

    # -- headers and footers -------------------------------------------------------------------

    def _draw_stories(self, page: Page, info) -> None:
        """The page's header and footer (:func:`paginate.story_kind`), each laid out as a
        document of its own from where Word puts it: a header's first line's box from
        ``w:pgMar/@w:header`` below the page's top, a footer's last from ``@w:footer``
        above its foot, growing upward (``make_story_probe.py``)."""
        page.info = info
        if info is None or info.blank:
            return
        headers, footers = self.references[info.section]
        section = self.document.sections[info.section]
        for kind, references in (("header", headers), ("footer", footers)):
            relationship = references.get(info.story)
            blocks = self.document.stories.get(relationship) if relationship else None
            if blocks:
                self._draw_story(page, kind, relationship, blocks, section, info)

    def _draw_story(self, page: Page, kind: str, relationship: str, blocks: list, section, info) -> None:
        part = self.document.story_parts.get(relationship, relationship)
        where = f"in {part}"
        texts: dict = {}

        def text_of(block: int, k: int, run) -> str | None:
            text = self._field_text(run, info, where)
            if text is not None:
                texts[(block, k)] = text
            return text

        substituted = substitute_blocks(blocks, text_of)
        cache = self.__dict__.setdefault("story_cache", {})
        key = (relationship, info.section, tuple(sorted(texts.items())))
        if key not in cache:
            # A floating drawing text does not wrap around is drawn with the story, from
            # where it is anchored (ROADMAP.md, "Floating drawings -- measured", F.6).
            clean, floating = paginate.floating_marks(substituted, keep=True)
            story = paginate.story_document(self.document, clean, section)
            items = paginate.flow(story, self.advances, self.metrics)
            placer = _Placer(story, None, self.advances, self.metrics, self.decorations)
            placer.layout = self.layout
            placer.items = items
            placer.references = []
            # A chart or a diagram in the story is related from the story's own part.
            placer.frames = self.frames
            placer.story_part = part
            details = placer._details(items)
            cache[key] = (placer, items, details, paginate.stack_height(story, items), floating)
        placer, items, details, height, floating = cache[key]
        obstacle = next((item for item in items if isinstance(item, paginate.Obstacle)), None)
        if obstacle is not None:
            self.layout.warn(f"story-stopped:{obstacle.reason}",
                             f"{part} is drawn up to {placer._obstacle_path(obstacle)} ({placer._why(obstacle)}); "
                             "the rest of it is not drawn", page.number + 1)
        if kind == "header":
            y0 = twips_to_px(section.margins.header)
        else:
            y0 = twips_to_px(section.page_size.height_twips - section.margins.footer) - height
        lines, pictures, floats = len(page.lines), len(page.pictures), len(page.floats)
        placer._walk(page, items, details, paginate.Position(0), paginate.Position(len(items)), y0,
                     bottom=Fraction(10 ** 7))
        for line in page.lines[lines:]:
            line.story, line.part = kind, part
        for picture in page.pictures[pictures:]:
            picture.part = part
        for placed in page.floats[floats:]:
            placed.story, placed.part = kind, part
            for line in placed.lines:
                line.story, line.part = kind, part

    def _notes(self, items) -> None:
        for section in self.document.sections:
            if section.vertical_alignment in ("center", "both", "bottom"):
                self.layout.warn("vertical-alignment-not-drawn",
                                 f"a section's body is aligned {section.vertical_alignment} between its margins "
                                 "(w:vAlign), which is not laid out: its body is drawn from the top margin; its "
                                 "header and footer, which it does not move, are where Word draws them")
        for keyword in sorted(set(self.document.cached_fields)):
            if keyword in RECOMPUTED_BY_WORD:
                self.layout.warn(f"field-cached:{keyword}",
                                 f"a {keyword} field is drawn as Word cached it in the file; Word computes it "
                                 "again when it exports, so the two can differ")

    # -- footnotes ------------------------------------------------------------------------------

    def _footnotes(self, page: Page, index: int) -> None:
        """The footnotes at the foot of page ``index`` (:attr:`paginate.Pagination.notes`):
        the separator's paragraph -- the continuation separator's on a page a note goes on
        to (:func:`notes.footnote_separator_kind`) -- and under it the notes' lines, the
        rest of a note from the page before first, laid out as a story's paragraphs are
        and stacked as the body stacks them, **the stack's foot on the text area's foot**
        (``make_footnote_draw_probe.py``: a page of short text keeps its notes at the
        bottom margin); under the section's ``w:pos`` ``beneathText``, its top where the
        body's stack ends."""
        frames = self.result.frames.get(index) if getattr(self, "result", None) is not None else None
        if frames and any(frame.multi for frame in frames) and page.info is not None:
            self._column_footnotes(page, frames)
            return
        entry = self.result.notes.get(index)
        if not entry or page.info is None:
            return
        carried, placed = entry
        self._note_stack(page, carried, placed, self.document.sections[page.info.section])

    def _note_stack(self, page: Page, carried: tuple, placed: list, section, *, clip: Fraction | None = None,
                    separator_only: bool = False) -> None:
        """The separator and the notes ``placed`` (the rest of a note going on, ``carried``,
        first) stacked at the foot of ``page.column``, laid out in ``section``'s text
        width (:meth:`_footnotes`); ``clip``, the right edge a separator's line ends at
        if it would reach past it."""
        from .notes import footnote_separator_kind, separator

        parts = []
        for k, line in enumerate(carried):
            # The rest of each note going on (more than one where a column's foot left a
            # line's later notes to the next).
            if k == 0 or line.note != carried[k - 1].note:
                parts.append((line.note, line, None))
        for lines, taken in placed:
            if taken and lines:
                parts.append((lines[0].note, None, lines[taken] if taken < len(lines) else None))
        if not parts and not separator_only:
            return
        document = self.document
        if placed and placed[-1][1] < len(placed[-1][0]) and any(
                run.text.strip() for block in document.footnote_separators.get("continuationNotice", [])
                for run in block.runs):
            self.layout.warn("footnote-continuation-notice-not-drawn", "the continuation notice under a footnote "
                             "that goes on to the next page is not drawn", page.number + 1)
        kind = footnote_separator_kind(document, "continuationSeparator" if carried else "separator")
        blocks = [block for block in separator(document, kind, "footnote") if isinstance(block, Paragraph)]
        separators = len(blocks)
        first_of = []
        for note_id, first, _ in parts:
            paragraphs = [p for p in document.footnotes.get(note_id, []) if isinstance(p, Paragraph)]
            skip = first.paragraph if first is not None else 0
            first_of.append(len(blocks) - skip)
            blocks.extend(paragraphs[skip:])
        story = paginate.story_document(document, blocks, section)
        items = paginate.flow(story, self.advances, self.metrics)
        if any(isinstance(item, paginate.Obstacle) for item in items):
            self.layout.warn("footnotes-not-drawn", "a footnote that cannot be laid out is not drawn",
                             page.number + 1)
            return
        placer = _Placer(story, None, self.advances, self.metrics, self.decorations)
        placer.layout = self.layout
        placer.frames = self.frames
        placer.items = items
        placer.references = []
        placer.separator_clip = clip
        details = placer._details(items)
        end = paginate.Position(len(items))
        last = parts[-1][2] if parts else None
        if last is not None:
            end = paginate.Position(first_of[-1] + last.paragraph, last.line)
        walks = [(paginate.Position(0), end)]
        if carried and carried[0].line:
            # The rest of a paragraph: the separator, then its lines from where they go on.
            walks = [(paginate.Position(0), paginate.Position(separators)),
                     (paginate.Position(separators, carried[0].line), end)]

        def stack(target: Page, y: Fraction) -> Fraction:
            for start, stop in walks:
                y, _ = placer._walk(target, items, details, start, stop, y, bottom=Fraction(10 ** 7))
            return y

        if section.footnote_position == "beneathText" and page.body_end is not None and clip is None:
            # Straight under the body's last line (``sectbeneath``); the document's own
            # ``w:pos`` is not taken (``beneath``).
            stack(page, page.body_end)
            return
        scratch = dataclasses.replace(page, lines=[], rules=[], pictures=[], floats=[], defs=None)
        height = stack(scratch, Fraction(0))
        if last is None and items and isinstance(items[-1], paginate.Para):
            # The last note's space after is in the stack (``make_footnote_draw_probe.py``,
            # ``after``: 160 and 240 twips put the notes 33 and 50 px higher, in every
            # setting); not where the note goes on to the next page.
            height += twips_to_px(items[-1].after)
        lines_before, rules_before = len(page.lines), len(page.rules)
        stack(page, page.column[3] - height)
        first_note = items[separators] if not carried and len(items) > separators else None
        before = first_note.before if isinstance(first_note, paginate.Para) else 0
        if before:
            # The first note's space before stands over the separator, not under it
            # (``make_footnote_draw_probe.py``, ``after``: 120 and 240 twips before put the
            # separator 25 and 50 px lower, its note where it was).
            shift = round_half_up(twips_to_px(before))
            paths = tuple(block.path for block in blocks[:separators])
            for line in page.lines[lines_before:]:
                if line.path.startswith(paths):
                    line.baseline += shift
                    line.top += shift
                    for span in line.spans:
                        span.y += shift
                    if line.band is not None:
                        line.band = (line.band[0] + shift, line.band[1] + shift)
            for rule in page.rules[rules_before:]:
                if rule.kind == "separator":
                    rule.y += shift

    def _column_footnotes(self, page: Page, frames: list) -> None:
        """The footnotes of a page of text columns (``make_columns_probe.py``, ``footnote``,
        ``notesright``, ``notes3``, ``notesunequal``, ``notesover``):

        * **below mode 15, each column's at its own foot**, stacked as a page's are
          (:meth:`_note_stack`) between the column's edges and broken in its width, under
          its own separator -- **its line cut at the column's right edge** (2,880 twips
          drawn 581 px in a column of 2,787, 458 in one of 2,200) -- or, where the notes go
          on from the column before, under the continuation separator, to the column's
          right edge;
        * **in mode 15, the page's note area** (:func:`paginate.note_area`): the notes run
          on through the columns, balanced, every column's top level, their foot on the
          text area's foot, under one separator at the first column's left, **2,880 twips
          whatever the column's width**."""
        document = self.document
        saved = page.column
        mode15 = (document.compatibility_mode or 0) >= 15
        try:
            if not mode15:
                for frame in frames:
                    if not frame.multi or not (frame.notes or frame.carried):
                        continue
                    section = document.sections[frame.section]
                    left, right = self.column_edges(section, frame.offset, frame.width)
                    page.column = (left, right, saved[2], saved[3])
                    placed = len(page.lines)
                    self._note_stack(page, frame.carried, frame.notes, paginate.column_section(section, frame.width),
                                     clip=right)
                    _in_column(page.lines[placed:], frame)
                return
            pooled = [frame for frame in frames if frame.pooled]
            ids = tuple(lines[0].note for frame in pooled for lines, _ in frame.notes)
            if not ids:
                return
            area = paginate.note_area(document, self.items, pooled[0].section, ids)
            if area is None:
                self.layout.warn("footnotes-not-drawn", "a footnote that cannot be laid out is not drawn",
                                 page.number + 1)
                return
            section = document.sections[pooled[0].section]
            top = saved[3] - area.height
            # The separator, over the first column.
            left, right = self.column_edges(section, 0, linebreak.column_width_twips(section))
            page.column = (left, right, saved[2], top)
            placed = len(page.lines)
            self._note_stack(page, (), [], paginate.column_section(section, linebreak.column_width_twips(section)),
                             separator_only=True)
            for line in page.lines[placed:]:
                line.column = 0
            placer = _Placer(area.story, None, self.advances, self.metrics, self.decorations)
            placer.layout = self.layout
            placer.frames = self.frames
            placer.items = area.items
            placer.references = []
            details = placer._details(area.items)
            for frame in area.frames:
                left, right = self.column_edges(section, frame.offset, frame.width)
                page.column = (left, right, saved[2], saved[3])
                placed = len(page.lines)
                placer._walk(page, area.items, details, frame.start, frame.end, top + frame.top,
                             bottom=Fraction(10 ** 7), first_on_page=frame.first_on_page, previous=frame.previous)
                _in_column(page.lines[placed:], frame)
        finally:
            page.column = saved

    def _details(self, items) -> dict[int, _Paragraph]:
        """Flow index -> the paragraph's pieces and lines, broken again exactly as the flow
        broke them (the same breaker, the same label)."""
        document = self.document
        counters = linebreak.ListCounters()
        labels: dict[int, str | None] = {}
        block = 0

        def walk(item):
            nonlocal block
            if isinstance(item, Table):
                for row in item.rows:
                    for cell in row:
                        for inner in cell:
                            walk(inner)
                return
            pp = resolve_paragraph(document, item)
            if pp.get("numPr.numId"):
                labels[block] = counters.label(document, item, pp)
            block += 1

        for item in document.body:
            walk(item)
        out: dict[int, _Paragraph] = {}
        by_block = {}
        index = 0

        def collect(item):
            nonlocal index
            if isinstance(item, Table):
                for row in item.rows:
                    for cell in row:
                        for inner in cell:
                            collect(inner)
                return
            by_block[index] = item
            index += 1

        for item in document.body:
            collect(item)
        # Whether each flow paragraph's neighbours in the body share its border box (a
        # table between them parts them), as ``paginate.flow`` decides its border heights.
        body = list(self.document.body)
        resolved_body = [resolve_paragraph(document, b) if not isinstance(b, Table) else None for b in body]
        position_of = {id(b): k for k, b in enumerate(body)}
        self.box_neighbours: dict[int, tuple[bool, bool]] = {}
        for flow_index, item in enumerate(items):
            if not isinstance(item, paginate.Para):
                continue
            paragraph = by_block[item.block]
            k = position_of.get(id(paragraph))
            if k is not None and resolved_body[k] is not None:
                box = vertical.border_box(resolved_body[k])
                drawn = any(vertical._drawn(side) for side in box[0])
                before = resolved_body[k - 1] if k > 0 else None
                after = resolved_body[k + 1] if k + 1 < len(body) else None
                self.box_neighbours[flow_index] = (
                    drawn and before is not None and vertical.border_box(before) == box,
                    drawn and after is not None and vertical.border_box(after) == box)
            section = document.sections[item.section]
            pp = resolve_paragraph(document, paragraph)
            pieces, lines = linebreak.break_paragraph(document, paragraph, section, self.advances, self.metrics,
                                                      labels.get(item.block))
            out[flow_index] = _Paragraph(paragraph, pp, pieces, lines,
                                         linebreak.geometry(document, paragraph, section, pp), labels.get(item.block))
            if item.source is not None:
                # Beside a drawing text wraps around, the lines as the paginator placed them.
                out[flow_index] = _Paragraph(paragraph, pp, item.source.pieces, list(item.lines),
                                             item.source.geometry, labels.get(item.block))
        return out

    # -- one page ---------------------------------------------------------------------------

    def _new_page(self, section_number: int, number: int, story: str | None) -> Page:
        section = self.document.sections[section_number]
        width_pt = device_page_extent_pt(section.page_size.width_twips)
        height_pt = device_page_extent_pt(section.page_size.height_twips)
        top, bottom = paginate._geometry(self.document, section_number, self.items, story=story or "default")
        # The text column starts at the left margin (and gutter) rounded to a whole device
        # pixel -- every line of the page, the body's and its header's and footer's
        # (``make_margin_probe.py``: margins 1440-1449 twips and gutters, every line's
        # start; ROADMAP.md, "Headers, footers and fields -- measured", H.5) -- as a table's
        # grid lines already hang from it (:func:`grid_line_px`).  The right edge moves
        # with it: a line's width is the column's, in twips.  The pixel is held in the
        # nearest whole layout unit: 120 px is 117,964.8 units, drawn at 117,965 -- 120.0002
        # px (``make_margin_probe.py``, margins of 576, 1152 and 1296 twips; the floor, first
        # adopted, agrees wherever the fraction is below a half, as at 1440 twips).
        exact = twips_to_px(section.margins.left + section.margins.gutter)
        left = round_half_up(round_half_up(exact) / LAYOUT_UNIT_PX) * LAYOUT_UNIT_PX
        right = twips_to_px(section.page_size.width_twips - section.margins.right) + (left - exact)
        return Page(number, round(width_pt * 300 / 72), round(height_pt * 300 / 72), width_pt, height_pt,
                    (left, right, top, bottom))

    def _blank(self, start, number: int, info=None) -> Page:
        item = self.items[start.item] if start.item < len(self.items) else None
        section = item.section if item is not None else len(self.document.sections) - 1
        page = self._new_page(section, number, "default")
        page.info = info
        return page

    def _page(self, items, details, start, end, number: int, stopped: int | None, stopped_line: int = 0,
              info=None) -> Page:
        document = self.document
        first_item = items[start.item] if start.item < len(items) else items[stopped]
        section_first = start.line == 0 and (start.item == 0 or items[start.item - 1].section != first_item.section)
        story = info.story if info is not None else ("first" if section_first else "default")
        page = self._new_page(first_item.section, number, story)
        y = page.column[2]
        if isinstance(first_item, paginate.Para) and first_item.continuation is not None and start.item < len(items):
            y = self._continuation(page, first_item, y)
        frames = self.result.frames.get(number) if getattr(self, "result", None) is not None else None
        if frames:
            y, previous = self._walk_columns(page, items, details, frames)
        else:
            y, previous = self._walk(page, items, details, start, end, y)
        page.body_end = y
        if stopped is not None and isinstance(items[stopped], paginate.TableItem) and items[stopped].floating:
            path = items[stopped].flow.table.path
            page.stop = Stop(path, "table", y, sum(1 for _ in items[stopped:]))
            self.layout.warn("layout-stopped:table",
                             f"the layout stops at {path} (a floating table that reaches below the bottom margin, "
                             f"which Word splits across pages in mode 15: not modelled); it and the "
                             f"{page.stop.remaining - 1} top-level item(s) after it are not drawn, and no page after "
                             f"this one is made", number + 1)
        elif stopped is not None and isinstance(items[stopped], paginate.TableItem):
            table_item = items[stopped]
            if stopped_line == 0 and previous is not None:
                y += paginate.table_gap_px(previous, table_item)
            path = f"{table_item.flow.table.path}/w:tr[{stopped_line + 1}]"
            page.stop = Stop(path, "table", y, table_item.count - stopped_line + sum(1 for _ in items[stopped + 1:]))
            self.layout.warn("layout-stopped:table",
                             f"the layout stops at {path} (a table row that cannot be placed on an empty page: a "
                             f"line taller than the page); it and the "
                             f"{page.stop.remaining - 1} row(s) and top-level item(s) after it are not drawn, and no "
                             f"page after this one is made",
                             number + 1)
        elif stopped is not None:
            obstacle = items[stopped]
            # Where the obstacle's own paragraph would start: past its space before, as
            # between any two paragraphs (a table's is not known: the band starts at y).
            start = y
            if obstacle.paragraph is not None:
                before = resolve_paragraph(document, obstacle.paragraph).get("spacing.before", 0) or 0
                after = previous.after if previous is not None else 0
                start = y + (vertical.paragraph_gap_px(after, before)[0] if previous is not None
                             else twips_to_px(before))
            page.stop = Stop(self._obstacle_path(obstacle), obstacle.reason, y,
                             sum(1 for _ in items[stopped:]), self._anchor_extent(page, obstacle, start))
            self.layout.warn(f"layout-stopped:{obstacle.reason}",
                             f"the layout stops at {page.stop.path} ({self._why(obstacle)}); it and the "
                             f"{page.stop.remaining - 1} top-level item(s) after it are not drawn, and no page "
                             f"after this one is made", number + 1)
        return page

    def column_edges(self, section, offset: int, width: int) -> tuple[Fraction, Fraction]:
        """A text column's left and right edges, device px (:func:`_new_page`'s for one
        column): the left at the margin, gutter and the columns before it, in twips,
        **rounded to a whole pixel** -- below mode 15 from the exact length, in mode 15
        from the length held in whole layout units (``make_columns_probe.py``,
        ``geometry``: a column 3,324 twips in, 692.5 px exactly, is drawn at 693 below mode
        15 and at 692 in it, 680,755.2 units held as 680,755) -- and held in the nearest
        whole unit; the right edge its width further."""
        total = section.margins.left + section.margins.gutter + offset
        if (self.document.compatibility_mode or 0) >= 15:
            px = round_half_up(units_px(linebreak.twips_to_units(total)))
        else:
            px = round_half_up(Fraction(total * 300, 1440))
        left = round_half_up(px / LAYOUT_UNIT_PX) * LAYOUT_UNIT_PX
        return left, left + twips_to_px(width)

    def _walk_columns(self, page: Page, items, details, frames: list):
        """The page's columns (:class:`paginate.Frame`), each walked as a page is, from its
        top and between its column's edges; then the lines between them
        (:meth:`_column_separators`).  Returns where the last column's stack ends and its
        last item (where a stop band starts)."""
        y, previous = page.column[2], None
        saved = page.column
        for frame in frames:
            section = self.document.sections[frame.section]
            left, right = self.column_edges(section, frame.offset, frame.width) if frame.multi else saved[:2]
            page.column = (left, right, saved[2], saved[3])
            # A drawing anchored in the column is positioned against it.
            self._column_box = (frame.offset, frame.width) if frame.multi else None
            placed = len(page.lines)
            try:
                y, previous = self._walk(page, items, details, frame.start, frame.end, frame.top, bottom=frame.bottom,
                                         first_on_page=frame.first_on_page, previous=frame.previous)
            finally:
                page.column = saved
                self._column_box = None
            _in_column(page.lines[placed:], frame)
        self._column_separators(page, frames)
        return y, previous

    def _column_separators(self, page: Page, frames: list) -> None:
        """``w:sep``: a line between each two columns of a region where the second holds
        text, **centred on the space between them, exactly** (the columns' unrounded
        edges), 0.75 pt wide, black, from the region's top to the foot of its longest
        column, its edges rounded to whole pixels (``make_columns_probe.py``, ``sep`` and
        ``balance``: 1253-1257 px between two columns 720 twips apart, 921-924 between
        three 401 apart, 809-812 after a column 2,000 wide and 900 before the next; from
        313 to 3166 on a page whose first column is full, 368 to 648 over five balanced
        lines; none on a page whose second column is empty)."""
        regions: list = []
        for frame in frames:
            if regions and regions[-1][0].section == frame.section and regions[-1][0].region_top == frame.region_top \
                    and frame.column > regions[-1][-1].column:
                regions[-1].append(frame)
            else:
                regions.append([frame])
        width = Fraction(3, 4) * PX_PER_PT
        for region in regions:
            section = self.document.sections[region[0].section]
            if not region[0].multi or not section.column_separator:
                continue
            filled = [frame for frame in region if frame.end > frame.start]
            if not filled:
                continue
            top = round_half_up(region[0].region_top)
            bottom = round_half_up(max(frame.y for frame in filled))
            boxes = linebreak.column_boxes(section)
            margin = section.margins.left + section.margins.gutter
            for frame in filled:
                if frame.column == 0:
                    continue
                offset, width_before, space = boxes[frame.column - 1]
                centre = twips_to_px(margin + offset + width_before + Fraction(space, 2))
                x0, x1 = round_half_up(centre - width / 2), round_half_up(centre + width / 2)
                page.rules.append(Rule("column-separator", f"w:sectPr[{region[0].section + 1}]/w:cols", Fraction(x0),
                                       Fraction(top), Fraction(x1 - x0), Fraction(bottom - top), "000000"))

    def _walk(self, page: Page, items, details, start, end, y: Fraction, *, bottom: Fraction | None = None,
              story_bottom: Fraction | None = None, first_on_page: bool = True, previous=None):
        """Place the flow's lines from ``start`` to ``end`` on ``page``, stacked from ``y``
        as the paginator stacks them; returns where the stack ends and its last item.
        ``bottom`` is where a table may reach (the page's text area, or anywhere for a
        story).  ``story_bottom``, for a text box, is the foot of its room: the last line
        of its last paragraph keeps everything down to it below its text, as space after
        in its line box, which moves its baseline as :func:`~docx2svg.vertical.baseline_in_box`
        rounds such a box (``make_text_box_top_probe.py``: 14 and 15 pt lines a pixel
        lower at the top of a box, 16 pt by the room's fraction of a pixel)."""
        document = self.document
        bottom = page.column[3] if bottom is None else bottom
        floated = False
        index, line_number = start.item, start.line
        while paginate.Position(index, line_number) < end and index < len(items):
            item = items[index]
            if isinstance(item, paginate.Obstacle):
                break
            if isinstance(item, paginate.TableItem) and item.floating is not None:
                # Drawn where the paginator positioned it (``paginate.table_wrap``); it
                # takes no room in the stack.
                from . import table as table_model

                wrap = next((w for w in getattr(self, "_wraps", ()) if w.key == ("table", item.block)), None)
                if wrap is not None:
                    laid = table_model.place_table(item.flow, wrap.y, wrap.y + Fraction(10 ** 7), first_on_page=False,
                                                   mode15=(document.compatibility_mode or 0) >= 15)
                    self._table(page, item, laid)
                floated = True
                index, line_number = index + 1, 0
                continue
            if isinstance(item, paginate.TableItem):
                from . import table as table_model

                if line_number == 0 and not first_on_page:
                    y += paginate.table_gap_px(previous, item)
                part = start.part if index == start.item else ()
                try:
                    laid = paginate.place_item(item, y, bottom, key=start, row=line_number, part=part,
                                               first_on_page=first_on_page,
                                               mode15=(document.compatibility_mode or 0) >= 15)
                except table_model.Overflow as overflow:
                    laid = overflow.args[1]
                if end.item == index:
                    # The paginator ends the page inside the table: draw what it placed -- a
                    # row it split where it split it, which a cell's footnotes can make
                    # sooner than the page's foot alone would (``paginate._table_notes``).
                    kept = []
                    for piece in laid.pieces:
                        if not piece.header and piece.row == end.line and piece.split and end.part:
                            kept.append(table_model.cut_piece(piece.flow or item.flow, piece, end.part)
                                        if piece.end != end.part else piece)
                        elif piece.header or (piece.row, piece.start if piece.continued else ()) < (
                                end.line, end.part) or (piece.row == end.line and piece.split):
                            kept.append(piece)
                    laid.pieces = kept
                self._table(page, item, laid)
                if any(band.top < laid.end and band.bottom > y for band in getattr(self, "_bands", ())):
                    self.layout.warn("wrap-table-not-modelled",
                                     "a table that reaches into the band of a drawing text wraps above and below "
                                     "(wrapTopAndBottom) is laid out where it is, not below the drawing", page.number + 1)
                y = laid.end
                first_on_page = first_on_page and not laid.pieces
                previous = item
                index, line_number = index + 1, 0
                continue
            # An empty section break's paragraph takes no room (paginate._scan): it is placed
            # where it stands, and the stack goes on as if it were not there.
            restore = (y, previous, first_on_page) if item.section_mark and not first_on_page and \
                not paginate.section_mark_room(previous, document.compatibility_mode) else None
            info = details[index]
            box_before = Fraction(0)
            band_top = band_bottom = None
            y_gap = y
            starts_here = line_number == 0
            # Where a floating drawing's ``paragraph`` is (:mod:`docx2svg.floating`): the
            # page's top for its first paragraph, else where the paragraph before ends
            # after its own space after.
            self._paragraph_top = y
            for number_ in range(line_number, item.count):
                if paginate.Position(index, number_) >= end:
                    break
                height = item.heights[number_]
                if number_ == 0:
                    if first_on_page:
                        keeps = vertical.keeps_space_before_at_page_top(
                            section_start=item.section_start, page_break_before=item.page_break_before,
                            compatibility_mode=document.compatibility_mode)
                        above = items[index - 1] if index else None
                        gap, box_before = vertical.page_top_gap_px(
                            item.before, keeps=keeps, previous_after=above.after if isinstance(above, paginate.Para)
                            else 0)
                    elif isinstance(previous, paginate.TableItem):
                        gap, box_before = twips_to_px(item.before), twips_to_px(item.before)
                    else:
                        same = item.style == previous.style
                        own_before = 0 if (item.contextual and same) else item.before
                        prev_after = 0 if (previous.contextual and same) else previous.after
                        gap, box_before = vertical.paragraph_gap_px(prev_after, own_before)
                        if floated:
                            # A floating table between them: the two do not collapse
                            # (``paginate._scan``).
                            gap, box_before = twips_to_px(prev_after) + twips_to_px(own_before), twips_to_px(own_before)
                        self._paragraph_top = y + twips_to_px(prev_after)
                    y += gap
                    y_gap = y
                    y += item.top_border
                    box_before += item.top_border
                bands = getattr(self, "_bands", ())
                if bands and getattr(self, "_column_box", None) is not None:
                    # In a column of several, only the bands of drawings across it.
                    bands = [band for band in bands if band.span is None
                             or (band.span[0] < page.column[1] and band.span[1] > page.column[0])]
                if bands:
                    reach = height.pitch + (twips_to_px(item.after) if number_ == item.count - 1
                                            and (document.compatibility_mode or 0) < 15 else 0)
                    y = paginate.clear_of_bands(y, reach, bands)
                drop = getattr(info.lines[number_], "drop", None) if number_ < len(info.lines) else None
                if drop is not None:
                    # A line beside a drawing that no segment took text at: it went down
                    # to the drawing's foot (``paginate.fit_beside``).
                    y = max(y, drop)
                last = number_ == item.count - 1
                space_after = Fraction(0)
                if last:
                    following = items[index + 1] if index + 1 < len(items) else None
                    same_next = isinstance(following, paginate.Para) and following.style == item.style
                    own_after = 0 if (item.contextual and same_next) else item.after
                    space_after = item.border + twips_to_px(own_after)
                    if story_bottom is not None and index == len(items) - 1:
                        space_after = max(space_after, story_bottom - (y + height.pitch))
                box = height.box(box_before if number_ == 0 else Fraction(0), space_after)
                baseline = baseline_in_box(y - box.space_before, box)
                band = text_band(y - box.space_before, box)
                band_top = band[0] if band_top is None else band_top
                band_bottom = band[1]
                if number_ == 0 and getattr(item, "drop_cap", None) is not None:
                    self._drop_cap(page, item, self._paragraph_top)
                if number_ not in item.undrawn:
                    self._line(page, info, item, number_, y, height, baseline, band)
                y += height.pitch
                first_on_page = False
            ends_here = paginate.Position(index, item.count) <= end
            if ends_here:
                y += item.border
            if band_top is not None:
                self._paragraph_rules(page, info, index, band_top, band_bottom, y_gap, y,
                                      starts_here=starts_here, ends_here=ends_here)
            previous = item
            floated = False
            if restore is not None:
                y, previous, first_on_page = restore
            index, line_number = index + 1, 0
        return y, previous

    def _anchor_extent(self, page: Page, obstacle: paginate.Obstacle, y: Fraction):
        """A floating drawing's box where the paragraph starting at the stop (``y``) holds
        one text wraps around and its position does not depend on the line (the page,
        the margins, the column, the paragraph): drawn outlined in the stop band.
        ``None`` otherwise."""
        if obstacle.paragraph is None:
            return None
        anchors = [a for run in obstacle.paragraph.runs for a in run.anchors if a.moves_text]
        if not anchors:
            return None
        anchor = anchors[0]
        if anchor.h.relative == "character" or anchor.v.relative == "line":
            return None
        frames = self._frames(page, obstacle.section, y, y, y, Fraction(0))
        return self._anchor_box(anchor, frames)

    def _frames(self, page: Page, section_number: int, paragraph_top: Fraction, line_top: Fraction,
                line_bottom: Fraction, character: Fraction):
        """What an anchor on ``page`` is positioned against (:class:`floating.Frames`), the
        paragraph's and line's tops and the character's x given in device px."""
        from .floating import Frames

        section = self.document.sections[section_number]
        m = section.margins
        box = getattr(self, "_column_box", None)
        return Frames(section.page_size.width_twips, section.page_size.height_twips, m.left + m.gutter, m.right,
                      m.top, m.bottom, m.header, m.footer, page.number + 1,
                      (self.document.compatibility_mode or 0) >= 15,
                      paragraph_top / TWIP_PX, line_top / TWIP_PX, line_bottom / TWIP_PX, character / TWIP_PX,
                      column=None if box is None else (m.left + m.gutter + box[0], box[1]))

    @staticmethod
    def _anchor_box(anchor, frames):
        """The drawing's box, device px: where :mod:`docx2svg.floating` puts it, at the size
        Word draws it (a picture's :func:`picture_size`, else the extent in whole twips)."""
        from .floating import emu_twips, horizontal, vertical

        x, y = horizontal(anchor, frames), vertical(anchor, frames)
        if x is None or y is None:
            return None
        graphic = anchor.graphic
        if graphic is not None and graphic.kind == "picture":
            width, height = picture_size(*anchor.extent)
        else:
            width, height = (emu_twips(v) * TWIP_PX for v in anchor.extent)
        return (x * TWIP_PX, y * TWIP_PX, width, height)

    def _anchors(self, page: Page, info: "_Paragraph", item: paginate.Para, number: int, top: Fraction, height,
                 positions: dict) -> None:
        """Place the floating drawings anchored on line ``number`` of ``item``: each in the
        line holding the character after its mark (the paragraph's last line when none
        follows), positioned by :mod:`docx2svg.floating` against this page, and drawn by
        :mod:`docx2svg.drawing` at the size Word draws it."""
        pieces, lines = info.pieces, info.lines
        broken = lines[number]
        section = self.document.sections[item.section]
        exact_left = twips_to_px(section.margins.left + section.margins.gutter)
        for source, run_index, k, anchor, here in item.anchors:
            if here != number:
                continue
            character = exact_left + units_px(anchor_character_units(pieces, broken, positions, source, info.geometry,
                                                                     first_line=number == 0, run=run_index))
            line_top = self._paragraph_top if number == 0 else top
            frames = self._frames(page, item.section, self._paragraph_top, line_top, top + height.pitch, character)
            run = info.paragraph.runs[run_index]
            path = f"{info.paragraph.path}/{run.path}/wp:anchor[{k + 1}]" if run.path else info.paragraph.path
            band = next((b for b in list(getattr(self, "_bands", ())) + list(getattr(self, "_wraps", ()))
                         if b.key == (item.block, source, run_index, k)), None)
            self._place_float(page, anchor, frames, path, info.paragraph.path, y=band.y if band else None,
                              x=getattr(band, "x", None))

    def _place_float(self, page: Page, anchor, frames, path: str, paragraph_path: str, story: str | None = None,
                     part: str | None = None, y: Fraction | None = None, x: Fraction | None = None) -> None:
        """Place and draw one floating drawing; ``y``, where given, is its top as the
        paginator positioned it before any line moved (a ``wrapTopAndBottom`` drawing's,
        or one text wraps beside), and ``x`` its left edge (one text wraps beside)."""
        box = self._anchor_box(anchor, frames)
        if box is not None and y is not None:
            box = (box[0] if x is None else x, y, box[2], box[3])
        if box is None:
            self.layout.warn("drawing-position-unknown",
                             f"a floating drawing ({path}) is positioned against {anchor.h.relative!r} / "
                             f"{anchor.v.relative!r}, which is not modelled: it is not drawn", page.number + 1)
            return
        x, y, width, height = box
        counter = self.__dict__.setdefault("_float_count", [0])
        counter[0] += 1
        placed = Float(path, x, y, width, height, anchor.behind, (anchor.relative_height, counter[0]),
                       story=story, part=part)
        if not anchor.allow_overlap:
            self.layout.warn("drawing-overlap-not-modelled",
                             "a floating drawing that may not overlap another (allowOverlap=\"0\") is drawn where it "
                             "is positioned; Word moves it clear of the drawings before it", page.number + 1)
        self._draw_graphic(page, placed, anchor.graphic, path, paragraph_path)

    def _draw_graphic(self, page: Page, placed: Float, graphic, path: str, paragraph_path: str) -> None:
        """Draw ``graphic`` in ``placed``'s box, lay out its text boxes, and add it to the page."""
        from . import drawing

        x, y, width, height = placed.x, placed.y, placed.width, placed.height
        if page.defs is None:
            from ooxml_common.drawingml.svg import Defs

            page.defs = Defs()
        drawn = drawing.draw(graphic, x, y, width, height, path, theme=self.document.theme_colors,
                             color_map=self.document.color_map, formats=self.document.theme_formats,
                             defs=page.defs, frames=self.frames,
                             part=placed.part or getattr(self, "story_part", None))
        for primitive in drawn.primitives:
            if primitive.path != path:
                primitive.path = f"{paragraph_path}/{primitive.path}"
        placed.primitives = drawn.primitives
        for code in dict.fromkeys(drawn.warnings):
            kind, _, what = code.partition(":")
            self.layout.warn(kind, f"a floating drawing's {what or 'content'} is not drawn faithfully ({path})",
                             page.number + 1)
        for box_graphic, bx, by, bw, bh in drawn.text_boxes:
            self._text_box(page, placed, box_graphic, bx, by, bw, bh, path, paragraph_path)
        page.floats.append(placed)

    def _text_box(self, page: Page, placed: Float, graphic, x: Fraction, y: Fraction, width: Fraction,
                  height: Fraction, path: str, paragraph_path: str) -> None:
        """A text box's content (``w:txbxContent``) laid out as a document of its own at the
        box's inner width -- the shape's box less ``wps:bodyPr``'s insets -- by the
        machinery a page's body uses, and placed at the inner box's top, middle or bottom
        (``anchor``); its lines and rules go with the drawing (:attr:`Float.lines`)."""
        import dataclasses

        from .floating import EMU_PER_TWIP

        body = graphic.body
        insets = [Fraction(body.get(key, default), EMU_PER_TWIP) for key, default in
                  (("lIns", 91440), ("tIns", 45720), ("rIns", 91440), ("bIns", 45720))]
        if body.get("vert", "horz") not in ("horz",) or body.get("rot", "0") not in ("0",):
            self.layout.warn("drawing-not-drawn", f"a text box's vertical or rotated text is not laid out ({path})",
                             page.number + 1)
            return
        # The insets count from the geometry's text rectangle, itself inside half the
        # outline's width (``drawing.text_area_edges``, measured).
        from . import drawing

        colours = drawing.Colours(self.document.theme_colors, self.document.color_map, self.document.theme_formats)
        edges = drawing.text_area_edges(graphic, colours, width / drawing.EMU_PX, height / drawing.EMU_PX)
        # Word holds a text box's text area in whole twips: each side's inset plus its
        # distance in from the shape's edge, rounded to the nearest twip, a half up
        # (``make_text_box_top_probe.py``: ``ins``, ``line`` and ``edge`` sweeps by a fifth
        # of a twip; a custom text rectangle on a half twip goes up).
        insets = [round_half_up(inset + edge / EMU_PER_TWIP) for inset, edge in zip(insets, edges)]
        inner_width = width / TWIP_PX - insets[0] - insets[2]
        section = self.document.sections[0]
        width_twips = max(1, math.floor(inner_width))
        one = dataclasses.replace(section, page_size=dataclasses.replace(section.page_size, width_twips=width_twips),
                                  margins=dataclasses.replace(section.margins, left=0, right=0, gutter=0))
        blocks = [dataclasses.replace(block, path=f"{paragraph_path}/{block.path}") for block in graphic.text]
        blocks, nested = paginate.floating_marks(blocks)
        story = paginate.story_document(self.document, blocks, one)
        cache = self.layout.__dict__.setdefault("_text_box_labels", {})
        if id(self.document) not in cache:
            cache[id(self.document)] = (self.document, linebreak.text_box_labels(self.document))
        story = dataclasses.replace(story, list_labels=cache[id(self.document)][1])
        items = paginate.flow(story, self.advances, self.metrics)
        placer = _Placer(story, None, self.advances, self.metrics, self.decorations)
        placer.layout = self.layout
        placer.frames = self.frames
        placer.story_part = placed.part or getattr(self, "story_part", None)
        placer.items = items
        placer.references = []
        details = placer._details(items)
        content = paginate.stack_height(story, items)
        if nested:
            self.layout.warn("drawing-not-drawn", f"a floating drawing inside a text box is not drawn ({path})",
                             page.number + 1)
        obstacle = next((item for item in items if isinstance(item, paginate.Obstacle)), None)
        if obstacle is not None:
            self.layout.warn(f"story-stopped:{obstacle.reason}", f"a text box ({path}) is drawn up to "
                             f"{placer._obstacle_path(obstacle)} ({placer._why(obstacle)})", page.number + 1)
        # Where the text starts (``make_text_box_top_probe.py``, ROADMAP F.14).  Word holds
        # the box in whole twips: its corner the nearest twip (a box against its paragraph
        # starts a fraction of one off it), plus the insets; that point in the nearest
        # layout unit, which decides an edge on exactly half a pixel.  The lines start on
        # its device pixel, a half up, held in the nearest unit (567 px is drawn at
        # 567.0003).  The content stacks down from its top as a page's body does from its
        # top margin, or centred or at the foot of the room between the insets when it
        # fits (from the top when it does not: ``make_drawing_probe.py``, ``taller than
        # its box``), that top again in the nearest unit.  Below mode 15 the room is a
        # twip shorter, truncated to the unit (204 of 204.8).  The last line's box reaches
        # down to the room's foot (:meth:`_walk`'s ``story_bottom``).
        origin_x = round_half_up(x / TWIP_PX) + insets[0]
        origin_y = round_half_up(y / TWIP_PX) + insets[1]
        top = vertical.quantise(origin_y * TWIP_PX)
        room = (height / TWIP_PX - insets[1] - insets[3]) * TWIP_PX
        if (self.document.compatibility_mode or 0) < 15:
            room -= vertical.quantise_down(TWIP_PX)
        anchor = body.get("anchor", "t")
        offset = Fraction(0)
        if anchor == "ctr" and content < room:
            offset = (room - content) / 2
        elif anchor == "b" and content < room:
            offset = room - content
        story_bottom = top + room
        moved = vertical.quantise(top + offset)
        story_bottom += moved - top
        top = moved
        left = round_half_up(vertical.quantise(origin_x * TWIP_PX))
        left = round_half_up(left / LAYOUT_UNIT_PX) * LAYOUT_UNIT_PX
        scratch = Page(page.number, page.width_px, page.height_px, page.width_pt, page.height_pt,
                       (left, left + width - (insets[0] + insets[2]) * TWIP_PX, top, top + room))
        scratch.info = page.info
        placer._walk(scratch, items, details, paginate.Position(0), paginate.Position(len(items)), top,
                     bottom=Fraction(10 ** 7), story_bottom=story_bottom)
        # Text whose colour is automatic takes the shape style's font colour, or white on
        # a dark fill (``drawing.automatic_text_colour``, measured).
        colour = drawing.automatic_text_colour(graphic, colours)
        if colour is not None:
            for line in scratch.lines:
                for span in line.spans:
                    if span.color is None:
                        span.color = colour
        placed.lines.extend(scratch.lines)
        placed.rules.extend(scratch.rules)
        placed.primitives.extend(drawing_picture(p) for p in scratch.pictures)
        # An inline drawing that is not a picture -- a chart, a shape -- is drawn on the
        # scratch page as a float of its line's box (:meth:`_draw_lines`); it goes with the
        # text box, as its lines do.
        for inner in scratch.floats:
            placed.primitives.extend(inner.primitives)
            placed.lines.extend(inner.lines)
            placed.rules.extend(inner.rules)

    def _why(self, obstacle: paginate.Obstacle) -> str:
        """The obstacle's reason, and for a paragraph the breaker could not measure, what
        it could not measure (a face that is not installed, a character it lacks)."""
        if obstacle.reason == "table" and obstacle.detail:
            return f"a table the layout does not model yet: {obstacle.detail}"
        if obstacle.reason == "columns" and obstacle.detail:
            return f"{obstacle.detail} in a section of several text columns: not laid out by column yet"
        reason = {"table": "a table: tables are not laid out yet",
                  "drawing": "a drawing this model does not lay text out around (a VML shape, an embedded object, a wrap it does not know)",
                  "columns": "a section of several text columns holding what is not laid out by column yet",
                  "field": "a field: its result is not read yet",
                  "frame": "a frame: frames are not laid out yet",
                  "drawing (picture bullet)": "a picture bullet: not laid out yet"}.get(obstacle.reason,
                                                                                         obstacle.reason)
        if obstacle.reason in ("unmeasurable", "no face metrics") and obstacle.paragraph is not None:
            section = self.document.sections[obstacle.section]
            try:
                linebreak.break_paragraph(self.document, obstacle.paragraph, section, self.advances, self.metrics)
            except linebreak.Unmeasurable as error:
                return f"a paragraph that cannot be measured: {error}"
            return "a paragraph whose faces have no metrics"
        return reason

    def _obstacle_path(self, obstacle: paginate.Obstacle) -> str:
        if obstacle.paragraph is not None:
            return obstacle.paragraph.path
        index = 0
        for item in self.document.body:
            if isinstance(item, Table):
                if index == obstacle.block:
                    return item.path
                index += self._cells(item)
            else:
                index += 1
        return ""

    def _cells(self, table: Table) -> int:
        total = 0
        for row in table.rows:
            for cell in row:
                for item in cell:
                    total += self._cells(item) if isinstance(item, Table) else 1
        return total

    # -- a table ------------------------------------------------------------------------------

    def _table(self, page: Page, item, laid) -> None:
        """The pieces of a table's rows on this page (:func:`docx2svg.table.place_table`):
        every cell paragraph's lines at the cell's text start and on the stack
        :func:`docx2svg.table.cell_stack` makes, each baseline rounded in its line box as
        the body's is; and the rows' borders and shading."""
        from . import table as table_model

        for index, piece in enumerate(laid.pieces):
            # A row beside a drawing text wraps around is drawn where the table was laid
            # out in the room beside it (``paginate.table_beside``).
            flow = piece.flow or item.flow
            row = flow.rows[piece.row]
            content_top = piece.top + piece.top_border
            # A merge going across the page: its lines here, from this piece's top down
            # through its rows (``table.place_table``, ``RowPiece.carry``).
            drawn = [(cell, None, None) for cell in row.cells] + [
                (flow.rows[r].cells[c], start, end) for r, c, start, end in piece.carry]
            for c, (cell, carried_start, carried_end) in enumerate(drawn):
                if cell.merged:
                    continue
                margin_top = twips_to_px(cell.cell.margins["top"])
                margin_bottom = twips_to_px(cell.cell.margins["bottom"])
                offset = Fraction(0)
                if carried_start is not None:
                    lines = range(carried_start, carried_end)
                else:
                    lines = range(piece.start[c], piece.end[c])
                if piece.whole and carried_start is None:
                    space = (cell.span_height if cell.span_height is not None else row.height) - margin_top - margin_bottom
                    align = cell.cell.properties.get("vAlign") or "top"
                    if align == "center":
                        offset = max(Fraction(0), (space - cell.content) / 2)
                    elif align == "bottom":
                        offset = max(Fraction(0), space - cell.content)
                line_y = content_top + margin_top + offset
                paragraph_top = line_y
                for k in lines:
                    placed = cell.stack[k]
                    paragraph = placed.paragraph
                    if paragraph.table is not None:
                        # A table nested in the cell (``table.nested_entry``), from its top.
                        top = line_y + placed.above
                        nested = table_model.place_table(paragraph.table, top, top + Fraction(10 ** 7),
                                                         first_on_page=False,
                                                         mode15=(self.document.compatibility_mode or 0) >= 15)
                        self._table(page, SimpleNamespace(flow=paragraph.table), nested)
                        line_y = top + placed.pitch + placed.below
                        continue
                    height = paragraph.heights[placed.number]
                    if placed.number == 0 and k > 0:
                        # Where a floating drawing's ``paragraph`` is: where the paragraph
                        # before ends after its own space after (the body's rule).
                        before = cell.stack[k - 1].paragraph
                        same = paragraph.style == before.style
                        paragraph_top = line_y + twips_to_px(0 if (before.contextual and same) else before.after)
                    top = line_y + placed.above
                    box = height.box(placed.box_before, placed.box_after)
                    baseline = baseline_in_box(top - box.space_before, box)
                    band = text_band(top - box.space_before, box)
                    # A line that starts below the page's edge (a row that may not split,
                    # taller than the page, below mode 15) is not drawn.
                    if placed.number not in paragraph.undrawn and top < page.height_px:
                        info = _Paragraph(paragraph.paragraph, paragraph.resolved, paragraph.pieces, paragraph.lines,
                                          paragraph.geometry, paragraph.label)
                        self._line(page, info, None, placed.number, top, height, baseline, band,
                                   left=cell.box.text_left)
                    if paragraph.anchors:
                        self._cell_anchors(page, item, flow, piece, cell, paragraph, placed.number, top, height,
                                           paragraph_top, content_top)
                    line_y = top + placed.pitch + placed.below
            last_here = index == len(laid.pieces) - 1
            # Where the rows below go on beside a drawing, or below it, the rows above end
            # as a table does: with the edge under them (``make_wrap_table_probe.py``,
            # below mode 15, a drawing that starts inside a table).
            following = None if last_here else laid.pieces[index + 1]
            parted = following is not None and ((following.flow or item.flow) is not flow
                                                 or following.top != content_top + piece.height)
            self._row_borders(page, flow, piece, content_top + piece.height,
                              last=last_here or parted or piece.row == len(flow.rows) - 1)

    def _cell_anchors(self, page: Page, item, flow, piece, cell, paragraph, number: int, top: Fraction, height,
                      paragraph_top: Fraction, cell_top: Fraction) -> None:
        """Place the floating drawings anchored on line ``number`` of a cell paragraph
        (``make_cell_anchor_probe.py``; ROADMAP.md, "Floating drawings -- measured", F.16).

        **In the cell** (``layoutInCell``, and in mode 15 whatever it says): ``column`` and
        ``margin`` are the cell's text area -- its grid lines less its margins, exact --
        ``page`` and ``leftMargin`` start inside its left border, ``paragraph`` is the cell
        paragraph's (the body's rule: the first starts below the cell's top margin, the
        others where the one before ends after its space after), ``line`` the line's, and
        ``margin``, ``page`` and ``topMargin`` start below the row's top border.  **Below
        mode 15 without ``layoutInCell``** every frame is the page's, but ``paragraph`` is
        the row's top (above its top border, whichever of the cell's paragraphs anchors it)
        and ``line`` the line's."""
        from . import table as table_model
        from .floating import Frames, in_cell

        section = self.document.sections[item.section]
        m = section.margins
        mode15 = (self.document.compatibility_mode or 0) >= 15
        info = _Paragraph(paragraph.paragraph, paragraph.resolved, paragraph.pieces, paragraph.lines,
                          paragraph.geometry, paragraph.label)
        broken = paragraph.lines[number]
        positions = wrapped_positions(paragraph.pieces, broken, paragraph.geometry, self.advances,
                                      first_line=number == 0, alignment=paragraph.resolved.get("jc"),
                                      last_line=number == len(paragraph.lines) - 1)
        grid = flow.grid_twips
        c = cell.cell
        text_left = flow.page_left_twips + grid[c.column] + table_model.left_offset(c)
        text_right = flow.page_left_twips + grid[min(c.column + c.span, len(grid) - 1)] - table_model.right_offset(c)
        box_left = flow.page_left_twips + grid[c.column] + table_model.inner_part(c.left_edge[0])
        width, height_twips = section.page_size.width_twips, section.page_size.height_twips
        line_top = paragraph_top if number == 0 else top
        for source, run_index, k, anchor, here in paragraph.anchors:
            if here != number:
                continue
            character = twips_to_px(text_left) + units_px(anchor_character_units(
                paragraph.pieces, broken, positions, source, paragraph.geometry, first_line=number == 0, run=run_index))
            if in_cell(anchor, mode15):
                frames = Frames(width, height_twips, text_left, width - text_right, cell_top / TWIP_PX,
                                height_twips - cell_top / TWIP_PX, m.header, m.footer, page.number + 1, mode15,
                                paragraph_top / TWIP_PX, line_top / TWIP_PX, (top + height.pitch) / TWIP_PX,
                                character / TWIP_PX, box_left, cell_top / TWIP_PX, cell=True)
            else:
                frames = self._frames(page, item.section, piece.top, line_top, top + height.pitch, character)
            run = info.paragraph.runs[run_index]
            path = f"{info.paragraph.path}/{run.path}/wp:anchor[{k + 1}]" if run.path else info.paragraph.path
            positioned = paragraph.positioned.get((source, run_index, k))
            if positioned is not None:
                # One text wraps around: where the cell's layout without it put it
                # (``table.wrap_cell``).
                self._place_float(page, anchor, frames, path, info.paragraph.path, x=positioned[0],
                                  y=cell_top + positioned[1])
                continue
            self._place_float(page, anchor, frames, path, info.paragraph.path)

    def _row_borders(self, page: Page, flow, piece, bottom: Fraction, *, last: bool) -> None:
        """A row's borders and shading as filled rectangles, as Word's PDF draws them
        (``make_table_border_probe.py``, ``make_table_line_probe.py``):

        * each **vertical edge** is the winner of the two cells' borders, its width in
          whole px (the larger half to the left), centred on its grid line **rounded
          against the text column's left edge** -- that edge rounded to a whole px, the
          grid line's distance from it rounded half away from zero
          (:func:`grid_line_px`) -- down the row; where it crosses the band above, it
          fills the band's height;
        * a cell's **shading** (``w:shd``) fills it between its grid lines less the inner
          half of each vertical border there (the smaller half of an odd width in
          twips), each edge rounded to the nearest px from the page's edge -- so where a
          line's width in px is less than its exact width, a column of white is left
          beside it, as Word leaves it -- from below the band above the row; and where
          the line drawn above it is short of its exact width by half a px or more
          (``w:sz`` 24: 12.5 px drawn 12), a row of it directly under the line, as wide;
        * the **edge above the row** is a band as tall as its widest border; across each
          cell the border that wins there (:func:`docx2svg.table.edge_winner`, the upper
          cell's bottom against this one's top) is drawn from the band's top, its own
          width tall, between the vertical borders;
        * the **edge below the last row on the page** (the table's bottom, or a row at the
          page's foot) is drawn up from the band's *bottom* rounded, and the vertical
          lines stop where it starts;
        * a line is its width in whole twips, truncated to whole device px; a ``double``
          border is three such widths, a line, a gap and a line.
        """
        from . import table as table_model

        r = piece.row
        row_top = piece.top
        content_top = piece.top + piece.top_border
        row = flow.rows[r]
        grid = flow.grid_twips
        columns = len(grid) - 1

        def line_px(border) -> int:
            return math.floor(twips_to_px(border[1] * 20 // 8)) if table_model.drawn(border) else 0

        def total_px(border) -> int:
            return line_px(border) * (3 if border and border[0] == "double" else 1)

        def colour(border) -> str:
            return _color(border[3]) or "000000"

        def vertical_span(border, index: int) -> tuple[int, int]:
            width = total_px(border)
            x0 = grid_line_px(flow.page_left_twips, grid[index]) - (width + 1) // 2
            return x0, x0 + width

        def shading_edge(index: int, border, sign: int) -> int:
            inner = table_model.inner_part(border) if table_model.drawn(border) else 0
            return round_half_up((flow.page_left_twips + grid[index] + sign * inner) * TWIP_PX)

        def band(kind: str, path: str, border, x0: int, x1: int, y0: int, y1: int, vertical: bool) -> None:
            """A border line, or a double's two lines, filling ``[x0, x1) x [y0, y1)``."""
            if x1 <= x0 or y1 <= y0:
                return
            if border[0] == "double":
                w = line_px(border)
                if vertical:
                    parts = [(x0, x0 + w), (x1 - w, x1)]
                    for a, b in parts:
                        page.rules.append(Rule(kind, path, Fraction(a), Fraction(y0), Fraction(b - a),
                                               Fraction(y1 - y0), colour(border), border[0]))
                else:
                    for a, b in ((y0, y0 + w), (y1 - w, y1)):
                        page.rules.append(Rule(kind, path, Fraction(x0), Fraction(a), Fraction(x1 - x0),
                                               Fraction(b - a), colour(border), border[0]))
                return
            page.rules.append(Rule(kind, path, Fraction(x0), Fraction(y0), Fraction(x1 - x0), Fraction(y1 - y0),
                                   colour(border), border[0]))

        # A merge going on at the page's top has its top border there, as a row going on
        # does (``make_table_foot_probe.py``).
        carried = {flow.rows[r0].cells[c0].cell.column for r0, c0, _start, _end in piece.carry}

        def edge_cells(cells_above, cell):
            """The winner on the horizontal edge above ``cell``."""
            if cell.merged and cell.cell.column in carried:
                return cell.cell.borders.get("top")
            own = (cell.cell.borders.get("top"), cell.cell.own.get("top")) if not cell.merged else (None, False)
            above = [c for c in cells_above if c.cell.column < cell.cell.column + cell.cell.span
                     and cell.cell.column < c.cell.column + c.cell.span] if not cell.merged else []
            winner = own[0]
            for other in above:
                winner = table_model.edge_winner((other.cell.borders.get("bottom"), other.cell.own.get("bottom")),
                                                 (winner, own[1]))[0]
            return winner

        def right_index(cell) -> int:
            return min(cell.cell.column + cell.cell.span, columns)

        def vertical_edges(cell):
            """The vertical edges drawn with ``cell``: its left one, and the row's last right one."""
            out = [(cell.cell.left_edge[0], cell.cell.column)]
            if cell is cells[-1]:
                out.append((cell.cell.right_edge[0], right_index(cell)))
            return out

        def line_extent(cell) -> tuple[int, int]:
            """A horizontal line's extent across ``cell``: between the vertical lines."""
            left_border = cell.cell.left_edge[0]
            right_border = cell.cell.right_edge[0]
            x0 = vertical_span(left_border, cell.cell.column)[1] if table_model.drawn(left_border) \
                else grid_line_px(flow.page_left_twips, grid[cell.cell.column])
            right = right_index(cell)
            x1 = vertical_span(right_border, right)[0] if table_model.drawn(right_border) \
                else grid_line_px(flow.page_left_twips, grid[right])
            return x0, x1

        def horizontal_edge(y0: int, y1: int, cells, winners, *, from_bottom: bool) -> None:
            for cell, winner in zip(cells, winners):
                if not table_model.drawn(winner):
                    continue
                x0, x1 = line_extent(cell)
                if from_bottom:
                    band("table-border", cell.cell.path, winner, x0, x1, y1 - total_px(winner), y1, False)
                else:
                    band("table-border", cell.cell.path, winner, x0, x1, y0, y0 + total_px(winner), False)
            # Where each vertical edge crosses the band: its own colour, the band's height.
            for cell in cells:
                for border, index in vertical_edges(cell):
                    if table_model.drawn(border) and y1 > y0:
                        a, b = vertical_span(border, index)
                        band("table-border", cell.cell.path, border, a, b, y0, y1, True)

        cells = [c for c in row.cells]
        # A row at a page's top -- a repeated header, a row going on from the page
        # before -- has no row above it there: its own top border.
        above = flow.rows[r - 1].cells if (r and not piece.header and not piece.continued) else []
        top_winners = [edge_cells(above, cell) for cell in cells]
        horizontal_edge(round_half_up(row_top), round_half_up(row_top + row.top_border), cells, top_winners,
                        from_bottom=False)
        y0 = round_half_up(content_top)
        y1 = round_half_up(bottom)
        if last:
            below = flow.rows[r + 1].top_border if r + 1 < len(flow.rows) else flow.bottom_border
            winners = [cell.cell.borders.get("bottom") for cell in cells]
            band_bottom = round_half_up(bottom + below)
            drawn_winners = [total_px(w) for w in winners if table_model.drawn(w)]
            band_top = band_bottom - max(drawn_winners) if drawn_winners else round_half_up(bottom)
            y1 = band_top
        for cell, winner in zip(cells, top_winners):
            fill = _shading(cell.cell.properties.get("shd"))
            if fill and not cell.merged:
                x0 = shading_edge(cell.cell.column, cell.cell.left_edge[0], 1)
                x1 = shading_edge(right_index(cell), cell.cell.right_edge[0], -1)
                span_bottom = round_half_up(content_top + cell.span_height) if piece.whole and cell.span_height \
                    else y1
                page.rules.insert(0, Rule("cell-shading", cell.cell.path, Fraction(x0), Fraction(y0),
                                          Fraction(x1 - x0), Fraction(span_bottom - y0), fill))
                # What the line above leaves of its exact width, rounded: the cell's
                # colour, directly under the line and as wide (a w:sz 24 line, 12.5 px,
                # is drawn 12 px and a row of shading).
                if table_model.drawn(winner):
                    rest = round_half_up(table_model.border_px(winner)) - total_px(winner)
                    if rest > 0:
                        a, b = line_extent(cell)
                        under = round_half_up(row_top) + total_px(winner)
                        page.rules.insert(0, Rule("cell-shading", cell.cell.path, Fraction(a), Fraction(under),
                                                  Fraction(b - a), Fraction(rest), fill))
            for border, index in vertical_edges(cell):
                if table_model.drawn(border):
                    a, b = vertical_span(border, index)
                    band("table-border", cell.cell.path, border, a, b, y0, y1, True)
        if last:
            horizontal_edge(band_top, band_bottom, cells, winners, from_bottom=True)

    # -- decorations of a paragraph ---------------------------------------------------------

    def _paragraph_rules(self, page: Page, info: _Paragraph, flow_index: int, band_top: int, band_bottom: int,
                         y_gap: Fraction, y_after: Fraction, *, starts_here: bool, ends_here: bool) -> None:
        """A paragraph's borders and shading on this page: :func:`paragraph_rects`, with
        whether its neighbours share its border box (``vertical.border_box``)."""
        before_same, after_same = self.box_neighbours.get(flow_index, (False, False))
        pp = info.resolved
        left = page.column[0] + twips_to_px(pp.get("ind.left", 0) or 0)
        right = page.column[1] - twips_to_px(pp.get("ind.right", 0) or 0)
        for rule in paragraph_rects(pp, info.paragraph.path, left, right, band_top, band_bottom, y_gap, y_after,
                                    top_side=starts_here and not before_same,
                                    between_above=starts_here and before_same,
                                    bottom_side=ends_here and not after_same):
            if rule.kind == "paragraph-shading":
                page.rules.insert(0, rule)
            else:
                page.rules.append(rule)

    # -- one line ---------------------------------------------------------------------------

    def _join_drop(self, info: _Paragraph, number: int) -> int:
        """How much lower than laid out Word draws line ``number`` of a paragraph that
        paragraphs whose marks are deleted were joined into (:attr:`Paragraph.joins`).

        Below mode 15 Word's final view draws a line that begins one of the joined
        paragraphs' text -- a line other than the first, the join falling at its start --
        lower by the joined paragraph's space before, and every later line of the paragraph
        with it, cumulatively; what follows the paragraph moves down by one fewer of them
        (:func:`paginate.join_starts`), so its last line overlaps it.  In mode 15, and
        with every revision accepted, nothing is drawn lower (ROADMAP.md, "Revisions --
        measured", R.4)."""
        if number == 0 or (self.document.compatibility_mode or 0) >= 15:
            return 0
        before = info.resolved.get("spacing.before", 0) or 0
        if not before or not info.paragraph.joins:
            return 0
        count = sum(1 for m in paginate.join_starts(info.paragraph, info.pieces, info.lines) if m <= number)
        return count * round_half_up(twips_to_px(before))

    def _line(self, page: Page, info: _Paragraph, item: paginate.Para | None, number: int, top: Fraction,
              height, baseline: int, band: tuple[int, int], *, left: Fraction | int | None = None) -> None:
        document = self.document
        paragraph = info.paragraph
        pieces = info.pieces
        broken = info.lines[number]
        alignment = info.resolved.get("jc")
        leaders: dict = {}
        positions = wrapped_positions(pieces, broken, info.geometry, self.advances, first_line=number == 0,
                                      alignment=alignment, last_line=number == len(info.lines) - 1, leaders=leaders)
        if alignment == "distribute":
            self.layout.warn("alignment-distribute-approximate",
                             "a distributed paragraph's last line is spread evenly between its glyphs; Word "
                             "spreads it by a rule not measured")
        column_left = page.column[0] if left is None else left
        drop = self._join_drop(info, number)
        if drop:
            top, baseline, band = top + drop, baseline + drop, (band[0] + drop, band[1] + drop)
        line = Line(f"{paragraph.path}", number, baseline, top, height.pitch, paragraph_id=paragraph.para_id,
                    band=band)
        page.lines.append(line)
        runs = paragraph.runs
        mark = None
        drawing_counts: dict[int, int] = {}
        current: Span | None = None

        for j in range(broken.start, broken.end):
            piece = pieces[j]
            if piece.kind == linebreak.TAB and piece.run >= 0 and j in leaders:
                current = None
                _, x0, x1 = leaders[j]
                self._tab_underline(line, info, piece, column_left + units_px(positions[j]),
                                    column_left + units_px(x1 + positions[j] - x0), baseline, height)
                continue
            if piece.kind in (linebreak.TAB, linebreak.BREAK):
                current = None
                continue
            if piece.kind == linebreak.SOFT_HYPHEN and not (broken.hyphenated and j == broken.end - 1):
                continue
            if piece.source < 0 and piece.char == linebreak.LABEL_SHIFT:
                continue
            x = column_left + units_px(positions[j])
            if piece.kind == linebreak.OBJECT:
                current = None
                k = piece.run
                count = drawing_counts.get(k, 0)
                drawing_counts[k] = count + 1
                run = runs[k]
                marks = [kind for _, kind in run.breaks if kind.startswith("drawing:")]
                relationship = run.drawings[count] if count < len(run.drawings) else None
                cx, cy, left, top_, right, bottom = (int(v) for v in marks[count].split(":")[1:])
                emu = Fraction(PX_PER_PT, 12700)
                width, depth = picture_size(cx, cy)
                stands = top + height.object_drop if height.object_drop is not None else baseline
                where = run.path and f"{paragraph.path}/{run.path}"
                if relationship is not None and not isinstance(relationship, str):
                    # An inline shape or group: drawn as a floating drawing is, in its
                    # line's box, in front of the text (Word paints it in the text's order;
                    # nothing inline overlaps the text, so the layer does not show).
                    counter = self.__dict__.setdefault("_float_count", [0])
                    counter[0] += 1
                    placed = Float(where or paragraph.path, x + left * emu, stands - bottom * emu - depth, width,
                                   depth, False, (0, counter[0]))
                    self._draw_graphic(page, placed, relationship, placed.path, paragraph.path)
                    continue
                picture = Picture(where, relationship, x + left * emu, stands - bottom * emu - depth, width, depth)
                page.pictures.append(picture)
                if relationship is None:
                    self.layout.warn("drawing-not-drawn", "an inline drawing that is not a picture (a chart, a "
                                     "shape) is drawn as a placeholder of its extent")
                continue
            if piece.source >= 0 and piece.char in (linebreak.NOTE_SEPARATOR, linebreak.NOTE_CONTINUATION):
                current = None
                k = piece.run
                resolved = self.run_cache.get(id(runs[k]))
                if resolved is None:
                    resolved = resolve_run(document, paragraph, runs[k])
                    self.run_cache[id(runs[k])] = resolved
                end = None
                if piece.char == linebreak.NOTE_CONTINUATION:
                    # From where the line starts to the right edge, whatever the alignment.
                    x = column_left + units_px(info.geometry.first_start if number == 0 else info.geometry.start)
                    end = column_left + units_px(info.geometry.right)
                self._separator(page, runs[k], resolved, f"{paragraph.path}/{runs[k].path}", x, end, baseline)
                continue
            if piece.source < 0:
                if mark is None:
                    mark = resolve_mark(document, paragraph)
                values = dict(mark.values)
                num_id, ilvl = info.resolved.get("numPr.numId"), info.resolved.get("numPr.ilvl", 0) or 0
                level = document.numbering.get(num_id, {}).get(ilvl) if num_id else None
                if level is not None:
                    values.update(level.run)
                resolved = Resolved(values, dict(mark.origins))
                path, kind, key = f"{paragraph.path}/label", "label", ("label",)
                if piece.char == linebreak.PICTURE_BULLET and level is not None:
                    current = None
                    page.pictures.append(self._picture_bullet(level, resolved, path, x, units_px(piece.width),
                                                              baseline))
                    continue
            else:
                k = piece.run
                resolved = self.run_cache.get(id(runs[k]))
                if resolved is None:
                    resolved = resolve_run(document, paragraph, runs[k])
                    self.run_cache[id(runs[k])] = resolved
                path = f"{paragraph.path}/{runs[k].path}" if runs[k].path else paragraph.path
                kind = "hyphen" if piece.kind == linebreak.SOFT_HYPHEN else "text"
                key = ("run", k, kind)
            char = "-" if piece.kind == linebreak.SOFT_HYPHEN else (
                "-" if piece.char == "‑" else piece.char)
            fmt = character_format(resolved, char if piece.source >= 0 else piece.char, document)
            face = piece.face[0] if piece.face else fmt.face
            drawn_hp = fmt.half_points
            offset = 0
            if fmt.vertical_align:
                face_metrics = self.metrics(fmt.face, fmt.bold, fmt.italic) if self.metrics else None
                drawn_hp = script_half_points(fmt.half_points, fmt.vertical_align, face_metrics) or fmt.half_points
                raised = half_points_offset_px(script_raise_half_points(fmt.half_points, fmt.vertical_align,
                                                                        face_metrics))
                offset += raised if fmt.vertical_align == "superscript" else -raised
            y = baseline - half_points_offset_px(fmt.position) - offset
            key = key + (face, fmt.bold, fmt.italic, drawn_hp, y)
            if current is None or getattr(current, "_key", None) != key:
                current = Span(path, [], [], y, x, face or "", fmt.bold, fmt.italic, drawn_hp,
                               color=_color(resolved.get("color")),
                               underline=resolved.get("u") if resolved.get("u") not in (None, "none") else None,
                               strike=bool(resolved.get("strike")), double_strike=bool(resolved.get("dstrike")),
                               highlight=HIGHLIGHT.get(resolved.get("highlight") or ""),
                               shading=_shading(resolved.get("shd")), kind=kind, line_baseline=baseline,
                               line_above=height.text_above, line_below=height.text_below,
                               vertical_align=fmt.vertical_align,
                               paint=self._text_paint(resolved))
                current._key = key  # type: ignore[attr-defined]
                line.spans.append(current)
            width = piece.hyphen if piece.kind == linebreak.SOFT_HYPHEN else piece.width
            current.chars.append(char if piece.kind != linebreak.SPACE else piece.char)
            # A note's number is drawn on whole device px (``linebreak.note_number_units``).
            current.xs.append(Fraction(round_half_up(x)) if piece.source >= 0 and runs[piece.run].note_number else x)
            current.end = x + units_px(width)
        for j, (leader, x0, x1) in sorted(leaders.items()):
            if leader is None:
                continue
            shift = positions[j] - x0
            self._leader(line, info, pieces[j], leader, column_left, positions[j], x1 + shift, baseline, height)
        self._decorate(page, line)
        if item is not None and item.anchors:
            self._anchors(page, info, item, number, top, height, positions)

    def _tab_underline(self, line: Line, info: _Paragraph, piece, start: Fraction, end: Fraction, baseline: int,
                       height) -> None:
        """An underlined tab's underline across its room, from the pen at the tab to the pen
        after it -- its stop, where it ends its line -- as a run's under a space
        (``make_tab_leader_probe.py``, ``runs``: a tab run underlined between text that is
        not, with no leader, under a dot and an underscore leader, and ending its line)."""
        run = info.paragraph.runs[piece.run]
        resolved = self.run_cache.get(id(run))
        if resolved is None:
            resolved = resolve_run(self.document, info.paragraph, run)
            self.run_cache[id(run)] = resolved
        underline = resolved.get("u")
        if underline in (None, "none") or end <= start:
            return
        fmt = character_format(resolved, " ", self.document)
        if fmt.face is None:
            return
        path = f"{info.paragraph.path}/{run.path}" if run.path else info.paragraph.path
        line.spans.append(Span(path, [" "], [start], baseline - half_points_offset_px(fmt.position), end, fmt.face,
                               fmt.bold, fmt.italic, fmt.half_points, color=_color(resolved.get("color")),
                               underline=underline, kind="tab", line_baseline=baseline,
                               line_above=height.text_above, line_below=height.text_below))

    def _leader(self, line: Line, info: _Paragraph, piece, leader: str, column_left, start, end, baseline: int,
                height) -> None:
        """A tab's leader (``w:tab/@w:leader``), as the glyphs Word draws it with
        (:data:`LEADER_GLYPHS`) in the tab's run's format -- face, size, weight, slant,
        colour, a script's size and offset (``make_tab_leader_probe.py``, ``runs``) -- each
        at a whole multiple of the glyph's advance (rounded to the layout unit) from the
        page's left edge, from the first at or after the pen at the tab to the last that
        ends at or before the pen after it (``phase``, ``end``, ``kinds``, ``faces``)."""
        document = self.document
        paragraph = info.paragraph
        char = LEADER_GLYPHS.get(leader)
        if char is None or piece.run < 0:
            self.layout.warn("tab-leader-not-drawn", f"a {leader} tab leader is not drawn")
            return
        run = paragraph.runs[piece.run]
        resolved = self.run_cache.get(id(run))
        if resolved is None:
            resolved = resolve_run(document, paragraph, run)
            self.run_cache[id(run)] = resolved
        fmt = character_format(resolved, char, document)
        if fmt.face is None:
            return
        drawn_hp = fmt.half_points
        offset = 0
        if fmt.vertical_align:
            face_metrics = self.metrics(fmt.face, fmt.bold, fmt.italic) if self.metrics else None
            drawn_hp = script_half_points(fmt.half_points, fmt.vertical_align, face_metrics) or fmt.half_points
            raised = half_points_offset_px(script_raise_half_points(fmt.half_points, fmt.vertical_align,
                                                                    face_metrics))
            offset += raised if fmt.vertical_align == "superscript" else -raised
        advance = self.advances.advance(fmt.face, fmt.bold, fmt.italic, char)
        if advance is None:
            self.layout.warn("tab-leader-not-drawn", f"a {leader} tab leader in {fmt.face}, whose advance is not "
                             "known, is not drawn")
            return
        step = round_half_up(Fraction(advance[0] * drawn_hp * 2048, advance[1]))
        if step <= 0:
            return
        origin = column_left / LAYOUT_UNIT_PX
        first = math.ceil((origin + start) / step)
        last = math.floor((origin + end) / step) - 1
        if last < first:
            return
        y = baseline - half_points_offset_px(fmt.position) - offset
        path = f"{paragraph.path}/{run.path}" if run.path else paragraph.path
        span = Span(path, [char] * (last - first + 1), [step * k * LAYOUT_UNIT_PX for k in range(first, last + 1)],
                    y, (last + 1) * step * LAYOUT_UNIT_PX, fmt.face, fmt.bold, fmt.italic, drawn_hp,
                    color=_color(resolved.get("color")), kind="leader", line_baseline=baseline,
                    line_above=height.text_above, line_below=height.text_below, vertical_align=fmt.vertical_align,
                    paint=self._text_paint(resolved))
        line.spans.append(span)

    def _drop_cap(self, page: Page, item: paginate.Para, top: Fraction) -> None:
        """A drop cap's lines (:class:`paginate.DropCap`), in its frame at the top of the
        paragraph it drops into: after its own space before, each line in its box as a
        paragraph's is; at the column's start, or -- a ``margin`` cap anchored to the page --
        its width left of it, rounded (``make_drop_cap_probe.py``: 172 px for a 128.17 px
        letter against a column at 300)."""
        drop = item.drop_cap
        section = self.document.sections[item.section]
        info = _Paragraph(drop.paragraph, drop.pp, drop.pieces, drop.lines,
                          linebreak.geometry(self.document, drop.paragraph, section, drop.pp), None)
        left = page.column[0]
        if not drop.beside:
            left = page.column[0] - round_half_up(units_px(drop.width))
        y = top + twips_to_px(drop.before)
        for number, height in enumerate(drop.heights):
            box = height.box()
            # Its baseline from the foot of its line, whatever its fraction of a pixel:
            # the line's bottom rounded, less what is below the text, rounded
            # (``make_drop_cap_probe.py``: exact lines of 400 to 1,600 twips, single lines
            # of 11 to 50 pt, the lowering (``w:position``) then as a run's).
            below = vertical._line_extent(box)[1]
            baseline = round_half_up(y + height.pitch) - round_half_up(below)
            self._line(page, info, None, number, y, height, baseline, text_band(y, box), left=left)
            y += height.pitch

    def _continuation(self, page: Page, item: paginate.Para, y: Fraction) -> Fraction:
        """The continuation separator's line at the top of a page the endnotes go on to
        (``paginate._scan``); where the stack goes on under it."""
        from .notes import separator

        paragraph = next((b for b in separator(self.document, "continuationSeparator") if isinstance(b, Paragraph)),
                         None)
        if paragraph is None:
            return y
        section = self.document.sections[item.section]
        pp = resolve_paragraph(self.document, paragraph)
        pieces, lines = linebreak.break_paragraph(self.document, paragraph, section, self.advances, self.metrics)
        info = _Paragraph(paragraph, pp, pieces, lines, linebreak.geometry(self.document, paragraph, section, pp), None)
        height = item.continuation
        box = height.box()
        self._line(page, info, None, 0, y, height, baseline_in_box(y, box), text_band(y, box))
        return y + height.pitch

    def _separator(self, page: Page, run, resolved, path: str, x: Fraction, end: Fraction | None,
                   baseline: int) -> None:
        """A note separator's line (:mod:`docx2svg.notes`): where its run's face would
        strike through its text, as thick -- ``OS/2`` strikeout position and size at the
        size the face is drawn at, each rounded (``make_endnote_probe.py``: Calibri 8 and
        11 pt, Times New Roman 20 pt, 3 / 3) -- from the pen's x, 2,880 twips long, or for
        the continuation separator (``end``) to the column's right edge."""
        fmt = character_format(resolved, " ", self.document)
        decorations = self.decorations(fmt.face, fmt.bold, fmt.italic) if self.decorations and fmt.face else None
        if decorations is None or decorations.strikeout_position is None:
            self.layout.warn("note-separator-not-drawn", "a note separator whose face has no strikeout metrics is "
                             "not drawn")
            return
        scale = Fraction(round_half_up(Fraction(fmt.half_points, 2) * PX_PER_PT)) / decorations.units_per_em
        thick = max(1, round_half_up(decorations.strikeout_size * scale))
        top = baseline - round_half_up(decorations.strikeout_position * scale)
        x0 = round_half_up(x)
        x1 = round_half_up(end if end is not None else x + 2880 * TWIP_PX)
        clip = getattr(self, "separator_clip", None)
        if end is None and clip is not None:
            # Below mode 15, in a column narrower than the line, cut at its right edge.
            x1 = min(x1, round_half_up(clip))
        page.rules.append(Rule("separator", path, Fraction(x0), Fraction(top), Fraction(x1 - x0), Fraction(thick),
                               _color(resolved.get("color"))))

    def _picture_bullet(self, level, resolved, path: str, x: Fraction, size: Fraction, baseline: int) -> Picture:
        """A picture bullet (``w:lvlPicBulletId``) drawn at its label's pen position: a
        square of its size (:func:`linebreak.picture_bullet_units`) standing on the
        baseline, its left edge and its side each rounded to a whole device pixel (a 24 pt
        bullet of 103.125 px centred on 375 px is drawn from 323 to 426) -- and moved by the level's
        ``w:position`` twice (``make_picture_bullet_probe.py``: +-6 half points, 13 px each, draw
        the picture 26 px up or down, where the line grows by the raise once), and by a
        ``w:vertAlign`` superscript's raise once (17 px at 11 pt), at its full size."""
        bullet = self.document.picture_bullets.get(level.picture_bullet)
        fmt = character_format(resolved, " ", self.document)
        bottom = baseline - 2 * half_points_offset_px(fmt.position)
        if fmt.vertical_align:
            face = self.metrics(fmt.face, fmt.bold, fmt.italic) if self.metrics and fmt.face else None
            raised = half_points_offset_px(script_raise_half_points(fmt.half_points, fmt.vertical_align, face))
            bottom -= raised if fmt.vertical_align == "superscript" else -raised
        side = round_half_up(size)
        return Picture(path, bullet.relationship if bullet else None, Fraction(round_half_up(x)),
                       Fraction(bottom - side), Fraction(side), Fraction(side), part=bullet.part if bullet else None)

    def _text_paint(self, resolved):
        """A run's Word 2010 text effects (``w14:textFill``, ``w14:textOutline``) as the
        shared renderers' fill and outline, or ``None`` when it states neither."""
        fill, outline = resolved.get("textFill"), resolved.get("textOutline")
        if fill is None and outline is None:
            return None
        from . import drawing

        colours = drawing.Colours(self.document.theme_colors, self.document.color_map)
        return colours.fill(fill), colours.outline(outline) if outline is not None else None

    # -- decorations of a line ---------------------------------------------------------------

    def _face_px(self, span: Span) -> Fraction:
        """The size a span's decorations scale with: its drawn size in whole device px,
        the size Word draws its ink at (Phase 2)."""
        return Fraction(round_half_up(Fraction(span.half_points, 2) * PX_PER_PT))

    def _decorate(self, page: Page, line: Line) -> None:
        """Underline, strikethrough, highlight and run shading for each span of the line,
        in device pixels: see :func:`decoration_rects`."""
        spans = [span for span in line.spans if span.chars]
        for index, span in enumerate(spans):
            trailing = index == len(spans) - 1
            decorations = self.decorations(span.face, span.bold, span.italic) if self.decorations else None
            for rule in decoration_rects(span, decorations, self._face_px(span), line_end=trailing, band=line.band):
                page.rules.append(rule)
            if span.underline and span.underline not in ("single", "words", "dotted", "dash"):
                self.layout.warn("underline-style-approximate",
                                 f"a {span.underline} underline is drawn approximately (only single, words, dotted and dash "
                                 f"underlines are measured to the pixel)")


def drawing_picture(picture: Picture):
    """An inline picture of a text box, as a drawing primitive."""
    from .drawing import Primitive

    kind = "image" if picture.relationship else "placeholder"
    return Primitive(kind, picture.path, x=picture.x, y=picture.y, width=picture.width, height=picture.height,
                     relationship=picture.relationship, what="drawing")


def anchor_character_units(pieces, line, positions: dict, source: int, geometry, *, first_line: bool,
                           run: int = 0):
    """The pen x, in layout units from the column's left edge, of the character a
    floating drawing anchored at text index ``source`` of run ``run`` is positioned
    against (``relativeFrom="character"``): the pen position of the piece before the
    anchor's mark on its line in document order -- a character, a space, a tab (where
    the tab starts) -- or the line's start when none is before it
    (``make_anchor_probe.py``, ``char``)."""
    before = [j for j in range(line.start, line.end) if j in positions and pieces[j].source >= 0
              and (pieces[j].source, pieces[j].run) < (source, run)]
    if before:
        return positions[before[-1]]
    first = next((j for j in range(line.start, line.end) if j in positions), None)
    if first is not None:
        return positions[first]
    return geometry.first_start if first_line else geometry.start


def _extent(span: Span, *, line_end: bool) -> tuple[Fraction, Fraction] | None:
    """The span's pen extent, less trailing spaces at the end of its line."""
    last = len(span.chars) - 1
    if line_end:
        while last >= 0 and span.chars[last].isspace():
            last -= 1
    if last < 0:
        return None
    end = span.end if last == len(span.chars) - 1 else span.xs[last + 1]
    return span.xs[0], end


def text_band(top: Fraction, box: LineBox) -> tuple[int, int]:
    """The rows a line's *text* occupies, as Word rounds them: what a highlight, run
    shading or paragraph shading fills, and where a paragraph's top border is measured
    from.

    The text line of the box (:func:`vertical.baseline_in_box`'s): below an ``atLeast``
    line's extra space, above an ``auto`` multiple's extra, the whole pitch of an
    ``exact`` line.  With nothing below the text its edges round on their own; with
    space below (a multiple's extra, space after, a bottom border) it rounds in the
    three parts the baseline does: the space above whole, the text keeping the rest.
    Measured by ``make_render_probe.py`` (``fill``: highlight and shading at 11 and
    20 pt under ``auto`` 240 and 360, ``exact`` 400, ``atLeast`` 480, and with space
    before and after -- 20 of 20 rectangles; the text extent about the rounded baseline,
    the first rule tried, 12 of 20).
    """
    inner, t, _above, _below = vertical._text_line(box)
    xa = box.space_before + inner
    if box.below_text == 0:
        start = top + xa
        return round_half_up(start), round_half_up(start + t)
    first = round_half_up(top) + round_half_up(xa)
    rest = t + (xa - round_half_up(xa)) + (box.below_text - round_half_up(box.below_text))
    return first, first + round_half_up(rest)


#: How far outside the text a paragraph's border box is drawn, beyond its ``w:space``, in
#: device px: the side borders' inner edges are this far out, and top and bottom borders
#: without side borders run this far past the text.  Measured by
#: ``make_render_probe.py`` (``box``): 6 at every width (``w:sz`` 4, 12, 24), space (0, 4,
#: 10 pt) and indent tried, on both sides.
BORDER_OUTSET_PX = 6


def _border_width_px(border) -> int:
    """A border line's drawn width: ``w:sz`` eighths of a point, truncated to whole device
    px (``w:sz`` 4, 8, 12, 24 draw 2, 4, 6, 12: 2.08, 4.17, 6.25 and 12.5 px)."""
    return max(1, math.floor(Fraction(border[1], 8) * PX_PER_PT))


def _space_px(border) -> int:
    """A border's ``w:space`` (points) as the drawing uses it: whole device px, truncated."""
    return math.floor(border[2] * PX_PER_PT) if border else 0


def paragraph_rects(pp, path: str, left: Fraction, right: Fraction, band_top: int, band_bottom: int,
                    y_gap: Fraction, y_after: Fraction, *, top_side: bool, between_above: bool,
                    bottom_side: bool) -> list[Rule]:
    """A paragraph's border lines and shading, in whole device px.

    ``left``/``right`` are its text edges (the column less its indents), ``band_top`` and
    ``band_bottom`` its first and last lines' text rows (:func:`text_band`), ``y_gap``
    where it starts below the space before it and ``y_after`` where the next paragraph's
    space starts (its bottom border included).  Measured by ``make_render_probe.py``
    (``box``, 36 boxes and 8 single sides, shading alone, indented and boxed, a group):

    * a **side** border's inner edge is ``w:space`` (whole px, truncated) and
      :data:`BORDER_OUTSET_PX` outside the text edge (the right one truncated to whole px
      first);
    * a **top** border's inner edge is ``w:space`` above the first line's text row;
    * a **bottom** border's inner edge is ``w:space`` below the last line's text row
      (which, the border being in that line's box, rounds as a line with space below
      its text does);
    * side borders run from the top border's outer edge to the bottom border's;
    * a **between** border (the upper of two paragraphs in one box) starts where the
      lower one starts, below its space before;
    * top, bottom and between lines run between the side borders' outer edges, or
      :data:`BORDER_OUTSET_PX` past the text where there is no side border;
    * **shading** fills the box inside the borders, or the text rows widened by
      :data:`BORDER_OUTSET_PX` where there is none.
    """
    out: list[Rule] = []
    sides = {name: pp.get(f"pBdr.{name}") for name in ("top", "left", "bottom", "right", "between")}
    drawn = {name: vertical._drawn(side) for name, side in sides.items()}
    text_left = round_half_up(left)
    text_right = math.floor(right)
    inner_left = text_left - _space_px(sides["left"]) - BORDER_OUTSET_PX if drawn["left"] else None
    inner_right = text_right + _space_px(sides["right"]) + BORDER_OUTSET_PX if drawn["right"] else None
    outer_left = inner_left - _border_width_px(sides["left"]) if drawn["left"] else text_left - BORDER_OUTSET_PX
    outer_right = inner_right + _border_width_px(sides["right"]) if drawn["right"] else text_right + BORDER_OUTSET_PX

    def line(kind: str, side, y0: int) -> None:
        out.append(Rule(kind, path, Fraction(outer_left), Fraction(y0), Fraction(outer_right - outer_left),
                        Fraction(_border_width_px(side)), _color(side[3]) or "000000", side[0]))

    box_top = band_top
    box_bottom = band_bottom
    if top_side and drawn["top"]:
        inner_top = band_top - _space_px(sides["top"])
        line("border-top", sides["top"], inner_top - _border_width_px(sides["top"]))
        box_top = inner_top
    elif between_above and drawn["between"]:
        start = round_half_up(y_gap)
        line("border-between", sides["between"], start)
        box_top = start + _border_width_px(sides["between"])
    outer_top, outer_bottom = box_top, box_bottom
    if top_side and drawn["top"]:
        outer_top = box_top - _border_width_px(sides["top"])
    if bottom_side and drawn["bottom"]:
        inner = band_bottom + _space_px(sides["bottom"])
        line("border-bottom", sides["bottom"], inner)
        box_bottom = inner
        outer_bottom = inner + _border_width_px(sides["bottom"])
    elif not bottom_side and any(drawn.values()):
        box_bottom = outer_bottom = round_half_up(y_after)
    for name, inner, width_sign in (("left", inner_left, -1), ("right", inner_right, 1)):
        if drawn[name]:
            width = _border_width_px(sides[name])
            x0 = inner - width if width_sign < 0 else inner
            out.append(Rule(f"border-{name}", path, Fraction(x0), Fraction(outer_top), Fraction(width),
                            Fraction(outer_bottom - outer_top), _color(sides[name][3]) or "000000", sides[name][0]))
    fill = _shading(pp.get("shd"))
    if fill:
        x0 = inner_left if drawn["left"] else text_left - BORDER_OUTSET_PX
        x1 = inner_right if drawn["right"] else text_right + BORDER_OUTSET_PX
        out.append(Rule("paragraph-shading", path, Fraction(x0), Fraction(box_top), Fraction(x1 - x0),
                        Fraction(box_bottom - box_top), fill))
    return out


def decoration_rects(span: Span, decorations, size_px: Fraction, *, line_end: bool,
                     band: tuple[int, int] | None = None) -> list[Rule]:
    """The rectangles Word fills for a span's decorations, in whole device pixels.

    Measured by ``make_render_probe.py`` (``decor``: eight faces at five sizes each):

    * **single underline**: ``post.underlinePosition`` below the (drawn) baseline and
      ``post.underlineThickness`` thick (at least 1), each scaled by the size Word draws
      the glyphs at -- whole device px -- and rounded; from the first glyph's pen x to the
      last's advance, rounded, trailing spaces at a line's end left out -- 40 / 40;
      ``words`` the same under each word -- ;
    * **strikethrough**: ``OS/2.yStrikeoutPosition`` above the baseline, ``yStrikeoutSize``
      thick, likewise -- 40 / 40; **double** one such line above and one below it -- 40 / 40;
    * **highlight** and **run shading**: the line's text rows (:func:`text_band`) across
      the span's pen extent rounded -- 20 / 20 under every line rule;
    * **thick** and **double** underlines are *approximate* (ROADMAP.md, "Phase 5 --
      measured"): the band from half the underline position (rounded up) to the single
      underline's middle, at least 5 px tall (23 / 40 exact), a double's two lines each
      half the single's thickness at its edges;
      **dotted** and **dash** are the single underline's band cut into a pattern of
      whole device px from the span's first pen x rounded, the last piece cut at its
      end: 6 px on, 6 off; 16 on, 8 off -- at 11 and 20 pt in Calibri and Times New Roman
      alike (``make_render_probe.py``: 8 / 8 underlines, every piece); the other patterns
      are drawn solid.
    """
    out: list[Rule] = []
    extent = _extent(span, line_end=line_end)
    if extent is None:
        return out
    x0, x1 = Fraction(round_half_up(extent[0])), Fraction(round_half_up(extent[1]))
    if span.highlight or span.shading:
        if band is None:
            band = (span.line_baseline - round_half_up(span.line_above),
                    span.line_baseline + round_half_up(span.line_below))
        fx0, fx1 = round_half_up(span.xs[0]), round_half_up(span.end)
        color = span.highlight or span.shading
        out.append(Rule("highlight" if span.highlight else "shading", span.path, Fraction(fx0), Fraction(band[0]),
                        Fraction(fx1 - fx0), Fraction(band[1] - band[0]), color))
    if decorations is None:
        return out
    scale = size_px / decorations.units_per_em
    if span.underline and decorations.underline_position is not None:
        position = -decorations.underline_position * scale
        thick = max(1, round_half_up(decorations.underline_thickness * scale))
        top = span.y + round_half_up(position)
        style = span.underline
        if style == "words":
            for a, b in _words(span, line_end=line_end):
                out.append(Rule("underline", span.path, Fraction(round_half_up(a)), Fraction(top),
                                Fraction(round_half_up(b) - round_half_up(a)), Fraction(thick), span.color, style))
        elif style in ("double", "thick") or style.endswith("Heavy") or style == "wavyDouble":
            first_offset = math.ceil(position / 2)
            first = span.y + first_offset
            exact_thick = decorations.underline_thickness * scale
            height = max(5, round_half_up(position + exact_thick / 2) - first_offset)
            if style == "double" or style == "wavyDouble":
                half = math.ceil(Fraction(thick, 2))
                out.append(Rule("underline", span.path, x0, Fraction(first), x1 - x0, Fraction(half), span.color,
                                style))
                out.append(Rule("underline", span.path, x0, Fraction(first + height), x1 - x0, Fraction(half),
                                span.color, style))
            else:
                out.append(Rule("underline", span.path, x0, Fraction(first), x1 - x0, Fraction(height), span.color,
                                style))
        elif style in UNDERLINE_DASHES:
            on, period = UNDERLINE_DASHES[style]
            start = x0
            while start < x1:
                end = min(start + on, x1)
                out.append(Rule("underline", span.path, start, Fraction(top), end - start, Fraction(thick), span.color,
                                style))
                start += period
        else:
            out.append(Rule("underline", span.path, x0, Fraction(top), x1 - x0, Fraction(thick), span.color, style))
    if (span.strike or span.double_strike) and decorations.strikeout_position is not None:
        thick = max(1, round_half_up(decorations.strikeout_size * scale))
        top = span.y - round_half_up(decorations.strikeout_position * scale)
        if span.double_strike:
            out.append(Rule("strike", span.path, x0, Fraction(top - thick), x1 - x0, Fraction(thick), span.color,
                            "double"))
            out.append(Rule("strike", span.path, x0, Fraction(top + thick), x1 - x0, Fraction(thick), span.color,
                            "double"))
        else:
            out.append(Rule("strike", span.path, x0, Fraction(top), x1 - x0, Fraction(thick), span.color, "single"))
    return out


def _words(span: Span, *, line_end: bool) -> list[tuple[Fraction, Fraction]]:
    """The pen extent of every word of a span (spaces left out)."""
    out = []
    start = None
    for k, char in enumerate(span.chars):
        end = span.xs[k + 1] if k + 1 < len(span.chars) else span.end
        if char.isspace():
            if start is not None:
                out.append((start, span.xs[k]))
                start = None
        elif start is None:
            start = span.xs[k]
    if start is not None:
        out.append((start, span.end))
    del end
    return out


#: Fields Word computes again when it exports (``make_field_probe.py``: each cached as a
#: result Word could not have computed, and drawn otherwise): their cached result, which
#: is what is drawn, can differ from Word's page.  ``AUTHOR``, ``FILENAME``, ``TITLE``,
#: ``NUMWORDS``, ``=``, ``HYPERLINK``, ``DOCPROPERTY``, ``QUOTE`` and ``TOC`` were drawn as
#: cached.
RECOMPUTED_BY_WORD = frozenset({"DATE", "TIME", "REF", "PAGEREF", "SEQ", "IF"})
#: The computed fields whose text depends on the pages: computed again with each layout.
PAGE_FIELDS = frozenset({"PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF"})


def _all_paragraphs(blocks) -> list[Paragraph]:
    """Every paragraph of ``blocks`` in ``tools/baselines.blocks`` order (cells included)."""
    out: list[Paragraph] = []

    def walk(item):
        if isinstance(item, Table):
            for row in item.rows:
                for cell in row:
                    for inner in cell:
                        walk(inner)
        else:
            out.append(item)

    for block in blocks:
        walk(block)
    return out


def substitute_blocks(blocks, text_of) -> list:
    """``blocks`` with each computed field's run holding ``text_of(paragraph index, run
    index, run)`` (``baselines.blocks`` order) -- its cached result where that is ``None``."""
    import dataclasses

    counter = [0]

    def paragraph(item: Paragraph) -> Paragraph:
        block = counter[0]
        counter[0] += 1
        if not any(run.field is not None for run in item.runs):
            return item
        runs = []
        for k, run in enumerate(item.runs):
            text = text_of(block, k, run) if run.field is not None else None
            runs.append(dataclasses.replace(run, text=text) if text is not None and text != run.text else run)
        return dataclasses.replace(item, runs=tuple(runs))

    def walk(item):
        if isinstance(item, Table):
            rows = tuple(tuple(tuple(walk(inner) for inner in cell) for cell in row) for row in item.rows)
            return dataclasses.replace(item, rows=rows)
        return paragraph(item)

    return [walk(block) for block in blocks]


def _footnote_pages(items, result) -> dict:
    """Footnote id -> the page its reference's line is on."""
    out = {}
    for index, item in enumerate(items):
        if isinstance(item, paginate.Para) and item.notes:
            found = result.pages.get(index, [])
            for line, laid in item.notes.items():
                if line < len(found):
                    for lines in laid:
                        if lines:
                            out.setdefault(lines[0].note, found[line])
    return out


def substitute_document(document: Document, values: dict) -> Document:
    """The document with its body's computed fields holding ``values[(block, run)]``."""
    import dataclasses

    body = substitute_blocks(document.body, lambda block, k, run: values.get((block, k)))
    return dataclasses.replace(document, body=body, paragraphs=[b for b in body if isinstance(b, Paragraph)])


def lay_out(document: Document, advances, metrics, *, package: bytes | None = None, decorations=None) -> Layout:
    """Every page of ``document``, placed: see the module docstring."""
    return _Placer(document, package, advances, metrics, decorations).run()

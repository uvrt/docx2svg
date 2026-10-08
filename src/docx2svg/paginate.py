"""Where Word ends a page: which line starts each page, from the file alone.

ROADMAP.md, "Phase 4 -- measured", holds the probes behind every rule here, their
scores and the hypotheses they refuted.  The model:

* **The flow.**  Every top-level paragraph, in order, as its lines (the model's line
  breaker, :mod:`docx2svg.linebreak`) with each line's pitch (:mod:`docx2svg.lines`,
  :mod:`docx2svg.vertical`), the space around it and its keep properties.  A table, a
  drawing the model cannot size, or a paragraph the breaker cannot measure is an
  :class:`Obstacle`: nothing after it can be placed, and :func:`paginate` stops there
  (:class:`Unplaceable`) rather than guess.
* **The stack** is the baseline model's: a page starts at its section's top margin; a
  paragraph that starts a page keeps its space before only where
  :func:`vertical.keeps_space_before_at_page_top` says; between paragraphs the gap is
  :func:`vertical.paragraph_gap_px`, contextual spacing and autospacing applied; a
  bottom border follows the last line.  Everything is exact, in device px on the
  1/4096 pt grid.
* **The break** (:func:`fits`): a line fits when the bottom of its pitch is at or above
  the bottom margin -- :data:`FIT` says which part of its box counts.
* **The keeps** (:func:`_break_before`): widow and orphan control, ``keepLines``,
  ``keepNext`` chains, ``pageBreakBefore``, manual page breaks and section breaks.
* **Columns** (:func:`_lay_columns`): a section of several text columns fills them in
  turn, each as a page under the keeps; one a continuous section follows on its page is
  balanced (:func:`_balance`), and the stack goes on under it (:func:`region_foot`).
  Footnotes stand at each column's foot below mode 15 (:func:`_scan`); in mode 15 they run
  on through the columns at the page's foot (:func:`note_area`), every column ending
  above them (:func:`_lay_columns_page`).

Standard library only.
"""

from __future__ import annotations

import dataclasses
import math
from dataclasses import dataclass, field
from fractions import Fraction

from . import lines as line_model
from . import table as table_model
from .linebreak import Line as BrokenLine, ListCounters, column_boxes, is_multi_column
from .model import Document, Paragraph, Table
from .resolve import autospaced, resolve_paragraph
from .vertical import (
    autospace_kept, keeps_space_before_at_page_top, page_top_gap_px, paragraph_borders_px, paragraph_gap_px,
    twips_to_px,
)

#: What of a line must be above the bottom margin for the line to stay on the page
#: (:func:`fits`): its pitch less an ``auto`` multiple's extra, and on a paragraph's last
#: line its bottom border -- ``text-border``, measured by ``make_page_fit_probe.py``.
#: Refuted alternatives, kept for scoring: ``pitch`` (the whole pitch, no border),
#: ``pitch-border``, ``box`` (the pitch, border and space after) and ``text`` (no border).
FIT = "text-border"


class Unplaceable(Exception):
    """The flow reached something the model cannot size; ``args[0]`` is the flow index,
    ``args[1]`` (when given) the line -- a table's row -- it stopped at."""


@dataclass
class Obstacle:
    """A top-level item whose height is not modelled: a table, a drawing the model
    cannot size, a paragraph the line breaker cannot measure."""

    block: int
    reason: str
    section: int
    #: The paragraph (``None`` for a table), for its keep properties.
    paragraph: Paragraph | None = None
    #: What about it is not modelled, where the reason alone does not say (a table's
    #: vertical merge, a floating table...).
    detail: str | None = None
    #: A table's row the model stopped at (``None``: the whole item).
    row: int | None = None


@dataclass
class Para:
    """One top-level paragraph as the paginator sees it."""

    #: Index in ``tools/baselines.blocks`` order (table cells counted), for scoring.
    block: int
    section: int
    heights: list[line_model.LineHeight]
    #: Lines ended by a ``w:br w:type="page"`` (or a column break in a one-column section).
    page_breaks: frozenset[int] = frozenset()
    #: Lines ended by a ``w:br w:type="column"`` in a section of several columns: each
    #: ends its column (:class:`Frame`).
    column_breaks: frozenset[int] = frozenset()
    before: int = 0
    after: int = 0
    contextual: bool = False
    style: str | None = None
    #: What the paragraph's borders take below its last line and above its first
    #: (``vertical.paragraph_borders_px``: one box with its neighbours when theirs match).
    border: Fraction = Fraction(0)
    top_border: Fraction = Fraction(0)
    keep_next: bool = False
    keep_lines: bool = False
    page_break_before: bool = False
    #: ``None`` when neither the paragraph nor its styles state it (on: :class:`Rules`).
    widow_control: bool | None = None
    section_start: bool = False
    #: Lines that hold nothing Word draws: no character but spaces, and not the
    #: paragraph's last (which draws its mark) -- the empty line before a page break at
    #: a paragraph's start.  They are placed like any line; a scorer matching drawn
    #: lines leaves them out (:func:`drawn_pages`).
    undrawn: frozenset[int] = frozenset()
    #: Line -> the footnotes referenced on it, each as its laid-out lines.
    notes: dict = field(default_factory=dict)
    #: The floating drawings text does not wrap around (:func:`anchors_of`).
    anchors: tuple = ()
    #: An empty paragraph holding a section break (a ``w:sectPr`` other than the body's):
    #: it never goes to the next page, and takes no room on its own (:func:`section_mark_room`).
    section_mark: bool = False
    #: In a document with drawings text wraps beside, what the paragraph's lines are
    #: broken again from where one is beside a drawing (:class:`ParaSource`); its lines
    #: as the paginator last placed them, and the first of them that is broken as a line
    #: in the whole column is, with every line after it.
    source: "ParaSource | None" = None
    lines: list = field(default_factory=list)
    plain_from: int = 0
    #: An endnote's paragraph (:mod:`docx2svg.notes`): the line of the continuation
    #: separator a page it starts goes under (:func:`continuation_line`).
    continuation: "line_model.LineHeight | None" = None
    #: The drop cap that drops into this paragraph (:class:`DropCap`).
    drop_cap: "DropCap | None" = None
    #: Lines that end in an automatic hyphen (``w:autoHyphenation``; :func:`hyphen_at_page_end`).
    auto_hyphen_ends: frozenset[int] = frozenset()

    @property
    def count(self) -> int:
        return len(self.heights)

    def set_lines(self, lines: list, plain_from: int) -> bool:
        """Take ``lines`` as the paragraph's (:class:`ParaSource`: every quantity that
        follows from them again); ``False`` where a line's height cannot be measured."""
        derived = self.source.derive(lines)
        if derived is None:
            return False
        self.lines = list(lines)
        self.heights, self.page_breaks, self.undrawn, self.notes, self.anchors, self.column_breaks = derived
        self.auto_hyphen_ends = _auto_hyphen_ends(self.source.pieces, lines)
        self.plain_from = plain_from
        return True

    def fit_width(self, line: int, width: int) -> None:
        """Break the paragraph again from line ``line`` on for a column ``width`` twips
        wide, where its lines were broken for another (one of unequal columns): a line
        is as wide as the column it is in."""
        if self.source is None or self.source.section is None or line >= len(self.lines) or \
                self.source.width_of(self.lines[line]) == width:
            return
        start = self.lines[line].start
        self.set_lines(self.lines[:line] + self.source.rest(start, first=line == 0, width=width), self.plain_from)

    def move_hyphenated_word(self, line: int) -> bool:
        """Line ``line`` ends a page in an automatic hyphen and Word moves only the word
        (:func:`hyphen_at_page_end`, ``"word"``): the line ends before the word, and the
        paragraph is broken again from it.  ``False`` where the word is the whole line."""
        broken = self.lines[line]
        pieces = self.source.pieces
        k = broken.end - 1  # the automatic hyphen
        while k > broken.start and not pieces[k - 1].char.isspace():
            k -= 1
        if k <= broken.start:
            return False
        kept = BrokenLine(broken.start, k)
        return self.set_lines(self.lines[:line] + [kept] + self.source.rest(k, first=False),
                              max(self.plain_from, line + 1))

    def plain(self, line: int) -> None:
        """Break the paragraph again from line ``line`` on as lines of the whole column,
        where an earlier placement broke them beside a drawing."""
        if self.source is None or self.plain_from <= line:
            return
        start = self.lines[line].start if line < len(self.lines) else len(self.source.pieces)
        self.set_lines(self.lines[:line] + self.source.rest(start, first=line == 0), line)


@dataclass
class TableItem:
    """A top-level table as the paginator sees it: its rows, each placed whole
    (:mod:`docx2svg.table`).  ``count`` is its number of rows, as a :class:`Para`'s is of
    lines, so a :class:`Position` inside it names a row."""

    #: The index of its first cell paragraph in ``tools/baselines.blocks`` order.
    block: int
    section: int
    flow: object
    section_start: bool = False
    # What the stack asks of any item, none of which a table has.
    before: int = 0
    after: int = 0
    border: Fraction = Fraction(0)
    top_border: Fraction = Fraction(0)
    contextual: bool = False
    style: str | None = None
    keep_next: bool = False
    keep_lines: bool = False
    page_break_before: bool = False
    widow_control: bool | None = None
    page_breaks: frozenset = frozenset()
    undrawn: frozenset = frozenset()
    notes: dict = field(default_factory=dict)
    #: The table laid out again in a narrower column (``(section) -> TableFlow``), for a
    #: room beside a drawing (:func:`table_beside`); ``None`` where it cannot be.
    reflow: object = None
    #: A page's first position -> where the table goes beside the drawings text wraps
    #: around on that page (:class:`TableBeside`), as the paginator last placed it.
    beside: dict = field(default_factory=dict)
    #: ``(left, width)`` twips -> the table laid out in that room (:meth:`room`).
    rooms: dict = field(default_factory=dict)

    @property
    def count(self) -> int:
        return len(self.flow.rows)

    @property
    def floating(self):
        """Where the table floats (:class:`docx2svg.table.Floating`); ``None`` in the flow."""
        return getattr(self.flow, "floating", None)

    def room(self, document: Document, left, width):
        """The table laid out in a column ``left`` twips right of its own and ``width``
        twips wide (``None`` where it cannot be)."""
        key = (left, width)
        if key not in self.rooms:
            self.rooms[key] = None
            if self.reflow is not None and left == int(left) and width == int(width):
                section = document.sections[self.section]
                m = section.margins
                new_left = m.left + int(left)
                margins = dataclasses.replace(m, left=new_left, right=section.page_size.width_twips - new_left
                                              - m.gutter - int(width))
                try:
                    self.rooms[key] = self.reflow(dataclasses.replace(section, margins=margins, column_count=1,
                                                                      columns=()))
                except table_model.Unsupported:
                    pass
        return self.rooms[key]

    def row_blocks(self, row: int) -> list[tuple[int, int]]:
        """``(block, lines)`` of every paragraph of a row's cells."""
        return [(paragraph.block, len(paragraph.lines)) for cell in self.flow.rows[row].cells
                for paragraph in cell.paragraphs]


#: The wraps text goes beside a drawing in (:mod:`docx2svg.wrap`).
SIDE_WRAPS = ("wrapSquare", "wrapTight", "wrapThrough")


@dataclass
class ParaSource:
    """What a paragraph's lines are broken from, kept so that the paginator can break them
    again where a line is beside a drawing (:mod:`docx2svg.wrap`), and what follows from
    its lines -- their heights, page breaks, undrawn lines, footnotes and anchors -- be
    derived again exactly as :func:`flow` derives them."""

    document: Document
    paragraph: Paragraph
    pp: object
    pieces: list
    geometry: object
    advances: object
    metrics: object
    cache: dict
    one_column: bool
    #: ``(text index, the footnote's laid-out lines)`` for every footnote reference.
    notes: list = field(default_factory=list)
    chars: dict = field(default_factory=dict)
    #: The section, where its columns are of unequal widths: a line broken for a column
    #: other than the first is a :class:`ColumnLine` (:meth:`rest`).
    section: object = None

    def __post_init__(self) -> None:
        from .resolve import resolve_run

        index = 0
        for run in self.paragraph.runs:
            resolved = resolve_run(self.document, self.paragraph, run)
            if not resolved.get("vanish"):
                for offset, char in enumerate(run.text):
                    self.chars[index + offset] = (resolved, char)
            index += len(run.text)

    def rest(self, start: int, *, first: bool, width: int | None = None) -> list:
        """The lines the paragraph's rest from piece ``start`` breaks into -- in a column
        ``width`` twips wide where given, each a :class:`ColumnLine` carrying that
        column's geometry."""
        from .linebreak import break_rest, geometry as line_geometry

        if width is None or self.section is None:
            return break_rest(self.pieces, start, self.geometry, self.advances, first_line=first,
                              columns=not self.one_column)
        geometry = line_geometry(self.document, self.paragraph, self.section, self.pp, width=width)
        lines = break_rest(self.pieces, start, geometry, self.advances, first_line=first, columns=not self.one_column)
        return [ColumnLine(line.start, line.end, line.forced, line.emergency, line.hyphenated, geometry=geometry,
                           width=width) for line in lines]

    def width_of(self, line) -> int:
        """The width (twips) of the column ``line`` was broken for."""
        from .linebreak import column_width_twips

        if self.section is None:
            return 0
        return getattr(line, "width", None) or column_width_twips(self.section)

    def height(self, line, first: bool, chars=None):
        """One line's height (:func:`lines.line_height`, or the object line's)."""
        from .linebreak import BREAK, OBJECT, TAB

        pieces = self.pieces[line.start:line.end]
        if chars is None:
            chars = [self.chars[p.source] for p in pieces
                     if p.source >= 0 and p.kind not in (TAB, BREAK, OBJECT) and p.source in self.chars]
        objects = [p.height for p in pieces if p.kind == OBJECT]
        if objects:
            runs = sorted({p.run for p in pieces if p.kind == OBJECT})
            return line_model.object_line_height(self.document, self.paragraph, self.pp, chars, self.metrics,
                                                 objects, first_line=first, runs=runs)
        return line_model.line_height(self.document, self.paragraph, self.pp, chars, self.metrics,
                                      first_line=first, cache=self.cache)

    def derive(self, broken: list):
        """``(heights, page breaks, undrawn lines, notes, anchors)`` of ``broken``, as
        :func:`flow` derives them; ``None`` where a height cannot be measured."""
        shares = line_model.shares_of(self.document, self.paragraph, self.pieces, broken)
        heights = []
        for number, (line, chars) in enumerate(zip(broken, shares)):
            height = self.height(line, number == 0, chars)
            if height is None:
                return None
            heights.append(height)
        pieces = self.pieces
        page_breaks = frozenset(
            number for number, line in enumerate(broken)
            if line.forced and pieces[line.end - 1].char in ("\f",) + (("\v",) if self.one_column else ()))
        column_breaks = frozenset(
            number for number, line in enumerate(broken)
            if line.forced and not self.one_column and pieces[line.end - 1].char == "\v")
        undrawn = frozenset(number for number, chars in enumerate(shares)
                            if number < len(shares) - 1 and all(char.isspace() for _, char in chars))
        notes: dict[int, list] = {}
        line_of = _line_of_sources(pieces, broken)
        for source, laid in self.notes:
            line = line_of.get(source, line_of.get(source - 1, len(broken) - 1))
            notes.setdefault(line, []).append(laid)
        anchors = anchors_of(self.paragraph, line_of, len(broken))
        return heights, page_breaks, undrawn, notes, anchors, column_breaks


@dataclass
class ColumnLine(BrokenLine):
    """A line broken for a column other than its paragraph's first (one of unequal
    columns): a line with the column's geometry, which the renderer sets it in."""

    geometry: object = None
    width: int = 0


@dataclass(frozen=True, order=True)
class Position:
    """A line of the flow: ``item`` indexes the flow, ``line`` the item's lines -- a
    table's rows -- and ``part``, inside a table row split across pages, each cell's first
    stacked line on the page (:func:`docx2svg.table.place_table`)."""

    item: int
    line: int = 0
    part: tuple = ()


@dataclass(frozen=True)
class PageInfo:
    """What a page is, beyond its lines -- **public**: every laid-out page carries one as
    :attr:`docx2svg.layout.Page.info` (also importable as ``docx2svg.PageInfo``), and it is
    what a ``PAGE``-like result needs (``make_story_select_probe.py``, ROADMAP.md H.2).

    * ``number`` -- the page's number as ``PAGE`` shows it, before formatting: the
      section's ``w:pgNumType/@w:start`` where it restarts, else the page before's plus
      one, blank pages counted.  Format it with the section's ``page_number_format`` (and a
      field's switches) by :func:`docx2svg.fields.field_text`.
    * ``section`` -- the 0-based index into ``Document.sections`` of the section the page
      belongs to: its first line's (a continuous section starting mid-page does not
      change it).
    * ``section_first`` -- whether that section starts on this page.
    * ``story`` -- the header and footer kind the page shows: ``"first"``, ``"even"`` or
      ``"default"`` (:func:`story_kind`).
    * ``blank`` -- a page Word inserts (an ``oddPage`` / ``evenPage`` section on the wrong
      parity): it has a number, counts in ``NUMPAGES``, and shows nothing.

    ``Page.number`` is the page's 0-based place in the document; ``info.number`` is the
    number Word prints on it.  A page the layout made after it stopped has no info
    (``None``)."""

    number: int
    section: int
    section_first: bool
    story: str
    blank: bool = False


@dataclass
class Pagination:
    #: The flow position each page starts at.
    starts: list[Position] = field(default_factory=list)
    #: Every page, blank ones included, by its 0-based place in the document.
    info: dict[int, PageInfo] = field(default_factory=dict)
    #: Flow index -> the page of each of its lines.
    pages: dict[int, list[int]] = field(default_factory=dict)
    #: Where the model stopped, when it met an :class:`Obstacle` (flow index).
    stopped: int | None = None
    #: The line (a table's row) of the item at ``stopped`` it stopped at: 0 for an
    #: obstacle, a row for a table whose row it cannot place.
    stopped_line: int = 0
    #: Every page's start and end, in order, for the page walk.
    ends: list[Position] = field(default_factory=list)
    #: Page -> the bands its ``wrapTopAndBottom`` drawings keep text off (:class:`Band`).
    bands: dict[int, list] = field(default_factory=dict)
    #: Page -> the drawings text wraps beside on it (:class:`docx2svg.wrap.Wrap`).
    wraps: dict[int, list] = field(default_factory=dict)
    #: Page -> its footnotes at the foot: ``(the lines of a note the page before could
    #: not hold, [(a note's lines, how many of them are on this page)])``.
    notes: dict[int, tuple] = field(default_factory=dict)
    #: Page -> its columns, in the flow's order (:class:`Frame`), in a document with a
    #: section of several columns.
    frames: dict[int, list] = field(default_factory=dict)


@dataclass
class Frame:
    """One text column of a page, as laid out: the flow from ``start`` to ``end`` stacked
    from ``top`` down to where it ends (``y``), above ``bottom``.  A region is a run of
    one section's columns side by side (a one-column section's is one frame); a page
    holds regions one under the other where continuous sections share it
    (:func:`_lay_columns`).  ``why`` says what ended it: ``full``, a ``column`` or ``page``
    break, the ``region`` (a continuous section of other columns), the ``end`` of the
    flow."""

    section: int
    column: int
    top: Fraction
    bottom: Fraction
    #: Whether the column is empty above its first line (its first line is placed
    #: whatever its height; its space before as at a page's top).
    first_on_page: bool = True
    #: The item before its first, where it goes on from a region above (``None`` at a
    #: column's top).
    previous: object = None
    #: Whether its section has several columns.
    multi: bool = False
    #: Its column's offset from the text area's left edge and its width, twips.
    offset: int = 0
    width: int = 0
    #: Whether it starts at the page's top.
    page_top: bool = False
    start: Position = Position(0)
    end: Position = Position(0)
    y: Fraction = Fraction(0)
    why: str = ""
    #: The top of its region, and whether the region was balanced (:func:`_balance`).
    region_top: Fraction = Fraction(0)
    balanced: bool = False
    #: The top of its first paragraph, past the space between it and the one before.
    first_top: Fraction | None = None
    #: The item the stack's next paragraph spaces itself from (a paragraph holding only a
    #: section break is passed over).
    last: object = None
    #: Below mode 15, the footnotes at the column's own foot (:func:`_scan`'s
    #: ``(lines, how many of them are in the column)``), and the lines of notes the column
    #: before (or the page before) could not hold, which open it under the continuation
    #: separator (``make_columns_probe.py``, ``notesover``, ``notes3``).  In mode 15 the
    #: notes referenced in the column, every line of them: the page's notes run on through
    #: its columns at its foot (:func:`note_area`).
    notes: list = field(default_factory=list)
    carried: tuple = ()
    #: In mode 15, a column of several whose notes go to the page's note area.
    pooled: bool = False


@dataclass(frozen=True)
class Band:
    """What a drawing text wraps top and bottom of keeps lines off: device px from the
    page's top, its box widened by ``distT`` above and ``distB`` below; and the
    drawing's own top, where it was positioned before any line moved (``key``: the
    paragraph's block, the anchor's text index, its run and its index in the run)."""

    top: Fraction
    bottom: Fraction
    y: Fraction
    key: tuple
    #: Its drawing's left and right edges where it is in a section of several columns
    #: (``None``: across the page's text).
    span: tuple | None = None


def clear_of_bands(y: Fraction, pitch: Fraction, bands) -> Fraction:
    """Where a line that would start at ``y`` and reach ``pitch`` down (its pitch, and on a
    paragraph's last line its space after) starts: below every band it would reach into
    (a line touching a band's edge does not), until none (ROADMAP.md, "Floating drawings
    -- measured", F.7)."""
    moved = True
    while moved:
        moved = False
        for band in bands:
            if y < band.bottom and y + pitch > band.top:
                y = band.bottom
                moved = True
    return y


class Flow(list):
    """The flow's items, with what the page's edges need: per section, the height of the
    footnote separator and of the continuation separator (``separators``), and of its
    header and footer on a page and on its first page (``stories``)."""

    separators: dict
    stories: dict
    #: Whether a paragraph anchors a drawing text wraps beside: its paragraphs keep what
    #: their lines are broken from (:class:`ParaSource`).
    wraps: bool
    #: Whether a section has more than one text column: its pages are laid out by
    #: column (:func:`_lay_columns`).
    columns: bool
    #: ``(note id, section, width in twips) ->`` a footnote's lines broken in a column of
    #: that width (``None`` where it cannot be measured): what a column of unequal ones
    #: holds of it at its foot.
    note_at: object
    #: ``(advances, metrics)`` the flow was measured with: a page's note area in mode 15
    #: (:func:`note_area`) is a flow of its own.
    measure: tuple
    #: ``(section, note ids) ->`` :class:`NoteArea` (``None``: cannot be laid out).
    areas: dict

    def __init__(self, items=()):
        super().__init__(items)
        self.separators = {}
        self.stories = {}
        self.wraps = False
        self.columns = False
        self.note_at = None
        self.measure = ()
        self.areas = {}


def story_document(document: Document, blocks: list, section) -> Document:
    """A header or footer as a document of its own: its blocks as the body of one section
    with the page's size and margins and **one text column** (a story spans the text
    width however many columns the section has), so the paginator's flow and the
    renderer's page walk lay it out exactly as they lay out a body."""
    import dataclasses

    paragraphs = [block for block in blocks if isinstance(block, Paragraph)]
    one = dataclasses.replace(section, first_paragraph=0, last_paragraph=len(paragraphs), column_count=1,
                              columns=(), start_type=None, headers=(), footers=(), title_page=False)
    return dataclasses.replace(document, body=list(blocks), paragraphs=paragraphs, sections=[one], footnotes={},
                               footnote_separators={}, stories={}, story_parts={})


def floating_marks(blocks: list, *, keep: bool = False) -> tuple[list, list[str]]:
    """A story's blocks with the mark of every floating drawing that text does not wrap
    around (``wp:anchor`` with ``wrapNone``: behind or in front of the text) taken out --
    such a drawing takes no room in its line or its story -- and what was taken out, as
    warnings; with ``keep``, those marks stay (the layout draws their drawings) and
    nothing is taken out.  A drawing text wraps around becomes a ``drawing`` mark, and
    makes its paragraph an obstacle."""
    import dataclasses

    taken: list[str] = []

    def paragraph(item: Paragraph) -> Paragraph:
        runs = []
        for run in item.runs:
            breaks = []
            for position, kind in run.breaks:
                if kind.startswith("anchor:"):
                    k = int(kind.split(":")[1])
                    if k < len(run.anchors) and not run.anchors[k].moves_text:
                        if keep:
                            breaks.append((position, kind))
                            continue
                        taken.append(item.path)
                        continue
                    kind = "drawing"
                breaks.append((position, kind))
            runs.append(dataclasses.replace(run, breaks=tuple(breaks)) if tuple(breaks) != run.breaks else run)
        return dataclasses.replace(item, runs=tuple(runs))

    def walk(item):
        if isinstance(item, Table):
            rows = tuple(tuple(tuple(walk(inner) for inner in cell) for cell in row) for row in item.rows)
            return dataclasses.replace(item, rows=rows)
        return paragraph(item)

    return [walk(block) for block in blocks], taken


def story_flow(document: Document, blocks: list, section, advances, metrics) -> tuple[Document, list]:
    """A header's or footer's flow (:func:`flow` over :func:`story_document`)."""
    blocks, _ = floating_marks(blocks)
    story = story_document(document, blocks, section)
    return story, flow(story, advances, metrics)


def stack_height(document: Document, items: list) -> Fraction:
    """How far a story reaches from its edge distance: its items stacked as a page stacks
    them from its top -- the first paragraph's space before kept (a story's first
    paragraph starts its section), the gaps and borders between them, a table's rows --
    and the last paragraph's space after.  Measured (``make_header_probe.py``): a header
    of two lines with 240 twips after starts the body 57 twips below the margin, one line
    with 480 before 29.  An obstacle ends the stack."""
    y = Fraction(0)
    previous = None
    first = True
    for item in items:
        if isinstance(item, Obstacle):
            break
        if isinstance(item, TableItem):
            if not first:
                y += table_gap_px(previous, item)
            laid = table_model.place_table(item.flow, y, y + 10 ** 7, row=0, part=(), first_on_page=first,
                                           mode15=(document.compatibility_mode or 0) >= 15)
            y = laid.end
            first = False
            previous = item
            continue
        for number in range(item.count):
            if number == 0:
                if first:
                    keeps = keeps_space_before_at_page_top(section_start=True,
                                                           compatibility_mode=document.compatibility_mode)
                    gap, _ = page_top_gap_px(item.before, keeps=keeps)
                elif isinstance(previous, TableItem):
                    gap = twips_to_px(item.before)
                else:
                    same = item.style == previous.style
                    own_before = 0 if (item.contextual and same) else item.before
                    prev_after = 0 if (previous.contextual and same) else previous.after
                    gap, _ = paragraph_gap_px(prev_after, own_before)
                y += gap + item.top_border
            y += item.heights[number].pitch
            first = False
        y += item.border
        previous = item
    if isinstance(previous, Para):
        y += twips_to_px(previous.after)
    return y


def story_height(document: Document, blocks: list, section, advances, metrics, cache: dict | None = None
                 ) -> Fraction | None:
    """How far a header or footer reaches from its edge distance (:func:`stack_height` of
    its :func:`story_flow`); ``None`` when its first item cannot be laid out."""
    story, items = story_flow(document, blocks, section, advances, metrics)
    if items and isinstance(items[0], Obstacle):
        return None
    return stack_height(story, items)


#: The stories a page may take, in ``w:headerReference/@w:type``'s words.
STORY_KINDS = ("default", "first", "even")


def story_references(document: Document) -> list[tuple[dict, dict]]:
    """Per section, its header and footer relationship id of each type: what it states,
    and what it does not state inherited from the section before -- type by type, whatever
    the section before used (``make_story_select_probe.py``, ``inherit``: a section under
    ``w:titlePg`` stating no ``first`` shows the ``first`` of the last section that stated
    one, although that section did not use it)."""
    out = []
    headers: dict[str, str] = {}
    footers: dict[str, str] = {}
    for section in document.sections:
        headers = {**headers, **dict(section.headers)}
        footers = {**footers, **dict(section.footers)}
        out.append((dict(headers), dict(footers)))
    return out


def story_kind(document: Document, section: int, *, section_first: bool, number: int | None) -> str:
    """Which story a page shows (``make_story_select_probe.py``, every page of six
    documents): ``first`` on the first page of a section under ``w:titlePg``; else
    ``even`` on a page whose **number** is even (restarts counted) when the settings state
    ``w:evenAndOddHeaders``; else ``default``.  A kind no section has stated is drawn empty
    -- not the ``default`` -- (``undefined``)."""
    if section_first and document.sections[section].title_page:
        return "first"
    if document.even_and_odd_headers and number is not None and number % 2 == 0:
        return "even"
    return "default"


def _stories(document: Document, advances, metrics, cache: dict) -> dict:
    """Section -> story kind -> (header, footer) height (:func:`story_height`)."""
    out = {}
    heights: dict = {}
    for number, (headers, footers) in enumerate(story_references(document)):
        section = document.sections[number]

        def height(references: dict, kind: str):
            relationship = references.get(kind)
            blocks = document.stories.get(relationship) if relationship else None
            if not blocks:
                return Fraction(0)
            key = (relationship, section.page_size, section.margins)
            if key not in heights:
                heights[key] = story_height(document, blocks, section, advances, metrics, cache) or Fraction(0)
            return heights[key]

        kinds = {kind: (height(headers, kind), height(footers, kind)) for kind in STORY_KINDS}
        if any(any(pair) for pair in kinds.values()):
            out[number] = kinds
    return out


@dataclass(frozen=True)
class NoteLine:
    """One line of a footnote as the page foot stacks it: its pitch and the paragraph
    spacing above it, and what its paragraph's widow control allows."""

    height: Fraction
    #: Its line in its paragraph, of how many; whether that paragraph has widow control.
    line: int
    count: int
    widow: bool
    #: The note it is a line of (its id; ``None`` for a separator's), its paragraph in
    #: the note and its place among the note's lines: what the layout draws of it
    #: (:attr:`Pagination.notes`).
    note: int | None = None
    paragraph: int = 0
    index: int = 0


def note_spacing(before: int, after: int, compatibility_mode: int | None) -> int:
    """The twips a footnote's paragraph spacing adds to the room it takes: its first
    paragraph's space before and its last one's space after, together.

    Measured on one-paragraph notes by ``make_footnote_probe.py`` (120/0, 0/120,
    120/120, 120/240 and 240/120 twips, swept by single units): in mode 15 the larger of
    the two; below 15 and unstated, the space before less the space after, and nothing
    when that is negative (the 240/120 threshold is not monotone there, and is recorded
    as such).  Between two paragraphs of a note the body's collapse is assumed, not
    measured.
    """
    if compatibility_mode is not None and compatibility_mode >= 15:
        return max(before, after)
    return max(0, before - after)


def _note_lines(document: Document, paragraphs: list, section, advances, metrics, cache: dict,
                widow_default: bool = True, note: int | None = None) -> list[NoteLine] | None:
    """A footnote's (or a separator's) paragraphs as lines, stacked as the body stacks
    them, with the note's own spacing (:func:`note_spacing`) above its first line;
    ``None`` when a paragraph cannot be measured."""
    out: list[NoteLine] = []
    previous_after = None
    first_before = last_after = 0
    for index, paragraph in enumerate(paragraphs):
        pp = resolve_paragraph(document, paragraph)
        pieces, broken = line_model.broken(document, paragraph, section, advances, metrics)
        if pieces is None:
            return None
        shares = line_model.shares_of(document, paragraph, pieces, broken)
        # A line holding an inline picture or drawing is as tall as it, as the body's
        # (:func:`flow`) and the note as drawn are.
        objects = line_model.objects_of(pieces, broken)
        object_runs = line_model.object_runs_of(pieces, broken)
        before = pp.get("spacing.before", 0) or 0
        if index == 0:
            first_before, gap = before, Fraction(0)
        else:
            gap = paragraph_gap_px(previous_after, before)[0]
        widow = pp.get("widowControl")
        for number, chars in enumerate(shares):
            if number < len(objects) and objects[number]:
                height = line_model.object_line_height(document, paragraph, pp, chars, metrics, objects[number],
                                                       first_line=number == 0, runs=object_runs[number])
            else:
                height = line_model.line_height(document, paragraph, pp, chars, metrics, first_line=number == 0,
                                                cache=cache)
            if height is None:
                return None
            out.append(NoteLine(height.pitch + (gap if number == 0 else 0), number, len(shares),
                                widow_default if widow is None else bool(widow), note, index, len(out)))
        previous_after = last_after = pp.get("spacing.after", 0) or 0
    if out:
        spacing = twips_to_px(note_spacing(first_before, last_after, document.compatibility_mode))
        first = out[0]
        out[0] = dataclasses.replace(first, height=first.height + spacing)
    return out


def _split_allowed(lines: list[NoteLine], k: int) -> bool:
    """Whether a note may end its page part after its first ``k`` lines."""
    if k >= len(lines):
        return True
    if k <= 0:
        return False
    before, after = lines[k - 1], lines[k]
    if after.line == 0 or not after.widow:
        return True
    return before.line + 1 >= 2 and after.count - after.line >= 2


def note_minimum(lines: list[NoteLine]) -> Fraction:
    """The least of a note that must go on its reference's page: its first lines up to
    the first place widow control lets it split (all of it, if nowhere)."""
    k = next(k for k in range(1, len(lines) + 1) if _split_allowed(lines, k))
    return sum((line.height for line in lines[:k]), Fraction(0))


def note_prefix(lines: list[NoteLine], room: Fraction) -> int:
    """How many of a note's lines go in ``room``: the most that fit where it may split."""
    best, total = 0, Fraction(0)
    for k in range(1, len(lines) + 1):
        total += lines[k - 1].height
        if total > room:
            break
        if _split_allowed(lines, k):
            best = k
    return best


def column_section(section, width: int):
    """``section`` as one text column ``width`` twips wide: what a footnote in one of its
    columns is broken in (its right margin moved in)."""
    page = section.page_size.width_twips - section.margins.left - section.margins.gutter
    margins = dataclasses.replace(section.margins, right=page - width)
    return dataclasses.replace(section, margins=margins, column_count=1, columns=())


def shares_page(document: Document, section_number: int) -> bool:
    """Whether a continuous section break joins section ``section_number`` to the one
    before or after it, so that the two may share a page."""
    sections = document.sections
    if section_number > 0 and sections[section_number].start_type == "continuous":
        return True
    return section_number + 1 < len(sections) and sections[section_number + 1].start_type == "continuous"


def _blocks_of_body(document: Document) -> list[tuple[object, int]]:
    """Each top-level item with its index in ``baselines.blocks`` order."""
    out = []
    index = 0

    def count(table: Table) -> int:
        total = 0
        for row in table.rows:
            for cell in row:
                for item in cell:
                    total += count(item) if isinstance(item, Table) else 1
        return total

    for item in document.body:
        out.append((item, index))
        index += count(item) if isinstance(item, Table) else 1
    return out


def _run_offsets(paragraph: Paragraph) -> list[int]:
    """Each run's first index into the paragraph's text."""
    out, offset = [], 0
    for run in paragraph.runs:
        out.append(offset)
        offset += len(run.text)
    return out


def _line_of_sources(pieces, broken) -> dict[int, int]:
    """Text index -> the line its character is on."""
    out: dict[int, int] = {}
    for number, line in enumerate(broken):
        for piece in pieces[line.start:line.end]:
            if piece.source >= 0 and piece.kind not in ("tab", "break"):
                out.setdefault(piece.source, number)
    return out


def ends_section(document: Document, paragraph_index: int) -> bool:
    """Whether the top-level paragraph ``paragraph_index`` holds a section break: the
    ``w:sectPr`` of a section other than the last (which is the body's own)."""
    return any(section.last_paragraph - 1 == paragraph_index > section.first_paragraph - 1
               for section in document.sections[:-1])


def join_starts(paragraph: Paragraph, pieces, broken) -> list[int]:
    """The lines of ``paragraph`` (other than its first) that begin the text of a paragraph
    joined into it because the mark before was deleted (:attr:`Paragraph.joins`): the
    lines Word's final view draws lower below mode 15 (ROADMAP.md, "Revisions --
    measured", R.4)."""
    joins = paragraph.joins
    if not joins:
        return []

    def segment(run: int) -> int:
        return sum(1 for join in joins if join <= run)

    out = []
    for number in range(1, len(broken)):
        start, end = broken[number].start, broken[number].end
        first = next((piece.run for piece in pieces[start:end] if piece.run >= 0), None)
        previous = next((piece.run for piece in reversed(pieces[:start]) if piece.run >= 0), None)
        if first is not None and previous is not None and segment(first) > segment(previous):
            out.append(number)
    return out


def is_empty(paragraph: Paragraph) -> bool:
    """No text, tab, break, picture or drawing: nothing but the paragraph's mark."""
    return not any(run.text or run.breaks or run.anchors for run in paragraph.runs)


def section_mark_room(previous, compatibility_mode: int | None) -> bool:
    """Whether an empty paragraph that holds a section break takes room on its page.

    Measured by ``make_section_end_probe.py`` (192 cases, no ``settings.xml``, modes 14
    and 15): such a paragraph **stays on its page** however far past the bottom margin it
    would reach -- by one unit, by its whole height, its top 60 pt below the margin -- with
    a 11, 16 or 48 pt mark, space before or after, after a paragraph or a table, before a
    ``nextPage`` or a ``continuous`` break; a paragraph with text there, an empty one that
    does not end its section, and the document's own last paragraph go to the next page
    as any line does.  And it **takes no room**: the next section's first line, after a
    ``continuous`` break, starts where the mark would have (its space before and after
    not counted either) -- except in mode 15 straight after a table, where the next line
    does not start beside it.
    """
    return (compatibility_mode or 0) >= 15 and isinstance(previous, TableItem)


def _section_of(document: Document, paragraph_index: int) -> int:
    for number, section in enumerate(document.sections):
        if section.first_paragraph <= paragraph_index < section.last_paragraph:
            return number
    return len(document.sections) - 1


#: Device px per twip.
TWIP_PX = Fraction(300, 1440)
#: What Word adds to a drop cap's width where the text beside it starts: 103 layout units
#: (``make_drop_cap_probe.py``, every case of mode 15 to 0.0001 px; half a twip, 102.4
#: units, rounded up).
DROP_CAP_EXTRA_UNITS = 103


def is_drop_cap(pp) -> bool:
    """Whether a paragraph is a drop cap's: in a frame (``w:framePr``) whose ``w:dropCap``
    is ``drop`` or ``margin``."""
    frame = pp.get("framePr")
    return isinstance(frame, dict) and frame.get("dropCap") in ("drop", "margin")


@dataclass
class DropCap:
    """A drop cap (``w:framePr`` ``w:dropCap``) as the paragraph it drops into carries it:
    its own paragraph laid out as a line of its own (not in the stack), how wide its text
    is, and what its frame states (``w:dropCap``, ``w:hSpace``, ``w:hAnchor``).

    Measured by ``make_drop_cap_probe.py`` (ROADMAP.md, "Floating drawings -- measured",
    F.19): the frame takes no room; its top is the paragraph's top (where the one before
    ends after its own space after -- not after this one's space before), it is as tall
    as its paragraph's space before and lines, and its text is set there as a paragraph's
    is -- its size, exact line and lowering as written (Word does not size the letter
    from ``w:lines``, which counts for nothing: a line is beside the frame when its top is
    above the frame's foot).  The text beside it starts after it (:meth:`text_start`)."""

    paragraph: Paragraph
    pp: object
    pieces: list
    lines: list
    heights: list
    before: int
    #: How far its widest line reaches from the column's start, its indent included, in
    #: the layout unit.
    width: int | Fraction
    kind: str
    hspace: int
    anchor: str

    @property
    def height(self) -> Fraction:
        return twips_to_px(self.before) + sum((h.pitch for h in self.heights), Fraction(0))

    @property
    def beside(self) -> bool:
        """Whether text goes beside it: a ``drop`` cap's, and a ``margin`` cap's anchored
        to the text or the margin (then it stands at the text's start as a drop does);
        a ``margin`` cap anchored to the page stands left of the column."""
        return self.kind == "drop" or self.anchor in ("text", "margin")

    def text_start(self, column_start: int, margin_exact: Fraction, mode15: bool) -> tuple[Fraction, Fraction]:
        """Where the text beside it starts, device px on the page, and how far its text is
        then moved (``make_drop_cap_probe.py``: 36 widths, two margins, ``w:hSpace`` 144
        and 432; right-aligned and justified text beside it).

        From the column's start (a whole pixel) its width, ``w:hSpace`` and
        :data:`DROP_CAP_EXTRA_UNITS`: ``S``.  In mode 15 the text is set from ``S`` and
        moved a pixel right where the margin's exact edge is further into its pixel than
        the width is into its own.  Below it, ``S`` is on the nearest whole twip of the
        page, and the text is moved from there to that pixel, rounded half down: a
        right-aligned line two thirds of a pixel past the column's edge."""
        from .vertical import LAYOUT_UNIT_PX

        width = Fraction(self.width) * LAYOUT_UNIT_PX + twips_to_px(self.hspace)
        start = column_start + width + DROP_CAP_EXTRA_UNITS * LAYOUT_UNIT_PX
        step = 1 if width - math.floor(width) < margin_exact - math.floor(margin_exact) else 0
        if mode15:
            return start, Fraction(step)
        twips = math.floor(start / TWIP_PX + Fraction(1, 2)) * TWIP_PX
        return twips, math.ceil(twips + step - Fraction(1, 2)) - twips


def drop_cap_of(document: Document, paragraph: Paragraph, pp, section, advances, metrics, cache) -> DropCap | None:
    """``paragraph`` (a drop cap's) laid out as a :class:`DropCap`; ``None`` where it
    cannot be measured."""
    from .linebreak import _kern

    pieces, broken = line_model.broken(document, paragraph, section, advances, metrics)
    if pieces is None:
        return None
    shares = line_model.shares_of(document, paragraph, pieces, broken)
    heights = []
    for number, chars in enumerate(shares):
        height = line_model.line_height(document, paragraph, pp, chars, metrics, first_line=number == 0, cache=cache)
        if height is None:
            return None
        heights.append(height)
    from .linebreak import geometry as line_geometry

    where = line_geometry(document, paragraph, section, pp)
    widths = []
    for number, line in enumerate(broken):
        # From the column's start: its indent (the first line's) and its text.
        total: int | Fraction = where.first_start if number == 0 else where.start
        for j in range(line.start, line.end):
            piece = pieces[j]
            if piece.kind in ("tab", "break", "soft hyphen"):
                continue
            total += piece.width + (_kern(piece, pieces[j + 1], advances) if j + 1 < line.end else 0)
        widths.append(total)
    frame = pp.get("framePr") or {}
    try:
        hspace = int(frame.get("hSpace") or 0)
    except ValueError:
        hspace = 0
    return DropCap(paragraph, pp, pieces, list(broken), heights, pp.get("spacing.before", 0) or 0,
                   max(widths, default=0), frame.get("dropCap"), hspace, frame.get("hAnchor") or "page")


def flow(document: Document, advances, metrics, *, features: list[str | None] | None = None) -> list:
    """The document's top-level items as :class:`Para` and :class:`Obstacle`.

    ``features`` -- per block, why the parsed model cannot stand for a paragraph (a
    field, a note reference, a frame), as ``tools/baselines.paragraph_features`` finds
    them -- makes those paragraphs obstacles too.
    """
    counters = ListCounters()
    heights_cache: dict = {}
    out = Flow()
    note_cache: dict = {}

    def note(note_id: int, section_number: int, width: int | None = None):
        key = (note_id, section_number, width)
        if key not in note_cache:
            paragraphs = document.footnotes.get(note_id)
            section = document.sections[section_number]
            if width is not None:
                section = column_section(section, width)
            note_cache[key] = None if paragraphs is None else _note_lines(
                document, paragraphs, section, advances, metrics, heights_cache, note=note_id)
        return note_cache[key]
    out.note_at = note
    out.measure = (advances, metrics)
    continuations: dict = {}

    def continuation(section_number: int):
        if section_number not in continuations:
            continuations[section_number] = continuation_line(document, document.sections[section_number], advances,
                                                              metrics, heights_cache)
        return continuations[section_number]
    top_level = 0
    numbered: dict[int, bool] = {}
    section_seen: set[int] = set()
    items = _blocks_of_body(document)
    # Numbering counts every paragraph in order, table cells included.
    labels: dict[int, str | None] = {}
    all_paragraphs: list[tuple[int, Paragraph]] = []

    def walk(item, index):
        if isinstance(item, Table):
            for row in item.rows:
                for cell in row:
                    for inner in cell:
                        index = walk(inner, index)
            return index
        all_paragraphs.append((index, item))
        return index + 1

    for item, index in items:
        walk(item, index)
    for index, paragraph in all_paragraphs:
        pp = resolve_paragraph(document, paragraph)
        numbered[index] = bool(pp.get("numPr.numId"))
        if numbered[index]:
            labels[index] = counters.label(document, paragraph, pp)

    out.stories = _stories(document, advances, metrics, heights_cache)
    # The drop caps the flow lays out (a paragraph in a frame whose ``w:dropCap`` is
    # set, followed by a paragraph, which ``features`` does not leave out as a frame).
    blocks_of = dict((id(item), block) for item, block in items)
    body_list = list(document.body)
    drop_caps: set[int] = set()
    for k, item in enumerate(body_list):
        if isinstance(item, Paragraph) and k + 1 < len(body_list) and isinstance(body_list[k + 1], Paragraph):
            block = blocks_of[id(item)]
            reason = features[block] if features and block < len(features) else None
            if reason is None and is_drop_cap(resolve_paragraph(document, item)):
                drop_caps.add(block)
    out.wraps = any(anchor.wrap in SIDE_WRAPS for item in document.body if isinstance(item, Paragraph)
                    for run in item.runs for anchor in run.anchors) or any(
        isinstance(item, Table) and "tblpPr" in item.properties for item in document.body) or bool(drop_caps)
    out.columns = any(is_multi_column(section) for section in document.sections)
    pending_drop: DropCap | None = None
    # Every top-level paragraph resolved, for its neighbours' borders (a table between
    # two paragraphs parts them).
    resolved_body = [resolve_paragraph(document, item) if not isinstance(item, Table) else None
                     for item in document.body]
    body_index = {id(item): k for k, item in enumerate(document.body)}
    for item, block in items:
        if isinstance(item, Table):
            section_number = _section_of(document, top_level)
            try:
                position = body_index[id(item)]
                before = document.body[position - 1] if position else None
                laid = table_model.flow_table(document, item, document.sections[section_number], advances, metrics,
                                              block=block, labels=labels, features=features, cache=heights_cache,
                                              previous=before if isinstance(before, Paragraph) else None)
            except table_model.Unsupported as why:
                out.append(Obstacle(block, "table", section_number, detail=str(why.args[0])))
                continue
            section_start = section_number not in section_seen
            section_seen.add(section_number)
            if is_multi_column(document.sections[section_number]) and laid.floating is not None:
                out.append(Obstacle(block, "columns", section_number, detail="a floating table"))
                continue

            def reflow(section, item=item, block=block, before=before):
                return table_model.flow_table(document, item, section, advances, metrics, block=block, labels=labels,
                                              features=features, cache=heights_cache,
                                              previous=before if isinstance(before, Paragraph) else None)

            out.append(TableItem(block, section_number, laid, section_start=section_start,
                                 reflow=reflow if out.wraps or out.columns else None))
            continue
        section_number = _section_of(document, top_level)
        section_mark = ends_section(document, top_level) and is_empty(item)
        top_level += 1
        section = document.sections[section_number]
        section_start = section_number not in section_seen
        section_seen.add(section_number)
        pp = resolve_paragraph(document, item)
        if block in drop_caps:
            # A drop cap: not in the stack; the paragraph after it carries it.
            pending_drop = drop_cap_of(document, item, pp, section, advances, metrics, heights_cache)
            if pending_drop is not None:
                continue
            out.append(Obstacle(block, "unmeasurable", section_number, item))
            continue
        drop, pending_drop = pending_drop, None
        reason = features[block] if features and block < len(features) else None
        if reason == "note reference":
            reason = None  # laid out: the page foot takes the note (``page_end``)
        if reason is None and any(kind == "drawing" for run in item.runs for _, kind in run.breaks):
            reason = "drawing"  # VML, an embedded object: not modelled
        if reason is None and any(anchor.moves_text and anchor.wrap not in ("wrapTopAndBottom",) + SIDE_WRAPS
                                  for run in item.runs for anchor in run.anchors):
            reason = "drawing"  # a wrap this model does not know
        detail = None
        if reason is None and is_multi_column(section):
            # In columns, what is not laid out by column yet (ROADMAP.md, "Phase 4 --
            # measured", 4.14 to 4.17): a floating drawing not positioned against its
            # column, a footnote in a section a continuous break joins to another, an
            # endnote in mode 15.
            if any((anchor.h.relative or "column") != "column" or anchor.simple is not None
                   for run in item.runs for anchor in run.anchors):
                # Measured against its column only (``make_columns_probe.py``,
                # ``floating``): one against the margin, the page or a character is not.
                reason, detail = "columns", "a floating drawing not positioned against its column"
            elif any(kind.startswith("footnoteReference:") for run in item.runs for _, kind in run.breaks) \
                    and shares_page(document, section_number):
                # Not measured enough: below mode 15 Word moves such a section to a page
                # of its own, in mode 15 it sets the page's notes in one column
                # (``make_columns_probe.py``, ``notesregion``).
                reason, detail = "columns", "a footnote, in a section a continuous break joins to another,"
            elif item.note and (document.compatibility_mode or 0) >= 15:
                # Below mode 15 the endnotes go on in the columns after the text; in mode 15
                # Word balances the text above them and sets them as a balanced region of
                # their own under their separator (``make_columns_probe.py``, ``endnotes``):
                # not modelled.
                reason, detail = "columns", "an endnote"
        if reason is None:
            pieces, broken = line_model.broken(document, item, section, advances, metrics, labels.get(block))
            if pieces is None:
                reason = "unmeasurable"
        if reason is not None:
            out.append(Obstacle(block, reason, section_number, item, detail=detail))
            continue
        shares = line_model.shares_of(document, item, pieces, broken)
        objects = line_model.objects_of(pieces, broken)
        object_runs = line_model.object_runs_of(pieces, broken)
        heights = []
        for number, chars in enumerate(shares):
            if objects[number]:
                height = line_model.object_line_height(document, item, pp, chars, metrics, objects[number],
                                                       first_line=number == 0, runs=object_runs[number])
            else:
                height = line_model.line_height(document, item, pp, chars, metrics, first_line=number == 0,
                                                cache=heights_cache)
            if height is None:
                break
            heights.append(height)
        if len(heights) != len(shares):
            out.append(Obstacle(block, "no face metrics", section_number, item))
            continue
        one_column = not is_multi_column(section)
        page_breaks = frozenset(
            number for number, line in enumerate(broken)
            if line.forced and pieces[line.end - 1].char in ("\f",) + (("\v",) if one_column else ()))
        column_breaks = frozenset(
            number for number, line in enumerate(broken)
            if line.forced and not one_column and pieces[line.end - 1].char == "\v")

        before = pp.get("spacing.before", 0) or 0
        after = pp.get("spacing.after", 0) or 0
        if autospaced(document, pp, "before") and not autospace_kept(
                numbered=numbered[block], neighbour_numbered=numbered.get(block - 1) if block else None):
            before = 0
        if autospaced(document, pp, "after") and not autospace_kept(
                numbered=numbered[block], neighbour_numbered=numbered.get(block + 1, False)):
            after = 0
        if (document.compatibility_mode or 0) < 15 and before:
            # Below mode 15, every line of a joined paragraph that begins a joined
            # paragraph's text is drawn lower by the space before, but what follows moves
            # down by one fewer of them (R.4: ``make_revision_probe.py``, chains of breaks).
            starts = len(join_starts(item, pieces, broken))
            after += max(0, starts - 1) * before
        if item.note == "separator" and (document.compatibility_mode or 0) >= 15:
            # The endnotes' separator keeps no space before in mode 15; below it, it
            # does (``make_endnote_probe.py``, ``separators``: 120 twips, 25 px).
            before = 0
        position = body_index[id(item)]
        neighbours = [resolved_body[k] if 0 <= k < len(resolved_body) else None for k in (position - 1, position + 1)]
        top_border, border_height = paragraph_borders_px(pp, *neighbours)
        widow = pp.get("widowControl")
        notes: dict[int, list] = {}
        note_sources: list = []
        references = [(offset + position, int(kind.split(":")[1]))
                      for offset, run in zip(_run_offsets(item), item.runs)
                      for position, kind in run.breaks if kind.startswith("footnoteReference:")]
        if references:
            line_of = _line_of_sources(pieces, broken)
            for source, note_id in references:
                laid = note(note_id, section_number)
                if laid is None:
                    reason = "footnote not measurable"
                    break
                line = line_of.get(source, line_of.get(source - 1, len(broken) - 1))
                notes.setdefault(line, []).append(laid)
                note_sources.append((source, laid))
            if reason is not None:
                out.append(Obstacle(block, reason, section_number, item))
                continue
            if section_number not in out.separators:
                from .notes import footnote_separator_kind, separator as note_separator

                out.separators[section_number] = tuple(
                    sum((line.height for line in (_note_lines(
                        document, note_separator(document, footnote_separator_kind(document, kind), "footnote"),
                        section, advances, metrics, heights_cache) or [])), Fraction(0))
                    for kind in ("separator", "continuationSeparator"))
        out.append(Para(
            block=block,
            section=section_number,
            heights=heights,
            page_breaks=page_breaks,
            column_breaks=column_breaks,
            before=before,
            after=after,
            contextual=bool(pp.get("contextualSpacing")),
            style=item.properties.style_id if item.properties else None,
            border=border_height,
            top_border=top_border,
            keep_next=bool(pp.get("keepNext")),
            keep_lines=bool(pp.get("keepLines")),
            page_break_before=bool(pp.get("pageBreakBefore")),
            widow_control=None if widow is None else bool(widow),
            section_start=section_start,
            undrawn=frozenset(number for number, chars in enumerate(shares)
                              if number < len(shares) - 1 and all(char.isspace() for _, char in chars)),
            notes=notes,
            anchors=anchors_of(item, _line_of_sources(pieces, broken), len(broken)),
            section_mark=section_mark and len(heights) == 1,
            continuation=continuation(section_number) if item.note == "endnote" else None,
            drop_cap=drop,
            auto_hyphen_ends=_auto_hyphen_ends(pieces, broken),
        ))
        unequal = len({width for _, width, _ in column_boxes(section)}) > 1
        word_moves = out[-1].auto_hyphen_ends and hyphen_at_page_end(document) == "word"
        if out.wraps or unequal or word_moves:
            from .linebreak import geometry as line_geometry

            out[-1].source = ParaSource(document, item, pp, pieces, line_geometry(document, item, section, pp),
                                        advances, metrics, heights_cache, one_column, note_sources,
                                        section=section if unequal else None)
            out[-1].lines = list(broken)
            out[-1].plain_from = 0
    return out


def continuation_line(document: Document, section, advances, metrics, cache: dict | None = None):
    """The continuation separator's line (:func:`docx2svg.notes.separator`): what a page
    the endnotes go on to holds above them (``make_endnote_probe.py``, ``overflow``: its
    line, then the notes stacked under it as under any paragraph).  ``None`` where it
    cannot be measured."""
    from .notes import separator

    paragraph = next((block for block in separator(document, "continuationSeparator")
                      if isinstance(block, Paragraph)), None)
    if paragraph is None:
        return None
    pieces, broken = line_model.broken(document, paragraph, section, advances, metrics)
    if pieces is None:
        return None
    shares = line_model.shares_of(document, paragraph, pieces, broken)
    return line_model.line_height(document, paragraph, resolve_paragraph(document, paragraph), shares[0], metrics,
                                  first_line=True, cache=cache)


def anchors_of(paragraph: Paragraph, line_of: dict | None = None, lines: int = 1) -> tuple:
    """Every floating drawing of ``paragraph`` that text does not wrap beside, where it is
    in the text: ``(index into the paragraph's text, run index, index in the run's
    anchors, the anchor, its line)``.  Such a drawing takes no room in its line (its mark
    is no piece of the line breaker's): one text does not wrap around is laid out with
    the lines, by :mod:`layout`; one text wraps top and bottom of also keeps lines off
    its band (:class:`Band`)."""
    out = []
    line_of = line_of or {}
    for run_index, (start, run) in enumerate(zip(_run_offsets(paragraph), paragraph.runs)):
        for position, kind in run.breaks:
            if kind.startswith("anchor:"):
                k = int(kind.split(":")[1])
                if k < len(run.anchors) and (not run.anchors[k].moves_text
                                             or run.anchors[k].wrap in ("wrapTopAndBottom",) + SIDE_WRAPS):
                    source = start + position
                    line = line_of.get(source, line_of.get(source - 1, lines - 1))
                    out.append((source, run_index, k, run.anchors[k], line))
    return tuple(out)


def multiple_extra(height: line_model.LineHeight) -> Fraction:
    """The part of an ``auto`` multiple above one that is below the line's text."""
    if height.rule == "auto" and height.line > 240:
        return height.pitch - height.text_above - height.text_below
    return Fraction(0)


def border_ends_every_page(compatibility_mode: int | None) -> bool:
    """Whether a paragraph's bottom border must fit below *any* of its lines that ends a
    page, not only below its last: below compatibility mode 15, and when none is stated.

    Measured by ``make_page_fit_probe.py`` (``border-first``: the first line of a
    two-line bordered paragraph, widow control off, swept by single units): with no
    ``settings.xml`` and in modes 12 and 14 it stays on the page only when the border
    fits under it too; in mode 15 when its own text does.  A last line needs its border
    in every mode.
    """
    return compatibility_mode is None or compatibility_mode < 15


def fits(y: Fraction, height: line_model.LineHeight, bottom: Fraction, *, border: Fraction = Fraction(0),
         after: Fraction = Fraction(0), rule: str = FIT) -> bool:
    """Whether a line whose pitch starts at ``y`` stays on a page whose text ends at
    ``bottom`` (inclusive: ending exactly on the margin fits).

    ``border`` and ``after`` are the paragraph's bottom border and space after, passed
    for its last line (the border, below mode 15, for every line:
    :func:`border_ends_every_page`).  What must fit is :data:`FIT`'s: the line's pitch without an
    ``auto`` multiple's extra (which sits below the text), and the bottom border; not the
    space after.
    """
    if rule == "text-border":
        end = y + height.pitch - multiple_extra(height) + border
    elif rule == "pitch":
        end = y + height.pitch
    elif rule == "pitch-border":
        end = y + height.pitch + border
    elif rule == "box":
        end = y + height.pitch + border + after
    elif rule == "text":
        end = y + height.pitch - multiple_extra(height)
    else:
        raise ValueError(rule)
    return end <= bottom


@dataclass
class Rules:
    """The measured behaviour, one field per rule, so a refuted alternative can be scored."""

    fit: str = FIT
    #: Widow control where nothing states it: on (measured, with and without a
    #: ``settings.xml`` or a styles part).
    widow_default: bool = True
    #: Widow and orphan control: the lines it keeps together at each end of a paragraph.
    widow_lines: int = 2
    #: What ``keepNext`` moves when the next paragraph cannot start on the page: ``mode``
    #: (measured: :func:`keep_next_moves`), or always ``line`` or ``paragraph`` (refuted).
    keep_next_moves: str = "mode"
    #: Whether widow control ends a paragraph at a manual page break: ``mode``
    #: (measured: :func:`page_break_ends_paragraph`), ``always`` or ``never`` (refuted).
    segment: str = "mode"
    #: When the keeps would leave the page empty: ``ignore`` them and break where the
    #: page is full.
    oversize: str = "ignore"
    #: Whether a header or footer that reaches past its margin moves it (measured: it
    #: does; ``False`` is the refuted alternative).
    stories: bool = True


def _geometry(document: Document, section: int, items: list | None = None,
              first_page: bool = False, story: str | None = None) -> tuple[Fraction, Fraction]:
    """Where the page's text starts and ends: the margins, pushed in by the page's header
    or footer where it reaches past them (:func:`story_height`).  ``story`` is the kind
    the page shows (:func:`story_kind`); without it, ``first`` on a section's first page
    under ``w:titlePg`` and ``default`` otherwise.  A negative top or bottom margin is
    kept whatever the header or footer (its magnitude is the margin)."""
    s = document.sections[section]
    stories = getattr(items, "stories", {}).get(section) if items is not None else None
    header = footer = Fraction(0)
    if story is None:
        story = "first" if first_page and s.title_page else "default"
    if stories is not None:
        header, footer = stories[story]
    top = twips_to_px(abs(s.margins.top))
    if s.margins.top >= 0 and header:
        top = max(top, twips_to_px(s.margins.header) + header)
    bottom = twips_to_px(s.page_size.height_twips - abs(s.margins.bottom))
    if s.margins.bottom >= 0 and footer:
        bottom = min(bottom, twips_to_px(s.page_size.height_twips - s.margins.footer) - footer)
    return top, bottom


def _new_page_section(document: Document, section: int) -> bool:
    """Whether a section starts a new page: every kind but ``continuous`` (and
    ``nextColumn``), and a continuous one whose page size differs from the section
    before's (``make_section_probe.py``: Word starts it on a new page; a continuous
    section that changes only the margins stays on the page)."""
    here = document.sections[section]
    kind = here.start_type or "nextPage"
    if kind in ("nextPage", "oddPage", "evenPage"):
        return True
    return section > 0 and here.page_size != document.sections[section - 1].page_size


def blank_page_before(document: Document, section: int, page: int) -> bool:
    """Whether an ``oddPage`` or ``evenPage`` section starting on page index ``page``
    (numbered from 1 as ``page + 1``) needs a blank page first: when that number has the
    wrong parity.  Measured by ``make_section_probe.py`` (each kind after each parity);
    ``w:pgNumType/@w:start`` (a restarted count) is not read."""
    kind = document.sections[section].start_type
    if kind == "oddPage":
        return (page + 1) % 2 == 0
    if kind == "evenPage":
        return (page + 1) % 2 == 1
    return False


def section_start_number(document: Document, section: int, previous: int, physical: int) -> tuple[int, int]:
    """A page-starting section's first page number, and how many blank pages Word puts
    before it, given the number of the page before (``previous``) and the 0-based index
    the section's page would have (``physical``).  Measured by
    ``make_story_select_probe.py`` (``even``, ``blank``, ``restart even``, ``restart odd``;
    every page's ``PAGE`` and every blank page) and ``make_section_probe.py``:

    * the number is ``w:pgNumType/@w:start``, or the page before's plus one;
    * an ``oddPage`` or ``evenPage`` section wants that **number**'s parity: a number
      that ``@w:start`` restarted is raised by one (no blank page: ``start 2`` on an
      ``oddPage`` section is drawn 3), any other gets a blank page first, which takes
      the number;
    * under ``w:evenAndOddHeaders`` a page's number and its place keep one parity: where
      they would differ (a restart), a blank page comes first -- ``start 1`` on the
      eighth page puts page 1 on the ninth.
    """
    here = document.sections[section]
    restarted = here.page_number_start is not None
    number = here.page_number_start if restarted else previous + 1
    blanks = 0
    kind = here.start_type
    if kind in ("oddPage", "evenPage") and number % 2 != (1 if kind == "oddPage" else 0):
        if restarted:
            number += 1
        else:
            blanks, number = 1, number + 1
    if document.even_and_odd_headers and (physical + blanks + 1) % 2 != number % 2:
        blanks += 1
    return number, blanks


def page_end(document: Document, items: list, start: Position, rules: Rules | None = None,
             carry: tuple = ()) -> Position:
    """The position the page starting at ``start`` ends before (``Position(len(items))``
    at the end of the document).  Raises :class:`Unplaceable` at an :class:`Obstacle`.
    ``carry`` is the rest of a footnote the previous page could not hold."""
    return lay_page(document, items, start, rules, carry)[0]


def lay_page(document: Document, items: list, start: Position, rules: Rules | None = None,
             carry: tuple = (), story: str | None = None) -> tuple[Position, tuple]:
    """:func:`page_end`, and the lines of a footnote that go on to the next page."""
    end, carry, _, _ = _lay_page(document, items, start, rules, carry, story)
    return end, carry


def _lay_page(document: Document, items: list, start: Position, rules: Rules | None, carry: tuple,
              story: str | None = None, page: int = 0, bands_out: list | None = None,
              wraps_out: list | None = None, frames_out: list | None = None):
    """A page, laid out once without the bands of its ``wrapTopAndBottom`` drawings and
    the drawings text wraps beside, which are positioned from that layout, and again with
    them (ROADMAP.md, "Floating drawings -- measured", F.7, F.8)."""
    rules = rules or Rules()
    if getattr(items, "columns", False):
        return _lay_columns_page(document, items, start, rules, carry, story, page, bands_out, wraps_out, frames_out)
    tops: dict = {}
    stop = None
    try:
        end, placed, forced = _scan(document, items, start, rules, carry, story, tops=tops)
    except Unplaceable as unplaceable:
        # The page the layout stops on is laid out with its drawings too, up to the stop.
        stop = unplaceable
    bands = page_bands(document, items, tops, page)
    if len(bands) > 1 and stop is None:
        bands = _stacked_bands(document, items, start, rules, carry, story, page, bands)
    wraps = page_wraps(document, items, tops, page) if getattr(items, "wraps", False) else []
    if bands or wraps:
        if bands_out is not None:
            bands_out.extend(bands)
        if wraps_out is not None:
            wraps_out.extend(wraps)
        end, placed, forced = _scan(document, items, start, rules, carry, story, bands=bands, wraps=wraps)
    if stop is not None:
        raise stop
    return end, _carried(placed, end), forced, _page_notes(placed, end)


def _stacked_bands(document: Document, items: list, start: Position, rules: Rules, carry: tuple,
                   story: str | None, page: int, bands: list) -> list:
    """The page's ``wrapTopAndBottom`` bands, each drawing positioned from the page as laid
    out with the bands of the drawings **before** it -- not its own, nor those after it
    (ROADMAP.md, F.24): a drawing anchored in the paragraph after another's stands below
    the text the first pushed down.  ``bands`` are the drawings positioned from the page
    with none, in document order; a drawing whose anchor the bands before it push off the
    page goes with its paragraph, to be positioned on the next page, and so do those after
    it."""
    fixed = bands[:1]
    for band in bands[1:]:
        found: dict = {}
        try:
            _scan(document, items, start, rules, carry, story, bands=fixed, tops=found)
        except Unplaceable:
            break
        again = {other.key: other for other in page_bands(document, items, found, page)}
        if band.key not in again:
            break
        fixed.append(again[band.key])
    return fixed


def _lay_columns_page(document: Document, items: list, start: Position, rules: Rules, carry: tuple,
                      story: str | None, page: int, bands_out: list | None, wraps_out: list | None,
                      frames_out: list | None):
    """:func:`_lay_page` for a document with a section of several columns: the page as its
    regions (:func:`_lay_columns`), laid out once without its drawings' bands and wraps and
    again with them, as :func:`_lay_page` lays a page out.

    **The page's footnotes, in mode 15**, run on through its columns at its foot under one
    separator (:func:`note_area`), and **every column ends above them**: the page is laid
    out again with the room the notes its columns reference take, until that room holds
    (``make_columns_probe.py``, ``notesright``: a first column with no reference ends six
    lines short, over the second's notes; ``footnote``: a column ends 46 lines down where
    its own notes alone, balanced, would leave it 47).  Should the room not settle, the
    larger is kept.  Below mode 15 each column's notes stand at its own foot
    (:func:`_scan`)."""
    tops: dict = Tops()
    stop = None
    frames: list = []
    room = Fraction(0)
    try:
        end, placed, forced, room = _lay_noted(document, items, start, rules, carry, story, tops, frames)
    except Unplaceable as unplaceable:
        stop = unplaceable
    bands = page_bands(document, items, tops, page)
    wraps = page_wraps(document, items, tops, page) if getattr(items, "wraps", False) else []
    if bands or wraps:
        if bands_out is not None:
            bands_out.extend(bands)
        for _attempt in range(4):
            frames = []
            found = Tops()
            try:
                end, placed, forced = _lay_columns(document, items, start, rules, carry, story, bands=bands,
                                                   wraps=wraps, frames=frames, note_room=room, tops=found)
            except Unplaceable as unplaceable:
                stop = unplaceable
                break
            # A drop cap is positioned where the text above it ends with the drop caps
            # above it in place (``make_columns_probe.py``, ``dropcap``: one under another
            # in the second column, a line lower than without the first): again until
            # they hold.
            again = page_wraps(document, items, found, page) if getattr(items, "wraps", False) else []
            in_columns = any(w.key[0] == "drop cap" and any(items[index].block == w.key[1] for index, _ in found.columns)
                             for w in again) if found.columns else False
            if not in_columns or [(w.top, w.bottom, w.left, w.right) for w in again] == [
                    (w.top, w.bottom, w.left, w.right) for w in wraps] or _attempt == 3:
                break
            wraps = again
        if wraps_out is not None:
            wraps_out.extend(wraps)
    if frames_out is not None:
        frames_out.extend(frames)
    if stop is not None:
        raise stop
    if any(frame.multi for frame in frames):
        # The notes are the columns' own (:attr:`Frame.notes`); below mode 15 what the
        # last column could not hold goes on to the next page.
        last = frames[-1]
        rest = note_rest(last.notes) if last.multi and not last.pooled else _carried(placed, end)
        return end, rest, forced, []
    notes = [note for frame in frames for note in frame.notes]
    return end, _carried(placed, end), forced, notes


def _lay_noted(document: Document, items: list, start: Position, rules: Rules, carry: tuple, story: str | None,
               tops: dict, frames: list):
    """:func:`_lay_columns`, with the room of the page's note area in mode 15 found
    (:func:`_lay_columns_page`): ``(end, placed, forced, room)``."""
    room = Fraction(0)
    seen: list = []
    while True:
        laid: list = []
        found: dict = Tops()
        try:
            end, placed, forced = _lay_columns(document, items, start, rules, carry, story, tops=found, frames=laid,
                                               note_room=room)
        except Unplaceable:
            frames.extend(laid)
            tops.update(found)
            if hasattr(tops, "columns"):
                tops.columns.update(found.columns)
            raise
        pooled = [frame for frame in laid if frame.pooled]
        if not pooled:
            break
        ids = tuple(lines[0].note for frame in pooled for lines, _ in frame.notes)
        if ids and carry:
            # A note going on from the page before onto a page whose notes run through
            # its columns: not measured.
            _notes_stop(items, start, "a footnote going on from the page before")
        need = Fraction(0)
        if ids:
            area = note_area(document, items, pooled[0].section, ids)
            if area is None:
                _notes_stop(items, start, "a footnote that cannot be laid out")
            need = area.room
            if need >= pooled[0].bottom + room - pooled[0].top:
                _notes_stop(items, start, "footnotes taller than the page")
        if need == room:
            break
        seen.append(room)
        if need in seen or len(seen) > 8:
            room = max(seen + [need])
            laid, found = [], Tops()
            end, placed, forced = _lay_columns(document, items, start, rules, carry, story, tops=found, frames=laid,
                                               note_room=room)
            break
        room = need
    frames.extend(laid)
    tops.update(found)
    if hasattr(tops, "columns"):
        tops.columns.update(found.columns)
    return end, placed, forced, room


def _notes_stop(items: list, start: Position, detail: str) -> None:
    """Stop the layout at the page's first paragraph that starts on it, as at an
    :class:`Obstacle` of the columns."""
    index = start.item + (1 if start.line or start.part else 0)
    index = min(index, len(items) - 1)
    item = items[index]
    items[index] = Obstacle(item.block, "columns", item.section, None, detail=f"{detail} (its notes in columns)")
    raise Unplaceable(index)


@dataclass
class NoteArea:
    """A page's footnotes in mode 15, as Word sets them under a section of several
    columns: their paragraphs a flow of their own (``story``, ``items``) in the section's
    columns, **balanced** -- the least height at which the columns, filled in turn under
    the keeps (widow control splitting a note), hold them all, a note going on into a
    column of another width broken again in it -- under **one separator** above the first
    column (``make_columns_probe.py``, ``footnote``: seven lines go four and three;
    ``notes3``: eleven lines four, four and three; ``notesunequal``: a note going on from
    a column of 2,200 twips into one of 4,000 broken again there).  ``frames`` are its
    columns from 0 down; ``height`` the tallest's foot; ``room`` that and the separator's
    line: what the page keeps free above its foot."""

    story: Document
    items: list
    frames: list
    height: Fraction
    room: Fraction


def note_area(document: Document, items: list, section_number: int, ids: tuple) -> NoteArea | None:
    """The page's note area for the notes ``ids`` in section ``section_number``'s columns
    (:class:`NoteArea`); ``None`` where a note cannot be laid out.  Cached on the flow."""
    cache = getattr(items, "areas", None)
    key = (section_number, ids)
    if cache is not None and key in cache:
        return cache[key]
    advances, metrics = items.measure
    blocks = [block for note_id in ids for block in document.footnotes.get(note_id, [])
              if isinstance(block, Paragraph)]
    section = document.sections[section_number]
    one = dataclasses.replace(section, first_paragraph=0, last_paragraph=len(blocks), start_type=None, headers=(),
                              footers=(), title_page=False)
    story = dataclasses.replace(document, body=list(blocks), paragraphs=list(blocks), sections=[one], footnotes={},
                                footnote_separators={}, stories={}, story_parts={})
    flowed = flow(story, advances, metrics)
    area = None
    if flowed and not any(isinstance(item, Obstacle) for item in flowed):
        def fill(height):
            try:
                laid, _ = _fill(story, flowed, Position(0), Rules(), (), None, Fraction(0), height, None, True, 0, [])
            except Unplaceable:
                return None
            return laid if laid[-1].end.item >= len(flowed) else None

        lo, hi = Fraction(0), Fraction(10 ** 5)
        if fill(hi) is not None:
            while hi - lo > Fraction(1, 65536):
                mid = (lo + hi) / 2
                if fill(mid) is None:
                    lo = mid
                else:
                    hi = mid
            laid = fill(hi)
            height = max(frame.y for frame in laid)
            separator = items.separators.get(section_number, (Fraction(0), Fraction(0)))[0]
            area = NoteArea(story, flowed, laid, height, height + separator)
    if cache is not None:
        cache[key] = area
    return area


def _lay_columns(document: Document, items: list, start: Position, rules: Rules, carry: tuple,
                 story: str | None, *, bands=(), tops: dict | None = None, wraps=(), frames: list,
                 note_room: Fraction = Fraction(0)):
    """One page from ``start`` as Word lays a page of text columns out
    (``make_columns_probe.py``; ROADMAP.md, "Phase 4 -- measured", 4.14): ``(end, placed,
    forced)`` as :func:`_scan`'s, every column laid out appended to ``frames``.

    * **Regions.**  The page's sections, one under the other where continuous ones share
      it: a one-column section's text goes on as one stack, a section of several columns
      fills its first column from the region's top, then the next from that top, and so
      on; the page ends when its last column is full, at a page break, or at a section
      that starts a page.  A column break ends its column (in the last, the page).
    * **The keeps** -- widow and orphan control, ``keepNext``, ``keepLines`` -- hold at a
      column's foot as at a page's, and a column's first line is placed whatever its
      height, as a page's is.
    * **Balance.**  A section of several columns that ends on the page, a continuous
      section after it, is balanced (:func:`_balance`); the next region starts under its
      longest column.
    """
    first = items[start.item]
    section_first = start.line == 0 and (start.item == 0 or items[start.item - 1].section != first.section)
    top, bottom = _geometry(document, first.section, items if rules.stories else None, section_first, story)
    # In mode 15, the room the page's note area keeps (:func:`_lay_columns_page`).
    bottom -= note_room
    y = top
    if isinstance(first, Para) and first.continuation is not None:
        y += first.continuation.pitch
    position = start
    previous = None
    first_on_page = True
    placed_all: list = []
    kw = dict(bands=bands, tops=tops, wraps=wraps)
    while True:
        section_number = items[position.item].section
        laid, placed = _fill(document, items, position, rules, carry if not frames else (), story, y, bottom,
                             previous, first_on_page, section_number, frames, **kw)
        last = laid[-1]
        if last.multi and last.why == "region":
            laid, placed = _balance(document, items, position, rules, story, y, bottom, previous, first_on_page,
                                    section_number, frames, laid, placed, kw)
            last = laid[-1]
        frames.extend(laid)
        placed_all.extend(placed)
        if last.why != "region":
            return last.end, placed_all, last.why in ("page", "column")
        position = last.end
        first_on_page = first_on_page and all(frame.end == frame.start for frame in laid)
        if last.multi:
            y, previous = region_foot(document, items, laid)
        else:
            y, previous = last.y, last.last


def _ended_after(items: list, frame: Frame) -> tuple[Fraction, object]:
    """The space after (px) of the last paragraph that ends in ``frame`` -- a paragraph
    holding a section break and nothing else left out -- and that paragraph; ``0`` where
    the column's last paragraph goes on in the next, or it ends in a table."""
    end = frame.end
    if end.line or end.part:
        return Fraction(0), None
    index = end.item - 1
    while index >= frame.start.item:
        item = items[index]
        if isinstance(item, Para) and item.section_mark and index > frame.start.item:
            index -= 1
            continue
        if isinstance(item, Para):
            return twips_to_px(item.after), item
        return Fraction(0), item
    return Fraction(0), None


def region_foot(document: Document, items: list, frames: list) -> tuple[Fraction, object]:
    """Where the stack goes on under a region of several columns, and the item the next
    paragraph's space before collapses with (``make_columns_probe.py``, ``regions``):

    * **in mode 15, under its tallest column with that column's last space after** -- the
      stack of each column and the space after of the paragraph that ends it;
    * **below mode 15, the taller of its tallest column (no space after) and the mean of
      the columns' stacks each with its last space after** -- four lines and 480 twips
      after over two columns end two lines and 240 twips down; three lines and 360 after
      beside two and 120 after, 190 px down where the columns reach 168 and 112;
    * the next paragraph's space before collapses with the space after of the region's
      last paragraph, which the foot already holds.

    A paragraph holding only the section break takes no part (its space after neither).
    """
    top = frames[0].region_top
    stacks = []
    for frame in frames:
        after, _ = _ended_after(items, frame)
        stacks.append((frame.y - top if frame.end > frame.start else Fraction(0), after))
    if (document.compatibility_mode or 0) >= 15:
        height = max(stack + after for stack, after in stacks)
    else:
        height = max(max(stack for stack, _ in stacks), sum(stack + after for stack, after in stacks) / len(frames))
    last = next((frame for frame in reversed(frames) if frame.end > frame.start), frames[0])
    after, previous = _ended_after(items, last)
    if previous is None:
        previous = items[frames[-1].end.item - 1]
        after = Fraction(0)
    return top + height - after, previous


def _fill(document: Document, items: list, start: Position, rules: Rules, carry: tuple, story: str | None,
          top: Fraction, bottom: Fraction, previous, first_on_page: bool, section_number: int, done: list,
          **kw) -> tuple[list, list]:
    """One region from ``start``: its section's columns filled in turn, each from ``top``
    down to ``bottom``, until the region, the page or the flow ends; ``(frames, placed)``.
    ``done`` holds the page's frames so far (a page's top is where none is)."""
    section = document.sections[section_number]
    multi = is_multi_column(section)
    boxes = column_boxes(section)
    frames: list = []
    placed_all: list = []
    position = start
    region_top = top
    pooled = multi and (document.compatibility_mode or 0) >= 15
    for column, (offset, width, _space) in enumerate(boxes):
        frame = Frame(section_number, column, region_top if column else top, bottom,
                      first_on_page if column == 0 else True, previous if column == 0 else None, multi, offset, width,
                      page_top=not done and column == 0 and first_on_page, start=position, region_top=region_top,
                      pooled=pooled)
        frames.append(frame)
        if column and multi and not pooled:
            # Below mode 15 the notes the column before could not hold go on at this
            # one's foot (``make_columns_probe.py``, ``notesover``, ``notes3``).
            carry = note_rest(frames[-2].notes)
        elif column:
            carry = ()
        try:
            end, placed, _forced = _scan(document, items, position, rules, carry, story, frame=frame, **kw)
        except Unplaceable as stop:
            frame.end = Position(stop.args[0], stop.args[1] if len(stop.args) > 1 else 0)
            done.extend(frames)
            raise
        placed_all.extend(placed)
        position = end
        if column == 0 and not first_on_page and frame.first_top is not None:
            # A region under another starts every column where its first paragraph
            # starts, past the space between the two (``make_columns_probe.py``,
            # ``regions``: the second column's first line level with the first's).
            region_top = frame.first_top
            frame.region_top = region_top
        if frame.why not in ("full", "column") or position.item >= len(items):
            break
    return frames, placed_all


def _balance(document: Document, items: list, start: Position, rules: Rules, story: str | None, top: Fraction,
             bottom: Fraction, previous, first_on_page: bool, section_number: int, done: list, laid: list,
             placed: list, kw: dict) -> tuple[list, list]:
    """A region whose section ends on the page, a continuous section after it, laid out
    again with its columns balanced: **the least height at which its columns, each filled
    in turn under the keeps as a page's column is, hold it all** (``make_columns_probe.py``,
    ``balance``: 7 lines go 4 and 3, 10 one-line paragraphs in three columns 4, 4 and 2,
    a two-line paragraph stays whole in the first column, 9 lines and 7 go 9 and 7 where
    8 would leave a widow, a ``keepLines`` paragraph moves whole).  The height is found
    by bisection to a 1/65536 px: every placement it decides is decided by a line's foot,
    which is never that close to another."""
    region_top = laid[0].region_top
    lo, hi = Fraction(0), bottom - region_top
    scratch_kw = dict(kw, tops=None)
    while hi - lo > Fraction(1, 65536):
        mid = (lo + hi) / 2
        try:
            frames, found = _fill(document, items, start, rules, (), story, top, region_top + mid, previous,
                                  first_on_page, section_number, done, **scratch_kw)
        except Unplaceable:
            lo = mid
            continue
        # Every column filled in turn holds a line: none is passed over empty.
        if frames[-1].why == "region" and all(frame.end > frame.start for frame in frames if frame.why == "full"):
            hi = mid
        else:
            lo = mid
    # The last layout is the one kept (a paragraph broken again for a column keeps the
    # lines its last placement gave it).
    frames, found = _fill(document, items, start, rules, (), story, top, region_top + hi, previous, first_on_page,
                          section_number, done, **kw)
    if frames[-1].why != "region":
        frames, found = _fill(document, items, start, rules, (), story, top, bottom, previous, first_on_page,
                              section_number, done, **kw)
        return frames, found
    for frame in frames:
        frame.balanced = True
    return frames, found


def _frames(document: Document, item, number: int, top: Fraction, pitch: Fraction, paragraph_top: Fraction,
            page: int, character: Fraction = Fraction(0), column: tuple | None = None):
    """What a drawing anchored on line ``number`` of ``item``, at ``top`` with ``pitch``
    (device px), is positioned against, as a drawing that moves text is: below mode 15 the
    paragraph's and the line's tops taken in whole twips, the nearest
    (``make_wrap_anchor_probe.py``, no settings: every picture); ``character`` in px."""
    from . import floating

    section = document.sections[item.section]
    m = section.margins
    px = floating.TWIP_PX
    mode15 = (document.compatibility_mode or 0) >= 15
    whole = (lambda v: Fraction(math.floor(v + Fraction(1, 2)))) if not mode15 else (lambda v: v)
    return floating.Frames(section.page_size.width_twips, section.page_size.height_twips, m.left + m.gutter,
                           m.right, m.top, m.bottom, m.header, m.footer, page + 1, mode15,
                           whole(paragraph_top / px), whole((paragraph_top if number == 0 else top) / px),
                           whole((top + pitch) / px), character / px,
                           column=None if column is None else (m.left + m.gutter + column[0], column[1]))


def _character_px(document: Document, item, number: int, source: int, run: int) -> Fraction:
    """The pen x (device px from the page's left edge, from the margin's exact edge) of the
    character a drawing anchored at text index ``source`` of run ``run`` on line ``number``
    of ``item`` is positioned against (:func:`docx2svg.layout.anchor_character_units`)."""
    from .layout import anchor_character_units, line_positions, units_px

    src = item.source
    if src is None or number >= len(item.lines):
        return Fraction(0)
    line = item.lines[number]
    positions = line_positions(src.pieces, line, src.geometry, src.advances, first_line=number == 0,
                               alignment=src.pp.get("jc"), last_line=number == len(item.lines) - 1)
    section = document.sections[item.section]
    exact = twips_to_px(section.margins.left + section.margins.gutter)
    return exact + units_px(anchor_character_units(src.pieces, line, positions, source, src.geometry,
                                                   first_line=number == 0, run=run))


def page_wraps(document: Document, items: list, tops: dict, page: int) -> list:
    """The drawings text wraps beside (``wrapSquare``, ``wrapTight``, ``wrapThrough``)
    anchored on lines at ``tops``, each positioned as :mod:`docx2svg.floating` positions a
    drawing, from the page laid out without them (:class:`docx2svg.wrap.Wrap`)."""
    from . import floating
    from .wrap import Wrap

    out = []
    px = floating.TWIP_PX
    for (index, number), (top, pitch, paragraph_top) in tops.items():
        item = items[index]
        drop = getattr(item, "drop_cap", None)
        if drop is not None and number == 0 and drop.beside:
            # A drop cap's frame, from the paragraph's top down its height, from the
            # column's exact edge to where the text beside it starts.
            column = wrap_column(document, item.section, *getattr(tops, "columns", {}).get((index, number), (0, None)))
            start, shift = drop.text_start(column.start, column.exact, column.mode15)
            out.append(Wrap(paragraph_top, paragraph_top + drop.height, column.exact,
                            column.exact + start - column.start, Fraction(0), Fraction(0), "bothSides",
                            paragraph_top, ("drop cap", item.block), text_shift=shift))
        if _floats(item):
            wrap = table_wrap(document, item, paragraph_top, page)
            section = document.sections[item.section]
            foot = twips_to_px(section.page_size.height_twips - abs(section.margins.bottom))
            if (document.compatibility_mode or 0) >= 15 and wrap.bottom - twips_to_px(item.floating.bottom) > foot:
                # Mode 15 splits a floating table that reaches below the bottom margin
                # across two pages (``make_float_table_probe.py``, the last case): not
                # modelled.  Below it, it stays whole where it is positioned.
                raise Unplaceable(index)
            out.append(wrap)
            continue
        for source, run, k, anchor, line in getattr(item, "anchors", ()):
            if line != number or anchor.wrap not in SIDE_WRAPS:
                continue
            character = _character_px(document, item, number, source, run) if anchor.h.relative == "character" \
                else Fraction(0)
            frames = _frames(document, item, number, top, pitch, paragraph_top, page, character,
                             getattr(tops, "columns", {}).get((index, number)))
            if not frames.mode15 and anchor.simple is None and (anchor.h.relative or "column") == "column" \
                    and anchor.h.offset is None and item.source is not None:
                # Below mode 15 a drawing aligned in the column is aligned between the
                # paragraph's indents (``make_wrap_side_probe.py``, no settings).
                pp = item.source.pp
                frames = dataclasses.replace(frames, left=frames.left + (pp.get("ind.left", 0) or 0),
                                             right=frames.right + (pp.get("ind.right", 0) or 0))
            x, y = floating.horizontal(anchor, frames), floating.vertical(anchor, frames)
            if x is None or y is None:
                continue
            if not frames.mode15:
                # ... and on a whole twip, the nearest (a centred drawing's half twip up).
                x = Fraction(math.floor(x + Fraction(1, 2)))
            cx, cy = (floating.emu_twips(v) for v in anchor.extent)
            el, et, er, eb = (floating.emu_twips(v) for v in anchor.effect)
            dt, db, dl, dr = (floating.emu_twips(v) * px for v in anchor.distance)
            polygon = None
            if anchor.wrap in ("wrapTight", "wrapThrough") and anchor.polygon:
                # Each point in whole twips from the drawing's corner, truncated.
                polygon = tuple(((x + int(Fraction(cx * a, 21600))) * px, (y + int(Fraction(cy * b, 21600))) * px)
                                for a, b in anchor.polygon)
            if polygon is not None and not frames.mode15:
                # Below mode 15 a wrap polygon keeps text off its own height only: no
                # distT or distB (``make_wrap_side_probe.py``, ``wrapTight, distT and distB``).
                dt = db = Fraction(0)
            if polygon is not None:
                top_px = min(b for _a, b in polygon) - dt
                bottom_px = max(b for _a, b in polygon) + db
            else:
                top_px, bottom_px = (y - et) * px - dt, (y + cy + eb) * px + db
            out.append(Wrap(top_px, bottom_px, (x - el) * px, (x + cx + er) * px, dl, dr, anchor.wrap_text or "bothSides",
                            y * px, (item.block, source, run, k), polygon, dt, db, x * px))
    return out


def page_bands(document: Document, items: list, tops: dict, page: int) -> list[Band]:
    """The bands of the ``wrapTopAndBottom`` drawings anchored on lines at ``tops``
    (``(item, line) -> (top, pitch, paragraph top)``, device px), each positioned as
    :mod:`docx2svg.floating` positions a drawing, as tall as its extent in whole twips."""
    from . import floating

    out = []
    for (index, number), (top, pitch, paragraph_top) in tops.items():
        item = items[index]
        for source, run, k, anchor, line in getattr(item, "anchors", ()):
            if line != number or anchor.wrap != "wrapTopAndBottom":
                continue
            px = floating.TWIP_PX
            column = getattr(tops, "columns", {}).get((index, number))
            frames = _frames(document, item, number, top, pitch, paragraph_top, page, column=column)
            y = floating.vertical(anchor, frames)
            if y is None:
                continue
            y = y * px
            span = None
            if column is not None:
                # In a column of several: the band keeps lines off the columns the drawing
                # is across, not the page's (``make_columns_probe.py``, ``floating``).
                x = floating.horizontal(anchor, frames)
                if x is not None:
                    span = (x * px, (x + floating.emu_twips(anchor.extent[0])) * px)
            # The extent in whole twips, as an alignment takes it -- not a picture's drawn
            # height, which is scaled to fit it (up to 0.06 px less).
            height = floating.emu_twips(anchor.extent[1]) * px
            above, below = (floating.emu_twips(v) * px for v in anchor.distance[:2])
            out.append(Band(y - above, y + height + below, y, (item.block, source, run, k), span))
    return out


def continuation_page(compatibility_mode: int | None) -> bool:
    """Whether the rest of a footnote that a page ended by a forced break (a page break,
    ``pageBreakBefore``, a new section) could not hold goes on a page of its own before
    the next page of text: in mode 15.  Below 15 it goes at the foot of that next page.
    Measured by ``make_footnote_probe.py`` (five-line notes split two and three, every
    case's page led by ``pageBreakBefore``): mode 15 puts the three lines alone on a page;
    modes 12 and 14, and no settings, put them under the next case's text.
    """
    return compatibility_mode is not None and compatibility_mode >= 15


def _heights(lines) -> Fraction:
    return sum((line.height for line in lines), Fraction(0))


def _scan(document: Document, items: list, start: Position, rules: Rules, carry: tuple, story: str | None = None,
          bands=(), tops: dict | None = None, wraps=(), frame: "Frame | None" = None):
    """Lay one page out from ``start``: ``(end, placed, forced)``, ``placed`` holding per
    line kept ``(position, the page foot's notes after it)``.

    **The page foot.**  A line with footnote references fits when its end, and under it
    the separator (the continuation separator and the rest of a note from the page
    before, when there is one), every note already on the page as placed, every new
    note but the last whole, and of the last the least widow control lets stay
    (:func:`note_minimum`), all fit above the bottom margin.  When it is placed, the last
    new note takes as many lines as then fit (:func:`note_prefix`), and keeps them: a
    later line fits above the notes as placed, and does not shorten them.  Measured by
    ``make_footnote_probe.py`` in mode 15, every case; below 15, where a line with two
    notes may leave the second to the next page, not modelled on a page.

    **In a column** (``frame``, one of several; ``make_columns_probe.py``, ``footnote``,
    ``notes3``, ``notesover``, ``notesunequal``): below mode 15 the column's foot is a
    page's, its notes broken in its width, but **of a line's new notes only the first
    need keep its least**: the line goes in when that fits, and the notes after it take
    what is left in turn, the first that does not fit whole and every one after it going
    on to the next column's foot (``notes3``: the second note of a line at the column's
    foot two lines in the column and four in the next, the third wholly in the next).  In
    mode 15 (``frame.pooled``) a line's notes take no room of the column's: the page's
    note area is reserved below every column (:func:`_lay_columns_page`).
    """
    first = items[start.item]
    section_first = start.line == 0 and (start.item == 0 or items[start.item - 1].section != first.section)
    if frame is None:
        top, bottom = _geometry(document, first.section, items if rules.stories else None, section_first, story)
    else:
        top, bottom = frame.top, frame.bottom
    separator, continuation = getattr(items, "separators", {}).get(first.section, (Fraction(0), Fraction(0)))
    pooled = frame is not None and frame.pooled
    in_column = frame is not None and frame.multi and not pooled
    if frame is not None and frame.multi and bands:
        # Only the bands of drawings across this column.
        section = document.sections[frame.section]
        left = twips_to_px(section.margins.left + section.margins.gutter + frame.offset)
        right = left + twips_to_px(frame.width)
        bands = [band for band in bands if band.span is None or (band.span[0] < right and band.span[1] > left)]
    unequal_notes = in_column and getattr(items, "note_at", None) is not None and len(
        {width for _, width, _ in column_boxes(document.sections[frame.section])}) > 1
    y = top
    if isinstance(first, Para) and first.continuation is not None and (frame is None or frame.page_top):
        # A page the endnotes go on to: the continuation separator's line first.
        y += first.continuation.pitch
    previous: Para | None = None if frame is None else frame.previous
    first_on_page = True if frame is None else frame.first_on_page
    #: Where the stack ends after each line placed (for a :class:`Frame`'s foot).
    stack_ends: list = []
    #: Whether a floating table stands between the paragraph and the one before it.
    floated = False
    index, line = start.item, start.line
    #: The notes at the page foot, as ``(lines, how many of them are on this page)``.
    notes: list = []
    foot = (continuation + _heights(carry)) if carry else Fraction(0)
    placed: list = []

    def done(end: Position, forced: bool = False, why: str = "full"):
        # A floating table goes to the page of the paragraph it is positioned against.
        while end.line == 0 and not end.part and end.item - 1 > start.item and _floats(items[end.item - 1]):
            end = Position(end.item - 1)
        if frame is not None:
            frame.why = why
            frame.end = end
            frame.last = previous
            frame.notes = _page_notes(placed, end)
            frame.carried = tuple(carry)
            frame.y = next((after for position, after in reversed(stack_ends) if position < end), frame.top)
        return end, placed, forced

    while index < len(items):
        item = items[index]
        if frame is not None and line == 0 and item.section != frame.section:
            # A section that starts a new page ends the column and the page; a continuous
            # one whose columns differ, the region (:func:`_lay_columns`).
            if _new_page_section(document, item.section):
                return done(Position(index), True, "page")
            if frame.multi or is_multi_column(document.sections[item.section]):
                return done(Position(index), False, "region")
        if isinstance(item, Obstacle):
            raise Unplaceable(index)
        if isinstance(item, TableItem) and item.floating is not None:
            # A floating table takes no room in the stack: it is positioned against the
            # paragraph after it (the top of that paragraph: where the one before ends
            # after its space after), and the lines beside it go around it
            # (:func:`page_wraps`).
            if not first_on_page and item.section != previous.section and _new_page_section(document, item.section):
                return done(Position(index), True, "page")
            paragraph_top = y
            if not first_on_page and isinstance(previous, Para):
                paragraph_top = y + twips_to_px(previous.after)
            placed.append((Position(index, 0), list(notes)))
            if tops is not None:
                tops[(index, 0)] = (paragraph_top, Fraction(0), paragraph_top)
            floated = True
            index, line = index + 1, 0
            continue
        if isinstance(item, TableItem):
            if not first_on_page and line == 0 and not start.part and item.section != previous.section and \
                    _new_page_section(document, item.section):
                return done(Position(index), True, "page")
            if line == 0 and not first_on_page:
                y += table_gap_px(previous, item)
            if frame is not None and frame.first_top is None:
                frame.first_top = y
            part = start.part if index == start.item else ()
            if frame is not None and frame.offset:
                # In a column right of the first: the table laid out in that column.
                item.beside[start] = TableBeside(0, item.room(document, frame.offset, frame.width))
            elif line == 0 and not part:
                # Beside the drawings text wraps around on this page, as this layout of it
                # places them (none: across the column).
                item.beside[start] = table_beside(document, item, y, wraps, first_on_page) if wraps else None
            try:
                laid = place_item(item, y, bottom - foot, key=start, row=line, part=part,
                                  first_on_page=first_on_page, mode15=(document.compatibility_mode or 0) >= 15)
                if frame is None or not frame.multi:
                    # Footnotes referenced in its cells: each line holding one fits as a body
                    # line does, its notes under it at the page's foot (``make_cell_footnote_probe.py``).
                    laid, notes, foot = _table_notes(document, items, item, laid, y, bottom, foot, notes, carry,
                                                     separator, start, line, part, first_on_page)
            except table_model.Overflow as overflow:
                raise Unplaceable(index, overflow.args[0]) from None
            for piece in laid.pieces:
                if not piece.header:
                    placed.append((Position(index, piece.row, piece.start if piece.continued else ()), list(notes)))
                    stack_ends.append((Position(index, piece.row, piece.start if piece.continued else ()),
                                       piece.top + piece.top_border + piece.height))
            if laid.next is not None:
                next_row, next_part = laid.next
                if not any(not piece.header for piece in laid.pieces):
                    # Nothing of the table on this page: it starts the next, and a
                    # keepNext paragraph before it goes with it.
                    return done(_break_before(items, start, Position(index, next_row), rules,
                                              document.compatibility_mode))
                return done(Position(index, next_row, tuple(next_part) if next_part else ()))
            y = laid.end
            first_on_page = False
            previous = item
            index, line = index + 1, 0
            continue
        if not first_on_page and line == 0:
            if item.section != previous.section and _new_page_section(document, item.section):
                return done(Position(index), True, "page")
            if item.page_break_before:
                return done(Position(index), True, "page")
        if item.section_mark and not first_on_page and not section_mark_room(previous, document.compatibility_mode):
            # On this page, and taking no room: the stack goes on as if it were not there.
            placed.append((Position(index, 0), list(notes)))
            if tops is not None:
                tops[(index, 0)] = (y, item.heights[0].pitch, y)
            index, line = index + 1, 0
            continue
        paragraph_top = y
        if item.source is not None:
            # Lines an earlier layout of this page broke beside a drawing are broken
            # again in the whole column, and beside this layout's drawings as placed.
            item.plain(line)
        number = line - 1
        while number + 1 < item.count:
            number += 1
            if frame is not None and frame.multi:
                # One of unequal columns: the line is as wide as its column.
                item.fit_width(number, frame.width)
            height = item.heights[number]
            if number == 0:
                if first_on_page:
                    keeps = keeps_space_before_at_page_top(
                        section_start=item.section_start, page_break_before=item.page_break_before,
                        compatibility_mode=document.compatibility_mode)
                    above = items[index - 1] if index else None
                    gap, _ = page_top_gap_px(item.before, keeps=keeps,
                                             previous_after=above.after if isinstance(above, Para) else 0)
                else:
                    if isinstance(previous, TableItem):
                        # After a table, a paragraph's space before is kept whole.
                        gap = twips_to_px(item.before)
                    else:
                        same = item.style == previous.style
                        own_before = 0 if (item.contextual and same) else item.before
                        prev_after = 0 if (previous.contextual and same) else previous.after
                        gap, _ = paragraph_gap_px(prev_after, own_before)
                        if floated:
                            # A floating table between them: the two do not collapse.
                            gap = twips_to_px(prev_after) + twips_to_px(own_before)
                        paragraph_top = y + twips_to_px(prev_after)
                y += gap + item.top_border
                if frame is not None and frame.first_top is None:
                    frame.first_top = y - item.top_border
            if bands:
                # Below mode 15 a paragraph's last line reaches down to the end of its space
                # after; in mode 15 to the end of its pitch.
                reach = height.pitch + (twips_to_px(item.after) if number == item.count - 1
                                        and (document.compatibility_mode or 0) < 15 else 0)
                y = clear_of_bands(y, reach, bands)
            if wraps and item.source is not None:
                y = fit_beside(document, item, number, y, wraps,
                               (frame.offset, frame.width) if frame is not None and frame.multi else (0, None))
                height = item.heights[number]
            border = after = Fraction(0)
            if number == item.count - 1:
                border, after = item.border, twips_to_px(item.after)
            elif border_ends_every_page(document.compatibility_mode):
                border = item.border
            new_notes = item.notes.get(number, [])
            if new_notes and unequal_notes:
                # Broken in this column's width.
                new_notes = [items.note_at(lines[0].note, item.section, frame.width) or lines
                             for lines in new_notes]
            below = foot
            if new_notes and pooled:
                need = below
            elif new_notes and in_column:
                below += separator if not notes and not carry else 0
                partial = bool(notes) and notes[-1][1] < len(notes[-1][0])
                need = below + note_minimum(new_notes[0]) if not partial else bottom - y + 1
            elif new_notes:
                below += (separator if not notes and not carry else 0) + _heights(
                    line for note in new_notes[:-1] for line in note)
                partial = bool(notes) and notes[-1][1] < len(notes[-1][0])
                need = below + note_minimum(new_notes[-1]) if not partial else bottom - y + 1
            else:
                need = below
            if after and (document.compatibility_mode or 0) >= 15 and not pooled and not in_column and (
                    notes or carry or new_notes):
                # In mode 15 a paragraph's last line keeps its space after above the page's
                # footnotes, where it need not above the bottom margin
                # (``make_note_after_probe.py``: 240 twips after moved the line on 240
                # twips sooner, to the 20-twip step; no settings and mode 14 as without).
                need += after
            if not first_on_page and not item.section_mark and not fits(
                    y, height, bottom - need, border=border, after=after, rule=rules.fit):
                overflow = Position(index, number)
                if frame is None and number - 1 in item.auto_hyphen_ends and Position(index, number - 1) > start:
                    moves = hyphen_at_page_end(document)
                    if moves == "word" and item.source is not None and item.source.section is None and \
                            item.move_hyphenated_word(number - 1):
                        pass  # the line, its word gone, stays; the next starts with the word
                    elif moves is not None:
                        overflow = Position(index, number - 1)
                return done(_break_before(items, start, overflow, rules,
                                          document.compatibility_mode,
                                          empty_ok=frame is not None and not frame.first_on_page))
            if new_notes and pooled:
                notes.extend((note, len(note)) for note in new_notes)
            elif new_notes and in_column:
                room = bottom - (y + height.pitch - multiple_extra(height) + border) - below
                stopped = False
                for note in new_notes:
                    taken = 0 if stopped else note_prefix(note, room)
                    stopped = stopped or taken < len(note)
                    room -= _heights(note[:taken])
                    notes.append((note, taken))
                    below += _heights(note[:taken])
                foot = below
            elif new_notes:
                end_of_line = y + height.pitch - multiple_extra(height) + border
                for note in new_notes[:-1]:
                    notes.append((note, len(note)))
                taken = note_prefix(new_notes[-1], bottom - end_of_line - below)
                notes.append((new_notes[-1], taken))
                foot = below + _heights(new_notes[-1][:taken])
            placed.append((Position(index, number), list(notes)))
            if tops is not None:
                tops[(index, number)] = (y, height.pitch, paragraph_top)
                if frame is not None and frame.multi and hasattr(tops, "columns"):
                    tops.columns[(index, number)] = (frame.offset, frame.width)
            y += height.pitch
            stack_ends.append((Position(index, number), y + (item.border if number == item.count - 1 else 0)))
            first_on_page = False
            if number in item.page_breaks:
                return done(Position(index, number + 1) if number + 1 < item.count else Position(index + 1), True,
                            "page")
            if number in item.column_breaks:
                return done(Position(index, number + 1) if number + 1 < item.count else Position(index + 1), True,
                            "column")
        y += item.border
        previous = item
        floated = False
        index, line = index + 1, 0
    return done(Position(len(items)), False, "end")


def _auto_hyphen_ends(pieces: list, broken: list) -> frozenset[int]:
    from .linebreak import _auto_hyphenated

    return frozenset(number for number, line in enumerate(broken) if _auto_hyphenated(pieces, line))


def hyphen_at_page_end(document: Document) -> str | None:
    """What Word does with a page whose last line would end in an automatic hyphen:
    ``"line"`` -- that line goes to the next page; ``"word"`` -- only the hyphenated word
    does, the rest of its line staying; ``None`` -- nothing, the page ends in the hyphen.

    Measured by ``make_hyphen_probe.py`` (``bottom``: a hyphenated paragraph, the page
    ending after each of its lines; ROADMAP.md, 3.9.2), as [MS-DOCX] 2.3.7 and 2.3.8 say:
    in mode 15 with neither compatSetting the line moves -- one line: the line before it
    may end in a hyphen and stay -- and the keeps then apply as to any line that does not
    fit; with ``useWord2013TrackBottomHyphenation`` false the word moves; with
    ``allowHyphenationAtTrackBottom`` true, and below mode 15 with neither, nothing
    moves.  Below mode 15 with either setting is not measured, nor a column's end."""
    if (document.compatibility_mode or 0) < 15:
        return None
    on = ("1", "true", "on")
    settings = document.compat_settings
    if str(settings.get("allowHyphenationAtTrackBottom", "0")).lower() in on:
        return None
    word2013 = settings.get("useWord2013TrackBottomHyphenation")
    return "line" if word2013 is None or str(word2013).lower() in on else "word"


def _floats(item) -> bool:
    return isinstance(item, TableItem) and item.floating is not None


def table_wrap(document: Document, item: TableItem, paragraph_top: Fraction, page: int):
    """The floating table ``item`` as text wraps beside it on its page
    (:class:`docx2svg.wrap.Wrap`), positioned against ``paragraph_top`` -- the top of the
    paragraph after it -- as :func:`docx2svg.table.floating_top` says.

    Measured by ``make_float_table_probe.py`` (ROADMAP.md, "Floating tables"): text keeps
    off the table's box -- its outer grid lines and the outer halves of its borders, or
    15 twips where it has none (:func:`docx2svg.table.floating_edges`), from its top to
    below its bottom border -- widened by ``topFromText`` and ``bottomFromText`` and,
    across, by ``leftFromText`` and ``rightFromText``, each at least 10 twips; and goes
    beside it as beside a ``wrapSquare`` drawing on both sides (:mod:`docx2svg.wrap`):
    which lines, the 360-twip narrowest segment, the pixel rounding of the text's start,
    lines that go below a table too wide to leave room, lines before the table in the
    document that it reaches up to, and a table in the flow beside it."""
    from .wrap import Wrap

    flow = item.flow
    floating = flow.floating
    mode15 = (document.compatibility_mode or 0) >= 15
    far = Fraction(10 ** 7)
    height = table_model.place_table(flow, Fraction(0), far, first_on_page=False, mode15=mode15).end
    y = table_model.floating_top(floating, document.sections[item.section], paragraph_top, height)
    left, right = table_model.floating_edges(flow)
    least = table_model.MIN_FROM_TEXT
    # Across in exact twips, as a drawing's edges are (a room beside it is laid out in
    # whole twips: ``TableItem.room``).
    px = Fraction(300, 1440)
    first, last = ((flow.page_left_twips + flow.grid_twips[k]) * px for k in (0, -1))
    return Wrap(y - twips_to_px(floating.top), y + height + twips_to_px(floating.bottom), first - left * px,
                last + right * px, max(floating.left, least) * px, max(floating.right, least) * px, "bothSides", y,
                ("table", item.block), x=first, half_up=True)


def wrap_column(document: Document, section_number: int, offset: int = 0, width: int | None = None):
    """The text column of a section's page as :mod:`docx2svg.wrap` measures from it; one of
    several ``offset`` twips from the text area's left and ``width`` wide, its start
    rounded as the layout rounds a column's (:meth:`docx2svg.layout._Placer.column_edges`)."""
    from .linebreak import column_width_twips, twips_to_units
    from .vertical import LAYOUT_UNIT_PX, round_half_up
    from .wrap import Column

    section = document.sections[section_number]
    total = section.margins.left + section.margins.gutter + offset
    exact = Fraction(total * 300, 1440)
    width = column_width_twips(section) if width is None else width
    mode15 = (document.compatibility_mode or 0) >= 15
    start = round_half_up(exact)
    if offset and mode15:
        start = round_half_up(twips_to_units(total) * LAYOUT_UNIT_PX)
    return Column(exact, start, Fraction(twips_to_units(width)), mode15, Fraction(width * 300, 1440))


class Tops(dict):
    """``(flow index, line) -> (top, pitch, paragraph top)`` of the lines a page placed,
    and in ``columns`` the column of each line placed in one of several: ``(offset,
    width)`` in twips."""

    def __init__(self):
        super().__init__()
        self.columns: dict = {}


def fit_beside(document: Document, item: "Para", number: int, y: Fraction, wraps, box: tuple = (0, None)
               ) -> Fraction:
    """Break line ``number`` of ``item``, whose top is at ``y``, beside the drawings
    ``wraps`` it reaches into (:mod:`docx2svg.wrap`), take it and the lines after it as
    the paragraph's, and return where its top is -- lower than ``y`` where no segment at
    ``y`` took text, and it went down to the foot of the drawings it is beside.  A line
    reaches down its pitch, in mode 15 less an ``auto`` multiple's extra."""
    from . import wrap as wrap_model
    from .linebreak import Line, break_line

    src = item.source
    pieces = src.pieces
    n = len(pieces)
    start = item.lines[number].start
    first = number == 0
    column = wrap_column(document, item.section, *box)

    def reach(height) -> Fraction:
        if column.mode15:
            return height.pitch - multiple_extra(height)
        return height.pitch + (twips_to_px(item.after) if number == item.count - 1 else 0)

    plain = Line(n, n) if start >= n else break_line(pieces, start, src.geometry, src.advances, first=first)
    moved = None
    line = plain
    for _attempt in range(64):
        height = src.height(plain, first)
        if height is None:
            return y
        segs, beside = wrap_model.segments(wraps, y, y + reach(height), column, y + height.pitch
                                           - multiple_extra(height))
        if segs is None:
            line = plain
            break
        found = wrap_model.break_across(pieces, start, segs, src.geometry, src.advances, first=first)
        if found is not None:
            line = found
            again = src.height(found, first)
            if again is not None and again.pitch != height.pitch:
                # The line's own height, not the whole column's, says what it is beside.
                others, _ = wrap_model.segments(wraps, y, y + reach(again), column,
                                                y + again.pitch - multiple_extra(again))
                if others is None:
                    line = plain
                elif others != segs:
                    line = wrap_model.break_across(pieces, start, others, src.geometry, src.advances,
                                                   first=first) or found
            break
        below = min(w.bottom for w in beside)
        if below <= y:
            break
        y = moved = below
    if moved is not None:
        if not isinstance(line, wrap_model.WrappedLine):
            line = wrap_model.WrappedLine(line.start, line.end, forced=line.forced, emergency=line.emergency,
                                          hyphenated=line.hyphenated)
        line.drop = moved
    old = item.lines[number]
    if (not isinstance(line, wrap_model.WrappedLine) and not isinstance(old, wrap_model.WrappedLine)
            and (old.start, old.end) == (line.start, line.end) and item.plain_from <= number):
        return y
    tail: list = []
    if line.end < n or (line.forced and n and pieces[n - 1].char == "\n"):
        tail = src.rest(line.end, first=False)
    item.set_lines(item.lines[:number] + [line] + tail,
                   number + 1 if isinstance(line, wrap_model.WrappedLine) else number)
    return y


@dataclass(frozen=True)
class TableBeside:
    """Where a table goes beside the drawings text wraps around on its page
    (:func:`table_beside`): from row ``row`` on, laid out in the room beside them
    (``flow``, the table in a narrower column) or, where it fits in none, moved down to
    ``drop``, the foot of the drawings -- and there beside them again, or across the
    column."""

    row: int
    flow: object = None
    drop: Fraction | None = None


def table_fits(document: Document, flow, width) -> bool:
    """Whether a table laid out in a room ``width`` twips wide keeps within it: in mode 15
    its right border's outer edge, below mode 15 its last cell's text (its grid line less
    the cell's right margin) -- measured by ``make_wrap_table_probe.py`` (``fit``): beside a
    room of 5,205 twips a table of 5,195 twips and ``w:sz`` 4 borders fits in mode 15 and
    5,196 does not, 5,145 with ``w:sz`` 24, 5,205 with none, 4,195 indented 1,000; below it
    5,313 fits and 5,314 does not, 108 twips of right cell margin past it."""
    resolved = flow.resolved
    last = resolved.rows[0].cells[-1] if resolved.rows and resolved.rows[0].cells else None
    right = flow.grid_twips[-1]
    if last is not None:
        if table_model.compatible_mode(document):
            right += table_model.outer_part(last.right_edge[0])
        else:
            right -= table_model.right_offset(last)
    return right <= width


def table_beside(document: Document, item: TableItem, y: Fraction, wraps, first_on_page: bool
                 ) -> TableBeside | None:
    """Where a table that starts at ``y`` goes beside the drawings text wraps around on its
    page (``wraps``), or ``None`` where it is beside none.  Measured by
    ``make_wrap_table_probe.py`` (ROADMAP.md, "Floating drawings -- measured", F.15):

    * **Which rows.**  A row is beside a drawing when it reaches into the drawing's box
      widened by ``distT`` and ``distB`` (touching is not).  **In mode 15 the whole table
      goes beside it** when any of its rows does (a drawing that starts in its third row
      moves the first two too); **below mode 15 the rows from the first that is beside it
      on** (the rows above stay where they are).
    * **Where.**  In what the drawings leave of the column, as a line's segments
      (:func:`docx2svg.wrap.free_spans`: ``distL`` / ``distR``, ``wrapText``, a span
      narrower than 360 twips not counted): the table is laid out as in a column that is
      that span -- its indent, ``w:jc`` and the pixel rounding all from the span's left
      edge -- when it fits there (:func:`table_fits`); in mode 15 the first span it fits
      in, below mode 15 the first span only.
    * **Otherwise it goes down** to the foot of the drawings it is beside, and is tried
      again there: a table too wide for the room starts below the drawing, at its place
      in the column.
    * Autofit tables are not narrowed to fit, and ``wrapTight`` / ``wrapThrough`` keep a
      table off their box, as ``wrapSquare`` does (the rectangles measured).
    """
    from .wrap import MIN_SEGMENT_TWIPS, TWIP_PX, free_spans

    mode15 = (document.compatibility_mode or 0) >= 15
    flow = item.flow
    column = wrap_column(document, item.section)
    far = Fraction(10 ** 7)
    laid = table_model.place_table(flow, y, y + far, row=0, part=(), first_on_page=first_on_page, mode15=mode15)
    pieces = [piece for piece in laid.pieces if not piece.header]
    if not pieces:
        return None
    if mode15:
        start = 0
        if not free_spans(wraps, y, laid.end, column)[1]:
            return None
    else:
        start = next((k for k, piece in enumerate(pieces)
                      if free_spans(wraps, piece.top, piece.top + piece.top_border + piece.height, column)[1]), None)
        if start is None:
            return None
    top = pieces[start].top
    drop = None
    for _attempt in range(64):
        rest = table_model.place_table(flow, top, top + far, row=pieces[start].row, part=(),
                                       first_on_page=first_on_page and start == 0, mode15=mode15)
        spans, near = free_spans(wraps, top, rest.end, column)
        if not near:
            break
        rooms = [(a, b) for a, b, _edge in spans if (b - a) / TWIP_PX >= MIN_SEGMENT_TWIPS]
        for a, b in rooms if mode15 else rooms[:1]:
            placed = item.room(document, a / TWIP_PX, (b - a) / TWIP_PX)
            if placed is not None and table_fits(document, placed, (b - a) / TWIP_PX):
                return TableBeside(pieces[start].row, placed, drop)
        below = min(wrap.bottom for wrap in near)
        if below <= top:
            break
        top = drop = below
    return TableBeside(pieces[start].row, None, drop) if drop is not None else None


def place_item(item: TableItem, y: Fraction, bottom: Fraction, *, key, row: int = 0, part: tuple = (),
               first_on_page: bool, mode15: bool):
    """:func:`docx2svg.table.place_table` for a table item, beside the drawings text wraps
    around on the page that starts at ``key`` where :func:`table_beside` put it there."""
    beside = item.beside.get(key)
    if beside is None:
        return table_model.place_table(item.flow, y, bottom, row=row, part=part, first_on_page=first_on_page,
                                       mode15=mode15)
    pieces: list = []
    if row < beside.row:
        head = table_model.place_table(item.flow, y, bottom, row=row, part=part, first_on_page=first_on_page,
                                       mode15=mode15, stop=beside.row)
        if head.next != (beside.row, ()):
            return head
        pieces, y, row, part = list(head.pieces), head.end, beside.row, ()
        first_on_page = first_on_page and not pieces
    flow = beside.flow or item.flow
    if beside.drop is not None and row == beside.row and not part:
        y = max(y, beside.drop)
    rest = table_model.place_table(flow, y, bottom, row=row, part=part, first_on_page=first_on_page, mode15=mode15)
    if flow is not item.flow:
        rest.pieces = [dataclasses.replace(piece, flow=flow) for piece in rest.pieces]
    return table_model.TablePage(pieces + rest.pieces, rest.end, rest.next)


def cell_notes(items, item: TableItem, flow) -> dict:
    """``(row, cell, stacked line) -> [laid note lines]`` for every footnote referenced in
    the table's cells (``items.note_at``, laid out in the section's width), in document
    order; empty where it has none."""
    cache = item.__dict__.setdefault("cell_notes", {})
    if id(flow) in cache:
        return cache[id(flow)]
    out: dict = {}
    note_at = getattr(items, "note_at", None)
    for r, row in enumerate(flow.rows):
        for c, cell in enumerate(row.cells):
            for k, stacked in enumerate(cell.stack):
                paragraph = stacked.paragraph
                model = getattr(paragraph, "paragraph", None)
                if model is None or not paragraph.lines:
                    continue
                references = [(offset + position, int(kind.split(":")[1]))
                              for offset, run in zip(_run_offsets(model), model.runs)
                              for position, kind in run.breaks if kind.startswith("footnoteReference:")]
                if not references or note_at is None:
                    continue
                line_of = _line_of_sources(paragraph.pieces, paragraph.lines)
                for source, note_id in references:
                    if line_of.get(source, line_of.get(source - 1, len(paragraph.lines) - 1)) != stacked.number:
                        continue
                    laid = note_at(note_id, item.section)
                    if laid:
                        out.setdefault((r, c, k), []).append(laid)
    cache[id(flow)] = out
    return out


def _table_notes(document: Document, items, item: TableItem, laid, y: Fraction, bottom: Fraction, foot: Fraction,
                 notes: list, carry: tuple, separator: Fraction, start: Position, row: int, part: tuple,
                 first_on_page: bool):
    """The table's part of the page (``laid``) with its cells' footnotes at the foot:
    ``(laid, notes, foot)``.

    A cell line holding a reference fits as a body line with one does (4.3, 4.13): its end,
    with the cell's bottom margin and the edge under its row, above the separator, the
    notes on the page, its new notes but the last whole and the least of the last; when it
    goes in, the last note takes what then fits.  A later line fits above the notes as
    placed.  Where a line does not, the table is placed again ending above it: its row
    splits there or moves (``make_cell_footnote_probe.py``: a one-line row with a note
    moves to the next page with it, a cell's lines split before the one holding the
    reference).  Measured in one column; notes in cells otherwise are not drawn."""
    mode15 = (document.compatibility_mode or 0) >= 15
    limit = bottom
    for _attempt in range(64):
        flowed = []
        for piece in laid.pieces:
            if piece.header:
                continue
            flow = piece.flow or item.flow
            found = cell_notes(items, item, flow)
            current = flow.rows[piece.row]
            edge = table_model._edge_below(flow, piece.row)
            for c, cell in enumerate(current.cells):
                top = piece.top + piece.top_border + twips_to_px(cell.cell.margins["top"])
                for k in range(piece.start[c], piece.end[c]):
                    end = top + table_model.slice_height(cell, piece.start[c], k + 1, piece.continued) + twips_to_px(
                        cell.cell.margins["bottom"]) + edge
                    flowed.append((end, piece.row, c, k, found.get((piece.row, c, k), [])))
        if not any(new for *_, new in flowed):
            return laid, notes, foot
        flowed.sort(key=lambda entry: (entry[0], entry[1], entry[2], entry[3]))
        trial_notes, trial_foot = list(notes), foot
        violation = None
        for number, (end, r, c, k, new) in enumerate(flowed):
            below = trial_foot
            if new:
                below += (separator if not trial_notes and not carry else 0) + _heights(
                    line for note in new[:-1] for line in note)
                partial = bool(trial_notes) and trial_notes[-1][1] < len(trial_notes[-1][0])
                need = below + note_minimum(new[-1]) if not partial else bottom - y + 1
            else:
                need = below
            first = first_on_page and number == 0
            if not first and end > bottom - need:
                violation = end
                break
            if new:
                for note in new[:-1]:
                    trial_notes.append((note, len(note)))
                taken = note_prefix(new[-1], bottom - end - below)
                trial_notes.append((new[-1], taken))
                trial_foot = below + _heights(new[-1][:taken])
        if violation is None:
            return laid, trial_notes, trial_foot
        # Placed again, ending just above the line that does not fit.
        limit = min(limit, violation - Fraction(1, 10 ** 6))
        laid = place_item(item, y, limit, key=start, row=row, part=part, first_on_page=first_on_page,
                          mode15=mode15)
    return laid, notes, foot


def table_gap_px(previous, item: TableItem) -> Fraction:
    """The space between a paragraph and the table after it: the paragraph's space
    after, whole -- no collapse with the first cell's space before, which is inside the
    cell -- unless it is contextual and the table's first cell starts with a paragraph of
    its style (``make_table_content_probe.py``, ``spacing contextual style, the same
    outside``).

    Between two tables with nothing between them the edge is **one**: the second starts
    the narrower of the first's bottom border and its own top border higher, so the two
    stand on one line (``make_wrap_tables_probe.py``, ``adjacent``: the second table and
    everything after it 2 px higher in Word, its 0.5 pt borders; ROADMAP.md, F.22)."""
    if isinstance(previous, TableItem):
        rows = item.flow.rows
        if not rows:
            return Fraction(0)
        return -min(previous.flow.bottom_border, rows[0].top_border)
    if not isinstance(previous, Para):
        return Fraction(0)
    rows = item.flow.rows
    first = rows[0].cells[0].paragraphs[0] if rows and rows[0].cells and rows[0].cells[0].paragraphs else None
    if "content" in table_model.STAGES and previous.contextual and first is not None and first.style == previous.style:
        return Fraction(0)
    return twips_to_px(previous.after)


def _page_notes(placed: list, end: Position) -> list:
    """The notes at the foot of the page, as ``(lines, how many are on it)``."""
    kept = [entry for entry in placed if entry[0] < end]
    return list(kept[-1][1]) if kept else []


def note_rest(notes: list) -> tuple:
    """The lines of ``notes`` (``(lines, how many are placed)``) not placed: what goes on to
    the next column's or page's foot."""
    return tuple(line for lines, taken in notes for line in lines[taken:])


def _carried(placed: list, end: Position) -> tuple:
    """The lines of the page's last footnote that stay off it, for the next page."""
    kept = [entry for entry in placed if entry[0] < end]
    if not kept or not kept[-1][1]:
        return ()
    lines, taken = kept[-1][1][-1]
    return tuple(lines[taken:])


def _break_before(items: list, start: Position, overflow: Position, rules: Rules,
                  compatibility_mode: int | None = None, *, empty_ok: bool = False) -> Position:
    """Where the page ends, given that the line at ``overflow`` does not fit on it.
    ``empty_ok``: the column did not start empty (a region under another on its page), so
    the keeps may leave it holding nothing."""
    target = _keeps(items, start, overflow, rules, compatibility_mode)
    if empty_ok:
        return max(target, start)
    if target <= start:
        return overflow if rules.oversize == "ignore" else target
    return target


def page_break_ends_paragraph(compatibility_mode: int | None) -> bool:
    """Whether widow control treats the line a manual page break ends as a paragraph's
    last: in mode 15 it does -- three lines and a page break, with room for two, go to the
    next page together, as three lines ending a paragraph would; below 15 and unstated
    the page break is not an end, and the two stay.  Measured by ``make_keep_probe.py``
    (``manual segment``: 2..5 lines, a page break, two lines; every room)."""
    return compatibility_mode is not None and compatibility_mode >= 15


def keep_next_moves(rules: Rules, compatibility_mode: int | None) -> str:
    """What a ``keepNext`` paragraph gives the page after it: in compatibility mode 15
    its **last line** (with that line's keeps), below 15 and when unstated the **whole
    paragraph**.  Measured by ``make_keep_probe.py`` (``keepnext``: a four-line kept
    paragraph that fits, before one that does not): mode 15 moves one line (two with
    widow control), every other setting all four.
    """
    if rules.keep_next_moves != "mode":
        return rules.keep_next_moves
    return "line" if compatibility_mode is not None and compatibility_mode >= 15 else "paragraph"


def _keeps(items: list, start: Position, overflow: Position, rules: Rules,
           compatibility_mode: int | None = None) -> Position:
    """The first line that must go to the next page when the line at ``overflow`` does
    not fit: that line, moved back by the paragraph's keeps, and then back over every
    ``keepNext`` paragraph before it on the page.

    * **Widow and orphan control** (on unless the paragraph or its style turns it off --
      also with no ``settings.xml`` and no styles part): a paragraph leaves at least two
      lines at the bottom of a page and takes at least two to the next; a paragraph that
      would leave one line behind moves whole.  In mode 15 the line a manual page break
      ends counts as a last line (:func:`page_break_ends_paragraph`).
    * **``keepLines``**: a paragraph that does not fit moves whole -- also one taller
      than a page, which then starts the next page and breaks there.
    * **``keepNext``**: a paragraph whose next one starts the next page goes with it --
      whole below mode 15; in mode 15 only its *last line*, with that line's keeps
      (widow control makes that two lines, ``keepLines`` the whole paragraph)
      (:func:`keep_next_moves`).  A one-line paragraph moves whole either way, and the
      paragraph before it is asked in turn: a chain moves together, even one taller
      than a page.
    """
    index, number = overflow.item, overflow.line
    item = items[index]
    first = start.line if index == start.item else 0
    target = overflow
    if number > 0:
        if item.keep_lines and first == 0:
            target = Position(index)
        elif (rules.widow_default if item.widow_control is None else item.widow_control) and rules.widow_lines > 1:
            end = item.count
            if (page_break_ends_paragraph(compatibility_mode) if rules.segment == "mode"
                    else rules.segment == "always"):
                end = min((k + 1 for k in item.page_breaks if k >= number), default=item.count)
            below = end - number
            if below < rules.widow_lines:
                number = max(first, end - rules.widow_lines)
            if first == 0 and 0 < number < rules.widow_lines:
                number = 0
            target = Position(index, number)
    if target.line == 0 and target.item - 1 >= start.item:
        previous = items[target.item - 1]
        if isinstance(previous, Para) and previous.keep_next:
            if keep_next_moves(rules, compatibility_mode) == "line":
                return _keeps(items, start, Position(target.item - 1, previous.count - 1), rules,
                              compatibility_mode)
            return _keeps(items, start, Position(target.item - 1, 0), rules, compatibility_mode)
    return target


def paginate(document: Document, items: list, rules: Rules | None = None, *,
             start: Position = Position(0)) -> Pagination:
    """Every page of ``items`` (:func:`flow`) from ``start``, until the end or an obstacle."""
    rules = rules or Rules()
    out = Pagination()
    position = start
    page = 0
    number = 0
    carry: tuple = ()
    while position.item < len(items):
        item = items[position.item]
        section_first = position.line == 0 and not position.part and (
            position.item == 0 or items[position.item - 1].section != item.section)
        if section_first and position.item > 0:
            number, blanks = section_start_number(document, item.section, number, page)
            for _ in range(blanks):
                out.info[page] = PageInfo(number - 1, items[position.item - 1].section, False, "default", True)
                page += 1
        elif section_first and document.sections[item.section].page_number_start is not None:
            number = document.sections[item.section].page_number_start
        else:
            number += 1
        kind = story_kind(document, item.section, section_first=section_first, number=number)
        out.info[page] = PageInfo(number, item.section, section_first, kind)
        out.starts.append(position)
        try:
            bands: list = []
            wraps: list = []
            frames: list = []
            carried = carry
            end, carry, forced, notes = _lay_page(document, items, position, rules, carry, kind, page, bands, wraps,
                                                  frames)
            if frames:
                out.frames[page] = frames
            if carried or notes:
                out.notes[page] = (carried, notes)
            if bands:
                out.bands[page] = bands
            if wraps:
                out.wraps[page] = wraps
        except Unplaceable as stop:
            if frames:
                out.frames[page] = frames
            if bands:
                out.bands[page] = bands
            if wraps:
                out.wraps[page] = wraps
            out.stopped = stop.args[0]
            out.stopped_line = stop.args[1] if len(stop.args) > 1 else 0
            _assign(items, out, position, Position(stop.args[0], out.stopped_line), page)
            return out
        _assign(items, out, position, end, page)
        out.ends.append(end)
        position = end
        page += 1
        if carry and forced and continuation_page(document.compatibility_mode):
            # The rest of the note takes a page of its own before the forced break's.
            number += 1
            out.info[page] = PageInfo(number, item.section, False, story_kind(
                document, item.section, section_first=False, number=number), True)
            out.notes[page] = (carry, [])
            page += 1
            carry = ()
    return out


def _assign(items: list, out: Pagination, start: Position, end: Position, page: int) -> None:
    """Record the page of every line from ``start`` to ``end``; a table row split across
    pages counts on the page it starts on."""
    index, line = start.item, start.line + (1 if start.part else 0)
    while Position(index, line) < end and index < len(items):
        item = items[index]
        count = item.count if isinstance(item, (Para, TableItem)) else 1
        stop = (end.line + (1 if end.part else 0)) if index == end.item else count
        out.pages.setdefault(index, []).extend([page] * (stop - line))
        index, line = index + 1, 0


def drawn_pages(items: list, result: Pagination) -> dict[int, list[int]]:
    """Block index -> the page of each line Word would draw (:attr:`Para.undrawn` left out);
    a table's cell paragraphs each line on the page its part of its row is on."""
    out = {}
    for index, found in result.pages.items():
        item = items[index]
        if isinstance(item, TableItem):
            continue
        undrawn = item.undrawn if isinstance(item, Para) else frozenset()
        out[item.block] = [page for number, page in enumerate(found) if number not in undrawn]
    ends = list(result.ends)
    if result.stopped is not None:
        ends.append(Position(result.stopped, result.stopped_line))
    for k, (start, end) in enumerate(zip(result.starts, ends)):
        page = _page_number(result, k)
        for index in range(start.item, min(end.item + 1, len(items))):
            item = items[index]
            if not isinstance(item, TableItem):
                continue
            first_row = start.line if index == start.item else 0
            last_row = end.line if index == end.item else item.count
            for r in range(first_row, min(last_row + 1, item.count)):
                row = item.flow.rows[r]
                for c, cell in enumerate(row.cells):
                    lo = start.part[c] if (index == start.item and r == start.line and start.part) else 0
                    hi = end.part[c] if (index == end.item and r == end.line and end.part) else (
                        0 if (index == end.item and r == end.line) else len(cell.stack))
                    for placed in cell.stack[lo:hi]:
                        paragraph = placed.paragraph
                        if placed.number in paragraph.undrawn:
                            continue
                        out.setdefault(paragraph.block, []).append(page)
    return out


def _page_number(result: Pagination, k: int) -> int:
    """The page (0-based, blank pages counted) that ``result.starts[k]`` begins: the page
    of its first line -- or, going on with a row split across pages, the page after the
    one before."""
    start = result.starts[k]
    if start.part and k:
        return _page_number(result, k - 1) + 1
    found = result.pages.get(start.item)
    if found and start.line < len(found):
        return found[start.line]
    return k

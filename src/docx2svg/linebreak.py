"""Word's line breaker, as measured: which characters of a paragraph go on which line.

ROADMAP.md, "Phase 3 -- measured", holds the observations behind every rule here and the
hypotheses they refuted.  In short:

* **The budget is exact, inclusive, with no slack.**  A line holds what it holds when the
  pen position after its last glyph is at or before the right edge -- ``x <= right`` in
  Word's layout unit, 1/4096 pt, to the unit (``make_wrap_budget_probe.py``: 3,120 of
  3,120 cases, δ swept through zero in steps of one unit, eight faces, four
  compatibility settings).  The right edge is the column less ``w:ind/@w:right``; the
  line starts at ``w:ind/@w:left``, plus ``@w:firstLine`` or less ``@w:hanging`` on the
  paragraph's first line.
* **Spaces at the end of a line do not count**, however many; nor does the paragraph mark.
* **Greedy first fit**: each line takes as much as fits, and breaks at the last
  opportunity that does; a word that does not fit on an empty line is broken after the
  last character that does.
* **Justified text in compatibility mode 15 may run past the edge** by a quarter of each
  space on the line, rounded to the unit, inclusive: Word compresses the spaces
  (:func:`squeeze`, ``make_justify_probe.py``).  Not below mode 15, and not in text that
  is not justified.
* **Automatic hyphenation** (``w:autoHyphenation``) puts a soft hyphen at every point of
  every word (:func:`auto_hyphens`, the points from :mod:`docx2svg.hyphenate`), which a
  line ends at when its hyphen fits and Word's rules allow it (:func:`auto_hyphen_allowed`:
  the zone below mode 15, a threshold of its own in mode 15, the letter after the point,
  ``w:consecutiveHyphenLimit``, ``w:noProof``; ``make_hyphen_probe.py``).

Everything is computed in the layout unit, 1/4096 pt, as an ``int`` wherever it is exact
(every 2048-upm face at a half-point size) and a :class:`~fractions.Fraction` otherwise.
Nothing horizontal rounds but twip lengths, which land on the unit half up (Phase 2).

Standard library only; widths come from an :class:`~docx2svg.measure.Advances`.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import NamedTuple
from fractions import Fraction

from .model import Document, Paragraph, Section
from .resolve import character_format, resolve_mark, resolve_paragraph, resolve_run, script_half_points

#: 1/4096 pt per twip.
UNITS_PER_TWIP = Fraction(4096, 20)

GLYPH = "glyph"
SPACE = "space"
TAB = "tab"
BREAK = "break"
SOFT_HYPHEN = "soft hyphen"
#: An inline picture (``wp:inline``): as wide as its extent, as tall as the line needs.
OBJECT = "object"

#: Characters that are a breaking space: a line may end after any number of them, and at
#: the end of a line they do not count toward its width.  The space, the en space, the em
#: space and the four-per-em space; the zero-width space too (it has no width to count,
#: and is not drawn).  *Not* the three- and six-per-em, punctuation, thin, hair, figure,
#: narrow no-break or no-break spaces: a line never breaks at those, and at the end of a
#: line they count like letters (``make_break_rules_probe.py``, families ``after``,
#: ``before`` and ``hang``, in Calibri and Times New Roman alike).
SPACES = frozenset(" \u2002\u2003\u2005\u200b")
#: Characters after which a line may break: the hyphen-minus, the en dash and the em
#: dash, between letters, between digits and between the two.  Not before them; and not
#: after or before any other punctuation measured -- ``/``, ``\\``, ``.``, ``,``, ``:``,
#: ``;``, ``?``, ``!``, brackets, quotes, ``%``, ``&``, ``+``, ``…``, ``•``, the figure dash
#: U+2012 or the horizontal bar U+2015 among them.
BREAK_AFTER = frozenset("-\u2013\u2014")
#: Word's tab-stop interval when the settings do not state ``w:defaultTabStop``: 708
#: twips (1.25 cm), measured on this machine's Word, whose ``Normal.dotm`` is metric.
#: Like the 12 pt default size (``resolve``), a fact about the application, not the file.
APPLICATION_DEFAULT_TAB = 708


def twips_to_units(twips: int | Fraction) -> int:
    """A twip length as Word holds it: in 1/4096 pt, half up (Phase 2, 2.1)."""
    return math.floor(Fraction(twips) * UNITS_PER_TWIP + Fraction(1, 2))


def emu_twips(emu: int) -> int:
    """An EMU length as Word lays a picture out: whole twips, truncated (635 EMU to the
    twip).  Measured on picture heights (``make_picture_probe.py``): 200 of 200 lines
    per setting against 169 rounding to the twip and 178 keeping the fraction."""
    return emu // 635


class Unmeasurable(Exception):
    """A piece of the paragraph has no known width (a face or character without an
    advance, a drawing, a superscript whose face has no script size)."""


@dataclass
class Piece:
    kind: str
    #: The character drawn (``"\\t"`` for a tab, ``"\\n"`` for a break, ``"\\f"`` for a
    #: page break and ``"\\v"`` for a column break).
    char: str
    #: Its advance, tracking included, in 1/4096 pt (0 for a tab or break).
    width: int | Fraction = 0
    #: Pieces with the same key kern against each other; ``None`` when kerning is off.
    kern_key: tuple | None = None
    #: Index into the paragraph's text (its runs' ``text``, concatenated), or -1 for the
    #: list label.
    source: int = -1
    #: What a soft hyphen costs when the line breaks at it (the hyphen drawn).
    hyphen: int | Fraction = 0
    #: ``(face, bold, italic)`` for kerning lookups.
    face: tuple | None = None
    #: An :data:`OBJECT`'s height, in whole twips (its extent and effect extent).
    height: int = 0
    #: The index of the run the piece comes from (-1 for the list label): what the
    #: renderer draws it with.  Breaking does not read it.
    run: int = -1
    #: A :data:`SOFT_HYPHEN` that automatic hyphenation put at a break point of a word
    #: (:func:`auto_hyphens`), not one the text holds.
    auto: bool = False
    #: For an automatic hyphen in a word that runs into a ``w:noProof`` run: the source
    #: index where that run starts.  Word takes the hyphen only when the line runs past the
    #: edge before it (``make_hyphen_probe.py``, ``noproof``).
    proof_end: int | None = None


@dataclass(frozen=True)
class TabStop:
    position: int  # 1/4096 pt from the column's left edge
    alignment: str = "left"
    #: ``w:leader`` (``dot``, ``hyphen``...; ``None`` or ``none``: no leader).
    leader: str | None = field(default=None, compare=False)


@dataclass(frozen=True)
class Geometry:
    """Where a paragraph's lines start and end, from the column's left edge, in 1/4096 pt."""

    first_start: int
    start: int
    right: int
    tabs: tuple[TabStop, ...] = ()
    default_tab: int = twips_to_units(APPLICATION_DEFAULT_TAB)
    #: ``w:ind/@w:left`` when the paragraph hangs: an implicit stop on its first line.
    hanging_stop: int | None = None
    #: Justified text in compatibility mode 15: a line may run past the right edge by a
    #: quarter of each space before it, rounded (:func:`squeeze`).
    squeeze: bool = False
    #: Where tab stops are measured from: the column's left edge, or below mode 15 a
    #: segment's start beside a floating drawing (:mod:`docx2svg.wrap`).
    tab_origin: int | Fraction = 0
    #: Automatic hyphenation (:func:`hyphen_rules`).  Below mode 15: the hyphenation zone,
    #: in units -- a word is broken at an automatic point only if it starts at least this
    #: far before the right edge; and whether the letter after the point must fit too.
    hyphen_zone: int | None = None
    hyphen_next: bool = False
    #: In mode 15, text not justified: the room after the line's text before the word
    #: must exceed this many units (the zone is not read).
    hyphen_gap: Fraction | None = None
    #: ``w:consecutiveHyphenLimit``: at most this many lines in a row end in an automatic
    #: hyphen (0: no limit).
    hyphen_limit: int = 0


@dataclass
class Line:
    start: int  # index of the first piece
    end: int  # one past the last
    #: Ended by a ``w:br`` / ``w:cr`` rather than by the budget or the paragraph.
    forced: bool = False
    #: Broken inside a word because the word alone did not fit.
    emergency: bool = False
    #: Broken at a soft hyphen, which is then drawn.
    hyphenated: bool = False


# -- geometry ---------------------------------------------------------------------------


def column_width_twips(section: Section, column: int = 0) -> int:
    """The width of one text column of ``section``, in twips: equal columns share the
    text width less the spaces between them, **truncated to a whole twip**
    (``make_columns_probe.py``, ``geometry``: the last column's right edge is its left
    edge plus that width -- 4,404 twips of 4,404.5, 1,632 of 1,632.8)."""
    page = section.page_size.width_twips - section.margins.left - section.margins.right - section.margins.gutter
    if section.columns:
        return section.columns[min(column, len(section.columns) - 1)][0]
    count = section.column_count
    if count <= 1:
        return page
    return (page - (count - 1) * section.column_space) // count


def is_multi_column(section: Section) -> bool:
    """Whether ``section`` lays its text out in more than one column."""
    return section.column_count > 1 or len(section.columns) > 1


def column_boxes(section: Section) -> tuple[tuple[int, int, int], ...]:
    """Every text column of ``section`` as ``(offset, width, space after)`` in twips, the
    offset from the text area's left edge (the left margin and gutter): each column
    starts where the one before ends plus the space after it -- ``w:space`` between
    equal columns, each ``w:col``'s own ``w:space`` between unequal ones, whose widths
    are taken as stated, whatever their sum (``make_columns_probe.py``, ``geometry``)."""
    if section.columns:
        widths = [(width, space) for width, space in section.columns]
    elif section.column_count > 1:
        width = column_width_twips(section)
        widths = [(width, section.column_space)] * section.column_count
    else:
        return ((0, column_width_twips(section), 0),)
    out = []
    offset = 0
    for width, space in widths:
        out.append((offset, width, space))
        offset += width + space
    return tuple(out)


def geometry(document: Document, paragraph: Paragraph, section: Section, pp=None, *,
             width: int | None = None) -> Geometry:
    """The paragraph's indents, tab stops and right edge, resolved; in a column ``width``
    twips wide where given (one of unequal columns), else the section's first."""
    pp = pp or resolve_paragraph(document, paragraph)
    left = pp.get("ind.left", 0) or 0
    right = pp.get("ind.right", 0) or 0
    first = pp.get("ind.firstLine", 0) or 0
    hanging = pp.get("ind.hanging", 0) or 0
    width = column_width_twips(section) if width is None else width
    tabs = tuple(TabStop(twips_to_units(stop.position_twips), stop.alignment, stop.leader)
                 for stop in pp.get("tabs", ()) or ()
                 if stop.alignment not in ("clear", "bar"))
    default = document.default_tab_stop or APPLICATION_DEFAULT_TAB
    return Geometry(
        first_start=twips_to_units(left + first - hanging),
        start=twips_to_units(left),
        right=twips_to_units(width - right),
        tabs=tabs,
        default_tab=twips_to_units(default) if default > 0 else twips_to_units(APPLICATION_DEFAULT_TAB),
        hanging_stop=twips_to_units(left) if hanging else None,
        squeeze=(document.compatibility_mode or 0) >= 15 and pp.get("jc") == "both",
        **hyphen_rules(document, pp),
    )


#: The hyphenation zone when ``settings.xml`` does not state one: 425 twips (0.75 cm),
#: measured on this machine's Word, whose ``Normal.dotm`` is metric -- an application
#: default like :data:`APPLICATION_DEFAULT_TAB`, not the schema's 360
#: (``make_hyphen_probe.py``, ``zone``: hyphenated from a word start exactly 425 twips
#: before the edge, not 2 units nearer, in no settings and mode 14).
APPLICATION_DEFAULT_ZONE = 425
#: Mode 15's own threshold in text that is not justified: the room after the line's text
#: (the edge less the end of the text before the word) must exceed 351.5 twips, whatever
#: ``w:hyphenationZone`` says (``zone``: 351.01 not, 352.00 hyphenated, at 11 pt in
#: Calibri and Times New Roman; the exploration found it between 350.6 and 353.2 at
#: other sizes and faces, a spread not explained).
MODE_15_HYPHEN_GAP = Fraction(3515, 10)


def hyphen_rules(document: Document, pp) -> dict:
    """The :class:`Geometry` fields of automatic hyphenation for a paragraph, or none
    when the document does not hyphenate or the paragraph suppresses it."""
    rules = document.hyphenation
    if rules is None or pp.get("suppressAutoHyphens"):
        return {}
    out: dict = {"hyphen_limit": rules.limit or 0}
    if (document.compatibility_mode or 0) >= 15:
        if pp.get("jc") not in ("both", "distribute"):
            out["hyphen_gap"] = MODE_15_HYPHEN_GAP * UNITS_PER_TWIP
    else:
        zone = APPLICATION_DEFAULT_ZONE if rules.zone is None else max(rules.zone, 0)
        out["hyphen_zone"] = twips_to_units(zone)
        out["hyphen_next"] = True
    return out


def tab_leader(x: int | Fraction, geometry: Geometry, first_line: bool) -> str | None:
    """The leader of the stop a tab at ``x`` jumps to (:func:`next_tab`): a custom stop's
    ``w:leader``; ``None`` for a default stop, the hanging indent's, or no leader."""
    stop, alignment, default = next_tab(x, geometry, first_line)
    if default:
        return None
    origin = geometry.tab_origin
    leader = next((s.leader for s in geometry.tabs if s.position + origin == stop and s.alignment == alignment), None)
    return None if leader in (None, "none") else leader


def squeezes(over, room: int, spaces: int, slack, spaces_before: int) -> bool:
    """Whether Word keeps a word ``over`` units past the right edge on a justified line in
    mode 15, compressing the line's ``spaces`` spaces, rather than send it down and
    stretch the line before it -- ``slack`` short of the edge, with ``spaces_before``
    spaces between its words.  Two conditions (``make_justify_probe.py``):

    * **no space shrinks by more than a quarter**: ``over`` is at most ``room``, the sum of
      the spaces' quarters (:func:`squeeze`), inclusive;
    * **the squeeze is less than half the stretch**: each space would shrink by at most
      half of what each would grow were the word sent down -- ``over / spaces <= slack /
      spaces_before / 2`` (the third round's 52 short words, the render probe's line
      that kept ``is`` and not ``the``: separated at 0.49 / 0.51); with no space before, the
      line could not stretch, and the first condition decides.
    """
    if over > room:
        return False
    return spaces_before == 0 or 2 * over * spaces_before <= slack * spaces


def squeeze(space: Piece) -> int:
    """How far a space lets a justified line in mode 15 run past the right edge: a quarter
    of its width, rounded half up to the layout unit (``make_justify_probe.py``: 1,214 of
    1,222 lines over two rounds -- δ at single units about the edge, 1-12 spaces, four
    faces at 11 and 14.5 pt, spaces of two sizes on one line, two spaces between words --
    the eight others Georgia's, one of them in a left-aligned control).  Word compresses
    the line's spaces to fit; below mode 15, and in text that is not justified, it does
    not (the controls: no line held a word past the edge)."""
    return math.floor(Fraction(space.width, 4) + Fraction(1, 2))


def next_tab(x: int | Fraction, geometry: Geometry, first_line: bool) -> tuple[int, str, bool]:
    """The stop a tab at ``x`` jumps to, its alignment, and whether it is a default stop.

    Custom stops first; the hanging indent is a stop on the first line; beyond the last
    custom stop, the default stops every ``default_tab`` from the column's left edge.
    """
    origin = geometry.tab_origin
    candidates = [(stop.position + origin, stop.alignment) for stop in geometry.tabs if stop.position + origin > x]
    if first_line and geometry.hanging_stop is not None and geometry.hanging_stop > x:
        candidates.append((geometry.hanging_stop, "left"))
    if candidates:
        return (*min(candidates), False)
    last = max([stop.position for stop in geometry.tabs], default=0) + origin
    base = max(x, last)
    step = geometry.default_tab
    return (math.floor((base - origin) / step) + 1) * step + origin, "left", True


# -- pieces -----------------------------------------------------------------------------


def _units(advance: tuple[int, int], half_points: int) -> int | Fraction:
    width, upm = advance
    value = Fraction(width * half_points * 2048, upm)
    return int(value) if value.denominator == 1 else value


def pieces(document: Document, paragraph: Paragraph, advances, metrics=None,
           label: list[Piece] | None = None) -> list[Piece]:
    """The paragraph's content as pieces, the list label's first.

    ``metrics(face, bold, italic)`` (a :class:`~docx2svg.vertical.FaceMetrics` or
    ``None``) is needed only for a superscript or subscript, whose glyphs are drawn at
    the face's ``OS/2`` script size.  Raises :class:`Unmeasurable`.
    """
    out: list[Piece] = list(label or [])
    source = 0
    cache: dict = {}
    hyphenating = document.hyphenation is not None and not resolve_paragraph(document, paragraph).get(
        "suppressAutoHyphens")
    #: Per run, for automatic hyphenation: its language, ``w:noProof``, and its hyphen.
    run_info: dict[int, tuple] = {}
    for run_index, run in enumerate(paragraph.runs):
        first_new = len(out)
        resolved = resolve_run(document, paragraph, run)
        marks: dict[int, list[str]] = {}
        for position, kind in run.breaks:
            marks.setdefault(position, []).append(kind)
        if resolved.get("vanish"):
            source += len(run.text)
            continue
        tracking = resolved.get("spacing") or 0
        scale = resolved.get("w") or 100
        kern_threshold = resolved.get("kern") or 0
        caps = bool(resolved.get("caps"))

        def glyph(char: str, kind: str = GLYPH, index: int = -1) -> Piece:
            # Keyed by the run, not by ``id(resolved)``: a run's resolved properties are
            # freed when the loop moves on, and a later run's may take the same id -- the
            # italic run of ``sample-simple.docx`` once took its "a" from the regular run
            # before it, depending on what else the process had allocated.
            key = (run_index, char)
            if key not in cache:
                fmt = character_format(resolved, char, document)
                drawn = char.upper() if caps else char
                if fmt.face is None:
                    raise Unmeasurable(f"no face for {char!r}")
                half_points = fmt.half_points
                if fmt.vertical_align:
                    face_metrics = metrics(fmt.face, fmt.bold, fmt.italic) if metrics else None
                    scripted = script_half_points(half_points, fmt.vertical_align, face_metrics) \
                        if face_metrics is not None else None
                    if scripted is None:
                        raise Unmeasurable(f"no script size for {fmt.face}")
                    half_points = scripted
                advance = advances.advance(fmt.face, fmt.bold, fmt.italic, drawn)
                if advance is None:
                    raise Unmeasurable(f"no advance for {drawn!r} in {fmt.face} b={fmt.bold} i={fmt.italic}")
                width = _units(advance, half_points)
                if run.note_number:
                    width = note_number_units(advance, half_points)
                if scale != 100:
                    width = Fraction(width) * scale / 100
                if tracking:
                    width += twips_to_units(tracking)
                kern_key = None
                if kern_threshold and fmt.half_points >= kern_threshold:
                    kern_key = (fmt.face, fmt.bold, fmt.italic, half_points)
                cache[key] = (drawn, width, kern_key, (fmt.face, fmt.bold, fmt.italic), half_points)
            drawn, width, kern_key, face, half_points = cache[key]
            return Piece(kind, drawn, width, kern_key, index, face=face)

        def mark(kind: str) -> None:
            if kind == "tab":
                out.append(Piece(TAB, "\t", 0, None, source))
            elif kind in ("br", "textWrapping", "page", "column", "cr"):
                # A page or column break ends its line as any break does; its character
                # says which, for the paginator (``\f``, ``\v``).
                char = {"page": "\f", "column": "\v"}.get(kind, "\n")
                out.append(Piece(BREAK, char, 0, None, source))
            elif kind == "noBreakHyphen":
                # Drawn as a hyphen, and as wide; but no line breaks after it.
                piece = glyph("-", GLYPH, source)
                piece.char = "\u2011"
                out.append(piece)
            elif kind == "softHyphen":
                hyphen = glyph("-", GLYPH, source)
                out.append(Piece(SOFT_HYPHEN, "\u00ad", 0, None, source, hyphen=hyphen.width))
            elif kind.startswith("drawing:"):
                cx, cy, left, top, right, bottom = (int(v) for v in kind.split(":")[1:])
                out.append(Piece(OBJECT, "\ufffc", twips_to_units(emu_twips(cx + left + right)), None, source,
                                 height=emu_twips(cy + top + bottom)))
            elif kind == "drawing":
                raise Unmeasurable("a drawing")
            elif kind == "separator":
                # A note separator's line: 2,880 twips (:mod:`docx2svg.notes`).
                out.append(Piece(GLYPH, NOTE_SEPARATOR, twips_to_units(2880), None, source))
            elif kind == "continuationSeparator":
                # Drawn to the column's right edge; it takes no width of its own.
                out.append(Piece(GLYPH, NOTE_CONTINUATION, 0, None, source))

        if hyphenating:
            try:
                hyphen_width = glyph("-").width
            except Unmeasurable:
                hyphen_width = None
            run_info[run_index] = (resolved.get("lang.val"), bool(resolved.get("noProof")), hyphen_width)
        for offset, char in enumerate(run.text):
            for kind in marks.get(offset, ()):
                mark(kind)
            if char in SPACES:
                piece = glyph(char, SPACE, source + offset)
            else:
                piece = glyph(char, GLYPH, source + offset)
            out.append(piece)
        for kind in marks.get(len(run.text), ()):
            mark(kind)
        for piece in out[first_new:]:
            piece.run = run_index
        source += len(run.text)
    if hyphenating:
        out = auto_hyphens(out, run_info, mode15=(document.compatibility_mode or 0) >= 15,
                           no_caps=document.hyphenation.no_caps)
    return out


#: Characters that join the parts of a compound: a word's parts either side are
#: hyphenated apart in mode 15, and only the first is below it.
_JOINERS = frozenset("-\u2011\u2013\u2014")
#: In French (``make_hyphen_probe.py``, ``french``): the typographic apostrophe joins as a
#: hyphen does -- ``aujourd\u2019hui`` breaks only in ``aujourd`` below mode 15, ``l\u2019organisation``
#: not at all, and in mode 15 ``l\u2019or-ga-ni-sa-tion`` -- though no line breaks at it; the
#: typewriter one is part of the word, which the patterns break (``jus-qu'\u00e0``,
#: ``l'or-ga-ni-sa-tion``), never at it.  Other languages' apostrophes are not measured.
_FRENCH_JOINERS = _JOINERS | {"\u2019"}
_FRENCH_LETTERS = frozenset("'")


def auto_hyphens(items: list[Piece], run_info: dict, *, mode15: bool, no_caps: bool) -> list[Piece]:
    """``items`` with an automatic :data:`SOFT_HYPHEN` (``auto``) at every point Word may
    break a word at (:mod:`docx2svg.hyphenate`; ROADMAP.md, 3.9).

    A word is a run of letters; its language is its first letter's run's ``w:lang``.
    Measured (``make_hyphen_probe.py``):

    * **the shortest fragment** is two letters before a break and two after it below mode
      15, three after it in mode 15;
    * **a word with a soft hyphen** keeps its own points in mode 15, the soft hyphen's
      among them; below mode 15 only those before its first soft hyphen;
    * **a compound** (``tele-communications``, a non-breaking hyphen too): in mode 15 each
      part is hyphenated as a word of its own; below it only the first part;
    * **``w:noProof``** ends the word where it starts; **``w:doNotHyphenateCaps``** leaves out a
      word whose letters are all capitals as drawn (typed so, or ``w:caps``).
    """
    from . import hyphenate

    right_min = 3 if mode15 else 2
    insert: dict[int, Piece] = {}  # piece index -> the automatic hyphen to put before it
    n = len(items)
    i = 0
    while i < n:
        # A token: what lies between two spaces, tabs or breaks.
        if items[i].kind not in (GLYPH, SOFT_HYPHEN) or items[i].source < 0:
            i += 1
            continue
        j = i
        while j < n and items[j].kind in (GLYPH, SOFT_HYPHEN) and items[j].source >= 0:
            j += 1
        _token_hyphens(items, i, j, run_info, mode15, no_caps, right_min, insert, hyphenate)
        i = j
    if not insert:
        return items
    out: list[Piece] = []
    for index, piece in enumerate(items):
        if index in insert:
            out.append(insert[index])
        out.append(piece)
    return out


def _token_hyphens(items, start, end, run_info, mode15, no_caps, right_min, insert, hyphenate) -> None:
    """The automatic hyphens of one token, ``items[start:end]`` (into ``insert``)."""
    first = next((k for k in range(start, end) if items[k].kind == GLYPH and items[k].char.isalpha()), None)
    language = run_info.get(items[first].run) if first is not None else None
    french = bool(language and language[0] and language[0].lower().startswith("fr"))
    joiners = _FRENCH_JOINERS if french else _JOINERS
    extra = _FRENCH_LETTERS if french else frozenset()

    def letter(k: int) -> bool:
        return items[k].char.isalpha() or items[k].char in extra

    segments: list[list[int]] = [[]]
    for k in range(start, end):
        if items[k].kind == GLYPH and items[k].char in joiners:
            segments.append([])
        else:
            segments[-1].append(k)
    for segment in segments if mode15 else segments[:1]:
        # Runs of letters; a soft hyphen inside one is part of it.
        words: list[list[int]] = [[]]
        for k in segment:
            if items[k].kind == SOFT_HYPHEN or letter(k):
                words[-1].append(k)
            else:
                words.append([k])  # a non-letter: kept to see digits beside a word
                words.append([])
        for index, word in enumerate(words):
            letters = [k for k in word if items[k].kind == GLYPH and letter(k)]
            if not letters:
                continue
            beside = [w[0] for w in (words[index - 1] if index else [], words[index + 1] if index + 1 < len(words)
                                     else []) if w]
            if any(items[k].char.isdigit() for k in beside):
                continue
            info = run_info.get(items[letters[0]].run)
            if info is None or info[2] is None:
                continue
            proof_end = None
            for count, k in enumerate(letters):
                if run_info.get(items[k].run, info)[1] and (count == 0 or not mode15):
                    # w:noProof: below mode 15 the word ends where it starts; in mode 15
                    # only the first letter's run decides.
                    proof_end = items[k].source
                    letters = letters[:count]
                    break
            text = "".join(items[k].char for k in letters)
            if not text or (no_caps and text.upper() == text and text.lower() != text):
                continue
            points = hyphenate.points(text, info[0], right_min=right_min, mode15=mode15)
            if not mode15:
                softs = [count for count, k in enumerate(letters) if items[k - 1].kind == SOFT_HYPHEN and count]
                if softs:
                    points = tuple(p for p in points if p < softs[0])
            for p in points:
                after = letters[p]
                if items[after - 1].kind == SOFT_HYPHEN:
                    continue  # the text breaks there itself
                width = run_info.get(items[letters[p - 1]].run, info)[2]
                if width is not None:
                    insert[after] = Piece(SOFT_HYPHEN, "\u00ad", 0, None, items[after].source, hyphen=width,
                                          run=items[letters[p - 1]].run, auto=True, proof_end=proof_end)


def label_pieces(document: Document, paragraph: Paragraph, text: str, advances, pp=None) -> list[Piece]:
    """The list label ``text`` (numbers already substituted) and what follows it."""
    pp = pp or resolve_paragraph(document, paragraph)
    num_id, ilvl = pp.get("numPr.numId"), pp.get("numPr.ilvl", 0) or 0
    level = document.numbering.get(num_id, {}).get(ilvl) if num_id else None
    if level is None:
        return []
    from .resolve.cascade import Resolved

    mark = resolve_mark(document, paragraph)
    values = dict(mark.values)
    values.update(level.run)
    resolved = Resolved(values, dict(mark.origins))
    out = []
    if level.picture_bullet is not None:
        pixels = picture_bullet_pixels(document, level)
        if pixels is None:
            raise Unmeasurable("a picture bullet whose picture is not read")
        fmt = character_format(resolved, " ", document)
        out.append(Piece(GLYPH, PICTURE_BULLET, picture_bullet_units(fmt.half_points, pixels), None, -1))
        text = ""
    for char in text:
        fmt = character_format(resolved, char, document)
        drawn = char.upper() if resolved.get("caps") else char
        advance = advances.advance(fmt.face, fmt.bold, fmt.italic, drawn) if fmt.face else None
        if advance is None:
            raise Unmeasurable(f"no advance for label {drawn!r} in {fmt.face}")
        out.append(Piece(GLYPH, drawn, _units(advance, fmt.half_points), None, -1,
                         face=(fmt.face, fmt.bold, fmt.italic)))
    shift = label_shift(level, out)
    if shift:
        out.insert(0, Piece(GLYPH, LABEL_SHIFT, -shift, None, -1))
    suffix = level.suffix or "tab"
    if suffix == "tab":
        out.append(Piece(TAB, "\t", 0, None, -1))
    elif suffix == "space":
        fmt = character_format(resolved, " ", document)
        face, bold, italic = (PICTURE_BULLET_SPACE_FACE, False, False) if level.picture_bullet is not None \
            else (fmt.face, fmt.bold, fmt.italic)
        advance = advances.advance(face, bold, italic, " ") if face else None
        if advance is None:
            raise Unmeasurable("no advance for the label's space")
        out.append(Piece(GLYPH, " ", _units(advance, fmt.half_points), None, -1))
    return out


def note_number_units(advance: tuple[int, int], half_points: int) -> Fraction:
    """A note number's glyph's advance, 1/4096 pt: its advance at the size it is drawn at
    (whole device px) rounded to whole device px -- the reference's and the note's own
    number alike (``make_endnote_probe.py``: Calibri ``i`` 7 px at 7 pt, 6 at 6.5, 15 at
    15.5, ``v`` 13 and 12, where the advance is 6.66, 6.20, 14.92, 13.14 and 12.23)."""
    ppem = math.floor(Fraction(half_points, 2) * Fraction(300, 72) + Fraction(1, 2))
    px = math.floor(Fraction(advance[0] * ppem, advance[1]) + Fraction(1, 2))
    return px * Fraction(72 * 4096, 300)


#: The characters a note separator's and a continuation separator's line stand in the
#: pieces as: the layout draws the line (:mod:`docx2svg.notes`).
NOTE_SEPARATOR = "\ue000"
NOTE_CONTINUATION = "\ue001"
#: The character of the piece that moves the pen back before a centred or right-aligned
#: label (:func:`label_shift`); it draws nothing.
LABEL_SHIFT = ""


def label_shift(level, label: list[Piece]) -> int | Fraction:
    """How far left of its start a label stands, by its level's ``w:lvlJc``: none (left),
    half its width (center) or all of it (right); what follows it (the tab, the space)
    starts where it ends.

    Measured by ``make_picture_bullet_probe.py`` (``jc``, and ``indent``): a text label and a
    picture bullet, at 11 and 24 pt, with a tab and with nothing after them -- the label
    centred on or ending at its start to the unit, and the tab or the text after it from
    its end.
    """
    width = sum((piece.width for piece in label), 0)
    alignment = getattr(level, "alignment", None)
    if alignment == "center":
        return Fraction(width, 2)
    if alignment in ("right", "end"):
        return width
    return 0


#: The character a picture bullet stands in the label's pieces as (a :data:`GLYPH` of the
#: label, source -1): the layout draws the picture where it is.
PICTURE_BULLET = "\ufffc"
#: The face the space after a picture bullet (``w:suff`` ``space``) is as wide as a space
#: of, at the label's size -- whatever the level names: 0.2778 em at 8, 11 and 24 pt, a
#: level in Symbol or in Calibri (``make_picture_bullet_probe.py``, ``suffix``).
PICTURE_BULLET_SPACE_FACE = "Arial"


def picture_bullet_pixels(document: Document, level) -> tuple[int, int] | None:
    """The pixels of the picture a level shows as its label, or ``None``."""
    bullet = document.picture_bullets.get(level.picture_bullet)
    return bullet.pixels if bullet is not None and bullet.pixels and all(bullet.pixels) else None


def picture_bullet_units(half_points: int, pixels: tuple[int, int]) -> Fraction:
    """How wide and tall Word draws a picture bullet, in 1/4096 pt: ``9 (s - 4) / h`` pt
    for a label of ``s`` half points and a picture ``h`` pixels high.

    Measured by ``make_picture_bullet_probe.py`` (no settings, modes 14 and 15 alike): the
    picture is a square -- whatever its own aspect or the size it states
    (``v:shape/@style``, ``wp:extent``: 4.5 to 72 pt, 18 x 9 and 9 x 18 all draw alike) --
    sized by the label's size alone (the level's ``w:sz`` over the mark's; not the face)
    and, inversely, by the picture's height in pixels (its ``pHYs`` density does not
    count): a 16 px picture is ``9/8 (pt - 2)`` pt at every size from 8 to 72 pt, 8, 12,
    24 and 48 px pictures twice, four thirds, half and a third of it, exactly.  Its
    advance is the same.  Pictures of 32, 64 and 128 px advance up to 18 units (0.02 px)
    more or less than this, and one wider than it is high about 460 units (0.47 px) less:
    not settled.
    """
    return Fraction(9 * (half_points - 4) * 4096, pixels[1])


# -- breaking ---------------------------------------------------------------------------


def _kern(previous: Piece | None, piece: Piece, advances) -> int | Fraction:
    if previous is None or piece.kern_key is None or previous.kern_key != piece.kern_key:
        return 0
    face, bold, italic, half_points = piece.kern_key
    value = advances.kern(face, bold, italic, previous.char, piece.char)
    if value is None:
        raise Unmeasurable(f"no kerning for {face}")
    if not value:
        return 0
    advance = advances.advance(face, bold, italic, piece.char)
    return _units((value, advance[1]), half_points)


def may_break(before: Piece, after: Piece) -> bool:
    """Whether a line may end between two adjacent pieces."""
    if after.kind == SPACE:
        return False
    if before.kind == SPACE:
        return True
    if after.kind == TAB:
        # A tab whose stop is past the edge goes down to the next line.
        return True
    if before.kind == SOFT_HYPHEN:
        return True
    if before.kind == GLYPH and before.char in BREAK_AFTER and before.source >= 0:
        return True
    return False


def _ends_on_its_own_line(char: str, columns: bool) -> bool:
    """Whether a break that ends the paragraph puts the paragraph's mark on a line of its
    own after it: a line break does (``make_break_end_probe.py``); so does a column break
    in a section of several columns -- the mark starts the next column
    (``make_columns_probe.py``, ``breaks`` and ``geometry``: the paragraph's mark drawn at
    the next column's top, the next paragraph a line below).  A page break's, and a
    column break's in one column, is the paginator's (:func:`paginate.flow`)."""
    return char == "\n" or (columns and char == "\v")


def break_pieces(items: list[Piece], geometry: Geometry, advances, *, first_line: bool = True,
                 columns: bool = False) -> list[Line]:
    """Greedy first fit: every line takes as much as fits (``x <= right``).

    ``first_line=False`` starts on a later line of the paragraph (its indent, no
    hanging stop): for breaking from the middle of a paragraph.  ``columns``: the
    paragraph is in a section of several columns (:func:`_ends_on_its_own_line`).
    """
    lines: list[Line] = []
    i, n = 0, len(items)
    hyphens = 0
    while i < n or not lines:
        line = break_line(items, i, geometry, advances, first=not lines and first_line, hyphens_before=hyphens)
        hyphens = hyphens + 1 if _auto_hyphenated(items, line) else 0
        lines.append(line)
        i = line.end
        if i >= n:
            if line.forced and _ends_on_its_own_line(items[n - 1].char, columns):
                lines.append(Line(n, n))
            break
    return lines


def break_rest(items: list[Piece], start: int, geometry: Geometry, advances, *, first_line: bool,
               columns: bool = False) -> list[Line]:
    """:func:`break_pieces` from piece ``start``: the lines the paragraph's rest breaks
    into when a line ends there -- the same lines greedy first fit gives, since a line
    depends on nothing before its first piece but whether it is the paragraph's first."""
    n = len(items)
    lines: list[Line] = []
    i = start
    hyphens = None
    while i < n or not lines:
        if i >= n:
            lines.append(Line(n, n))
            break
        line = break_line(items, i, geometry, advances, first=not lines and first_line, hyphens_before=hyphens)
        hyphens = (hyphens or 0) + 1 if _auto_hyphenated(items, line) else 0
        lines.append(line)
        i = line.end
        if i >= n:
            if line.forced and _ends_on_its_own_line(items[n - 1].char, columns):
                lines.append(Line(n, n))
            break
    return lines


def _auto_hyphenated(items: list[Piece], line: Line) -> bool:
    """Whether ``line`` ends in an automatic hyphen."""
    return line.hyphenated and line.end > line.start and items[line.end - 1].auto


def auto_hyphen_allowed(geometry: Geometry, word_x, text_x, next_fits: bool, hyphens_before: int) -> bool:
    """Whether a line may end at an automatic hyphen whose hyphen fits (ROADMAP.md, 3.9;
    ``make_hyphen_probe.py``):

    * **not past ``w:consecutiveHyphenLimit``** lines in a row ending in one;
    * **below mode 15, the word starts at least the hyphenation zone before the right
      edge** (``word_x``, inclusive, in every alignment), **and the letter after the
      point fits on the line too** -- Word breaks a word before the letter that runs past
      the edge, never at it;
    * **in mode 15**, where the zone is not read, the room after the line's text before
      the word (``text_x``) exceeds :data:`MODE_15_HYPHEN_GAP` unless the text is justified.
    """
    if geometry.hyphen_limit and hyphens_before >= geometry.hyphen_limit:
        return False
    if geometry.hyphen_zone is not None and geometry.right - word_x < geometry.hyphen_zone:
        return False
    if geometry.hyphen_next and not next_fits:
        return False
    if geometry.hyphen_gap is not None and geometry.right - text_x <= geometry.hyphen_gap:
        return False
    return True


def break_line(items: list[Piece], i: int, geometry: Geometry, advances, *, first: bool,
               hyphens_before: int | None = None) -> Line:
    """One line of greedy first fit from piece ``i``, at ``geometry``'s start (its first
    line's, where ``first``) and right edge.  ``hyphens_before``: how many lines in a row
    before it end in an automatic hyphen (``None``: one if the line before it does, which
    is all a caller that breaks from the middle of a paragraph knows)."""
    n = len(items)
    x: int | Fraction = geometry.first_start if first else geometry.start
    start = i
    if hyphens_before is None:
        hyphens_before = int(i > 0 and items[i - 1].auto)
    #: Every place the line may end so far, with where the word it is in starts and
    #: where the text before that word ends (for automatic hyphens).
    breaks: list[tuple[int, int | Fraction, int | Fraction]] = []
    word_x: int | Fraction = x
    text_x: int | Fraction = x
    text_end: int | Fraction = x
    line = None
    #: A right or centre stop past the edge carries the text after it there.
    limit = geometry.right
    #: How far past the edge the spaces on the line since its last tab let it run;
    #: how many spaces there are; and, at the last run of spaces, where the text
    #: before it ended and how many spaces were before that (:func:`squeezes`).
    room = spaces = 0
    before_x: int | Fraction = x
    before_spaces = 0
    j = i
    while j < n:
        piece = items[j]
        if j > start and may_break(items[j - 1], piece):
            before = items[j - 1]
            if before.kind != SOFT_HYPHEN or x + before.hyphen <= geometry.right or (
                    before.auto and geometry.squeeze
                    and squeezes(x + before.hyphen - geometry.right, room, spaces, limit - before_x, before_spaces)):
                breaks.append((j - 1, word_x, text_x))
            if before.kind != SOFT_HYPHEN:
                word_x, text_x = x, text_end
        if piece.kind == BREAK:
            line = Line(start, j + 1, forced=True)
            break
        if piece.kind == TAB:
            stop, alignment, default = next_tab(x, geometry, first)
            if alignment in ("right", "center", "decimal"):
                # The text after a right or centre stop moves left, as far as the
                # tab's own position, rather than wrap: for breaking it starts where
                # the tab does.  A stop past the edge carries the text out to it (a
                # centred text, as far again past it as the tab is before it).
                x_new = x
                if stop <= geometry.right:
                    limit = geometry.right
                else:
                    limit = stop if alignment != "center" else 2 * stop - x
            elif stop > geometry.right and not default:
                # A custom left stop past the edge: nothing after it wraps.
                x_new = stop
                limit = math.inf
            else:
                x_new = stop
                limit = geometry.right
        else:
            # A kern pair adjusts its left glyph's advance: a glyph at the end of a
            # line is charged its pair with what follows -- the space after it too.
            x_new = x + piece.width + (_kern(piece, items[j + 1], advances) if j + 1 < n else 0)
        if piece.kind == TAB:
            room = spaces = before_spaces = 0
            before_x = x_new
        elif piece.kind == SPACE and geometry.squeeze:
            if j == start or items[j - 1].kind != SPACE:
                before_x, before_spaces = x, spaces
            room += squeeze(piece)
            spaces += 1
        if piece.kind not in (SPACE, SOFT_HYPHEN) and x_new > limit and j > start and not (
                geometry.squeeze and squeezes(x_new - limit, room, spaces, limit - before_x, before_spaces)):
            chosen = None
            for at, at_word, at_text in reversed(breaks):
                if items[at].auto and (not auto_hyphen_allowed(geometry, at_word, at_text, at + 1 < j, hyphens_before)
                                       or (items[at].proof_end is not None and piece.source >= items[at].proof_end)):
                    continue
                chosen = at
                break
            if chosen is not None:
                line = Line(start, chosen + 1, hyphenated=items[chosen].kind == SOFT_HYPHEN)
            else:
                line = Line(start, j, emergency=True)
            break
        if piece.kind not in (SPACE, SOFT_HYPHEN, TAB, BREAK):
            text_end = x_new
        elif piece.kind == TAB:
            text_end = x_new
        x = x_new
        j += 1
    if line is None:
        line = Line(start, n)
    return line


# -- numbering --------------------------------------------------------------------------


def _roman(value: int) -> str:
    out = ""
    for number, letters in ((1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
                            (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")):
        while value >= number:
            out += letters
            value -= number
    return out


def format_number(value: int, fmt: str | None) -> str:
    if fmt in ("lowerLetter", "upperLetter"):
        letter = chr(ord("a") + (value - 1) % 26) * ((value - 1) // 26 + 1)
        return letter.upper() if fmt == "upperLetter" else letter
    if fmt in ("lowerRoman", "upperRoman"):
        roman = _roman(value)
        return roman.upper() if fmt == "upperRoman" else roman
    if fmt in ("bullet", "none"):
        return ""
    return str(value)


def _start(level) -> int:
    return 1 if level is None or level.start is None else level.start


class ListItem(NamedTuple):
    """One list item as :meth:`ListCounters.item` counts it: the label Word draws for it
    (``"2.1."``; ``""`` for a bullet or ``none`` level's number) and the item's own running
    number at its level.  Both are ``None`` where the instance has no such level, or the
    level no ``w:lvlText``."""

    label: str | None
    number: int | None


@dataclass
class ListCounters:
    """The running numbers of a story's lists, for label text, counted as Word counts
    them (``make_numbering_probe.py``; ROADMAP.md, "Numbering and lists"):

    * instances (``w:num``) over one abstract definition share its counts
      (:attr:`Document.numbering_lists`);
    * a ``w:startOverride`` restarts the level at the instance's first item there
      (:attr:`Document.numbering_restarts`), and is its start whenever it restarts;
    * an item sets every shallower level that has no count since it restarted to its
      start -- as if it had an item there -- and restarts each deeper level its
      ``w:lvlRestart`` allows (unstated: every deeper one; ``0``: none);
    * a ``w:isLgl`` level's label shows every number in decimal.

    **Public entry point**: :meth:`item` counts an item from the numbering part alone
    (:func:`docx2svg.parse.styles.parse_numbering`), its instance (``numId``) and level
    (``ilvl``) -- no document or paragraph of docx2svg's is needed::

        numbering = parse_numbering(numbering_xml, parse_styles(styles_xml))
        counters = ListCounters()            # one per story, asked in document order
        counters.item(numbering, 3, 0)       # ListItem(label='1.', number=1)
        counters.item(numbering, 3, 1)       # ListItem(label='1.1.', number=1)

    :meth:`label` is the layout's: the same count for a paragraph of a parsed document.
    One counter counts one story (the body, a header, the text boxes...), its items asked
    in document order; a story's lists never continue another's.
    """

    counts: dict = field(default_factory=dict)
    #: ``(numId, ilvl)`` that have had an item.
    used: set = field(default_factory=set)

    def item(self, numbering, num_id: int, ilvl: int = 0) -> ListItem:
        """Count the next item of instance ``num_id`` at level ``ilvl`` and return its
        label and number.  ``numbering`` is a :class:`docx2svg.parse.styles.Numbering`
        (``levels``, ``lists``, ``restarts``) or anything with those three attributes."""
        levels = numbering.levels.get(num_id, {}) if num_id else {}
        key = numbering.lists.get(num_id, ("num", num_id))
        label = self._count(levels, key, numbering.restarts, num_id, ilvl or 0)
        if label is None:
            return ListItem(None, None)
        return ListItem(label, self.counts[key][ilvl or 0])

    def label(self, document: Document, paragraph: Paragraph, pp=None) -> str | None:
        if paragraph.path in document.list_labels:
            return document.list_labels[paragraph.path]
        pp = pp or resolve_paragraph(document, paragraph)
        num_id, ilvl = pp.get("numPr.numId"), pp.get("numPr.ilvl", 0) or 0
        levels = document.numbering.get(num_id, {}) if num_id else {}
        return self._count(levels, document.numbering_lists.get(num_id, ("num", num_id)),
                           document.numbering_restarts, num_id, ilvl)

    def _count(self, levels: dict, key, restarts, num_id, ilvl: int) -> str | None:
        level = levels.get(ilvl)
        if level is None or level.text is None:
            return None
        counts = self.counts.setdefault(key, {})
        first = (num_id, ilvl) not in self.used
        self.used.add((num_id, ilvl))
        if first and (num_id, ilvl) in restarts:
            counts[ilvl] = _start(level)
        else:
            counts[ilvl] = counts.get(ilvl, _start(level) - 1) + 1
        for shallower in range(ilvl):
            counts.setdefault(shallower, _start(levels.get(shallower)))
        for deeper in [k for k in counts if k > ilvl]:
            restart = getattr(levels.get(deeper), "restart", None)
            if ilvl < (deeper if restart is None else restart):
                del counts[deeper]
        text = level.text
        for number in range(9):
            token = f"%{number + 1}"
            if token in text:
                other = levels.get(number)
                value = counts.get(number, _start(other))
                fmt = other.format if other else None
                if level.legal and fmt not in ("bullet", "none"):
                    fmt = "decimal"
                text = text.replace(token, format_number(value, fmt))
        return text


def text_box_labels(document: Document) -> dict:
    """Every list label in the story's text boxes, counted as Word counts them: the text
    boxes are one story of their own -- their lists continue from box to box, in the
    order of their anchors, and neither continue the body's nor are continued by it
    (``make_numbering_probe.py``, ``story``).  Keyed by the paths the paragraphs have
    once laid out in their box (``layout._Placer._text_box``): a top-level block's under
    its anchor paragraph's path."""
    from .model import Table

    counters = ListCounters()
    out: dict = {}

    def blocks(items, prefix: str | None) -> None:
        for item in items:
            if isinstance(item, Table):
                for row in item.rows:
                    for cell in row:
                        blocks(cell, None)
                continue
            path = f"{prefix}/{item.path}" if prefix is not None else item.path
            out[path] = counters.label(document, item)

    def graphic(found, anchor_path: str) -> None:
        if found is None or isinstance(found, str):
            return
        if found.text:
            blocks(found.text, anchor_path)
        for inner in found.children:
            graphic(inner, anchor_path)

    def paragraphs(items) -> None:
        for item in items:
            if isinstance(item, Table):
                for row in item.rows:
                    for cell in row:
                        paragraphs(cell)
                continue
            for run in item.runs:
                for found in run.drawings:
                    graphic(found, item.path)
                for anchor in run.anchors:
                    graphic(anchor.graphic, item.path)

    paragraphs(document.body)
    return out


def break_paragraph(document: Document, paragraph: Paragraph, section: Section, advances,
                    metrics=None, label: str | None = None) -> tuple[list[Piece], list[Line]]:
    """The paragraph's pieces and its lines.  Raises :class:`Unmeasurable`."""
    pp = resolve_paragraph(document, paragraph)
    label_items = label_pieces(document, paragraph, label, advances, pp) if label is not None else []
    items = pieces(document, paragraph, advances, metrics, label_items)
    return items, break_pieces(items, geometry(document, paragraph, section, pp), advances,
                               columns=is_multi_column(section))

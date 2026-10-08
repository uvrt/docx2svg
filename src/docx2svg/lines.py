"""A paragraph's lines as the page stack needs them: which characters, and how tall.

The rules are :mod:`docx2svg.linebreak`'s (which characters go on which line) and
:mod:`docx2svg.vertical`'s (how tall a line of given characters is); this module joins
them per paragraph, so that the baseline scorer (``tools/baselines.py``) and the
paginator (:mod:`docx2svg.paginate`) build a line from the same code.  Moved here from
``tools/baselines.py`` unchanged when pagination needed it without the oracle.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from fractions import Fraction

from . import linebreak, vertical
from .model import Document, Paragraph
from .resolve import character_format, resolve_mark, resolve_run
from .resolve.cascade import Resolved
from .vertical import (
    PX_PER_PT, LineBox, line_extent, mixed_line_pitch, natural_height_px, tallest_natural,
)


def label_format(document: Document, paragraph: Paragraph, pp: Resolved, mark: Resolved):
    """The list label: the mark's properties with the level's ``w:rPr`` on top."""
    num_id, ilvl = pp.get("numPr.numId"), pp.get("numPr.ilvl", 0)
    level = document.numbering.get(num_id, {}).get(ilvl or 0) if num_id else None
    if level is None or not level.text:
        return []
    values = dict(mark.values)
    origins = dict(mark.origins)
    values.update(level.run)
    text = re.sub(r"%\d", "1", level.text)
    label = Resolved(values, origins)
    if level.picture_bullet is not None:
        pixels = linebreak.picture_bullet_pixels(document, level)
        if pixels is None:
            return []
        fmt = character_format(label, " ", document)
        return [PictureLabel(fmt.half_points, fmt.position,
                             linebreak.picture_bullet_units(fmt.half_points, pixels) * vertical.LAYOUT_UNIT_PX)]
    return [character_format(label, char, document) for char in text]


@dataclass(frozen=True)
class PictureLabel:
    """A picture bullet as the line height sees it: it stands on the baseline and reaches
    ``picture_px`` above it (moved by the level's ``w:position``), with no descent
    (:func:`vertical.line_extent`)."""

    half_points: int
    position: int
    picture_px: Fraction
    face: None = None
    bold: bool = False
    italic: bool = False
    border: None = None


def visible_chars(document: Document, paragraph: Paragraph) -> list[tuple[Resolved, str]]:
    """``[(resolved run properties, character)]`` for every character not hidden."""
    return [(resolved, char) for run in paragraph.runs
            for resolved in [resolve_run(document, paragraph, run)] if not resolved.get("vanish")
            for char in run.text]


def model_line_shares(document: Document, paragraph: Paragraph, section, advances, metrics=None,
                      label: str | None = None) -> list[list[tuple]] | None:
    """Which of the paragraph's characters each of its lines holds, by the model's line
    breaker: ``[[(resolved run properties, character)]]`` per line; ``None`` when the
    breaker cannot measure the paragraph (:class:`linebreak.Unmeasurable`)."""
    return shares_of(document, paragraph, *broken(document, paragraph, section, advances, metrics, label))


def broken(document: Document, paragraph: Paragraph, section, advances, metrics=None, label=None):
    """``linebreak.break_paragraph``, or ``(None, None)`` when it cannot measure."""
    try:
        return linebreak.break_paragraph(document, paragraph, section, advances, metrics, label)
    except linebreak.Unmeasurable:
        return None, None


def shares_of(document: Document, paragraph: Paragraph, items, lines) -> list[list[tuple]] | None:
    if items is None:
        return None
    line_of: dict[int, int] = {}
    for number, line in enumerate(lines):
        for piece in items[line.start:line.end]:
            # A tab or a break carries its run's first index, not a character's: only
            # characters say which line a character is on.
            if piece.source >= 0 and piece.kind not in (linebreak.TAB, linebreak.BREAK, linebreak.OBJECT):
                line_of.setdefault(piece.source, number)
    out: list[list[tuple]] = [[] for _ in lines]
    index = current = 0
    for run in paragraph.runs:
        resolved = resolve_run(document, paragraph, run)
        if not resolved.get("vanish"):
            for offset, char in enumerate(run.text):
                current = line_of.get(index + offset, current)
                out[current].append((resolved, char))
        index += len(run.text)
    return out


def objects_of(items, lines) -> list[list[int]]:
    """Per line, the heights (twips) of the inline pictures on it."""
    if items is None:
        return []
    return [[piece.height for piece in items[line.start:line.end] if piece.kind == linebreak.OBJECT]
            for line in lines]


def object_runs_of(items, lines) -> list[list[int]]:
    """Per line, the indices of the runs holding its inline pictures."""
    if items is None:
        return []
    return [sorted({piece.run for piece in items[line.start:line.end] if piece.kind == linebreak.OBJECT})
            for line in lines]


def line_items(document: Document, paragraph: Paragraph, pp: Resolved, chars=None, *,
               first_line: bool = True):
    """Resolved character formats that set the line height: (text and mark, label).

    Which of them take part is the model's: no space, and the paragraph mark only on a
    line holding nothing but spaces (``vertical.line_items``); the list label on the
    paragraph's first line only (``vertical.label_items``).

    ``chars`` -- ``[(resolved run properties, character)]``, one line's share -- limits
    the text to what that line holds; by default the whole paragraph's.  ``first_line``
    says whether it is the paragraph's first.
    """
    if chars is None:
        chars = visible_chars(document, paragraph)
    mark = resolve_mark(document, paragraph)
    items = vertical.line_items([(character_format(resolved, char, document), char) for resolved, char in chars],
                                character_format(mark, " ", document))
    return items, vertical.label_items(label_format(document, paragraph, pp, mark), first_line=first_line)


def height_key(item):
    """Items that differ only in the character are one item as far as height goes."""
    return (item.face, item.bold, item.italic, item.half_points, item.position, item.border,
            getattr(item, "picture_px", None))


@dataclass(frozen=True)
class LineHeight:
    """One line's vertical facts, in exact device px: its pitch and its text's extent.

    ``mixed`` says the line holds more than one face or size (or a label, a raised run,
    a border) and its extent is the separate maxima of its items'.
    """

    pitch: Fraction
    text_above: Fraction
    text_below: Fraction
    natural: Fraction
    rule: str
    line: int
    mixed: bool = False
    #: A line holding inline pictures: how far below the line's top (below its space
    #: before) a picture's bottom, its effect extent included, stands
    #: (:func:`object_line_height`); ``None`` for a line of text.
    object_drop: Fraction | None = None

    def box(self, space_before: Fraction = Fraction(0), space_after: Fraction = Fraction(0)) -> LineBox:
        return LineBox(pitch=self.pitch, text_above=self.text_above, text_below=self.text_below,
                       natural=self.natural, rule=self.rule, line=self.line,
                       space_before=space_before, space_after=space_after)


def object_line_height(document: Document, paragraph: Paragraph, pp: Resolved, chars, metrics, objects,
                       *, first_line: bool, runs: list[int] | None = None) -> LineHeight | None:
    """A line holding inline pictures (``objects``: their heights in twips).

    Measured by ``make_picture_probe.py`` on pictures alone in their paragraph: **a
    picture stands on the baseline and nothing is below it** -- the line is the
    picture's height, whole twips (:func:`linebreak.emu_twips`), and the paragraph mark
    adds no descent; in modes below 15 and unstated the line is at least the mark's
    natural height (a 20 pt mark over a 98 px picture: 102 px; mode 15, 98).  An ``auto``
    multiple adds its extra over the mark's natural height (``auto`` 259, 276, 360 add
    0.079, 0.15, 0.5 of Calibri 11's 55.9 px); ``exact`` is the line; ``atLeast`` at least
    the line.  Text beside a picture is not measured: its extent is maxed with the
    picture's, as items on a line are.

    Where the pictures are drawn in the line is ``object_drop``, measured by
    ``make_picture_place_probe.py``.  **In mode 15 the floor is not the mark's but the
    natural height of the runs holding the pictures** (``runs``, their indices;
    ``make_picture_run_probe.py``: pictures of 33, 49 and 66 px in runs of 8-30 pt under
    marks of 11 and 20 pt, 48 / 48 lines; below mode 15 the run's size does not count).
    """
    from .vertical import FaceMetrics, natural_height_px, quantise, twips_to_px

    picture = twips_to_px(max(objects))
    mark = resolve_mark(document, paragraph)
    mark_item = character_format(mark, " ", document)
    mark_face = metrics(mark_item.face, mark_item.bold, mark_item.italic) if mark_item.face else None
    if mark_face is None:
        return None
    mark_natural = natural_height_px(mark_face, mark_item.half_points)
    text = line_height(document, paragraph, pp, chars, metrics, first_line=first_line) if any(
        not char.isspace() for _, char in chars) else None
    above = max(picture, text.text_above) if text else picture
    below = text.text_below if text else Fraction(0)
    natural = above + below
    if text is None and (document.compatibility_mode is None or document.compatibility_mode < 15):
        natural = max(natural, mark_natural)
    elif text is None and runs:
        for index in runs:
            item = character_format(resolve_run(document, paragraph, paragraph.runs[index]), " ", document)
            face = metrics(item.face, item.bold, item.italic) if item.face else None
            if face is not None:
                natural = max(natural, natural_height_px(face, item.half_points))
        above = natural
    rule = pp.get("spacing.lineRule", "auto")
    value = pp.get("spacing.line", 240)
    if rule == "exact":
        pitch = twips_to_px(value)
    elif rule == "atLeast":
        pitch = max(natural, twips_to_px(value))
    elif value != 240:
        pitch = quantise(natural + (text.natural if text else mark_natural) * (Fraction(value, 240) - 1))
    else:
        pitch = natural
    # Where the pictures stand (make_picture_place_probe.py, every picture to 0.001 px):
    # on the text's baseline -- under a taller mark, at the bottom of the mark's height,
    # and an auto multiple's extra below them; an atLeast line's extra above them; and in
    # an exact line four fifths of the way down, whatever the picture's height.
    descent = text.text_below if text else Fraction(0)
    if rule == "exact":
        drop = pitch * Fraction(4, 5)
    elif rule == "atLeast":
        drop = pitch - descent
    else:
        drop = natural - descent
    return LineHeight(pitch=pitch, text_above=above, text_below=natural - above, natural=natural, rule=rule,
                      line=value, mixed=True, object_drop=drop)


def line_height(document: Document, paragraph: Paragraph, pp: Resolved, chars, metrics, *,
                first_line: bool, cache: dict | None = None) -> LineHeight | None:
    """The height of one line holding ``chars``; ``None`` when no face on it has metrics.

    ``cache`` (a dict the caller keeps for one ``metrics``) remembers heights by what
    decides them: the distinct items on the line, the label's, and the line rule.
    """
    text_items, label_items = line_items(document, paragraph, pp, chars, first_line=first_line)
    text_items = list({height_key(item): item for item in text_items}.values())
    label_items = list({height_key(item): item for item in label_items}.values())
    if cache is not None:
        key = (tuple(height_key(item) for item in text_items), tuple(height_key(item) for item in label_items),
               pp.get("spacing.lineRule", "auto"), pp.get("spacing.line", 240))
        if key not in cache:
            cache[key] = _line_height(text_items, label_items, pp, metrics)
        return cache[key]
    return _line_height(text_items, label_items, pp, metrics)


def _line_height(text_items, label_items, pp: Resolved, metrics) -> LineHeight | None:
    combined = line_extent(text_items, metrics, label_items)
    if combined is None:
        return None
    text_only = line_extent(text_items, metrics) if label_items else None
    text_natural = tallest_natural(text_items, metrics)
    rule = pp.get("spacing.lineRule", "auto")
    value = pp.get("spacing.line", 240)
    face, half_points, mixed = combined
    em = Fraction(half_points, 2) * PX_PER_PT / face.units_per_em
    return LineHeight(
        pitch=mixed_line_pitch(combined, text_only, rule, value, text_natural),
        text_above=(face.ascent + face.line_gap) * em,
        text_below=face.descent * em,
        natural=natural_height_px(face, half_points),
        rule=rule,
        line=value,
        mixed=mixed,
    )

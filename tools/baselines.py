"""Predict every baseline of a document from its *resolved* properties, and score it.

This is where the cascade meets Phase 2's vertical model.  Nothing a paragraph states is
read directly: its spacing, line rule, face and size all come out of
``docx2svg.resolve``, and the faces' four metric integers come from a ``metrics``
callable (``face_metrics.metrics`` against installed fonts, or recorded integers in a
test).  So a document whose runs state nothing reproduces only if the cascade is right.

What is taken from the oracle rather than predicted, and why:

* **which paragraph each drawn line belongs to and the page it is on** -- pagination is
  Phase 4.  Lines are matched to paragraphs by their text (``match_lines``).
* **which of the paragraph's characters each line holds**, unless ``predict`` is given
  advance widths: then the model's line breaker decides (``model_line_shares``;
  ROADMAP.md, "Phase 3 -- measured"), and Word's lines only where the breaker cannot
  measure a paragraph or breaks it into another number of lines.  Otherwise each line's
  text is taken from the front of its paragraph's (``line_shares``).
* nothing else.  In particular every baseline's *position* is predicted.

The stack, per page (``top`` in exact device px, Word's 1/4096 pt grid):

1. a page starts at the section's top margin;
2. a paragraph adds the gap after the previous one -- its space after and this one's
   space before collapsed to the larger, in twips (``vertical.paragraph_gap_px``), each
   zeroed by ``w:contextualSpacing`` against a paragraph of the same style -- then each
   of its lines at the line rule's pitch, then any bottom border (``vertical.border_px``);
3. every line rounds its baseline in its **line box** (``vertical.baseline_in_box``):
   the first line's box carries what is left of the space before, above the text; the
   last line's carries the bottom border and the space after, below it;
4. a line's extent is that of everything on it -- every character Word put on that line
   but spaces, the list label on the paragraph's first line, and the paragraph mark on a
   line of nothing but spaces (``vertical.line_items``, ``vertical.label_items``), each
   resolved -- with the largest ascent and the largest descent
   taken separately, a label's descent left out (``vertical.line_extent``), and its
   ``auto`` multiple extra over the tallest text item (``vertical.mixed_line_pitch``).

Out of scope, and reported as such rather than scored: lines after a table or an inline
drawing on the same page (neither has measured geometry), and paragraphs inside tables.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from fractions import Fraction

from docx2svg.model import Document, Paragraph, Table
from docx2svg.resolve import (
    autospaced, character_format, resolve_mark, resolve_paragraph, resolve_run, script_half_points,
)
from docx2svg import linebreak, vertical
from docx2svg.lines import height_key as _height_key, label_format as _label_format  # noqa: F401
from docx2svg.lines import line_items, model_line_shares  # noqa: F401
from docx2svg.resolve.cascade import Resolved
from docx2svg.vertical import (
    PX_PER_PT, LineBox, autospace_kept, baseline_in_box, border_px, keeps_space_before_at_page_top,
    line_extent, mixed_line_pitch, natural_height_px, page_top_gap_px, paragraph_gap_px,
    tallest_natural, twips_to_px,
)


@dataclass
class Block:
    paragraph: Paragraph
    #: ``None`` when in scope; otherwise why its lines are not scored.
    excluded: str | None
    section: int


@dataclass
class LineResult:
    page: int
    block: int
    text: str
    observed: int
    predicted: int | None = None
    #: ``exact``, ``miss``, or the reason it is out of scope.
    status: str = ""
    notes: list[str] = field(default_factory=list)


def paragraph_features(package: bytes) -> list[str | None]:
    """For every paragraph in :func:`blocks` order, why the parsed model cannot stand for
    it (``None`` when nothing is wrong).  Now ``docx2svg.scope.paragraph_features``, which
    the renderer needs at run time; this name is kept for the tools that call it."""
    from docx2svg.scope import paragraph_features as features

    return features(package)


def blocks(document: Document, package: bytes | None = None) -> list[Block]:
    """Every paragraph in document order, cells included, with its section.

    With the ``.docx`` bytes, a paragraph :func:`paragraph_features` flags is out of scope.
    """
    out: list[Block] = []
    top_level = 0

    def section_of(index: int) -> int:
        for number, section in enumerate(document.sections):
            if section.first_paragraph <= index < section.last_paragraph:
                return number
        return len(document.sections) - 1

    def walk(items, excluded, section):
        for item in items:
            if isinstance(item, Table):
                for row in item.rows:
                    for cell in row:
                        walk(cell, "table", section)
            else:
                reason = excluded
                if reason is None and any(kind.startswith("drawing") for run in item.runs for _, kind in run.breaks):
                    reason = "drawing"
                # A floating drawing text does not wrap around takes no room in its line: in
                # scope (ROADMAP.md, "Floating drawings -- measured"); one text wraps around is not.
                if reason is None and any(anchor.moves_text for run in item.runs for anchor in run.anchors):
                    reason = "drawing"
                out.append(Block(item, reason, section))

    for item in document.body:
        if isinstance(item, Table):
            walk([item], None, section_of(top_level))
        else:
            walk([item], None, section_of(top_level))
            top_level += 1
    if package is not None:
        features = paragraph_features(package)
        if len(features) == len(out):
            for block, reason in zip(out, features):
                if reason and block.excluded is None:
                    block.excluded = reason
    return out


def _norm(text: str) -> str:
    return re.sub(r"\s+", "", text).lower()


def _visible_text(document: Document, paragraph: Paragraph) -> str:
    parts = []
    for run in paragraph.runs:
        resolved = resolve_run(document, paragraph, run)
        if not resolved.get("vanish"):
            parts.append(run.text)
    return "".join(parts)


#: The block index :func:`match_lines` gives a drawn line no body paragraph accounts for.
UNMATCHED = -1


def match_lines(document: Document, all_blocks: list[Block], drawn) -> list[tuple[int, object]]:
    """Assign each drawn line (``quartz_pdf.Line``) to a block by text: [(block, line)]."""
    out = []
    j = 0
    skip_until = -1
    texts = [_norm(_visible_text(document, block.paragraph)) for block in all_blocks]

    def starts(line_text: str, remaining: str) -> int | None:
        return next((k for k in range(len(line_text))
                     if line_text[k:] and remaining.startswith(line_text[k:])), None)

    for index, block in enumerate(all_blocks):
        if index <= skip_until:
            continue
        remaining = texts[index]
        if block.excluded == "table" and remaining and j < len(drawn) and (
                starts(_norm(drawn[j].text), remaining) is None):
            # A table row is drawn as one line holding every cell's text side by side, so
            # a cell whose text wraps cannot be matched line by line.  The table is out of
            # scope anyway: give it every drawn line up to the first that begins the next
            # paragraph outside it with text.
            after = next((k for k in range(index + 1, len(all_blocks))
                          if all_blocks[k].excluded != "table" and texts[k]), None)
            stop = len(drawn)
            if after is not None:
                stop = next((k for k in range(j, len(drawn))
                             if _begins(_norm(drawn[k].text), texts[after], starts)),
                            len(drawn))
            while j < stop:
                out.append((index, drawn[j]))
                j += 1
            skip_until = (after - 1) if after is not None else len(all_blocks)
            continue
        if not remaining:
            # An empty paragraph draws only its mark.  A page break alone in a paragraph
            # leaves the mark at the top of the next page.
            if j < len(drawn) and not _norm(drawn[j].text):
                out.append((index, drawn[j]))
                j += 1
            continue
        start, taken = j, len(out)
        if j < len(drawn) and starts(_norm(drawn[j].text), remaining) is None:
            # Something drawn here is no body paragraph's text (a footnote at the foot of
            # the page, a field result, a frame), or this paragraph is drawn in a way the
            # text does not match (a footnote reference mark).  Look ahead for the line
            # that begins this paragraph; what is skipped goes to UNMATCHED.  If none
            # does, this paragraph is left unmatched and the next one tries.
            found = next((k for k in range(j + 1, min(j + 80, len(drawn)))
                          if _begins(_norm(drawn[k].text), remaining, starts)), None)
            if found is None:
                continue
            while j < found:
                out.append((UNMATCHED, drawn[j]))
                j += 1
        while remaining and j < len(drawn):
            text = _norm(drawn[j].text)
            # A list label precedes the text on the first line; skip what does not match.
            cut = next((k for k in range(len(text)) if text[k:] and remaining.startswith(text[k:])), None)
            if cut is None:
                break
            out.append((index, drawn[j]))
            remaining = remaining[len(text) - cut:]
            j += 1
        if remaining and j < len(drawn):
            # Only part of the paragraph was found: a false start (a heading whose text
            # begins a table-of-contents entry) or text drawn differently from what the
            # paragraph holds.  Take nothing, not even the lines skipped to get here;
            # the next paragraph resynchronises.
            del out[taken:]
            j = start
    return out


def _begins(line_text: str, remaining: str, starts) -> bool:
    """Whether a drawn line is the first line of a paragraph: all of it, or all but a
    list label (at most ten characters), begins the paragraph's text (at least five
    characters of it, or all of a shorter one)."""
    cut = starts(line_text, remaining) if line_text else None
    return cut is not None and (cut == 0 or (cut <= 10 and len(line_text) - cut >= min(5, len(remaining))))


def line_shares(document: Document, paragraph: Paragraph, drawn_texts: list[str]) -> list[list[tuple]]:
    """Which of the paragraph's characters each of its drawn lines holds.

    Line breaking is Phase 3, so this is the oracle's, as which paragraph a line belongs
    to is: each drawn line's text (a list label skipped, as :func:`match_lines` skips it)
    is taken from the front of the paragraph's visible text.  Returns, per line,
    ``[(resolved run properties, character)]``; whitespace stays with the line that holds
    the character before it, and anything the drawn lines do not account for goes to
    the last line.
    """
    chars = []
    for run in paragraph.runs:
        resolved = resolve_run(document, paragraph, run)
        if not resolved.get("vanish"):
            chars += [(resolved, char) for char in run.text]
    remaining = _norm("".join(char for _, char in chars))
    quotas = []
    for text in drawn_texts:
        line = _norm(text)
        cut = next((k for k in range(len(line)) if line[k:] and remaining.startswith(line[k:])), None)
        taken = 0 if cut is None else len(line) - cut
        quotas.append(taken)
        remaining = remaining[taken:]
    out: list[list[tuple]] = [[] for _ in drawn_texts]
    line, used = 0, 0
    for resolved, char in chars:
        if not char.isspace():
            while line + 1 < len(out) and used >= quotas[line]:
                line, used = line + 1, 0
            used += 1
        out[line].append((resolved, char))
    return out


@dataclass(frozen=True)
class DrawnLine:
    """What the oracle says about one line: its page, baseline, and its text objects as
    ``(/BaseFont, drawn size in device px, text)``."""

    page: int
    y: float
    runs: tuple

    @property
    def text(self) -> str:
        return "".join(run[2] for run in self.runs)


def drawn_lines(pages) -> list[DrawnLine]:
    """``quartz_pdf.read(pdf)`` -> the lines :func:`predict` consumes (and tests record)."""
    import quartz_pdf

    return [
        DrawnLine(line.page, line.y, tuple((run.font, run.size_px, run.text) for run in line.runs))
        for line in merge_raised(quartz_pdf.lines(pages))
    ]


def merge_raised(lines):
    """Fold superscript and subscript runs back into the line they belong to.

    Quartz groups runs by baseline, and a ``w:vertAlign`` run is drawn smaller on a
    baseline of its own, so a line with a superscript arrives as two "lines" (the
    ``filesamples`` sample1: text at 836, "script" raised to 819 and lowered to 838).
    A group is folded into a neighbour on the same page when **every** run in it is
    smaller than the neighbour's largest and its baseline is within half that size;
    the neighbour keeps its baseline and the runs are re-sorted by pen x.  The size
    condition keeps apart a table row whose cells round to baselines 1 px apart in the
    same size.
    """
    import quartz_pdf

    out = list(lines)
    changed = True
    while changed:
        changed = False
        for i, line in enumerate(out):
            small = max(run.size_px for run in line.runs)
            best = None
            for j in (i - 1, i + 1):
                if not 0 <= j < len(out) or out[j].page != line.page:
                    continue
                big = max(run.size_px for run in out[j].runs)
                gap = abs(out[j].y - line.y)
                if small < big and gap < big / 2 and (best is None or gap < best[0]):
                    best = (gap, j)
            if best is not None:
                j = best[1]
                host = out[j]
                out[j] = quartz_pdf.Line(host.page, host.y,
                                         sorted(host.runs + line.runs, key=lambda r: r.x))
                del out[i]
                changed = True
                break
    return out


def check_glyphs(document: Document, drawn: list[DrawnLine], metrics=None) -> tuple[int, list[str]]:
    """Every drawn non-space glyph against the resolved face, weight, slant and size.

    A ``w:vertAlign`` run's size is the one its glyphs are drawn at
    (``resolve.script_half_points``), which needs the face's ``OS/2`` script size from
    ``metrics``; without it (or for a face recorded without one) the run's ``w:sz``.

    Walks the paragraphs and the drawn glyphs in step; a drawn glyph that no paragraph
    character accounts for (a list label, the tab after it) is skipped.  Returns the
    number of glyphs compared and a description of each disagreement.
    """
    stream = [(char, font, size) for line in drawn for font, size, text in line.runs
              for char in text if not char.isspace()]
    position = 0
    compared = 0
    problems: list[str] = []
    for block in blocks(document):
        paragraph = block.paragraph
        for run in paragraph.runs:
            resolved = resolve_run(document, paragraph, run)
            if resolved.get("vanish"):
                continue
            for char in run.text:
                if char.isspace():
                    continue
                want = char.upper() if resolved.get("caps") else char
                while position < len(stream) and stream[position][0] != want:
                    position += 1
                if position == len(stream):
                    problems.append(f"{char!r} of {paragraph.text[:30]!r} not drawn")
                    return compared, problems
                _, font, size = stream[position]
                position += 1
                compared += 1
                fmt = character_format(resolved, char, document)
                family, bold, italic = _drawn_face(font)
                expected = (fmt.face or "").replace(" ", "")
                half_points = fmt.half_points
                face = metrics(fmt.face, fmt.bold, fmt.italic) if metrics and fmt.face else None
                if fmt.vertical_align and face is not None:
                    half_points = script_half_points(fmt.half_points, fmt.vertical_align, face) or half_points
                size_px = int(Fraction(half_points, 2) * PX_PER_PT + Fraction(1, 2))
                if (family, bold, italic, int(size)) != (expected, fmt.bold, fmt.italic, size_px):
                    problems.append(f"{char!r} in {paragraph.text[:30]!r}: drawn {font} {size:g}px,"
                                    f" resolved {fmt.face} b={fmt.bold} i={fmt.italic} {size_px}px")
    return compared, problems


def _drawn_face(base: str) -> tuple[str, bool, bool]:
    """``TimesNewRomanPS-BoldMT`` -> ``("TimesNewRoman", True, False)``."""
    family, _, style = base.partition("-")
    for suffix in ("PSMT", "MT", "PS"):
        if family.endswith(suffix) and family != suffix:
            family = family[: -len(suffix)]
            break
    style = style.removesuffix("MT")
    bold = "Bold" in style
    italic = "Italic" in style or "Oblique" in style
    rest = style.replace("Bold", "").replace("Italic", "").replace("Oblique", "").replace("Regular", "")
    return family + rest, bold, italic


def predict(document: Document, drawn: list[DrawnLine], metrics, *, collapse: str = "max",
            before_at_page_top: bool = False, line_box: bool = True,
            package: bytes | None = None, advances=None) -> list[LineResult]:
    """Predicted against observed baselines for every drawn line.

    ``metrics(face, bold, italic)`` returns a :class:`FaceMetrics` or ``None``.

    The keyword arguments are the behaviours that were measured rather than assumed, and
    the defaults are what Word does (ROADMAP.md, "Style inheritance", "The line box" and
    "The page top"): the previous paragraph's space after and this one's space before
    **collapse to the larger** (``max``; ``sum`` is refuted -- except under
    ``w:doNotUseHTMLParagraphAutoSpacing``, where it is what Word does and which the
    model does not read yet: ROADMAP.md, "Autospacing"), what a paragraph that
    starts a page keeps of its space before is ``vertical.keeps_space_before_at_page_top``
    (``before_at_page_top=True``, keeping it everywhere, is refuted), and the paragraph
    spacing and bottom border are part of the **line box** the baseline rounds in
    (``line_box=False`` is Phase 2's model, which rounded in the line's pitch alone).

    ``advances`` (a :class:`docx2svg.measure.Advances`) lets the model's line breaker
    decide which characters each line holds, instead of Word's drawn lines.
    """
    all_blocks = blocks(document, package)
    labels: dict[int, str | None] = {}
    if advances is not None:
        counters = linebreak.ListCounters()
        for index, block in enumerate(all_blocks):
            resolved = resolve_paragraph(document, block.paragraph)
            if resolved.get("numPr.numId"):
                labels[index] = counters.label(document, block.paragraph, resolved)
    assignment = match_lines(document, all_blocks, drawn)
    extents: dict[int, tuple] = {}
    results: list[LineResult] = []

    page = None
    top = Fraction(0)
    pending_after = 0  # twips, already zeroed by contextual spacing where it applies
    previous_style = None
    previous_contextual = False
    stopped_page = None  # a page on which an out-of-scope block has moved everything below
    last_block = None
    last_line_of = {b: line for b, line in assignment}
    lines_of: dict[int, list] = {}
    for b, line in assignment:
        lines_of.setdefault(b, []).append(line)
    shares: dict[int, list] = {}
    order = list(dict.fromkeys(b for b, _ in assignment if b != UNMATCHED))
    following = {b: (order[i + 1] if i + 1 < len(order) else None) for i, b in enumerate(order)}
    section_starts = {min(i for i, b in enumerate(all_blocks) if b.section == s)
                      for s in {b.section for b in all_blocks}}

    def style_of(paragraph):
        return paragraph.properties.style_id if paragraph.properties else None

    numbered_cache: dict[int, bool] = {}
    borders_cache: dict[int, tuple] = {}

    def borders(index: int) -> tuple:
        """``(above, below)`` of block ``index``'s borders, its neighbours considered."""
        if index not in borders_cache:
            def resolved(k):
                return resolve_paragraph(document, all_blocks[k].paragraph) if 0 <= k < len(all_blocks) else None
            borders_cache[index] = vertical.paragraph_borders_px(
                resolved(index), resolved(index - 1), resolved(index + 1))
        return borders_cache[index]

    def numbered(index: int) -> bool:
        if index not in numbered_cache:
            resolved = resolve_paragraph(document, all_blocks[index].paragraph)
            numbered_cache[index] = bool(resolved.get("numPr.numId"))
        return numbered_cache[index]

    for block_index, line in assignment:
        if block_index == UNMATCHED:
            # Nothing is known about what this line is, so nothing below it on its page
            # can be predicted either.
            results.append(LineResult(line.page, UNMATCHED, line.text.strip()[:40], round(line.y),
                                      status="out of scope: not a body paragraph's text"))
            stopped_page = line.page
            continue
        block = all_blocks[block_index]
        paragraph = block.paragraph
        pp = resolve_paragraph(document, paragraph)
        section = document.sections[block.section]
        result = LineResult(line.page, block_index, line.text.strip()[:40], round(line.y))
        results.append(result)
        new_paragraph = block_index != last_block
        new_page = line.page != page
        style = style_of(paragraph)
        contextual = bool(pp.get("contextualSpacing"))

        before = pp.get("spacing.before", 0) or 0
        after = pp.get("spacing.after", 0) or 0
        # An autospaced side is 14 pt (the cascade), or nothing against its neighbour.
        if autospaced(document, pp, "before") and not autospace_kept(
                numbered=numbered(block_index),
                neighbour_numbered=numbered(block_index - 1) if block_index else None):
            before = 0
        if autospaced(document, pp, "after") and not autospace_kept(
                numbered=numbered(block_index),
                neighbour_numbered=numbered(block_index + 1) if block_index + 1 < len(all_blocks) else False):
            after = 0
        space_before = Fraction(0)
        if new_page:
            page = line.page
            top = twips_to_px(section.margins.top)
            above_break = pending_after
            pending_after = 0
            if new_paragraph and before_at_page_top:
                top += twips_to_px(before)
            elif new_paragraph:
                keeps = keeps_space_before_at_page_top(
                    section_start=block_index in section_starts,
                    page_break_before=bool(pp.get("pageBreakBefore")),
                    compatibility_mode=document.compatibility_mode)
                gap, space_before = page_top_gap_px(before, keeps=keeps, previous_after=above_break)
                top += gap
        elif new_paragraph:
            same = style == previous_style
            own_before = 0 if (contextual and same) else before
            prev_after = 0 if (previous_contextual and same) else pending_after
            if collapse == "max":
                gap, space_before = paragraph_gap_px(prev_after, own_before)
            else:
                gap, space_before = twips_to_px(prev_after) + twips_to_px(own_before), twips_to_px(own_before)
            top += gap

        if new_paragraph:
            above_border, _ = borders(block_index)
            top += above_border
            space_before += above_border

        if block.excluded:
            stopped_page = line.page
            result.status = f"out of scope: {block.excluded}"
        elif stopped_page == line.page:
            result.status = "out of scope: below a table or drawing"

        if block_index not in shares:
            # Each line's extent is that of the characters on it: the model's line breaker
            # decides which (``advances``), or -- without advances, or where the model's
            # lines are not as many as Word drew -- the oracle's lines do.
            drawn_texts = [l.text for l in lines_of[block_index]]
            found = None
            if advances is not None:
                found = model_line_shares(document, paragraph, section, advances, metrics,
                                          labels.get(block_index))
                if found is not None and len(found) != len(drawn_texts):
                    result.notes.append(f"the model breaks {len(found)} lines, Word drew {len(drawn_texts)}")
                    found = None
            shares[block_index] = found if found is not None else line_shares(document, paragraph, drawn_texts)
        line_number = lines_of[block_index].index(line)
        key = (block_index, line_number)
        if key not in extents:
            # Items that differ only in the character are one item as far as height goes.
            text_items, label_items = line_items(document, paragraph, pp, shares[block_index][line_number],
                                                 first_line=line_number == 0)
            text_items = list({_height_key(item): item for item in text_items}.values())
            label_items = list({_height_key(item): item for item in label_items}.values())
            extents[key] = (line_extent(text_items, metrics, label_items),
                            line_extent(text_items, metrics) if label_items else None,
                            tallest_natural(text_items, metrics))
        combined, text_only, text_natural = extents[key]
        rule = pp.get("spacing.lineRule", "auto")
        value = pp.get("spacing.line", 240)
        last_line = last_line_of[block_index] is line
        if combined is None:
            result.status = result.status or "out of scope: no face metrics"
            last_block = block_index
            continue
        face, half_points, mixed = combined
        if mixed:
            result.notes.append("mixed faces/sizes: ascent and descent maxed separately")

        # What sits below this line's pitch in its box: the bottom border and the space
        # after, on the paragraph's last line.  The space after is the paragraph's own,
        # unless contextual spacing removes it against the next paragraph.
        border_height = Fraction(0)
        if last_line:
            border_height = borders(block_index)[1]
            if border_height:
                result.notes.append("bottom border")
        own_after = 0
        if last_line:
            nxt = following.get(block_index)
            same_next = nxt is not None and style_of(all_blocks[nxt].paragraph) == style
            own_after = 0 if (contextual and same_next) else after

        em = Fraction(half_points, 2) * PX_PER_PT / face.units_per_em
        box = LineBox(
            pitch=mixed_line_pitch(combined, text_only, rule, value, text_natural),
            text_above=(face.ascent + face.line_gap) * em,
            text_below=face.descent * em,
            natural=natural_height_px(face, half_points),
            rule=rule,
            line=value,
            space_before=space_before if line_box else Fraction(0),
            space_after=(border_height + twips_to_px(own_after)) if line_box else Fraction(0),
        )
        box_top = top - box.space_before
        result.predicted = baseline_in_box(box_top, box)
        top += box.pitch
        if not result.status:
            result.status = "exact" if result.predicted == result.observed else "miss"

        if last_line:
            top += border_height
            pending_after = after
            previous_style = style
            previous_contextual = contextual
        last_block = block_index
    return results


def face_integers(face) -> list[int]:
    """What a recording keeps of a face: ``[upm, ascent, descent, lineGap]``, and the
    ``OS/2`` superscript and subscript sizes after them when the face has them."""
    out = [face.units_per_em, face.ascent, face.descent, face.line_gap]
    if face.superscript_size is not None:
        out += [face.superscript_size, face.subscript_size]
    return out


def summary(results: list[LineResult]) -> tuple[int, int, int]:
    scored = [r for r in results if r.status in ("exact", "miss")]
    return sum(r.status == "exact" for r in scored), len(scored), len(results) - len(scored)

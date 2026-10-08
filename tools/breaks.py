"""Score ``docx2svg.linebreak`` on whole documents: lines that break where Word breaks them.

The breaker's counterpart of ``baselines.py``, and it takes from the oracle exactly what
``baselines`` still does -- which paragraph each drawn line belongs to
(``baselines.match_lines``) -- and nothing else: which characters go on which of the
paragraph's lines is now the model's.

**What agrees.**  Each paragraph's lines are compared as spans of its drawn characters
(whitespace removed, the list label left out).  A line *agrees* when the model's line
of the same number starts and ends where Word's does, which is "breaks after the same
word" and also requires every line before it to agree.  Reported beside it, as a
diagnostic, is the *conditional* count: starting each line where Word started it, does
the model end it where Word did?  The two are equal when everything agrees; when they
differ, the conditional count says how many separate mistakes there are.

Out of scope, with the reason: what ``baselines.blocks`` excludes (tables, drawings,
fields, frames, note references), paragraphs no drawn line was matched to, paragraphs
Word drew in another order (a table of contents), and paragraphs with a piece the
breaker cannot measure (``linebreak.Unmeasurable``).  Justified paragraphs are scored
but counted apart: justification is not modelled (ROADMAP.md, Phase 3's "do not do").
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import baselines
from docx2svg import linebreak
from docx2svg.resolve import resolve_paragraph


@dataclass
class ParagraphResult:
    block: int
    text: str
    #: ``scored``, or why not.
    status: str
    observed: list[tuple[int, int]] = field(default_factory=list)  # spans of drawn chars
    predicted: list[tuple[int, int]] = field(default_factory=list)
    conditional: list[bool] = field(default_factory=list)
    justified: bool = False
    page: int = 0
    #: The paragraph's drawn characters, whitespace removed: what the spans index.
    content: str = ""

    def texts(self, which: str = "observed") -> list[str]:
        return [self.content[a:b] for a, b in getattr(self, which)]

    @property
    def agree(self) -> int:
        return sum(1 for a, b in zip(self.observed, self.predicted) if a == b)

    @property
    def lines(self) -> int:
        return len(self.observed)


def _norm(text: str) -> str:
    return re.sub(r"\s+", "", text)


#: What Word's PDF says for a character it draws as another: U+00AD in the text is drawn,
#: as a hyphen (``make_break_rules_probe.py``, family ``shy``).
DRAWN_AS = {"\u00ad": "-", "\u2010": "-", "\u2011": "-"}
#: Characters Word does not draw at all.
NOT_DRAWN = frozenset("\u200b")


def _content(items) -> list[tuple[int, str]]:
    """``(piece index, drawn character)`` for every non-space character of the content
    (not the label): what Word's drawn lines are compared with."""
    out = []
    for index, piece in enumerate(items):
        if piece.source < 0 or piece.kind in (linebreak.TAB, linebreak.BREAK, linebreak.SOFT_HYPHEN):
            continue
        if piece.char.isspace() or piece.char in NOT_DRAWN:
            continue
        out.append((index, DRAWN_AS.get(piece.char, piece.char)))
    return out


def observed_spans(drawn_texts: list[str], content: str, *, greedy: bool = False) -> list[tuple[int, int]] | None:
    """Each drawn line's span of the content, the label (a prefix that does not match)
    skipped on the first line; a hyphen Word drew at a soft hyphen is not content.

    ``greedy``: stop at the line that completes the content, and ignore the lines after
    it (the caller counts ``len`` of the result to know how many were consumed).
    """
    spans = []
    position = 0
    for number, text in enumerate(drawn_texts):
        if greedy and position == len(content) and spans:
            break
        line = _norm(text)
        remaining = content[position:]
        found = None
        for cut in range(len(line) + 1 if number == 0 else 1):
            piece = line[cut:]
            for candidate in (piece, piece[:-1] if piece.endswith(("-", "‐", "‑")) else None):
                if candidate is not None and candidate and remaining.startswith(candidate):
                    found = len(candidate)
                    break
            if found is not None:
                break
        if found is None:
            if not line:
                spans.append((position, position))
                continue
            return None
        spans.append((position, position + found))
        position += found
    if position != len(content):
        return None
    return spans


def _spans(lines, content_index: dict[int, int], total: int) -> list[tuple[int, int]]:
    """Predicted lines as spans of content characters."""
    out = []
    for line in lines:
        indices = [content_index[i] for i in range(line.start, line.end) if i in content_index]
        if indices:
            out.append((indices[0], indices[-1] + 1))
        else:
            start = out[-1][1] if out else 0
            out.append((start, start))
    return out


def score(document, drawn, advances, metrics=None, *, package: bytes | None = None,
          sequential: bool = False) -> list[ParagraphResult]:
    """``sequential``: every paragraph is drawn, in order, and nothing else is -- a
    generated probe -- so each takes the drawn lines that spell its content, from where
    the previous one stopped.  Otherwise lines are matched to paragraphs by
    ``baselines.match_lines``, as the vertical model's are."""
    all_blocks = baselines.blocks(document, package)
    assignment = [] if sequential else baselines.match_lines(document, all_blocks, drawn)
    cursor = 0
    lines_of: dict[int, list] = {}
    for block, line in assignment:
        if block != baselines.UNMATCHED:
            lines_of.setdefault(block, []).append(line)
    counters = linebreak.ListCounters()
    results = []
    for index, block in enumerate(all_blocks):
        paragraph = block.paragraph
        pp = resolve_paragraph(document, paragraph)
        label = counters.label(document, paragraph, pp) if pp.get("numPr.numId") else None
        drawn_lines = lines_of.get(index, []) if not sequential else drawn[cursor:cursor + 1]
        result = ParagraphResult(index, paragraph.text[:40], "scored",
                                 justified=pp.get("jc") in ("both", "distribute"),
                                 page=drawn_lines[0].page if drawn_lines else -1)
        results.append(result)
        if block.excluded:
            result.status = f"out of scope: {block.excluded}"
            continue
        if not drawn_lines:
            result.status = "out of scope: not drawn, or not matched"
            continue
        section = document.sections[block.section]
        try:
            items, lines = linebreak.break_paragraph(document, paragraph, section, advances, metrics, label)
        except linebreak.Unmeasurable as error:
            result.status = f"out of scope: unmeasurable ({error})"
            continue
        content = _content(items)
        text = "".join(char for _, char in content)
        if sequential:
            # From the cursor, or -- when something before was drawn in a way the text does
            # not match -- from the first line within the next fifty that begins this one.
            observed = None
            for start in range(cursor, min(cursor + 50, len(drawn))):
                observed = observed_spans([line.text for line in drawn[start:start + 200]], text, greedy=True)
                if observed is not None and (start == cursor or observed[0][1] > observed[0][0]):
                    cursor = start + len(observed)
                    break
                observed = None
        else:
            observed = observed_spans([line.text for line in drawn_lines], text)
        if observed is None:
            result.status = "out of scope: drawn text is not the paragraph's"
            continue
        content_index = {piece: position for position, (piece, _) in enumerate(content)}
        result.observed = observed
        result.content = text
        result.predicted = _spans(lines, content_index, len(text))
        # The conditional count: from where Word started each line, where does the model end it?
        piece_of = [piece for piece, _ in content]
        geo = linebreak.geometry(document, paragraph, section, pp)
        for number, (start, end) in enumerate(observed):
            if start == end:
                result.conditional.append(True)
                continue
            first_piece = 0 if number == 0 else _line_start_piece(items, piece_of[start])
            sub = linebreak.break_pieces(items[first_piece:], geo, advances, first_line=number == 0)
            ends = [content_index[i + first_piece] for i in range(sub[0].start, sub[0].end)
                    if i + first_piece in content_index]
            result.conditional.append(bool(ends) and ends[-1] + 1 == end)
    return results


def _line_start_piece(items, piece: int) -> int:
    """Word starts a line at the first piece after the previous line's trailing spaces;
    a line that begins after a forced break keeps the spaces after it."""
    start = piece
    while start > 0 and items[start - 1].kind == linebreak.SPACE:
        start -= 1
    if start > 0 and items[start - 1].kind == linebreak.BREAK:
        return start
    return piece


def summary(results: list[ParagraphResult], *, justified: bool | None = False) -> dict:
    """``{"lines": n, "agree": a, "conditional": c, "paragraphs": p, "out": o}`` over the
    scored paragraphs (``justified``: False = not justified, True = only justified,
    None = all)."""
    scored = [r for r in results if r.status == "scored" and (justified is None or r.justified == justified)]
    return {
        "paragraphs": len(scored),
        "lines": sum(r.lines for r in scored),
        "agree": sum(r.agree for r in scored),
        "conditional": sum(sum(r.conditional) for r in scored),
        "out": sum(1 for r in results if r.status != "scored"),
    }

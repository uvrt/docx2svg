"""Score ``docx2svg.paginate`` against Word: which line starts each page.

Word's side is its drawn lines, matched to paragraphs by text (``baselines.match_lines``):
the page of every line of every paragraph.  The model's side is ``paginate.paginate`` on
the file alone -- the line breaker, the vertical model and the keeps -- with nothing from
Word's export.

Three counts, each an equality:

* **page tops** -- Word's page tops (the first line of each page but the first, as
  ``(paragraph, line of it)``) that the model starts a page at, *running free* from the
  document's start.  One wrong break usually moves every later one, so this is the
  strict number, and the one the roadmap asks for.
* **conditional page ends** -- from each of Word's page tops, does the model end that
  page where Word did?  Phase 3's "conditional" count, for pages: it counts separate
  mistakes, and it is what is left to report past an obstacle the model cannot size.
  It takes Word's page *top* from the oracle, so it is never the headline number.
* **paragraphs** -- paragraphs whose every line is on the page Word put it on (running
  free).
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import baselines  # noqa: E402

from docx2svg import paginate  # noqa: E402


def body_lines(document, drawn) -> list:
    """Word's drawn lines less those of headers and footers: a baseline above every
    section's top margin or below every section's bottom margin is not the body's (a
    header that pushes the body down is not caught)."""
    from docx2svg.vertical import twips_to_px

    top = min(twips_to_px(abs(s.margins.top)) for s in document.sections)
    bottom = max(twips_to_px(s.page_size.height_twips - abs(s.margins.bottom)) for s in document.sections)
    return [line for line in drawn if top < line.y <= bottom]


def word_pages(document, data: bytes, drawn) -> dict[int, list[int]]:
    """Block index -> the page of each of its drawn lines, from Word's export."""
    all_blocks = baselines.blocks(document, data)
    drawn = body_lines(document, drawn)
    out: dict[int, list[int]] = {}
    for block, line in baselines.match_lines(document, all_blocks, drawn):
        if block != baselines.UNMATCHED:
            out.setdefault(block, []).append(line.page)
    return out


def tops(pages: dict[int, list[int]]) -> dict[int, tuple[int, int]]:
    """Page -> ``(block, line)`` of its first line, from a block -> pages mapping."""
    out: dict[int, tuple[int, int]] = {}
    for block in sorted(pages):
        for line, page in enumerate(pages[block]):
            if page not in out or (block, line) < out[page]:
                out[page] = (block, line)
    return out


@dataclass
class Score:
    name: str
    word_tops: int = 0
    tops_matched: int = 0
    #: Word's page tops the free-running model never reached (after an obstacle).
    tops_unreached: int = 0
    conditional: int = 0
    conditional_scored: int = 0
    paragraphs: int = 0
    paragraphs_matched: int = 0
    #: Word's paragraphs with another number of lines than the model breaks them into.
    line_count_differs: int = 0
    stopped: str | None = None
    misses: list = field(default_factory=list)
    conditional_misses: list = field(default_factory=list)

    def row(self) -> str:
        return (f"{self.name:34} page tops {self.tops_matched:4} / {self.word_tops:<4}"
                f" (unreached {self.tops_unreached}); conditional {self.conditional} / {self.conditional_scored};"
                f" paragraphs {self.paragraphs_matched} / {self.paragraphs}"
                + (f"; stopped at {self.stopped}" if self.stopped else ""))


def score(name: str, document, data: bytes, drawn, advances, metrics, rules=None, *,
          features=True, word: dict[int, list[int]] | None = None) -> Score:
    """The model against Word on one document.  ``word`` (block -> pages, as
    :func:`word_pages` finds them) replaces ``drawn`` where a recording keeps only that."""
    if word is None:
        word = word_pages(document, data, drawn)
    items = paginate.flow(document, advances, metrics,
                          features=baselines.paragraph_features(data) if features else None)
    result = paginate.paginate(document, items, rules)
    block_of = {i: item.block for i, item in enumerate(items)}
    model = paginate.drawn_pages(items, result)
    out = Score(name)
    if result.stopped is not None:
        obstacle = items[result.stopped]
        out.stopped = f"b{obstacle.block} ({obstacle.reason})"
    word_top = tops(word)
    # Both sides in drawn lines: Word draws nothing for a line holding only pictures
    # (not its paragraph's last) or the empty line before a page break.
    model_starts = set(tops(model).values())
    reached = max(model.keys(), default=-1) if result.stopped is None else items[result.stopped].block
    for page, top in sorted(word_top.items()):
        if page == min(word_top):
            continue
        out.word_tops += 1
        if top in model_starts:
            out.tops_matched += 1
        else:
            if top[0] >= reached and result.stopped is not None:
                out.tops_unreached += 1
            out.misses.append((page, top))
    for block, pages in word.items():
        if block not in model:
            continue
        out.paragraphs += 1
        if len(model[block]) != len(pages):
            out.line_count_differs += 1
        if model[block] == pages:
            out.paragraphs_matched += 1
    # Conditional: from each of Word's page tops, where does the model end the page?
    index_of = {item.block: i for i, item in enumerate(items)}
    ordered = sorted(word_top.items())
    for (page, top), (next_page, following) in zip(ordered, ordered[1:]):
        if top[0] not in index_of or following[0] not in index_of or next_page != page + 1:
            # A page between them drew no text (a page of pictures): its top is unknown.
            continue
        start = _model_position(items, index_of[top[0]], top[1])
        want = _model_position(items, index_of[following[0]], following[1])
        item = items[start.item]
        if isinstance(item, paginate.Obstacle):
            continue
        try:
            got = paginate.page_end(document, items, start, rules)
        except paginate.Unplaceable:
            continue
        if got.item < len(items) and isinstance(items[got.item], paginate.Para):
            # The first line of the next page that Word would draw.
            while got.item < len(items) and isinstance(items[got.item], paginate.Para) and (
                    got.line >= items[got.item].count or got.line in items[got.item].undrawn):
                got = paginate.Position(got.item + 1) if got.line >= items[got.item].count - 1 \
                    else paginate.Position(got.item, got.line + 1)
        out.conditional_scored += 1
        if got == want:
            out.conditional += 1
        else:
            gi = (items[got.item].block, got.line) if got.item < len(items) else ("end",)
            out.conditional_misses.append((page, top, following, gi))
    return out


def _model_position(items, index: int, drawn_line: int):
    """The model's position of a paragraph's ``drawn_line``-th drawn line."""
    item = items[index]
    if not isinstance(item, paginate.Para):
        return paginate.Position(index, drawn_line)
    drawn = [n for n in range(item.count) if n not in item.undrawn]
    return paginate.Position(index, drawn[drawn_line] if drawn_line < len(drawn) else drawn_line)


# -- probes -------------------------------------------------------------------------------


def record_probe(name: str, data: bytes, metrics_faces: dict, advance_faces: dict):
    """Word's lines of a generated probe (``oracle.py``, cached by content), and a model
    run on it that records every face metric and advance it asked for."""
    import face_advances
    import probe_documents

    drawn = probe_documents.measure(name, data)
    metrics = probe_documents.RecordingMetrics()
    metrics.faces = metrics_faces
    advances = face_advances.RecordingAdvances(face_advances.InstalledAdvances(data), advance_faces)
    return drawn, advances, metrics


def recorded(payload: dict):
    """``(advances, metrics)`` answering from a probe recording, without fonts."""
    import face_advances
    import probe_documents

    return (face_advances.RecordedAdvances(payload["advances"]),
            probe_documents.recorded_metrics(payload["faces"]))


def write_recording(path: Path, about: str, faces: dict, advances: dict, documents: dict) -> None:
    import json

    payload = {"_about": about, "faces": dict(sorted(faces.items())),
               "advances": dict(sorted(advances.items())), "documents": documents}
    path.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True) + "\n")
    print(f"wrote {path} ({path.stat().st_size} bytes)")


def compact_pages(word: dict[int, list[int]], blocks: int, drawn=None) -> dict:
    """What a probe recording keeps of Word's export, enough to rebuild :func:`word_pages`:
    the ``[block, line]`` each page starts at, and every block's number of lines where it
    is not one (0: no line of it matched).  With ``drawn``, also the baseline and text of
    every page's first drawn line (``first``), which shows a line no paragraph accounts
    for -- an empty paragraph mark -- at a page's top."""
    counts = {str(block): len(word.get(block, ())) for block in range(blocks) if len(word.get(block, ())) != 1}
    starts = sorted(tops(word).items())
    for block, found in word.items():
        if found != sorted(found):
            raise ValueError(f"block {block}'s lines are not in page order")
    out = {"blocks": blocks, "lines": counts, "tops": [[page, block, line] for page, (block, line) in starts]}
    if drawn is not None:
        first: dict[int, list] = {}
        for line in drawn:
            first.setdefault(line.page, [line.page, line.y, line.text])
        out["first"] = [first[page] for page in sorted(first)]
    return out


def expand_pages(recorded: dict) -> dict[int, list[int]]:
    out: dict[int, list[int]] = {}
    starts = {(block, line): page for page, block, line in recorded["tops"]}
    page = None
    for block in range(recorded["blocks"]):
        count = recorded["lines"].get(str(block), 1)
        if not count:
            continue
        out[block] = []
        for line in range(count):
            page = starts.get((block, line), page)
            out[block].append(page)
    return out

#!/usr/bin/env python3
"""Read ``make_hyphen_probe.py``: Word's lines of every case, the rules they show, and how
the library and its patterns score against them.

Word's PDF (``oracle.py``) is read with ``quartz_pdf``; every paragraph keeps its lines
together and has 18 pt after it, so a paragraph's lines are those between two gaps wider
than 1.6 em (or a page's top).  Each paragraph's drawn lines are checked against its text
-- the hyphens Word draws at a line's end taken off, and at most one letter changed where
a word broke (Dutch ``omaatje`` is drawn ``oma-`` / ``tje``).

What is recorded (``--record``, ``tests/fixtures/hyphen-observations.json``) is only what
is scored: per document, per paragraph, the text of the lines that decide -- the first for
the swept families, every line of running text, and page by page for ``bottom`` (whose
cases are told apart by their pages, not by gaps) -- and for running text the pen x of each
line's drawn hyphen and the end of its last glyph, in device px from the column's left.

Scored here and in ``tests/test_hyphenation.py``:

* **the model**: the paragraph laid out by the library (``docx2svg._render`` with the
  faces' recorded numbers), its lines against Word's, line by line;
* **the patterns** (``points`` only): for every probe word, Word's break points -- read off
  the sweep, the last point at or before each ``k`` -- against Liang's with the library's
  minimums, per language.

Usage::

    python tools/read_hyphen_probe.py [--record] [--only NAME...]
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_hyphen_probe as probe  # noqa: E402

OBSERVATIONS = HERE.parent / "tests" / "fixtures" / "hyphen-observations.json"
FACES = HERE.parent / "tests" / "fixtures" / "hyphen-faces.json"
#: Families whose every line is recorded; the rest record their first line.
ALL_LINES = ("text", "limit")
PX_PER_UNIT = 300 / 72 / 4096
#: A string this repository's files must not hold (case-insensitive): the letters i and c
#: and a hyphen, spelt here so that this file does not hold it either.
FORBIDDEN = re.compile("i" + "c" + "-", re.I)


def _norm(text: str) -> str:
    return re.sub(r"\s+", "", text).replace(probe.NBH, "-").lower()


def segment(lines) -> list[list]:
    """``quartz_pdf.lines`` grouped into paragraphs by the gaps between them."""
    out: list[list] = []
    previous = None
    for line in lines:
        em = max(run.size_px for run in line.runs)
        if previous is None or line.page != previous.page or line.y - previous.y > 1.6 * em:
            out.append([])
        out[-1].append(line)
        previous = line
    return out


def matches(text: str, drawn: list[str]) -> bool:
    """Whether ``drawn`` (a paragraph's line texts) is ``text``: a hyphen at a line's end
    may be Word's, and at most one letter may change about each such break."""
    full = _norm(text)
    target = "".join(_norm(t) for t in drawn)
    if target == full:
        return True
    stripped = "".join(_norm(t)[:-1] if _norm(t).endswith("-") and i < len(drawn) - 1 else _norm(t)
                       for i, t in enumerate(drawn))
    if stripped == full:
        return True
    # A spelling change at a break: allow a small edit.
    import difflib

    ops = [op for op in difflib.SequenceMatcher(None, full, stripped).get_opcodes() if op[0] != "equal"]
    return len(ops) <= len(drawn) and all(max(i2 - i1, j2 - j1) <= 2 for _, i1, i2, j1, j2 in ops)


def word_lines(name: str, data: bytes) -> list[list[str]]:
    """Word's drawn lines of every paragraph of document ``name``."""
    import oracle
    import quartz_pdf

    lines = quartz_pdf.lines(quartz_pdf.read(oracle.export(data, name=f"hyphen-{name}")))
    paragraphs = segment(lines)
    cases = probe.documents()[name][1]
    if len(paragraphs) != len(cases):
        raise ValueError(f"{name}: {len(paragraphs)} paragraphs drawn, {len(cases)} written")
    out = []
    for case, drawn in zip(cases, paragraphs):
        texts = [line.text.rstrip() for line in drawn]
        if not matches(case.text.replace(probe.SHY, ""), texts):
            raise ValueError(f"{name}: drawn {texts} for {case.text!r}")
        out.append(texts)
    return out


def word_pages(name: str, data: bytes) -> list[list[list[str]]]:
    """Word's drawn lines of every ``bottom`` case, page by page: a case's paragraph starts
    a page of its own (after its empty paragraph, which draws nothing) and runs on to the
    next."""
    import oracle
    import quartz_pdf

    pages: dict = {}
    for line in quartz_pdf.lines(quartz_pdf.read(oracle.export(data, name=f"hyphen-{name}"))):
        if line.text.strip():
            pages.setdefault(line.page, []).append(line.text.rstrip())
    out: list = []
    start = probe.TEXT["en-US"].split(" ")[0]
    for number in sorted(pages):
        if pages[number][0].startswith(start + " "):
            out.append([])
        out[-1].append(pages[number])
    cases = probe.documents()[name][1]
    if len(out) != len(cases):
        raise ValueError(f"{name}: {len(out)} paragraphs drawn, {len(cases)} written")
    for case, drawn in zip(cases, out):
        if not matches(case.text, [text for page in drawn for text in page]):
            raise ValueError(f"{name}: drawn {drawn} for {case.text!r}")
    return out


def model_pages(data: bytes, fonts) -> list[list[list[str]]]:
    """As :func:`word_pages`, from the library's layout: every paragraph that draws text,
    its lines page by page."""
    import render_record
    from docx2svg import _lay_out

    layout, _, _ = _lay_out(data, render_record.options(fonts))
    by_path: dict = {}
    for number, page in enumerate(layout.pages):
        for line in page.lines:
            text = "".join("".join(span.chars) for span in line.spans)
            if text.strip():
                by_path.setdefault(line.path, {}).setdefault(number, []).append(text)
    return [[pages[number] for number in sorted(pages)] for pages in by_path.values()]


def drawn(name: str, data: bytes) -> list:
    """Word's lines of document ``name``: per paragraph, page by page for ``bottom``."""
    return word_pages(name, data) if name.startswith("bottom-") else word_lines(name, data)


def model(name: str, data: bytes, fonts) -> list:
    """The library's lines of document ``name``, as :func:`drawn` gives Word's."""
    return model_pages(data, fonts) if name.startswith("bottom-") else model_lines(data, fonts)


def hyphen_glyphs(name: str, data: bytes) -> list[list]:
    """Every hyphen Word drew at a line's end: ``[page, baseline, x, exact, x of the glyph
    before it or None]``, device px to 1e-3 -- ``exact`` when the hyphen starts a text
    object (its x is Word's to 1e-4 px), else its step from the glyph before it is what
    Quartz's encoding keeps (``glyphs.py``)."""
    import glyphs
    import oracle

    out = []
    drawn = glyphs.word_glyphs(oracle.export(data, name=f"hyphen-{name}"))
    for k, glyph in enumerate(drawn):
        last = k + 1 == len(drawn) or (drawn[k + 1].page, drawn[k + 1].line_y) != (glyph.page, glyph.line_y)
        if glyph.char == "-" and last:
            before = drawn[k - 1] if k and not glyph.exact and drawn[k - 1].run == glyph.run else None
            out.append([glyph.page, round(glyph.y), round(glyph.x, 3), glyph.exact,
                        None if before is None else round(before.x, 3)])
    return out


def model_hyphen_glyphs(data: bytes, fonts) -> list[list]:
    """As :func:`hyphen_glyphs`, from the library's layout: ``[page, baseline, x, x of the
    glyph before it, the drawn size in px]`` of every hyphen it draws at a line's end."""
    import render_record
    from docx2svg import _lay_out
    from docx2svg.svg import font_size_px

    layout, _, _ = _lay_out(data, render_record.options(fonts))
    out = []
    for number, page in enumerate(layout.pages):
        for line in page.lines:
            glyphs = [(char, x, span) for span in line.spans for char, x in zip(span.chars, span.xs)
                      if not char.isspace()]
            for k, (char, x, span) in enumerate(glyphs):
                if span.kind == "hyphen":
                    before = glyphs[k - 1][1] if k else None
                    out.append([number, line.baseline, float(x), None if before is None else float(before),
                                float(font_size_px(span.half_points, "device"))])
    return out


def hyphens_placed(word: list[list], ours: list[list]) -> tuple[int, int]:
    """``(agreeing, Word's)``: Word's line-end hyphens the model draws at the same baseline
    and x -- to ``glyphs.exact_tolerance`` where Word's x is exact, else by the step from
    the glyph before it, to half of 1/1000 em -- matched in page and baseline order."""
    import glyphs

    mine = {(page, y): (x, before, size) for page, y, x, before, size in ours}
    agree = 0
    for page, y, x, exact, before in word:
        found = mine.get((page, y))
        if found is None:
            continue
        ox, obefore, size = found
        if exact:
            agree += abs(ox - x) <= glyphs.exact_tolerance(x)
        elif before is not None and obefore is not None:
            agree += abs((ox - obefore) - (x - before)) <= glyphs.STEP_EM * size + 2e-3
    return agree, len(word)


def recorded_lines(name: str, lines: list[list[str]]) -> list:
    cases = probe.documents()[name][1]
    out = []
    for case, texts in zip(cases, lines):
        out.append(texts if case.family in ALL_LINES + ("bottom",) else texts[0])
    return out


# -- the model ------------------------------------------------------------------------------


def model_layout(data: bytes, fonts) -> list[list]:
    """The library's lines of every paragraph of a probe document, as layout lines."""
    import render_record
    from docx2svg import _lay_out

    layout, _, _ = _lay_out(data, render_record.options(fonts))
    by_path: dict = {}
    for page in layout.pages:
        for line in page.lines:
            by_path.setdefault(line.path, []).append(line)
    return list(by_path.values())


def model_lines(data: bytes, fonts) -> list[list[str]]:
    """The library's lines of every paragraph of a probe document."""
    return [["".join("".join(span.chars) for span in line.spans) for line in lines]
            for lines in model_layout(data, fonts)]


def same(word: str, ours: str) -> bool:
    return _norm(word).replace("\u2011", "-") == _norm(ours).replace("\u2011", "-")


def score(name: str, recorded: list, ours: list[list[str]]) -> list[bool]:
    """Per paragraph, whether the model's lines are Word's (the recorded ones)."""
    cases = probe.documents()[name][1]
    if len(ours) != len(cases):
        raise ValueError(f"{name}: the model laid out {len(ours)} paragraphs of {len(cases)}")
    out = []
    for case, word, mine in zip(cases, recorded, ours):
        if case.family == "bottom":
            out.append([len(page) for page in word] == [len(page) for page in mine]
                       and all(same(a, b) for x, y in zip(word, mine) for a, b in zip(x, y)))
        elif case.family in ALL_LINES:
            out.append(len(word) == len(mine) and all(same(a, b) for a, b in zip(word, mine)))
        else:
            out.append(same(word, mine[0]))
    return out


# -- what Word's lines say -----------------------------------------------------------------


def pattern_agreement(name: str, firsts: list[str]) -> dict:
    """Per language: ``(words whose points are Word's, words, points in common, the
    library's, Word's)`` for the ``points`` sweeps of a document -- the library's points (Liang's, with
    the mode's minimums) against those Word's lines show (:func:`word_points`)."""
    from docx2svg import hyphenate

    cases = probe.documents()[name][1]
    mode15 = probe.documents()[name][0]["mode"] == 15
    out: dict = {}
    swept = [(case, first) for case, first in zip(cases, firsts) if case.family == "points"]
    for (lang, word), (points, _) in word_points(*zip(*swept)).items():
        ours = set(hyphenate.points(word, lang, right_min=3 if mode15 else 2, mode15=mode15))
        cell = out.setdefault(lang, [0, 0, 0, 0, 0])
        cell[0] += ours == points
        cell[1] += 1
        cell[2] += len(ours & points)
        cell[3] += len(ours)
        cell[4] += len(points)
    return {lang: tuple(cell) for lang, cell in out.items()}


def first_break(case, first: str) -> int | None:
    """Where Word broke the word under test on the first line: the number of its
    characters there (a hyphen drawn after them), 0 when it went down whole, ``None``
    when the line holds it all."""
    info = case.info()
    word = info.get("word", probe.ZONE_WORD)
    line = _norm(first)
    head = _norm(case.text.split(word)[0])
    if not line.startswith(head[:-1] if head.endswith("-") else head):
        raise ValueError((case.text, first))
    rest = line[len(head):]
    if not rest:
        return 0
    if _norm(word).startswith(rest) and not rest.endswith("-"):
        return len(rest)
    if rest.endswith("-"):
        body = rest[:-1]
        plain = word.replace(probe.SHY, "")
        if _norm(plain).startswith(body):
            # Index in the word as written (soft hyphens counted).
            count, seen = 0, 0
            for char in word:
                if seen == len(body):
                    break
                count += 1
                if char != probe.SHY:
                    seen += 1
            return count
        return -len(body)  # a spelling change: |n| letters drawn before the hyphen
    return len(rest)


def word_points(cases, firsts) -> dict:
    """``(lang, word) -> (Word's points, ks where the line held nothing of it)``, from the
    ``points`` sweep: a break after ``j`` letters at ``k`` says ``j`` is a point and no point
    lies in ``(j, k]``."""
    out: dict = {}
    for case, first in zip(cases, firsts):
        info = case.info()
        key = (info["lang"], info["word"])
        found = first_break(case, first)
        points, spelled = out.setdefault(key, (set(), {}))
        if found and found > 0 and found < len(info["word"]):
            points.add(found)
        elif found is not None and found < 0:
            spelled[info["k"]] = -found
    return out


#: The English word list of the probe's own words, as Word was observed breaking them.
EXCEPTIONS = HERE.parent / "src" / "docx2svg" / "patterns" / "en-observed-exceptions.txt"


def observed_exceptions(documents: dict) -> list[str]:
    """The lines of ``EXCEPTIONS``, from the recording alone: every English word of the
    ``points`` and ``words`` sweeps whose points Word's lines show, without compatibility
    mode or in mode 15, differ from those the Moby list and the patterns give (with the
    mode's minimums) -- spelt with Word's points below mode 15 and, after a space, those of
    mode 15 where they are not the same points less the last letter's."""
    from docx2svg import hyphenate

    lists = tuple(name for name in hyphenate.WORD_LISTS["en-us-knuth"] if name != EXCEPTIONS.name)
    out = {}
    for family in ("points", "words"):
        seen = {}
        for mode in ("none", "15"):
            name = f"{family}-{mode}"
            swept = [(case, first) for case, first in zip(probe.documents()[name][1], documents[name])
                     if case.family == "points"]
            for (lang, word), (points, _) in word_points(*zip(*swept)).items():
                if lang.startswith("en"):
                    seen.setdefault(word, {})[mode] = points
        for word, by_mode in seen.items():
            below, fifteen = by_mode["none"], by_mode["15"]
            mine = hyphenate.word_points(word, "en-us-knuth", False, lists)
            mine15 = hyphenate.word_points(word, "en-us-knuth", True, lists)
            if (below == {p for p in mine if 2 <= p <= len(word) - 2}
                    and fifteen == {p for p in mine15 if 2 <= p <= len(word) - 3}):
                continue
            spelt = _spelt(word, below)
            if fifteen != {p for p in below if p <= len(word) - 3}:
                spelt += " " + _spelt(word, fifteen)
            out[word] = spelt
    lines = [out[word] for word in sorted(out)]
    if any(FORBIDDEN.search(line) for line in lines):
        raise ValueError("an exception holds a string this repository keeps out of its files")
    return lines


def _spelt(word: str, points) -> str:
    from docx2svg.hyphenate import WORD_MARK

    return "".join((WORD_MARK if i in points else "") + c for i, c in enumerate(word))


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("--only", nargs="*")
    parser.add_argument("--misses", action="store_true", help="print every paragraph the model breaks otherwise")
    parser.add_argument("--exceptions", action="store_true",
                        help=f"write {EXCEPTIONS.name} from the recording (no Word) and stop")
    args = parser.parse_args(argv[1:])
    if args.exceptions:
        documents = json.loads(OBSERVATIONS.read_text(encoding="utf-8"))["documents"]
        EXCEPTIONS.write_text("".join(line + "\n" for line in observed_exceptions(documents)), encoding="utf-8")
        print(f"wrote {EXCEPTIONS}")
        return 0
    import render_record

    recorded: dict = {}
    positions: dict = {}
    faces: dict = {}
    for name in probe.documents():
        if args.only and not any(name.startswith(prefix) for prefix in args.only):
            continue
        data = probe.build(name)
        lines = drawn(name, data)
        recorded[name] = recorded_lines(name, lines)
        if name.startswith("text-"):
            positions[name] = hyphen_glyphs(name, data)
        fonts = render_record.RecordingFonts(data, faces)
        ours = model(name, data, fonts)
        verdicts = score(name, recorded[name], ours)
        by_family: dict = defaultdict(lambda: [0, 0])
        for case, verdict in zip(probe.documents()[name][1], verdicts):
            by_family[case.family][0] += verdict
            by_family[case.family][1] += 1
        print(f"{name:24} model {sum(verdicts)} / {len(verdicts)}: "
              + ", ".join(f"{family} {a}/{b}" for family, (a, b) in by_family.items()))
        if name in positions:
            placed, total = hyphens_placed(positions[name], model_hyphen_glyphs(data, fonts))
            print(f"{'':24} line-end hyphens drawn where Word draws them: {placed} / {total}")
        if name.startswith(("points-", "words-", "french-")):
            for lang, (same_words, words, common, ours_, words_) in pattern_agreement(name, recorded[name]).items():
                print(f"{'':24} patterns {lang}: {same_words} / {words} words break where Word breaks them; "
                      f"points {common} in common, {ours_} ours, {words_} Word's")
        if args.misses:
            for case, verdict, word, mine in zip(probe.documents()[name][1], verdicts, recorded[name], ours):
                if not verdict:
                    shown = mine if case.family in ALL_LINES else mine[0]
                    print(f"    {case.family} {case.info()} word {word!r} ours {shown!r}")
    for name, rows in recorded.items():
        blob = json.dumps(rows, ensure_ascii=False)
        if FORBIDDEN.search(blob):
            raise ValueError(f"{name}: a line holds a string this repository keeps out of its files")
    if args.record:
        payload = {
            "_about": ("Word 16.106's lines of tools/make_hyphen_probe.py's documents, read by "
                       "tools/read_hyphen_probe.py: per document, per paragraph, the text of its first "
                       "line, or of every line for running text (page by page for bottom); and for the "
                       "text documents every hyphen drawn at a line's end, [page, baseline, x, x exact, x "
                       "of the glyph before], device px. Measurements only; regenerate with --record."),
            "documents": recorded,
            "positions": positions,
        }
        OBSERVATIONS.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
                                + "\n")
        print(f"wrote {OBSERVATIONS}")
        render_record.dump(FACES, dict(sorted(faces.items())))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

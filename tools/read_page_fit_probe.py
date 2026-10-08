#!/usr/bin/env python3
"""Measure ``make_page_fit_probe.py``: which slacks Word keeps a candidate line on its page.

For every case the candidate's page is compared with its anchor's: the same page means
it fitted.  Each family's threshold is reported as the smallest slack that fitted and the
largest that did not, per compatibility setting; then every hypothesis is scored as the
number of cases it predicts, and the model (``docx2svg.paginate``, run on the file alone)
is scored as page tops placed where Word placed them.

``--record`` writes ``tests/fixtures/page-fit-observations.json`` for
``tests/test_paginate.py``.

Usage::

    python tools/read_page_fit_probe.py [--record]
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_page_fit_probe as probe  # noqa: E402
import pages  # noqa: E402
import probe_documents  # noqa: E402

from docx2svg import parse_package  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "page-fit-observations.json"


def fitted(word: dict[int, list[int]]) -> dict[int, bool]:
    """Case number -> whether its candidate line is on its anchor's page, from Word's
    pages per block (``pages.word_pages``)."""
    out = {}
    number = anchor = None
    for block, (role, _) in enumerate(probe.blocks()):
        if role == "anchor":
            number = 0 if number is None else number + 1
            anchor = word.get(block, [None])[0]
        elif role == "candidate" and block in word:
            line = probe.CASES[number].family.candidate
            if line < len(word[block]) and anchor is not None:
                out[number] = word[block][line] == anchor
    return out


def _props(family) -> dict:
    return dict(family.props)


def _last(family) -> bool:
    return family.candidate == family.lines - 1


def _after(family) -> int:
    """The candidate line's space after (its paragraph's last line only), in units."""
    spacing = _props(family).get("spacing") or {}
    return probe.twips_units(spacing.get("after", 0)) if _last(family) else 0


def _border(family, setting=None, every_line=True) -> int:
    """The candidate's bottom border if it counts: on its paragraph's last line, and --
    below mode 15, as measured -- on any line (``every_line``)."""
    if "pBdr" not in _props(family):
        return 0
    if _last(family) or (every_line and setting != "15"):
        return probe.BORDER_UNITS
    return 0


def _extra(family) -> int:
    """An auto multiple's extra below the candidate's text (negative: text past the pitch)."""
    spacing = _props(family).get("spacing") or {}
    if spacing.get("lineRule") == "auto" and spacing.get("line", 240) != 240:
        return family.pitch - probe.LINE
    return 0


def _text(c) -> int:
    return c.slack + max(_extra(c.family), 0)


HYPOTHESES = {
    "text and border (the model)": lambda c, s: _text(c) - _border(c.family, s) >= 0,
    "pitch, inclusive": lambda c, s: c.slack >= 0,
    "pitch, strict": lambda c, s: c.slack > 0,
    "pitch and border": lambda c, s: c.slack - _border(c.family, s) >= 0,
    "box: pitch, space after and border": lambda c, s: c.slack - _border(c.family, s) - _after(c.family) >= 0,
    "text, no border": lambda c, s: _text(c) >= 0,
    "text and border, last line only in every mode": lambda c, s: _text(c) - _border(c.family, s, False) >= 0,
    "text and border, on every line in every mode": lambda c, s: _text(c) - _border(c.family, "any") >= 0,
    "text and border, overhang too": lambda c, s: c.slack + _extra(c.family) - _border(c.family, s) >= 0,
    "text and border, strict": lambda c, s: _text(c) - _border(c.family, s) > 0,
}


def report(observed: dict[str, dict[int, bool]]) -> None:
    print(f"{'family':12} {'threshold':>10}" + "".join(f"{s:>22}" for s in probe.SETTINGS))
    for family in probe.FAMILIES:
        for threshold in family.thresholds:
            cells = []
            for setting in probe.SETTINGS:
                fits = observed[setting]
                cases = [(n, c) for n, c in enumerate(probe.CASES) if c.family is family and c.threshold == threshold]
                held = [c.slack - threshold for n, c in cases if fits.get(n)]
                lost = [c.slack - threshold for n, c in cases if n in fits and not fits[n]]
                cells.append(f"fits from {min(held) if held else '-'}, not {max(lost) if lost else '-'}")
            print(f"{family.name:12} {threshold:>+10}" + "".join(f"{c:>22}" for c in cells))
    for name, rule in HYPOTHESES.items():
        agree = sum(rule(probe.CASES[n], setting) == fit
                    for setting, fits in observed.items() for n, fit in fits.items())
        total = sum(len(fits) for fits in observed.values())
        print(f"  {name:45} {agree} / {total}")


def model_scores(documents: dict, advances, metrics, rules=None) -> dict[str, pages.Score]:
    """The model on each document, against Word's pages per block."""
    out = {}
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        out[setting] = pages.score(setting, parse_package(data), data, None, advances, metrics, rules,
                                   word=pages.expand_pages(documents[setting]))
    return out


def offline(data: dict | None = None):
    """``(observed fits, model scores)`` from the recording, without Word or fonts."""
    data = data or json.loads(OBSERVATIONS.read_text())
    advances, metrics = pages.recorded(data)
    words = {setting: pages.expand_pages(data["documents"][setting]) for setting in probe.SETTINGS}
    return {s: fitted(w) for s, w in words.items()}, model_scores(data["documents"], advances, metrics)


def main(argv: list[str]) -> int:
    faces: dict = {}
    advance_faces: dict = {}
    documents = {}
    count = len(probe.blocks())
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        drawn, advances, metrics = pages.record_probe(f"page-fit-{setting}", data, faces, advance_faces)
        documents[setting] = pages.compact_pages(pages.word_pages(parse_package(data), data, drawn), count)
    observed = {setting: fitted(pages.expand_pages(lines)) for setting, lines in documents.items()}
    report(observed)
    for setting, score in model_scores(documents, advances, metrics).items():
        print(score.row())
    if "--record" in argv[1:]:
        pages.write_recording(OBSERVATIONS, (
            "For each document of tools/make_page_fit_probe.py, the page Word 16.106 drew each line "
            "of each paragraph on (null: none matched); the four hhea/typo integers of every face "
            "the model asked for; and the advance widths the line breaker asked for, in font units. "
            "Measurements only; regenerate with tools/read_page_fit_probe.py --record."),
            faces, advance_faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

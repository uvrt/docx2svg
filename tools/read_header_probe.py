#!/usr/bin/env python3
"""Measure ``make_header_probe.py``: where each section's body starts and how many lines
its first page holds, against where its header and footer reach.

Usage::

    python tools/read_header_probe.py [--record]
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_header_probe as probe  # noqa: E402
import probe_documents  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "header-observations.json"


def observe(lines) -> dict[int, tuple]:
    """Case number -> (first body baseline, lines on the section's first page)."""
    pages: dict[int, dict[int, list]] = defaultdict(lambda: defaultdict(list))
    for page, y, text in lines:
        words = text.split()
        if len(words) >= 3 and words[0] == "Case":
            pages[int(words[1])][page].append(y)
    out = {}
    for number, by_page in pages.items():
        first = min(by_page)
        out[number] = (min(by_page[first]), len(by_page[first]))
    return out


def body(drawn) -> list:
    """The body's lines: the probe's own text (a header pushed into the body's area is
    drawn there too, and where it shares a baseline with a body line, as under the
    negative top margin, Quartz merges the two and that line is not counted)."""
    return [line for line in drawn if line.text.startswith("Case")]


def scores(documents: dict, advances, metrics) -> dict:
    """The model's page tops against Word's, per setting."""
    import pages

    from docx2svg import parse_package

    out = {}
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        out[setting] = pages.score(setting, parse_package(data), data, None, advances, metrics,
                                   word=pages.expand_pages(documents[setting]["pages"]))
    return out


def offline() -> dict:
    import pages

    data = json.loads(OBSERVATIONS.read_text(encoding="utf-8"))
    advances, metrics = pages.recorded(data)
    return scores(data["documents"], advances, metrics)


def main(argv: list[str]) -> int:
    import pages

    from docx2svg import parse_package

    faces: dict = {}
    advance_faces: dict = {}
    documents = {}
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        drawn, advances, metrics = pages.record_probe(f"header-{setting}", data, faces, advance_faces)
        lines = body(drawn)
        found = pages.word_pages(parse_package(data), data, lines)
        documents[setting] = {"pages": pages.compact_pages(found, len(probe.CASES)),
                              "first": observe([[line.page, line.y, line.text] for line in lines])}
    print(f"{'case':26}" + "".join(f"{s:>14}" for s in probe.SETTINGS))
    for number, case in enumerate(probe.CASES):
        print(f"{case.name:26}" + "".join(
            f"{str(tuple(documents[s]['first'].get(number, ()))):>14}" for s in probe.SETTINGS))
    for setting, score in scores(documents, advances, metrics).items():
        print(score.row(), score.misses, score.conditional_misses)
    if "--record" in argv[1:]:
        pages.write_recording(OBSERVATIONS, (
            "For each document of tools/make_header_probe.py, the page Word 16.106 drew each body line on, and "
            "per case (section) its first body baseline in device px and the lines on its first page; the "
            "face metrics and advance widths the model asked for. Measurements only; regenerate with "
            "tools/read_header_probe.py --record."), faces, advance_faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

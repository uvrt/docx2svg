#!/usr/bin/env python3
"""Measure ``make_section_probe.py``: the page of every line of every section, Word's and
the model's (``docx2svg.paginate``, from the file alone).

``--record`` writes ``tests/fixtures/section-observations.json``.

Usage::

    python tools/read_section_probe.py [--record]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_section_probe as probe  # noqa: E402
import pages  # noqa: E402
import probe_documents  # noqa: E402

from docx2svg import parse_package  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "section-observations.json"


def scores(documents: dict, advances, metrics, rules=None) -> dict:
    out = {}
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        out[setting] = pages.score(setting, parse_package(data), data, None, advances, metrics, rules,
                                   word=pages.expand_pages(documents[setting]))
    return out


def offline(rules=None) -> dict:
    data = json.loads(OBSERVATIONS.read_text(encoding="utf-8"))
    advances, metrics = pages.recorded(data)
    return scores(data["documents"], advances, metrics, rules)


def main(argv: list[str]) -> int:
    faces: dict = {}
    advance_faces: dict = {}
    documents = {}
    for setting in probe.SETTINGS:
        data = probe.build(setting)
        drawn, advances, metrics = pages.record_probe(f"section-{setting}", data, faces, advance_faces)
        found = pages.word_pages(parse_package(data), data, drawn)
        documents[setting] = pages.compact_pages(found, len(probe.SECTIONS))
        print(setting, "Word's first page of each section:", [found[k][0] for k in sorted(found)])
    for setting, score in scores(documents, advances, metrics).items():
        print(score.row(), score.misses)
    if "--record" in argv[1:]:
        pages.write_recording(OBSERVATIONS, (
            "For each document of tools/make_section_probe.py, the page Word 16.106 drew each line of each "
            "section on; the face metrics and advance widths the model asked for. Measurements only; regenerate "
            "with tools/read_section_probe.py --record."), faces, advance_faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

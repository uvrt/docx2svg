#!/usr/bin/env python3
"""Measure ``make_pagination_probe.py``: every paragraph on Word's page, by the model alone.

For each document the model (``docx2svg.paginate``, from the file alone) is scored against
the page Word drew every line of every paragraph on: page tops placed where Word placed
them, and paragraphs whose every line is on Word's page.

``--record`` writes ``tests/fixtures/pagination-observations.json`` for
``tests/test_paginate.py``.

Usage::

    python tools/read_pagination_probe.py [--record] [-v]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_pagination_probe as probe  # noqa: E402
import pages  # noqa: E402
import probe_documents  # noqa: E402

from docx2svg import parse_package  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "pagination-observations.json"


def scores(documents: dict, advances, metrics, rules=None) -> dict:
    out = {}
    for document in probe.DOCUMENTS:
        data = probe.build(document)
        out[document.name] = pages.score(document.name, parse_package(data), data, None, advances, metrics, rules,
                                         word=pages.expand_pages(documents[document.name]))
    return out


def offline(rules=None) -> dict:
    data = json.loads(OBSERVATIONS.read_text())
    advances, metrics = pages.recorded(data)
    return scores(data["documents"], advances, metrics, rules)


def main(argv: list[str]) -> int:
    faces: dict = {}
    advance_faces: dict = {}
    documents = {}
    for document in probe.DOCUMENTS:
        data = probe.build(document)
        drawn, advances, metrics = pages.record_probe(f"pagination-{document.name}", data, faces, advance_faces)
        parsed = parse_package(data)
        found = pages.word_pages(parsed, data, drawn)
        documents[document.name] = pages.compact_pages(found, len(parsed.body))
    for name, score in scores(documents, advances, metrics).items():
        print(score.row(), f"line counts differ: {score.line_count_differs}")
        if "-v" in argv[1:]:
            print("   misses", score.misses[:10])
            print("   conditional", score.conditional_misses[:10])
    if "--record" in argv[1:]:
        pages.write_recording(OBSERVATIONS, (
            "For each document of tools/make_pagination_probe.py, the page Word 16.106 drew each line of each "
            "paragraph on; the face metrics and advance widths the model asked for. Measurements only; "
            "regenerate with tools/read_pagination_probe.py --record."), faces, advance_faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

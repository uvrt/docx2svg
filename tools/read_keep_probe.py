#!/usr/bin/env python3
"""Measure ``make_keep_probe.py``: where Word ends the page, against the model.

For every case, the page of each line of each paragraph under test is taken as an offset
from its anchor's page -- Word's (its export) and the model's (``docx2svg.paginate`` on the
file alone).  A case agrees when every one of those offsets does.  The report lists each
family's cases as ``room: Word's offsets | the model's``, marking disagreements.

``--record`` writes ``tests/fixtures/keep-observations.json`` for ``tests/test_paginate.py``.

Usage::

    python tools/read_keep_probe.py [--record] [-v]
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_keep_probe as probe  # noqa: E402
import pages  # noqa: E402
import probe_documents  # noqa: E402

from docx2svg import paginate, parse_package  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "keep-observations.json"


def offsets(document: str, found: dict[int, list[int]]) -> dict[int, list]:
    """Case number -> per test paragraph, the page of each line less the anchor's page."""
    out: dict[int, list] = defaultdict(list)
    anchor = {}
    for block, (number, role, index, _) in enumerate(probe.blocks(document)):
        if role == "anchor":
            anchor[number] = found.get(block, [None])[0]
        elif role == "test":
            base = anchor.get(number)
            out[number].append(tuple(p - base for p in found.get(block, [])) if base is not None else None)
    return out


def model_pages(document: str, advances, metrics, rules=None) -> dict[int, list[int]]:
    data = probe.build(document)
    doc = parse_package(data)
    items = paginate.flow(doc, advances, metrics)
    return paginate.drawn_pages(items, paginate.paginate(doc, items, rules))


def compare(documents: dict, advances, metrics, rules=None) -> dict:
    """``{document: {case number: (Word's offsets, the model's)}}``."""
    out = {}
    for document in probe.DOCUMENTS:
        word = offsets(document, pages.expand_pages(documents[document]))
        model = offsets(document, model_pages(document, advances, metrics, rules))
        out[document] = {n: (word[n], model[n]) for n in word}
    return out


def tally(compared: dict) -> tuple[int, int]:
    cases = [w == m for per in compared.values() for w, m in per.values()]
    return sum(cases), len(cases)


def _format(offsets: list) -> str:
    """``0 01 1``: per paragraph, each line's page offset; for long cases, a count."""
    if len(offsets) >= 5:
        return f"{sum(1 for p in offsets if p and p[0] == 0)} paragraphs on the anchor's page"
    if len(offsets) == 1 and offsets[0] and len(offsets[0]) > 10:
        return f"{offsets[0].count(0)} of {len(offsets[0])} lines on the anchor's page"
    return " ".join("".join(str(x) for x in p) if p else "-" for p in offsets)


def report(compared: dict, verbose: bool = False) -> None:
    for document, per in compared.items():
        rows: dict = defaultdict(list)
        for number, (word, model) in per.items():
            case = probe.CASES[number]
            rows[(case.family, case.variant, tuple(p.lines for p in case.paragraphs))].append(
                (case.room, word, model))
        agree = sum(w == m for w, m in per.values())
        print(f"== {document}: {agree} / {len(per)} cases agree")
        for (family, variant, shape), cases in rows.items():
            bad = [c for c in cases if c[1] != c[2]]
            if not verbose and not bad:
                continue
            shape_text = "+".join(map(str, shape)) if len(shape) < 5 else f"{len(shape)} paragraphs"
            print(f"  {family}/{variant} [{shape_text}]")
            for room, word, model in cases:
                fmt = _format
                print(f"    {'!!' if word != model else '  '} room {room:2}: word {fmt(word):20} model {fmt(model)}")


def offline(data: dict | None = None, rules=None):
    data = data or json.loads(OBSERVATIONS.read_text())
    advances, metrics = pages.recorded(data)
    return compare(data["documents"], advances, metrics, rules)


def main(argv: list[str]) -> int:
    faces: dict = {}
    advance_faces: dict = {}
    documents = {}
    for document in probe.DOCUMENTS:
        data = probe.build(document)
        drawn, advances, metrics = pages.record_probe(f"keep-{document}", data, faces, advance_faces)
        found = pages.word_pages(parse_package(data), data, drawn)
        documents[document] = pages.compact_pages(found, len(probe.blocks(document)), drawn)
    compared = compare(documents, advances, metrics)
    report(compared, "-v" in argv[1:])
    print("total", tally(compared))
    if "--record" in argv[1:]:
        pages.write_recording(OBSERVATIONS, (
            "For each document of tools/make_keep_probe.py, the page Word 16.106 drew each line of "
            "each paragraph on, and every page's first drawn line [page, baseline, text]; the face "
            "metrics and advance widths the model asked for. Measurements only; regenerate with "
            "tools/read_keep_probe.py --record."), faces, advance_faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

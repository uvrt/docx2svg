#!/usr/bin/env python3
"""Measure ``make_table_keep_probe.py``: on which page Word puts each row of each case,
against the model.

Every marker a case draws -- a row's cell text (``R3a``), a nested table's (``N5``), the
paragraph before and after the table -- is taken with the page it is on, as an offset
from the page of the case's anchor line: Word's from its export, the model's from
``docx2svg.paginate`` on the file alone (with recorded face metrics and advances).  A case
agrees when every marker does.

``--record`` writes ``tests/fixtures/table-keep-observations.json`` for
``tests/test_table_keep.py``.

Usage::

    python tools/read_table_keep_probe.py [--record] [-v]
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import baselines  # noqa: E402
import make_table_keep_probe as probe  # noqa: E402
import pages  # noqa: E402
import probe_documents  # noqa: E402

from docx2svg import paginate, parse_package  # noqa: E402

OBSERVATIONS = probe_documents.FIXTURES / "table-keep-observations.json"
#: What a case draws that is followed: anchors, rows' cells, nested tables, the paragraphs
#: around the table (the body lines that fill the page, ``F<k>``, and wrapped words not).
MARKER = re.compile(r"Case \d+|After \d+|Lead|[RC]\d+[ab]\d?|N\d+|H[ab]|C[ab]")


def _cases(found: list[tuple[int, str]]) -> dict[str, dict[str, int]]:
    """Case name -> marker -> its page less the case's anchor's, from ``(page, text)`` in
    document order."""
    out: dict[str, dict[str, int]] = {}
    current = anchor = None
    for page, text in found:
        for marker in MARKER.findall(text):
            if marker.startswith("Case "):
                current, anchor = probe.CASES[int(marker[5:])][0], page
                out[current] = {}
            elif current is not None:
                # A header row repeated on a later page counts where it is first drawn.
                out[current].setdefault(marker, page - anchor)
    return out


def word_cases(drawn) -> dict[str, dict[str, int]]:
    return _cases([(line.page, line.text) for line in drawn])


def model_cases(data: bytes, advances, metrics, rules=None) -> dict[str, dict[str, int]]:
    document = parse_package(data)
    items = paginate.flow(document, advances, metrics)
    found = paginate.drawn_pages(items, paginate.paginate(document, items, rules))
    texts = ["".join(run.text for run in block.paragraph.runs) for block in baselines.blocks(document, data)]
    # A nested table's paragraphs are drawn with their cell, whole (the model does not
    # split a nested table): on the page of the next paragraph that has one.
    page = None
    placed = []
    for k in reversed(range(len(texts))):
        page = found[k][0] if found.get(k) else page
        placed.append((page, texts[k]))
    return _cases([(found_page, text) for found_page, text in reversed(placed) if found_page is not None])


def compare(documents: dict, advances, metrics, rules=None) -> dict:
    """``{setting: {case: (Word's markers, the model's)}}``."""
    out = {}
    for setting in probe.SETTINGS:
        model = model_cases(probe.build(setting), advances, metrics, rules)
        out[setting] = {name: (word, model.get(name)) for name, word in documents[setting].items()}
    return out


def tally(compared: dict) -> tuple[int, int]:
    cases = [w == m for per in compared.values() for w, m in per.values()]
    return sum(cases), len(cases)


def _format(markers: dict | None) -> str:
    if markers is None:
        return "-"
    by_page: dict[int, list[str]] = {}
    for marker, page in markers.items():
        if not marker.endswith("b") and "b" not in marker[1:]:
            by_page.setdefault(page, []).append(marker)
    return " | ".join(f"p{page}: {' '.join(found[:4])}{' ...' if len(found) > 4 else ''}"
                      for page, found in sorted(by_page.items()))


def report(compared: dict, verbose: bool = False) -> None:
    for setting, per in compared.items():
        agree = sum(w == m for w, m in per.values())
        print(f"== {setting}: {agree} / {len(per)} cases agree")
        for name, (word, model) in per.items():
            if verbose or word != model:
                print(f"  {'!!' if word != model else '  '} {name:26} word  {_format(word)}")
                if word != model:
                    print(f"     {'':26} model {_format(model)}")


def offline(data: dict | None = None, rules=None):
    data = data or json.loads(OBSERVATIONS.read_text(encoding="utf-8"))
    advances, metrics = pages.recorded(data)
    return compare(data["documents"], advances, metrics, rules)


def main(argv: list[str]) -> int:
    faces: dict = {}
    advance_faces: dict = {}
    documents = {}
    advances = metrics = None
    for setting in probe.SETTINGS:
        drawn, advances, metrics = pages.record_probe(f"table-keep-{setting}", probe.build(setting), faces,
                                                      advance_faces)
        documents[setting] = word_cases(drawn)
    compared = compare(documents, advances, metrics)
    report(compared, "-v" in argv[1:])
    print("total", tally(compared))
    if "--record" in argv[1:]:
        pages.write_recording(OBSERVATIONS, (
            "For each document of tools/make_table_keep_probe.py, the page Word 16.106 drew each row's "
            "cell text and the paragraphs around each table on, less its case's anchor's page; the face "
            "metrics and advance widths the model asked for. Measurements only; regenerate with "
            "tools/read_table_keep_probe.py --record."), faces, advance_faces, documents)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

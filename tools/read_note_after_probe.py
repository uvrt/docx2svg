#!/usr/bin/env python3
"""Where ``make_note_after_probe.py``'s candidates go, in Word and in docx2svg: for each
case, whether the candidate line stayed on its case's page (``0``) or went on to the next
(``1``), read off Word's PDF (the page its text is on, against the page its case's first
line is on) and off the layout.

``--record`` writes Word's side to ``tests/fixtures/note-after-observations.json``, which
``tests/test_footnote_draw.py`` holds the model to.

Usage::

    python tools/read_note_after_probe.py [--record]
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import read_render  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "note-after-observations.json"
_CASE = re.compile(r"(Case|Candidate) (\d+)")


def _moves(pages: list[str]) -> list[int]:
    """Per case, the candidate's page less its case's, from each page's text."""
    first: dict = {}
    candidate: dict = {}
    for index, text in enumerate(pages):
        for kind, number in _CASE.findall(text):
            (first if kind == "Case" else candidate).setdefault(int(number), index)
    return [candidate[k] - first[k] for k in sorted(first)]


def word_side(pdf: Path) -> list[int]:
    import pymupdf

    with pymupdf.open(str(pdf)) as document:
        return _moves([page.get_text() for page in document])


def model_side(data: bytes, fonts) -> list[int]:
    """The same off the layout, measured with ``fonts`` (``render_record``'s)."""
    import render_record
    from docx2svg import convert_docx_to_layout

    layout = convert_docx_to_layout(data, render_record.options(fonts))
    return _moves([" ".join("".join("".join(span.chars) for span in line.spans) for line in page.lines)
                   for page in layout.pages])


def documents() -> list[tuple[str, bytes]]:
    import make_note_after_probe as probe

    return [(name, probe.build(name)) for name in probe.DOCUMENTS]


def main(argv: list[str] | None = None) -> int:
    import make_note_after_probe as probe
    import oracle
    import render_record

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true")
    options = parser.parse_args(argv)
    recorded, faces = {}, {}
    for name, data in documents():
        word = word_side(oracle.export(data, name=name))
        model = model_side(data, render_record.RecordingFonts(data, faces))
        agree = sum(a == b for a, b in zip(word, model))
        print(f"{name:24s} candidates on Word's page {agree} / {len(word)}")
        print("   spacer twips", " ".join(f"{h}:{w}{'' if w == m else '*'}" for h, w, m in zip(probe.HEIGHTS, word,
                                                                                            model)))
        recorded[name] = word
    if options.record:
        render_record.dump(OBSERVATIONS, {"_about": "Per case of tools/make_note_after_probe.py, whether Word put "
                                                    "the candidate on its case's page (0) or the next (1).",
                                          "faces": dict(sorted(faces.items())), "documents": recorded})
        print("recorded", OBSERVATIONS)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

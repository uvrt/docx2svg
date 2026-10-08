#!/usr/bin/env python3
"""Score picture bullets against Word: ``make_picture_bullet_probe.py``.

Two scores per document:

* **lines** -- ``read_float_table_probe.py``'s: every line of Word's within half a pixel
  of the model's (every glyph on it matched on the same baseline), per family; the text
  after each bullet shows where the label ended, and the ``After`` line how tall the
  bullet's line was;
* **pictures** -- every image Word drew (the bullets) against the model's pictures on the
  same page, matched in reading order, a box agreeing when every edge is within
  ``read_picture_place_probe.TOLERANCE`` px.

``--record`` writes Word's side to ``tests/fixtures/picture-bullet-observations.json``;
``tests/test_picture_bullet.py`` holds the model to it offline.

Usage::

    python tools/read_picture_bullet_probe.py [-v] [--record] [NAME...]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_picture_bullet_probe as probe  # noqa: E402
import read_float_table_probe as reader  # noqa: E402
import read_picture_place_probe as pictures  # noqa: E402
import read_render  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "picture-bullet-observations.json"


def documents() -> list[tuple[str, bytes]]:
    return [(name, probe.build(name.rsplit("-", 1)[1])) for name in probe.DOCUMENTS]


def picture_score(layout, boxes: list, pages: dict) -> tuple[list, dict]:
    """``[agreeing, Word's pictures]`` and, per case, what disagrees."""
    model = pictures.model_boxes(layout)
    agree, problems = 0, {}
    for page in sorted({box[0] for box in boxes}):
        theirs = [box for box in boxes if box[0] == page]
        ours = [box for box in model if box[0] == page]
        for k, box in enumerate(theirs):
            other = ours[k] if k < len(ours) else None
            if other is not None and all(abs(a - b) <= pictures.TOLERANCE for a, b in zip(other[1:], box[1:])):
                agree += 1
            else:
                problems[f"case {pages.get(page)}"] = (box[1:], other[1:] if other else None)
    return [agree, len(boxes)], problems


def score(name: str, data: bytes, fonts, recorded: dict) -> dict:
    from docx2svg import _render

    result = reader.score(name, data, fonts, recorded, probe.CASES)
    _documents, layout, _data, _fonts = _render(data, render_record.options(fonts))
    pages = reader.case_pages(recorded["objects"])
    result["pictures"], result["picture_problems"] = picture_score(layout, recorded["images"], pages)
    return result


def main(argv: list[str]) -> int:
    import oracle

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("names", nargs="*")
    parser.add_argument("--record", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--limit", type=int, default=40)
    args = parser.parse_args(argv[1:])
    faces: dict = {}
    recorded: dict = {}
    try:
        for name, data in documents():
            if args.names and name not in args.names:
                continue
            pdf = oracle.export(data, name=name)
            word = {"objects": read_render.word_objects(pdf), "fills": read_render.word_fills(pdf),
                    "images": pictures.word_boxes(pdf)}
            fonts = render_record.RecordingFonts(data, faces)
            result = score(name, data, fonts, word)
            print(result["result"].row())
            print(f"{'':34} lines within half a pixel {result['lines'][0]} / {result['lines'][1]} "
                  f"({', '.join(f'{k} {a}/{b}' for k, (a, b) in result['families'].items())}); pictures "
                  f"{result['pictures'][0]} / {result['pictures'][1]}; pages {result['pages']}; "
                  f"warnings {result['warnings']}")
            if args.verbose:
                pages = reader.case_pages(word["objects"])
                for where, problems in list(result["problems"].items())[:args.limit]:
                    case = pages.get(int(where[1:].split()[0]) - 1)
                    note = probe.CASES[case].note if case is not None else ""
                    print(f"    {where} (case {case}: {note}): {problems[:3]}")
                for where, (theirs, ours) in list(result["picture_problems"].items())[:args.limit]:
                    case = int(where.split()[1]) if where.split()[1].isdigit() else None
                    note = probe.CASES[case].note if case is not None else ""
                    print(f"    picture {where} ({note}): Word {theirs}, model {ours}")
            recorded[name] = {**word, "glyphs": result["glyphs"], "lines": result["lines"],
                              "families": result["families"], "pictures": result["pictures"],
                              "warnings": result["warnings"]}
    finally:
        oracle.recover()
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for tools/make_picture_bullet_probe.py: every text object (as "
                       "render-observations.json records them), every filled rectangle and every image box, the "
                       "renderer's scores against them -- lines within half a pixel, per family; the glyph row; "
                       "the bullets' boxes -- and every face number the renderer asked for. Measurements only; "
                       "regenerate with tools/read_picture_bullet_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

#!/usr/bin/env python3
"""Measure where Word draws inline pictures: ``make_picture_place_probe.py`` and
``make_picture_probe.py``, every picture's box against the model's.

Word's PDF draws each picture as an image whose box (``pymupdf``'s image info, in points)
is read in device px, and matched in reading order with the model's pictures on the same
page.  A box agrees when every edge is within :data:`TOLERANCE` px -- the precision to
which Quartz writes an image's matrix.

``--record`` writes Word's boxes to ``tests/fixtures/picture-place-observations.json``;
``tests/test_picture_place.py`` holds the model to them offline.

Usage::

    python tools/read_picture_place_probe.py [--record] [-v]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_picture_place_probe  # noqa: E402
import make_picture_probe  # noqa: E402
import read_render  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "picture-place-observations.json"
#: Quartz writes an image's matrix in points to about 0.001 pt; a box agrees within this.
TOLERANCE = 0.03


def documents() -> list[tuple[str, bytes]]:
    out = [(f"picture-place-{s}", make_picture_place_probe.build(s)) for s in make_picture_place_probe.SETTINGS]
    out += [(f"picture-{s}", make_picture_probe.build(s)) for s in make_picture_probe.SETTINGS]
    return out


def word_boxes(pdf: Path) -> list:
    """``[page, x0, y0, x1, y1]`` of every image Word drew, device px, 1e-3."""
    import pymupdf

    out = []
    with pymupdf.open(str(pdf)) as document:
        for index, page in enumerate(document):
            for info in page.get_image_info():
                out.append([index] + [round(v * 300 / 72, 3) for v in info["bbox"]])
    return out


def model_boxes(layout) -> list:
    return [[index, float(p.x), float(p.y), float(p.x + p.width), float(p.y + p.height)]
            for index, page in enumerate(layout.pages) for p in page.pictures]


def compare(word: list, ours: list) -> tuple[int, int, list]:
    """``(agreeing, scored, disagreements)``, matching each page's boxes in reading order."""
    def by_page(boxes):
        out: dict = {}
        for box in boxes:
            out.setdefault(box[0], []).append(box)
        return {page: sorted(items, key=lambda b: (round(b[2] / 4), b[1])) for page, items in out.items()}

    theirs, mine = by_page(word), by_page(ours)
    agree = scored = 0
    problems = []
    for page in sorted(set(theirs) | set(mine)):
        a, b = theirs.get(page, []), mine.get(page, [])
        if len(a) != len(b):
            problems.append((page, "count", len(a), len(b)))
            scored += max(len(a), len(b))
            continue
        for w, m in zip(a, b):
            scored += 1
            if max(abs(x - y) for x, y in zip(w[1:], m[1:])) <= TOLERANCE:
                agree += 1
            else:
                problems.append((page, w[1:], [round(v, 3) for v in m[1:]]))
    return agree, scored, problems


def main(argv: list[str]) -> int:
    import oracle
    from docx2svg import _render

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--record", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv[1:])
    faces: dict = {}
    recorded: dict = {}
    for name, data in documents():
        word = word_boxes(oracle.export(data, name=name))
        fonts = render_record.RecordingFonts(data, faces)
        _, layout, _, _ = _render(data, render_record.options(fonts))
        agree, scored, problems = compare(word, model_boxes(layout))
        print(f"{name:22} pictures where Word drew them: {agree} / {scored}")
        if args.verbose:
            for problem in problems:
                print("   ", problem)
        recorded[name] = word
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("Where Word 16.106 drew every inline picture of tools/make_picture_place_probe.py and "
                       "tools/make_picture_probe.py: [page, x0, y0, x1, y1], device px, from its PDF's image "
                       "boxes (tools/read_picture_place_probe.py). Measurements only; regenerate with --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

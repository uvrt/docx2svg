#!/usr/bin/env python3
"""Score text beside a floating drawing against Word: ``make_wrap_side_probe.py``.

Word's PDF gives every text object and every picture (``read_render.word_objects``,
``read_anchor_probe.word_images``); the model lays each document out and draws it.  Three
scores, per document:

* **lines** -- a line of Word's agrees when every glyph on it is matched to the model's on
  the same baseline and the first glyph of every text object Word starts on it (where
  Word positions a glyph exactly: each segment's start, each run) is within half a device
  pixel of the model's, and every glyph inside an object steps from the one before as the
  model's does (``glyphs.STEP_EM``); counted per family (:data:`FAMILIES`);
* **glyphs** -- ``glyphs.compare``'s row, as every probe records it;
* **pictures** -- the drawings' boxes, in Word's paint order (``read_anchor_probe``).

``--record`` writes Word's side to ``tests/fixtures/wrap-side-observations.json``;
``tests/test_wrap_side.py`` holds the model to it offline.

Usage::

    python tools/read_wrap_side_probe.py [-v] [--record] [NAME...]
"""

from __future__ import annotations

import argparse
import difflib
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import glyphs  # noqa: E402
import make_wrap_side_probe as probe  # noqa: E402
import read_anchor_probe  # noqa: E402
import read_render  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "wrap-side-observations.json"
#: How far a text object's first glyph may be from the model's, device px.
HALF_PIXEL = 0.5


def documents() -> list[tuple[str, bytes]]:
    return [(name, probe.build(name)) for name in probe.DOCUMENTS]


def line_score(model: list, word: list) -> dict:
    """``{(page, baseline): problems}`` for every line of Word's (``[]`` where it agrees):
    see the module's docstring."""
    out: dict = {}
    for page in sorted({g.page for g in word}):
        ours = sorted((g for g in model if g.page == page), key=lambda g: (round(g.y), g.x))
        theirs = sorted((g for g in word if g.page == page), key=lambda g: (round(g.y), g.x))
        for g in theirs:
            out.setdefault((page, round(g.line_y or g.y)), [])
        matcher = difflib.SequenceMatcher(None, [glyphs._key(g) for g in ours], [glyphs._key(g) for g in theirs],
                                          autojunk=False)
        matched = set()
        for block in matcher.get_matching_blocks():
            previous = None
            for k in range(block.size):
                a, b = ours[block.a + k], theirs[block.b + k]
                matched.add(id(b))
                if b.exact:
                    ok = abs(a.x - b.x) <= HALF_PIXEL
                elif previous is not None and previous[1].run == b.run:
                    advances = b.index - previous[1].index
                    error = abs((a.x - previous[0].x) - (b.x - previous[1].x))
                    ok = error <= advances * glyphs.STEP_EM * b.size_px + 2e-3
                else:
                    ok = True
                previous = (a, b)
                if not ok or round(a.y) != round(b.y):
                    out[(page, round(b.line_y or b.y))].append(
                        f"{b.char!r} Word x={b.x:.2f} y={b.y:g}, model x={a.x:.2f} y={a.y:g}")
        for g in theirs:
            if id(g) not in matched:
                out[(page, round(g.line_y or g.y))].append(f"{g.char!r} at x={g.x:.2f} not drawn by the model")
    return out


def families(name: str, lines: dict) -> dict:
    """``{family: [agreeing lines, lines]}``."""
    cases = probe.cases_of(name)
    out: dict = {}
    for (page, _y), problems in lines.items():
        family = cases[page].family if page < len(cases) else "past the last case"
        row = out.setdefault(family, [0, 0])
        row[0] += not problems
        row[1] += 1
    return dict(sorted(out.items()))


def score(name: str, data: bytes, fonts, recorded: dict) -> dict:
    """Every score of one document against Word's recorded side."""
    from docx2svg import _render

    options = render_record.options(fonts)
    _documents, layout, _data, _fonts = _render(data, options)
    model = glyphs.exact_floats(glyphs.layout_glyphs(layout, options.glyph_size))
    word = read_render.glyphs_of(recorded["objects"])
    result = glyphs.compare(name, model, word)
    lines = line_score(model, word)
    images = read_anchor_probe.compare_boxes(recorded["images"], read_anchor_probe.model_images(layout))
    return {"glyphs": read_render.row(result), "lines": [sum(1 for p in lines.values() if not p), len(lines)],
            "families": families(name, lines), "images": list(images[:2]),
            "warnings": sorted({w.code for w in options.warnings}),
            "problems": {f"p{page + 1} y={y}": problems for (page, y), problems in lines.items() if problems},
            "image_problems": images[2], "pages": len(layout.pages), "result": result}


def main(argv: list[str]) -> int:
    import oracle

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("names", nargs="*")
    parser.add_argument("--record", action="store_true")
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--limit", type=int, default=30)
    args = parser.parse_args(argv[1:])
    faces: dict = {}
    recorded: dict = {}
    try:
        for name, data in documents():
            if args.names and name not in args.names:
                continue
            pdf = oracle.export(data, name=name)
            word = {"objects": read_render.word_objects(pdf), "images": read_anchor_probe.word_images(pdf)}
            fonts = render_record.RecordingFonts(data, faces)
            result = score(name, data, fonts, word)
            print(result["result"].row())
            print(f"{'':34} lines within half a pixel {result['lines'][0]} / {result['lines'][1]} "
                  f"({', '.join(f'{k} {a}/{b}' for k, (a, b) in result['families'].items())}); pictures "
                  f"{result['images'][0]} / {result['images'][1]}; pages {result['pages']}; warnings "
                  f"{result['warnings']}")
            if args.verbose:
                for where, problems in list(result["problems"].items())[:args.limit]:
                    print(f"    {where}: {problems[:3]}")
                for problem in result["image_problems"][:args.limit]:
                    print("    ", problem)
            recorded[name] = {**word, "glyphs": result["glyphs"], "lines": result["lines"],
                              "families": result["families"], "images_score": result["images"],
                              "warnings": result["warnings"]}
    finally:
        oracle.recover()
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for tools/make_wrap_side_probe.py: every text object (as "
                       "render-observations.json records them) and every picture ([page, x0, y0, x1, y1, layer], "
                       "device px, in paint order), the renderer's scores against them -- lines within half a "
                       "pixel, per family; the glyph row; pictures -- and every face number the renderer asked "
                       "for. Measurements only; regenerate with tools/read_wrap_side_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

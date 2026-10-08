#!/usr/bin/env python3
"""Score a table beside a floating drawing against Word: ``make_wrap_table_probe.py``;
and a drawing anchored in a table cell: ``make_cell_anchor_probe.py``.

Word's PDF gives every text object and every picture (``read_render.word_objects``,
``read_anchor_probe.word_images``); the model lays each document out and draws it.  Three
scores, per document, as ``read_wrap_side_probe.py`` keeps them:

* **lines** -- a line of Word's agrees when every glyph on it is matched to the model's on
  the same baseline and the first glyph of every text object Word starts on it is within
  half a device pixel of the model's (``read_wrap_side_probe.line_score``); counted per
  family;
* **glyphs** -- ``glyphs.compare``'s row, as every probe records it;
* **pictures** -- the drawings' boxes, in Word's paint order (``read_anchor_probe``).

``--record`` writes Word's side to ``tests/fixtures/wrap-table-observations.json``;
``tests/test_wrap_table.py`` holds the model to it offline.

Usage::

    python tools/read_wrap_table_probe.py [-v] [--record] [NAME...]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import glyphs  # noqa: E402
import make_cell_anchor_probe as cell_probe  # noqa: E402
import make_wrap_table_probe as table_probe  # noqa: E402
import read_anchor_probe  # noqa: E402
import read_render  # noqa: E402
import read_wrap_side_probe  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "wrap-table-observations.json"


def probe_of(name: str):
    return cell_probe if name.startswith("cell-anchor") else table_probe


def documents() -> list[tuple[str, bytes]]:
    out = [(name, table_probe.build(name.rsplit("-", 1)[1])) for name in table_probe.DOCUMENTS]
    out += [(name, cell_probe.build(name.rsplit("-", 1)[1])) for name in cell_probe.DOCUMENTS]
    return out


def families(name: str, lines: dict) -> dict:
    """``{family: [agreeing lines, lines]}``."""
    cases = probe_of(name).CASES
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
    lines = read_wrap_side_probe.line_score(model, word)
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
                cases = probe_of(name).CASES
                for where, problems in list(result["problems"].items())[:args.limit]:
                    page = int(where[1:].split()[0]) - 1
                    note = cases[page].note if page < len(cases) else ""
                    print(f"    {where} ({note}): {problems[:3]}")
                for problem in result["image_problems"][:args.limit]:
                    print("    ", problem)
            recorded[name] = {**word, "glyphs": result["glyphs"], "lines": result["lines"],
                              "families": result["families"], "images_score": result["images"],
                              "warnings": result["warnings"]}
    finally:
        oracle.recover()
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for tools/make_wrap_table_probe.py and tools/make_cell_anchor_probe.py: "
                       "every text object (as render-observations.json records them) and every picture ([page, x0, "
                       "y0, x1, y1, layer], device px, in paint order), the renderer's scores against them -- lines "
                       "within half a pixel, per family; the glyph row; pictures -- and every face number the "
                       "renderer asked for. Measurements only; regenerate with tools/read_wrap_table_probe.py "
                       "--record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

#!/usr/bin/env python3
"""Score endnotes against Word: ``make_endnote_probe.py``.

Word's PDF gives every text object (``read_render.word_objects``) and every filled
rectangle (``read_render.word_fills``: the separators); the model lays each document out
and draws it.  Scores, per document:

* **lines** -- a line of Word's agrees when every glyph on it is matched to the model's on
  the same baseline and the first glyph of every text object Word starts on it is within
  half a device pixel of the model's (``read_wrap_side_probe.line_score``): the body's
  lines, with their references, and every line of the notes, with their numbers;
* **glyphs** -- ``glyphs.compare``'s row, as every probe records it;
* **separators** -- the model's rules against Word's fills (``model only`` / ``Word
  only`` device pixels, :func:`rect_score`), summed over the pages;
* **pages** -- the model's page count and Word's.

``--record`` writes Word's side to ``tests/fixtures/endnote-observations.json``;
``tests/test_endnotes.py`` holds the model to it offline.

Usage::

    python tools/read_endnote_probe.py [-v] [--record] [NAME...]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import glyphs  # noqa: E402
import make_endnote_probe as probe  # noqa: E402
import read_render  # noqa: E402
import read_wrap_side_probe  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "endnote-observations.json"


def documents() -> list[tuple[str, bytes]]:
    return [(name, probe.build(name)) for name in probe.DOCUMENTS]


def rect_score(layout, fills: list) -> dict:
    """``read_render.rect_score``'s pixels over every page either side drew on and every
    colour either side drew in (that one counts only the pages and colours the model drew
    rules in, which would not see a separator the model leaves out)."""
    import numpy as np

    out: dict = {}
    total = [0, 0, 0, 0]
    pages = max(len(layout.pages), 1 + max((f[0] for f in fills), default=-1))
    for index in range(pages):
        page = layout.pages[index] if index < len(layout.pages) else None
        rules = page.rules if page is not None else []
        size = (page.height_px, page.width_px) if page is not None else (3508, 2480)
        colors = sorted({(rule.color or "000000").upper() for rule in rules} | {f[5] for f in fills if f[0] == index})
        for color in colors:
            ours = np.zeros(size, bool)
            theirs = np.zeros_like(ours)
            for rule in rules:
                if (rule.color or "000000").upper() == color:
                    ours[int(rule.y):int(rule.y + rule.height), int(rule.x):int(rule.x + rule.width)] = True
            for fill in fills:
                if fill[0] == index and fill[5] == color:
                    theirs[round(fill[2]):round(fill[4]), round(fill[1]):round(fill[3])] = True
            row = [int(ours.sum()), int(theirs.sum()), int((ours & ~theirs).sum()), int((theirs & ~ours).sum())]
            out[f"{index + 1}/{color}"] = row
            total = [a + b for a, b in zip(total, row)]
    out["all"] = total
    return out


def score(name: str, data: bytes, fonts, recorded: dict) -> dict:
    """Every score of one document against Word's recorded side."""
    from docx2svg import _render

    options = render_record.options(fonts)
    _documents, layout, _data, _fonts = _render(data, options)
    model = glyphs.exact_floats(glyphs.layout_glyphs(layout, options.glyph_size))
    word = read_render.glyphs_of(recorded["objects"])
    result = glyphs.compare(name, model, word)
    lines = read_wrap_side_probe.line_score(model, word)
    rects = rect_score(layout, recorded["fills"])
    word_pages = 1 + max((o[0] for o in recorded["objects"]), default=0)
    return {"glyphs": read_render.row(result), "lines": [sum(1 for p in lines.values() if not p), len(lines)],
            "separators": list(rects["all"][2:]), "pages": [len(layout.pages), word_pages],
            "warnings": sorted({w.code for w in options.warnings}),
            "problems": {f"p{page + 1} y={y}": problems for (page, y), problems in lines.items() if problems},
            "rect_pages": {key: row[2:] for key, row in rects.items() if key != "all" and any(row[2:])},
            "result": result}


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
            word = {"objects": read_render.word_objects(pdf), "fills": read_render.word_fills(pdf)}
            fonts = render_record.RecordingFonts(data, faces)
            result = score(name, data, fonts, word)
            print(result["result"].row())
            print(f"{'':34} lines within half a pixel {result['lines'][0]} / {result['lines'][1]}; separator px "
                  f"model only {result['separators'][0]}, Word only {result['separators'][1]}; pages "
                  f"{result['pages'][0]} (Word {result['pages'][1]}); warnings {result['warnings']}")
            if args.verbose:
                for where, problems in list(result["problems"].items())[:args.limit]:
                    print(f"    {where}: {problems[:3]}")
                for where, row in list(result["rect_pages"].items())[:args.limit]:
                    print(f"    rules {where}: model only {row[0]}, Word only {row[1]}")
            recorded[name] = {**word, "glyphs": result["glyphs"], "lines": result["lines"],
                              "separators": result["separators"], "pages": result["pages"],
                              "warnings": result["warnings"]}
    finally:
        oracle.recover()
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for tools/make_endnote_probe.py: every text object (as "
                       "render-observations.json records them) and every filled rectangle, the renderer's scores "
                       "against them -- lines within half a pixel, the glyph row, separator pixels, pages -- and "
                       "every face number the renderer asked for. Measurements only; regenerate with "
                       "tools/read_endnote_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

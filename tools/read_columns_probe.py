#!/usr/bin/env python3
"""Score text columns against Word: ``make_columns_probe.py``.

Word's PDF gives every text object that draws a glyph (``read_render.word_objects``) and
every filled rectangle (``read_render.word_fills``: the column separators, a table's borders, a
footnote's separator); the model lays each document out and draws it.  Scores, per
document:

* **lines** -- a line of Word's agrees when every glyph on it is matched to the model's on
  the same baseline and the first glyph of every text object Word starts on it is within
  half a device pixel of the model's (``read_wrap_side_probe.line_score``);
* **glyphs** -- ``glyphs.compare``'s row, as every probe records it;
* **rules** -- the model's rules against Word's fills (``model only`` / ``Word only``
  device pixels, ``read_endnote_probe.rect_score``), summed over the pages;
* **pages** -- the model's page count and Word's;
* **pictures** -- in ``floating``, the floating pictures' boxes and layers against Word's
  (``read_anchor_probe.compare_boxes``): ``[agreeing, scored]``.

``--record`` writes Word's side to ``tests/fixtures/columns-observations.json``;
``tests/test_columns.py`` holds the model to it offline.

Usage::

    python tools/read_columns_probe.py [-v] [--record] [NAME...]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import glyphs  # noqa: E402
import make_columns_probe as probe  # noqa: E402
import read_endnote_probe  # noqa: E402
import read_render  # noqa: E402
import read_wrap_side_probe  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "columns-observations.json"


def documents() -> list[tuple[str, bytes]]:
    return [(name, probe.build(name)) for name in probe.DOCUMENTS]


def score(name: str, data: bytes, fonts, recorded: dict) -> dict:
    """Every score of one document against Word's recorded side."""
    from docx2svg import _render

    options = render_record.options(fonts)
    _documents, layout, _data, _fonts = _render(data, options)
    model = glyphs.exact_floats(glyphs.layout_glyphs(layout, options.glyph_size))
    word = read_render.glyphs_of(recorded["objects"])
    result = glyphs.compare(name, model, word)
    lines = read_wrap_side_probe.line_score(model, word)
    rects = read_endnote_probe.rect_score(layout, recorded["fills"])
    word_pages = 1 + max((o[0] for o in recorded["objects"]), default=0)
    pictures = None
    if "images" in recorded:
        import read_anchor_probe

        agree, scored, _problems = read_anchor_probe.compare_boxes(recorded["images"],
                                                                 read_anchor_probe.model_images(layout))
        pictures = [agree, scored]
    return {"pictures": pictures, "glyphs": read_render.row(result), "lines": [sum(1 for p in lines.values() if not p), len(lines)],
            "rules": list(rects["all"][2:]), "pages": [len(layout.pages), word_pages],
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
            # The text objects that draw a glyph: a run of spaces Word ends a line with draws
            # none, and no score reads it.
            objects = [o for o in read_render.word_objects(pdf) if o[4].strip()]
            word = {"objects": objects, "fills": read_render.word_fills(pdf)}
            if probe.parts(name)[0] == "floating":
                import read_anchor_probe

                word["images"] = read_anchor_probe.word_images(pdf)
            fonts = render_record.RecordingFonts(data, faces)
            result = score(name, data, fonts, word)
            print(result["result"].row())
            print(f"{'':34} lines within half a pixel {result['lines'][0]} / {result['lines'][1]}; rule px "
                  f"model only {result['rules'][0]}, Word only {result['rules'][1]}; pages "
                  f"{result['pages'][0]} (Word {result['pages'][1]}); warnings {result['warnings']}"
                  + (f"; pictures {result['pictures'][0]} / {result['pictures'][1]}" if result["pictures"] else ""))
            if args.verbose:
                for where, problems in list(result["problems"].items())[:args.limit]:
                    print(f"    {where}: {problems[:3]}")
                for where, row in list(result["rect_pages"].items())[:args.limit]:
                    print(f"    rules {where}: model only {row[0]}, Word only {row[1]}")
            recorded[name] = {**word, "glyphs": result["glyphs"], "lines": result["lines"],
                              "rules": result["rules"], "pages": result["pages"],
                              "warnings": result["warnings"]}
            if result["pictures"] is not None:
                recorded[name]["pictures"] = result["pictures"]
    finally:
        oracle.recover()
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for tools/make_columns_probe.py: every text object that draws a "
                       "glyph (as render-observations.json records them) and every filled rectangle, the "
                       "renderer's scores against them -- lines within half a pixel, the glyph row, rule pixels, "
                       "pages -- and "
                       "every face number the renderer asked for. Measurements only; regenerate with "
                       "tools/read_columns_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

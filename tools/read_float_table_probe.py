#!/usr/bin/env python3
"""Score a floating table (``w:tblpPr``) and the text beside it against Word:
``make_float_table_probe.py``.

Word's PDF gives every text object (``read_render.word_objects``) and every filled
rectangle (``read_render.word_fills``: the tables' borders); the model lays each document
out and draws it.  Three scores, per document:

* **lines** -- a line of Word's agrees when every glyph on it is matched to the model's on
  the same baseline and the first glyph of every text object Word starts on it is within
  half a device pixel of the model's (``read_wrap_side_probe.line_score``): every cell's
  label and every line of the text beside the table and past it; counted per family;
* **glyphs** -- ``glyphs.compare``'s row, as every probe records it;
* **borders** -- ``read_render.rect_score``'s pixels, the model's rules against Word's
  fills (``model only`` / ``Word only``), summed over the pages.

``--record`` writes Word's side to ``tests/fixtures/float-table-observations.json``;
``tests/test_float_table.py`` holds the model to it offline.

Usage::

    python tools/read_float_table_probe.py [-v] [--record] [NAME...]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import glyphs  # noqa: E402
import make_float_table_probe as probe  # noqa: E402
import read_render  # noqa: E402
import read_wrap_side_probe  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "float-table-observations.json"


def documents() -> list[tuple[str, bytes]]:
    return [(name, probe.build(name.rsplit("-", 1)[1])) for name in probe.DOCUMENTS]


def case_pages(objects: list) -> dict[int, int]:
    """Page -> case, from the ``Case N`` line Word drew on it (a case whose table reaches
    past its page moves every later case a page on)."""
    out: dict[int, int] = {}
    for page, _y, _font, _size, text, _x, _steps in objects:
        words = text.split()
        if len(words) >= 2 and words[0] == "Case" and words[1].isdigit():
            out.setdefault(page, int(words[1]))
    return out


def families(lines: dict, pages: dict[int, int], cases=None) -> dict:
    """``{family: [agreeing lines, lines]}``."""
    cases = probe.CASES if cases is None else cases
    out: dict = {}
    for (page, _y), problems in lines.items():
        case = pages.get(page)
        family = cases[case].family if case is not None and case < len(cases) else "no case"
        row = out.setdefault(family, [0, 0])
        row[0] += not problems
        row[1] += 1
    return dict(sorted(out.items()))


def score(name: str, data: bytes, fonts, recorded: dict, cases=None) -> dict:
    """Every score of one document against Word's recorded side (``cases``: the probe's,
    :data:`make_float_table_probe.CASES` unless another probe's are given)."""
    from docx2svg import _render

    options = render_record.options(fonts)
    _documents, layout, _data, _fonts = _render(data, options)
    model = glyphs.exact_floats(glyphs.layout_glyphs(layout, options.glyph_size))
    word = read_render.glyphs_of(recorded["objects"])
    result = glyphs.compare(name, model, word)
    lines = read_wrap_side_probe.line_score(model, word)
    rects = read_render.rect_score(layout, recorded["fills"])
    borders = list(rects["all"][2:])
    return {"glyphs": read_render.row(result), "lines": [sum(1 for p in lines.values() if not p), len(lines)],
            "families": families(lines, case_pages(recorded["objects"]), cases), "borders": borders,
            "warnings": sorted({w.code for w in options.warnings}),
            "problems": {f"p{page + 1} y={y}": problems for (page, y), problems in lines.items() if problems},
            "border_pages": {key: row[2:] for key, row in rects.items() if key != "all" and any(row[2:])},
            "pages": len(layout.pages), "result": result}


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
            print(f"{'':34} lines within half a pixel {result['lines'][0]} / {result['lines'][1]} "
                  f"({', '.join(f'{k} {a}/{b}' for k, (a, b) in result['families'].items())}); border px model "
                  f"only {result['borders'][0]}, Word only {result['borders'][1]}; pages {result['pages']}; "
                  f"warnings {result['warnings']}")
            if args.verbose:
                pages = case_pages(word["objects"])
                for where, problems in list(result["problems"].items())[:args.limit]:
                    page = int(where[1:].split()[0]) - 1
                    case = pages.get(page)
                    note = probe.CASES[case].note if case is not None else ""
                    print(f"    {where} (case {case}: {note}): {problems[:3]}")
                for where, row in list(result["border_pages"].items())[:args.limit]:
                    case = pages.get(int(where.split("/")[0]) - 1)
                    print(f"    borders {where} (case {case}: {probe.CASES[case].note if case is not None else ''}): "
                          f"model only {row[0]}, Word only {row[1]}")
            recorded[name] = {**word, "glyphs": result["glyphs"], "lines": result["lines"],
                              "families": result["families"], "borders": result["borders"],
                              "warnings": result["warnings"]}
    finally:
        oracle.recover()
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for tools/make_float_table_probe.py: every text object (as "
                       "render-observations.json records them) and every filled rectangle, the renderer's scores "
                       "against them -- lines within half a pixel, per family; the glyph row; border pixels -- and "
                       "every face number the renderer asked for. Measurements only; regenerate with "
                       "tools/read_float_table_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
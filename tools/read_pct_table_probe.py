#!/usr/bin/env python3
"""Score a table whose widths are stated in percent against Word:
``make_pct_table_probe.py``.

The scores are ``read_float_table_probe.py``'s -- every cell's text within half a pixel
of Word's, per family; the glyph row; the border pixels only one side drew -- over this
probe's cases.  ``--record`` writes Word's side to
``tests/fixtures/pct-table-observations.json``; ``tests/test_pct_table.py`` holds the
model to it offline.

Usage::

    python tools/read_pct_table_probe.py [-v] [--record] [NAME...]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import make_pct_table_probe as probe  # noqa: E402
import read_float_table_probe as reader  # noqa: E402
import read_render  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "pct-table-observations.json"


def documents() -> list[tuple[str, bytes]]:
    return [(name, probe.build(name.rsplit("-", 1)[1])) for name in probe.DOCUMENTS]


def score(name: str, data: bytes, fonts, recorded: dict) -> dict:
    return reader.score(name, data, fonts, recorded, probe.CASES)


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
                pages = reader.case_pages(word["objects"])
                for where, problems in list(result["problems"].items())[:args.limit]:
                    case = pages.get(int(where[1:].split()[0]) - 1)
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
            "_about": ("What Word 16.106 drew for tools/make_pct_table_probe.py: every text object (as "
                       "render-observations.json records them) and every filled rectangle, the renderer's scores "
                       "against them -- lines within half a pixel, per family; the glyph row; border pixels -- and "
                       "every face number the renderer asked for. Measurements only; regenerate with "
                       "tools/read_pct_table_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

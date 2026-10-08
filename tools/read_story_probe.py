#!/usr/bin/env python3
"""Score the headers, footers and fields the renderer draws against Word's export.

Three generated probes, every page of each exported by Word 16.106 (``oracle.py``):

* ``make_story_probe.py`` -- where a header's and a footer's glyphs go (four settings);
* ``make_story_select_probe.py`` -- which story each page shows, and its number;
* ``make_field_probe.py`` -- what each field shows, and in which run's format.

Every glyph the renderer draws -- body, header and footer -- is matched to Word's by
``glyphs.compare`` (pen x, baseline, face and size, as the committed documents are), and
the header and footer glyphs are also counted apart; every filled rectangle is compared
pixel for pixel as ``read_render.py`` does.  With ``--record``, Word's text objects and
fills and every face number the renderer asked for are written to
``tests/fixtures/story-observations.json``, so ``tests/test_stories.py`` holds all of it
offline.

Usage::

    python tools/read_story_probe.py [-v] [--record] [NAME...]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))

import read_render  # noqa: E402
import render_record  # noqa: E402

OBSERVATIONS = read_render.FIXTURES / "story-observations.json"


def documents() -> list[tuple[str, bytes]]:
    import make_field_probe
    import make_story_probe
    import make_story_select_probe

    out = [(f"story-{setting}", make_story_probe.build(setting)) for setting in make_story_probe.SETTINGS]
    out += [(f"select-{doc.name.replace(' ', '-')}", make_story_select_probe.build(index))
            for index, doc in enumerate(make_story_select_probe.DOCS)]
    out += [(f"fields-{name}", make_field_probe.build(name)) for name in make_field_probe.DOCUMENTS]
    return out


def row(result) -> list[int]:
    return read_render.row(result) + [result.story_model, result.story_matched, result.story_exact]


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
    for name, data in documents():
        if args.names and name not in args.names:
            continue
        pdf = oracle.export(data, name=name)
        objects, fills = read_render.word_objects(pdf), read_render.word_fills(pdf)
        fonts = render_record.RecordingFonts(data, faces)
        result, rects, warnings = read_render.score(name, data, fonts, objects, fills)
        print(result.row())
        print(f"{'':34} stories: {result.story_exact} exact of {result.story_matched} matched, "
              f"{result.story_model} drawn; rectangles: model {rects['all'][0]} px, Word {rects['all'][1]} px, "
              f"model only {rects['all'][2]}, Word only {rects['all'][3]}; warnings {warnings}")
        if args.verbose:
            for problem in result.problems[:args.limit]:
                print("   ", problem.replace("|", ""))
        recorded[name] = {"objects": objects, "fills": fills, "glyphs": row(result), "rects": rects,
                          "warnings": warnings}
    if args.record:
        render_record.dump(OBSERVATIONS, {
            "_about": ("What Word 16.106 drew for each document of tools/make_story_probe.py, "
                       "make_story_select_probe.py and make_field_probe.py -- every text object and filled "
                       "rectangle, as render-observations.json records them -- the renderer's scores against "
                       "them (read_render.row, then header and footer glyphs drawn, matched and exact), and "
                       "every face number the renderer asked for. Measurements only; regenerate with "
                       "tools/read_story_probe.py --record."),
            "faces": dict(sorted(faces.items())),
            "documents": recorded,
        })
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
